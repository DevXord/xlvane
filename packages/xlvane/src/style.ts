// Cell styles as plain data, and their validation. The same types describe the style of a cell
// read from a file (`Cell.style`) and the style given to the writer (`NewSheet.set`). Allowed
// values follow the SpreadsheetML schema (docs/facts/sml-write.md W2-10 … W2-19).
import { XlvError } from './errors';
import { BUILTIN_NUMBER_FORMATS } from './sml/numfmt';
import { suggest } from './util/suggest';

/**
 * A colour. A string is an RGB colour as 6 hex digits (`"FF0000"`) or ARGB as 8 hex digits
 * (`"80FF0000"`); colours read from files are always 8 upper-case digits. Files can also refer
 * to theme colours, the legacy indexed palette or the automatic colour; those are kept as
 * objects so that they can be written back unchanged.
 *
 * @example
 * ```ts
 * import { createWorkbook, type Color } from 'xlvane';
 *
 * const red: Color = 'FF0000';
 * const accent: Color = { theme: 4, tint: 0.4 };
 * createWorkbook().addSheet('Data').set('A1', 'x', { font: { color: red }, fill: { color: accent } });
 * ```
 */
export type Color = string | { readonly theme: number; readonly tint?: number } | { readonly indexed: number } | { readonly auto: true };

/**
 * Underline styles (ST_UnderlineValues).
 *
 * @example
 * ```ts
 * import { createWorkbook, type UnderlineStyle } from 'xlvane';
 *
 * const underline: UnderlineStyle = 'double';
 * createWorkbook().addSheet('Data').set('A1', 'Total', { font: { underline } });
 * ```
 */
export type UnderlineStyle = 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';

/**
 * Vertical position of text (ST_VerticalAlignRun, ECMA-376 Part 1 §22.9.2.17): superscript or
 * subscript text is raised or lowered and made smaller if possible.
 *
 * @example
 * ```ts
 * import { createWorkbook, type VerticalAlignRun } from 'xlvane';
 *
 * const vertAlign: VerticalAlignRun = 'superscript';
 * createWorkbook().addSheet('Data').set('A1', { richText: [{ text: 'm' }, { text: '2', font: { vertAlign } }] });
 * ```
 */
export type VerticalAlignRun = 'baseline' | 'superscript' | 'subscript';

/**
 * Theme font scheme a font belongs to (ST_FontScheme, ECMA-376 Part 1 §18.18.33): `major`
 * (headings) or `minor` (body text) fonts follow the theme; `none` is not a theme font.
 *
 * @example
 * ```ts
 * import { createWorkbook, type FontScheme } from 'xlvane';
 *
 * const scheme: FontScheme = 'minor';
 * createWorkbook().addSheet('Data').set('A1', 'x', { font: { name: 'Calibri', scheme } });
 * ```
 */
export type FontScheme = 'major' | 'minor' | 'none';

/**
 * Font of a cell. Missing properties take the workbook's default font (Calibri 11 for new workbooks).
 *
 * @example
 * ```ts
 * import { createWorkbook, type FontStyle } from 'xlvane';
 *
 * const title: FontStyle = { name: 'Arial', size: 16, bold: true, color: '1F4E79' };
 * createWorkbook().addSheet('Report').set('A1', 'Quarterly report', { font: title });
 * ```
 */
export interface FontStyle {
  /** Font name, 1 to 31 characters (the length Office accepts), such as `"Arial"`. */
  readonly name?: string;
  /** Size in points, 1 to 409.55 (the range Office accepts). */
  readonly size?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  /** Underline style; `true` means `"single"`. Read styles always use the string form. */
  readonly underline?: UnderlineStyle | boolean;
  readonly strike?: boolean;
  readonly color?: Color;
  /** Superscript or subscript (ECMA-376 Part 1 §18.4.14). */
  readonly vertAlign?: VerticalAlignRun;
  /** Font family, 0 to 5 in Office ([MS-OI29500] §2.1.703; the schema allows 0 to 14). */
  readonly family?: number;
  /** Theme font scheme (ECMA-376 Part 1 §18.8.35). */
  readonly scheme?: FontScheme;
  /** Character set, 0 to 255, used to pick a substitute font (ECMA-376 Part 1 §18.4.1). */
  readonly charset?: number;
  /** Outline font (§18.4.2). */
  readonly outline?: boolean;
  /** Shadow (§18.8.36). */
  readonly shadow?: boolean;
  /** Condensed font (§18.8.12). */
  readonly condense?: boolean;
  /** Extended font (§18.8.17). */
  readonly extend?: boolean;
}

/**
 * Fill patterns (ST_PatternType), except `"none"`, which is the absence of a fill.
 *
 * @example
 * ```ts
 * import { createWorkbook, type PatternType } from 'xlvane';
 *
 * const pattern: PatternType = 'lightGrid';
 * createWorkbook().addSheet('Data').set('A1', 'x', { fill: { pattern, fgColor: 'C0C0C0' } });
 * ```
 */
export type PatternType =
  | 'solid'
  | 'mediumGray'
  | 'darkGray'
  | 'lightGray'
  | 'darkHorizontal'
  | 'darkVertical'
  | 'darkDown'
  | 'darkUp'
  | 'darkGrid'
  | 'darkTrellis'
  | 'lightHorizontal'
  | 'lightVertical'
  | 'lightDown'
  | 'lightUp'
  | 'lightGrid'
  | 'lightTrellis'
  | 'gray125'
  | 'gray0625';

/**
 * Background of a cell. `{ color }` is short for `{ pattern: "solid", fgColor: color }`; read
 * styles always use the long form.
 *
 * @example
 * ```ts
 * import { createWorkbook, type FillStyle } from 'xlvane';
 *
 * const highlight: FillStyle = { color: 'FFFF00' };
 * createWorkbook().addSheet('Data').set('A1', 'check me', { fill: highlight });
 * ```
 */
export type FillStyle =
  | { readonly color: Color }
  | { readonly pattern: PatternType; readonly fgColor?: Color; readonly bgColor?: Color }
  | GradientFill;

/**
 * A gradient fill (CT_GradientFill, sml.xsd): `linear` with an angle in `degree`, or `path`,
 * which spreads from the rectangle given by `left`, `right`, `top` and `bottom` (fractions of the
 * cell, 0 to 1). Each stop has a position from 0 to 1 and a colour. The meaning of the numbers is
 * PROVISIONAL until ECMA-376 Part 1 §18.8.24 is checked (docs/facts/sml-write.md W2-60).
 *
 * @example
 * ```ts
 * import { createWorkbook, type GradientFill } from 'xlvane';
 *
 * const fill: GradientFill = { type: 'gradient', gradient: 'linear', degree: 90, stops: [{ position: 0, color: 'FFFFFF' }, { position: 1, color: '4472C4' }] };
 * createWorkbook().addSheet('Data').set('A1', 'x', { fill });
 * ```
 */
export interface GradientFill {
  readonly type: 'gradient';
  readonly gradient: 'linear' | 'path';
  /** Angle of a linear gradient in degrees (default 0). */
  readonly degree?: number;
  readonly left?: number;
  readonly right?: number;
  readonly top?: number;
  readonly bottom?: number;
  readonly stops: readonly { readonly position: number; readonly color: Color }[];
}

/**
 * Border line styles (ST_BorderStyle), except `"none"`, which is the absence of a line.
 *
 * @example
 * ```ts
 * import { createWorkbook, type BorderLineStyle } from 'xlvane';
 *
 * const line: BorderLineStyle = 'double';
 * createWorkbook().addSheet('Data').set('A1', 100, { border: { bottom: { style: line } } });
 * ```
 */
export type BorderLineStyle =
  | 'thin'
  | 'medium'
  | 'dashed'
  | 'dotted'
  | 'thick'
  | 'double'
  | 'hair'
  | 'mediumDashed'
  | 'dashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot'
  | 'slantDashDot';

/**
 * One edge of a cell border.
 *
 * @example
 * ```ts
 * import { createWorkbook, type BorderEdge } from 'xlvane';
 *
 * const edge: BorderEdge = { style: 'thin', color: '808080' };
 * createWorkbook().addSheet('Data').set('A1', 'x', { border: { top: edge, bottom: edge } });
 * ```
 */
export interface BorderEdge {
  readonly style: BorderLineStyle;
  readonly color?: Color;
}

/**
 * Borders of a cell.
 *
 * @example
 * ```ts
 * import { createWorkbook, type BorderStyle } from 'xlvane';
 *
 * const box: BorderStyle = {
 *   top: { style: 'thin' },
 *   right: { style: 'thin' },
 *   bottom: { style: 'thin' },
 *   left: { style: 'thin' },
 * };
 * createWorkbook().addSheet('Data').style('A1:C3', { border: box });
 * ```
 */
export interface BorderStyle {
  readonly top?: BorderEdge;
  readonly right?: BorderEdge;
  readonly bottom?: BorderEdge;
  readonly left?: BorderEdge;
  /** The diagonal line; it is drawn in the directions that `diagonalUp` and `diagonalDown` enable. */
  readonly diagonal?: BorderEdge;
  /** Draws the diagonal from bottom left to top right. */
  readonly diagonalUp?: boolean;
  /** Draws the diagonal from top left to bottom right. */
  readonly diagonalDown?: boolean;
}

/**
 * Horizontal alignment (ST_HorizontalAlignment).
 *
 * @example
 * ```ts
 * import { createWorkbook, type HorizontalAlignment } from 'xlvane';
 *
 * const horizontal: HorizontalAlignment = 'center';
 * createWorkbook().addSheet('Data').set('A1', 'Title', { alignment: { horizontal } });
 * ```
 */
export type HorizontalAlignment = 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';

/**
 * Vertical alignment (ST_VerticalAlignment).
 *
 * @example
 * ```ts
 * import { createWorkbook, type VerticalAlignment } from 'xlvane';
 *
 * const vertical: VerticalAlignment = 'top';
 * createWorkbook().addSheet('Data').set('A1', 'Notes', { alignment: { vertical, wrapText: true } });
 * ```
 */
export type VerticalAlignment = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

/**
 * Alignment of a cell's content.
 *
 * @example
 * ```ts
 * import { createWorkbook, type AlignmentStyle } from 'xlvane';
 *
 * const wrapped: AlignmentStyle = { horizontal: 'left', vertical: 'top', wrapText: true, indent: 1 };
 * createWorkbook().addSheet('Data').set('A1', 'A long comment', { alignment: wrapped });
 * ```
 */
export interface AlignmentStyle {
  readonly horizontal?: HorizontalAlignment;
  readonly vertical?: VerticalAlignment;
  readonly wrapText?: boolean;
  /** Indent level, 0 to 255. */
  readonly indent?: number;
  /** Text rotation in degrees: 0 to 180, or 255 for vertical text. */
  readonly textRotation?: number;
  /** Shrinks the text to fit the cell. */
  readonly shrinkToFit?: boolean;
  /** Justifies the last line of distributed text. */
  readonly justifyLastLine?: boolean;
  /**
   * Reading order: stored as 0, 1 or 2 ([MS-OI29500] §2.1.693 c), here as `context`,
   * `leftToRight` and `rightToLeft` (meaning PROVISIONAL, docs/facts/sml-write.md W2-62).
   */
  readonly readingOrder?: ReadingOrder;
}

/**
 * Reading order of a cell's text, see {@link AlignmentStyle.readingOrder}.
 *
 * @example
 * ```ts
 * import { createWorkbook, type ReadingOrder } from 'xlvane';
 *
 * const readingOrder: ReadingOrder = 'rightToLeft';
 * createWorkbook().addSheet('Data').set('A1', 'x', { alignment: { readingOrder } });
 * ```
 */
export type ReadingOrder = 'context' | 'leftToRight' | 'rightToLeft';

/**
 * Cell protection (CT_CellProtection): `locked` cells cannot be changed and the formulas of
 * `hidden` cells are not shown, once the sheet is protected. Office's defaults are locked and
 * not hidden ([MS-OI29500] §2.1.714).
 *
 * @example
 * ```ts
 * import { createWorkbook, type ProtectionStyle } from 'xlvane';
 *
 * const input: ProtectionStyle = { locked: false };
 * createWorkbook().addSheet('Form').set('B2', null, { protection: input }).protect();
 * ```
 */
export interface ProtectionStyle {
  readonly locked?: boolean;
  readonly hidden?: boolean;
}

/**
 * The style of a cell. Every part is optional; an empty object is the default style.
 *
 * @example
 * ```ts
 * import { createWorkbook, type CellStyle } from 'xlvane';
 *
 * const header: CellStyle = {
 *   font: { bold: true, color: 'FFFFFF' },
 *   fill: { color: '4472C4' },
 *   border: { bottom: { style: 'thin' } },
 *   alignment: { horizontal: 'center' },
 * };
 * createWorkbook().addSheet('Report').set('A1', 'Total', header);
 * ```
 */
export interface CellStyle {
  readonly font?: FontStyle;
  readonly fill?: FillStyle;
  readonly border?: BorderStyle;
  readonly alignment?: AlignmentStyle;
  /** Cell protection, effective when the sheet is protected. */
  readonly protection?: ProtectionStyle;
  /**
   * Number format: a format code, for example `"#,##0.00"` or `"yyyy-mm-dd"`, or the id of a
   * built-in format of ECMA-376 Part 1 §18.8.30 that every language has (0–4, 9–22, 37–40,
   * 45–49), for example `14` for a date shown in the order of the reader's locale. An id is
   * written without a format code, so applications use their own code for it; a code string is
   * always stored as given. Reading gives the code (`14` reads back as `"mm-dd-yy"`).
   */
  readonly numFmt?: string | number;
}

export const UNDERLINES: readonly UnderlineStyle[] = ['single', 'double', 'singleAccounting', 'doubleAccounting'];
export const PATTERNS: readonly PatternType[] = [
  'solid',
  'mediumGray',
  'darkGray',
  'lightGray',
  'darkHorizontal',
  'darkVertical',
  'darkDown',
  'darkUp',
  'darkGrid',
  'darkTrellis',
  'lightHorizontal',
  'lightVertical',
  'lightDown',
  'lightUp',
  'lightGrid',
  'lightTrellis',
  'gray125',
  'gray0625',
];
export const BORDER_STYLES: readonly BorderLineStyle[] = [
  'thin',
  'medium',
  'dashed',
  'dotted',
  'thick',
  'double',
  'hair',
  'mediumDashed',
  'dashDot',
  'mediumDashDot',
  'dashDotDot',
  'mediumDashDotDot',
  'slantDashDot',
];
export const HORIZONTAL: readonly HorizontalAlignment[] = [
  'general',
  'left',
  'center',
  'right',
  'fill',
  'justify',
  'centerContinuous',
  'distributed',
];
export const VERTICAL: readonly VerticalAlignment[] = ['top', 'center', 'bottom', 'justify', 'distributed'];
export const EDGES = ['left', 'right', 'top', 'bottom'] as const;
export const VERT_ALIGN_RUNS: readonly VerticalAlignRun[] = ['baseline', 'superscript', 'subscript'];
export const FONT_SCHEMES: readonly FontScheme[] = ['major', 'minor', 'none'];
export const READING_ORDERS: readonly ReadingOrder[] = ['context', 'leftToRight', 'rightToLeft'];
/** Font boolean properties beyond bold, italic and strike, in the order Office writes them. */
export const FONT_FLAGS = ['condense', 'extend', 'outline', 'shadow'] as const;

// Limits that Office applies to the styles part ([MS-OI29500] rev. 25.0, docs/decisions.md D-1101).
// The writer enforces them for styles given by the caller; styles already in a file being edited
// are checked only against the looser caps below them (`'existing'` mode of normalizeStyle).
/** Smallest font size Office accepts, in points ([MS-OI29500] §2.1.677 c). */
export const MIN_FONT_SIZE: number = 1;
/** Largest font size Office accepts, in points ([MS-OI29500] §2.1.677 c). */
export const MAX_FONT_SIZE: number = 409.55;
/** Longest font name Office accepts, in characters before escaping ([MS-OI29500] §2.1.711). */
export const MAX_FONT_NAME_LENGTH: number = 31;
/** Longest number format code Office accepts: shorter than 255 characters ([MS-OI29500] §2.1.712 b). */
export const MAX_NUMFMT_LENGTH: number = 254;
/** Largest `indent` Office accepts ([MS-OI29500] §2.1.693 b). */
export const MAX_INDENT: number = 255;
/** Most custom number formats Office loads ([MS-OI29500] §2.1.713 g). */
export const MAX_CUSTOM_NUMFMTS: number = 206;
/** Most cell format (`xf`) records in `cellXfs` ([MS-OI29500] §2.1.700 a, c). */
export const MAX_CELL_XFS: number = 65430;
/** Most `font` records ([MS-OI29500] §2.1.707 a). */
export const MAX_FONTS: number = 65491;
/** Most `fill` records ([MS-OI29500] §2.1.705 a). */
export const MAX_FILLS: number = 65430;
/** Most `border` records ([MS-OI29500] §2.1.696). */
export const MAX_BORDERS: number = 65430;

/** Caps for styles read from a file that is being edited: the earlier implementation caps (D-115, D-207). */
const EXISTING_MAX_FONT_NAME_LENGTH = 255;
const EXISTING_MAX_NUMFMT_LENGTH = 1024;

/** The style records Office limits in number, with the limit and how to stay below it. */
export type StyleCountLimit = 'customNumFmts' | 'cellXfs' | 'fonts' | 'fills' | 'borders';
const STYLE_COUNT_LIMITS: Readonly<Record<StyleCountLimit, { readonly max: number; readonly what: string; readonly hint: string }>> = {
  customNumFmts: { max: MAX_CUSTOM_NUMFMTS, what: 'custom number formats', hint: 'Use fewer distinct number format codes, for example one code such as "#,##0.00" for all amounts.' },
  cellXfs: { max: MAX_CELL_XFS, what: 'cell formats (cellXfs records)', hint: 'Use fewer distinct styles, for example one style object such as { font: { bold: true } } for all header cells.' },
  fonts: { max: MAX_FONTS, what: 'fonts', hint: 'Use fewer distinct fonts, for example one font such as { name: "Arial", size: 11 } for all body cells.' },
  fills: { max: MAX_FILLS, what: 'fills', hint: 'Use fewer distinct fills, for example one fill such as { color: "DDEBF7" } for all highlighted cells.' },
  borders: { max: MAX_BORDERS, what: 'borders', hint: 'Use fewer distinct borders, for example one border such as { bottom: { style: "thin" } } for all underlined cells.' },
};

/**
 * Throws `XLV_E_LIMIT_EXCEEDED` when a styles part would have more records of one kind than
 * Office accepts ([MS-OI29500] §2.1.696, §2.1.700, §2.1.705, §2.1.707, §2.1.713 g).
 */
export function checkStyleCount(limit: StyleCountLimit, count: number): void {
  const l = STYLE_COUNT_LIMITS[limit];
  if (count <= l.max) return;
  throw new XlvError('XLV_E_LIMIT_EXCEEDED', `The workbook needs ${String(count)} ${l.what}; Office accepts at most ${String(l.max)}.`, {
    hint: l.hint,
    details: { limit, value: l.max },
  });
}

/** `'input'`: a style given by the caller (Office limits). `'existing'`: a style read from a file. */
export type StyleSource = 'input' | 'existing';

/** A validated gradient fill: stops in the given order, numbers finite. */
export interface NormalizedGradient {
  type: 'gradient';
  gradient: 'linear' | 'path';
  degree?: number;
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
  stops: { position: number; color: NormalizedColor }[];
}

/** A validated colour: an ARGB string or one of the object forms, with only known keys. */
export type NormalizedColor = Color;

/** A style after validation: only known keys, colours in ARGB form, defaults removed. */
export interface NormalizedStyle {
  font?: {
    name?: string;
    size?: number;
    bold?: true;
    italic?: true;
    underline?: UnderlineStyle;
    strike?: true;
    color?: NormalizedColor;
    vertAlign?: VerticalAlignRun;
    family?: number;
    scheme?: FontScheme;
    charset?: number;
    outline?: true;
    shadow?: true;
    condense?: true;
    extend?: true;
  };
  fill?: { pattern: PatternType; fgColor?: NormalizedColor; bgColor?: NormalizedColor } | NormalizedGradient;
  border?: {
    top?: BorderEdge;
    right?: BorderEdge;
    bottom?: BorderEdge;
    left?: BorderEdge;
    diagonal?: BorderEdge;
    diagonalUp?: true;
    diagonalDown?: true;
  };
  alignment?: {
    horizontal?: HorizontalAlignment;
    vertical?: VerticalAlignment;
    wrapText?: true;
    indent?: number;
    textRotation?: number;
    shrinkToFit?: true;
    justifyLastLine?: true;
    readingOrder?: ReadingOrder;
  };
  /** As given; the writer writes only values that differ from Office's defaults (locked, not hidden). */
  protection?: { locked?: boolean; hidden?: boolean };
  /** A format code, or a built-in format id (never 0). */
  numFmt?: string | number;
}

function invalid(path: string, message: string, hint: string): XlvError {
  return new XlvError('XLV_E_INVALID_STYLE', `Invalid style: ${path} ${message}`, { hint, details: { property: path } });
}

/** Reads the own enumerable properties of `value` once each, rejecting unknown keys. */
function readObject(value: unknown, path: string, known: readonly string[]): Map<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value) || value instanceof Date) {
    throw invalid(path, 'must be an object.', `Pass an object with some of these properties: ${known.join(', ')}.`);
  }
  const out = new Map<string, unknown>();
  for (const key of Object.keys(value)) {
    if (!known.includes(key)) {
      const close = suggest(key, known);
      throw invalid(
        `${path}.${key.slice(0, 40)}`,
        'is not a known style property.',
        `${close.length > 0 ? `Did you mean "${close[0] ?? ''}"? ` : ''}Known properties of ${path}: ${known.join(', ')}.`,
      );
    }
    const v = (value as Record<string, unknown>)[key];
    if (v !== undefined) out.set(key, v);
  }
  return out;
}

function oneOf<T extends string>(value: unknown, path: string, allowed: readonly T[]): T {
  if (typeof value === 'string' && (allowed as readonly string[]).includes(value)) return value as T;
  const close = typeof value === 'string' ? suggest(value, allowed) : [];
  throw invalid(
    path,
    `has the value ${describe(value)}.`,
    `${close.length > 0 ? `Did you mean "${close[0] ?? ''}"? ` : ''}Allowed values: ${allowed.join(', ')}.`,
  );
}

function bool(value: unknown, path: string): boolean {
  if (typeof value === 'boolean') return value;
  throw invalid(path, `has the value ${describe(value)}.`, 'Use true or false.');
}

function integer(value: unknown, path: string, min: number, max: number, extra?: number, example?: number): number {
  if (typeof value === 'number' && Number.isInteger(value) && ((value >= min && value <= max) || value === extra)) return value;
  throw invalid(
    path,
    `has the value ${describe(value)}.`,
    `Use a whole number from ${String(min)} to ${String(max)}${extra === undefined ? '' : ` or ${String(extra)}`}${example === undefined ? '' : `, such as ${String(example)}`}.`,
  );
}

function describe(value: unknown): string {
  if (typeof value === 'string') return JSON.stringify(value.length > 40 ? `${value.slice(0, 40)}…` : value);
  if (typeof value === 'number' || typeof value === 'boolean' || value === null || value === undefined) return String(value);
  return `of type ${typeof value}`;
}

/** Validates a colour and returns its normalised form (ARGB upper case for strings). */
export function normalizeColor(value: unknown, path: string): NormalizedColor {
  if (typeof value === 'string') {
    if (/^[0-9A-Fa-f]{6}$/.test(value)) return `FF${value.toUpperCase()}`;
    if (/^[0-9A-Fa-f]{8}$/.test(value)) return value.toUpperCase();
    throw invalid(path, `has the value ${describe(value)}.`, 'Use 6 hex digits (RRGGBB, for example "FF0000") or 8 (AARRGGBB).');
  }
  const hint = 'Use a hex string such as "FF0000", or { theme: 1 }, { indexed: 8 } or { auto: true }.';
  if (typeof value !== 'object' || value === null) throw invalid(path, `has the value ${describe(value)}.`, hint);
  const keys = Object.keys(value);
  if (keys.includes('theme')) {
    const o = readObject(value, path, ['theme', 'tint']);
    const theme = integer(o.get('theme'), `${path}.theme`, 0, 0xffff);
    const tint = o.get('tint');
    if (tint === undefined || tint === 0) return { theme };
    if (typeof tint !== 'number' || !(tint >= -1 && tint <= 1)) {
      throw invalid(`${path}.tint`, `has the value ${describe(tint)}.`, 'Use a number from -1 to 1.');
    }
    return { theme, tint };
  }
  if (keys.includes('indexed')) {
    const o = readObject(value, path, ['indexed']);
    return { indexed: integer(o.get('indexed'), `${path}.indexed`, 0, 0xffff) };
  }
  if (keys.includes('auto')) {
    const o = readObject(value, path, ['auto']);
    if (o.get('auto') !== true) throw invalid(`${path}.auto`, 'must be true.', hint);
    return { auto: true };
  }
  throw invalid(path, 'is not a colour.', hint);
}

function normalizeFont(value: unknown, source: StyleSource): NonNullable<NormalizedStyle['font']> {
  const o = readObject(value, 'font', ['name', 'size', 'bold', 'italic', 'underline', 'strike', 'color', 'vertAlign', 'family', 'scheme', 'charset', ...FONT_FLAGS]);
  const out: NonNullable<NormalizedStyle['font']> = {};
  const name = o.get('name');
  if (name !== undefined) {
    const max = source === 'input' ? MAX_FONT_NAME_LENGTH : EXISTING_MAX_FONT_NAME_LENGTH;
    if (typeof name !== 'string' || name.length === 0 || name.length > max) {
      throw invalid('font.name', `has the value ${describe(name)}.`, `Use a font name of 1 to ${String(max)} characters, such as "Arial".`);
    }
    out.name = name;
  }
  const size = o.get('size');
  if (size !== undefined) {
    const ok = typeof size === 'number' && (source === 'input' ? size >= MIN_FONT_SIZE && size <= MAX_FONT_SIZE : size > 0 && Number.isFinite(size));
    if (!ok) {
      throw invalid('font.size', `has the value ${describe(size)}.`, `Use a size in points from ${String(MIN_FONT_SIZE)} to ${String(MAX_FONT_SIZE)}, such as 11.`);
    }
    out.size = size;
  }
  for (const key of ['bold', 'italic', 'strike', ...FONT_FLAGS] as const) {
    const v = o.get(key);
    if (v !== undefined && bool(v, `font.${key}`)) out[key] = true;
  }
  const underline = o.get('underline');
  if (underline !== undefined) {
    if (typeof underline === 'boolean') {
      if (underline) out.underline = 'single';
    } else {
      out.underline = oneOf(underline, 'font.underline', UNDERLINES);
    }
  }
  const color = o.get('color');
  if (color !== undefined) out.color = normalizeColor(color, 'font.color');
  const vertAlign = o.get('vertAlign');
  if (vertAlign !== undefined) {
    const v = oneOf(vertAlign, 'font.vertAlign', VERT_ALIGN_RUNS);
    if (v !== 'baseline') out.vertAlign = v;
  }
  const family = o.get('family');
  // [MS-OI29500] §2.1.703: Office accepts 0 to 5; the schema's ST_FontFamily allows 0 to 14.
  if (family !== undefined) out.family = integer(family, 'font.family', 0, source === 'input' ? 5 : 14, undefined, 2);
  const scheme = o.get('scheme');
  if (scheme !== undefined) out.scheme = oneOf(scheme, 'font.scheme', FONT_SCHEMES);
  const charset = o.get('charset');
  // ECMA-376 Part 1 §18.4.1: one value per character set, 0 to 255.
  if (charset !== undefined) out.charset = integer(charset, 'font.charset', 0, 255, undefined, 1);
  return out;
}

function finite(value: unknown, path: string, min: number, max: number, example: number): number {
  if (typeof value === 'number' && Number.isFinite(value) && value >= min && value <= max) return value;
  throw invalid(path, `has the value ${describe(value)}.`, `Use a number from ${String(min)} to ${String(max)}, such as ${String(example)}.`);
}

function normalizeGradient(value: unknown): NormalizedGradient {
  const o = readObject(value, 'fill', ['type', 'gradient', 'degree', 'left', 'right', 'top', 'bottom', 'stops']);
  const gradient = oneOf(o.get('gradient'), 'fill.gradient', ['linear', 'path'] as const);
  const out: NormalizedGradient = { type: 'gradient', gradient, stops: [] };
  const degree = o.get('degree');
  if (degree !== undefined) {
    const d = finite(degree, 'fill.degree', -360, 360, 90);
    if (d !== 0) out.degree = d;
  }
  for (const side of ['left', 'right', 'top', 'bottom'] as const) {
    const v = o.get(side);
    if (v === undefined) continue;
    const n = finite(v, `fill.${side}`, 0, 1, 0.5);
    if (n !== 0) out[side] = n;
  }
  const stops = o.get('stops');
  if (!Array.isArray(stops) || stops.length < 2 || stops.length > 256) {
    throw invalid('fill.stops', 'must be an array of 2 to 256 stops.', 'Use stops such as [{ position: 0, color: "FFFFFF" }, { position: 1, color: "4472C4" }].');
  }
  stops.forEach((stop: unknown, i) => {
    const so = readObject(stop, `fill.stops[${String(i)}]`, ['position', 'color']);
    out.stops.push({
      position: finite(so.get('position'), `fill.stops[${String(i)}].position`, 0, 1, 0.5),
      color: normalizeColor(so.get('color'), `fill.stops[${String(i)}].color`),
    });
  });
  return out;
}

function normalizeFill(value: unknown): NonNullable<NormalizedStyle['fill']> {
  if (typeof value === 'object' && value !== null && Object.keys(value).includes('type')) {
    if ((value as { type?: unknown }).type !== 'gradient') {
      throw invalid('fill.type', `has the value ${describe((value as { type?: unknown }).type)}.`, 'Use "gradient" for a gradient fill, or leave type out for a pattern fill.');
    }
    return normalizeGradient(value);
  }
  if (typeof value === 'object' && value !== null && Object.keys(value).includes('color')) {
    const o = readObject(value, 'fill', ['color']);
    return { pattern: 'solid', fgColor: normalizeColor(o.get('color'), 'fill.color') };
  }
  const o = readObject(value, 'fill', ['pattern', 'fgColor', 'bgColor']);
  const out: NonNullable<NormalizedStyle['fill']> = { pattern: oneOf(o.get('pattern'), 'fill.pattern', PATTERNS) };
  const fg = o.get('fgColor');
  if (fg !== undefined) out.fgColor = normalizeColor(fg, 'fill.fgColor');
  const bg = o.get('bgColor');
  if (bg !== undefined) out.bgColor = normalizeColor(bg, 'fill.bgColor');
  return out;
}

function normalizeBorder(value: unknown): NonNullable<NormalizedStyle['border']> {
  const o = readObject(value, 'border', [...EDGES, 'diagonal', 'diagonalUp', 'diagonalDown']);
  const out: NonNullable<NormalizedStyle['border']> = {};
  for (const flag of ['diagonalUp', 'diagonalDown'] as const) {
    const v = o.get(flag);
    if (v !== undefined && bool(v, `border.${flag}`)) out[flag] = true;
  }
  for (const edge of [...EDGES, 'diagonal'] as const) {
    const e = o.get(edge);
    if (e === undefined) continue;
    const eo = readObject(e, `border.${edge}`, ['style', 'color']);
    const style = oneOf(eo.get('style'), `border.${edge}.style`, BORDER_STYLES);
    const color = eo.get('color');
    out[edge] = color === undefined ? { style } : { style, color: normalizeColor(color, `border.${edge}.color`) };
  }
  return out;
}

function normalizeAlignment(value: unknown, source: StyleSource): NonNullable<NormalizedStyle['alignment']> {
  const o = readObject(value, 'alignment', ['horizontal', 'vertical', 'wrapText', 'indent', 'textRotation', 'shrinkToFit', 'justifyLastLine', 'readingOrder']);
  const out: NonNullable<NormalizedStyle['alignment']> = {};
  const h = o.get('horizontal');
  if (h !== undefined) out.horizontal = oneOf(h, 'alignment.horizontal', HORIZONTAL);
  const v = o.get('vertical');
  if (v !== undefined) out.vertical = oneOf(v, 'alignment.vertical', VERTICAL);
  for (const key of ['wrapText', 'shrinkToFit', 'justifyLastLine'] as const) {
    const v = o.get(key);
    if (v !== undefined && bool(v, `alignment.${key}`)) out[key] = true;
  }
  const order = o.get('readingOrder');
  if (order !== undefined) {
    const r = oneOf(order, 'alignment.readingOrder', READING_ORDERS);
    if (r !== 'context') out.readingOrder = r;
  }
  const indent = o.get('indent');
  if (indent !== undefined) {
    const n = integer(indent, 'alignment.indent', 0, source === 'input' ? MAX_INDENT : Number.MAX_SAFE_INTEGER, undefined, 2);
    if (n !== 0) out.indent = n;
  }
  const rotation = o.get('textRotation');
  if (rotation !== undefined) {
    // [MS-OI29500] §2.1.693 a: 0 to 180, or 255 for vertical text.
    const n = integer(rotation, 'alignment.textRotation', 0, 180, 255, 45);
    if (n !== 0) out.textRotation = n;
  }
  return out;
}

/**
 * Validates a style and returns a normalised copy. Throws `XLV_E_INVALID_STYLE` for unknown keys
 * (including `__proto__`), wrong types and values outside the allowed sets. Each property of the
 * input is read exactly once. With `source` `'input'` (the default) the values must be within the
 * limits Office accepts; `'existing'` is for styles read from a file being edited, which are kept
 * even when they are outside those limits.
 */
export function normalizeStyle(value: unknown, source: StyleSource = 'input'): NormalizedStyle {
  const o = readObject(value, 'style', ['font', 'fill', 'border', 'alignment', 'protection', 'numFmt']);
  const out: NormalizedStyle = {};
  const font = o.get('font');
  if (font !== undefined) {
    const f = normalizeFont(font, source);
    if (Object.keys(f).length > 0) out.font = f;
  }
  const fill = o.get('fill');
  if (fill !== undefined) out.fill = normalizeFill(fill);
  const border = o.get('border');
  if (border !== undefined) {
    const b = normalizeBorder(border);
    if (Object.keys(b).length > 0) out.border = b;
  }
  const alignment = o.get('alignment');
  if (alignment !== undefined) {
    const a = normalizeAlignment(alignment, source);
    if (Object.keys(a).length > 0) out.alignment = a;
  }
  const protection = o.get('protection');
  if (protection !== undefined) {
    const po = readObject(protection, 'protection', ['locked', 'hidden']);
    const p: NonNullable<NormalizedStyle['protection']> = {};
    for (const key of ['locked', 'hidden'] as const) {
      const v = po.get(key);
      if (v !== undefined) p[key] = bool(v, `protection.${key}`);
    }
    if (Object.keys(p).length > 0) out.protection = p;
  }
  const numFmt = o.get('numFmt');
  if (typeof numFmt === 'number') {
    // Only the "All Languages" ids; the others are locale-dependent or need a formatCode (E-41).
    if (!BUILTIN_NUMBER_FORMATS.has(numFmt)) {
      throw invalid('numFmt', `has the value ${describe(numFmt)}, which is not a built-in number format id.`, `Use a built-in id (${[...BUILTIN_NUMBER_FORMATS.keys()].join(', ')}), such as 14 for a date, or a format code such as "yyyy-mm-dd".`);
    }
    if (numFmt !== 0) out.numFmt = numFmt;
  } else if (numFmt !== undefined) {
    const max = source === 'input' ? MAX_NUMFMT_LENGTH : EXISTING_MAX_NUMFMT_LENGTH;
    if (typeof numFmt !== 'string' || numFmt.length === 0 || numFmt.length > max) {
      throw invalid('numFmt', `has the value ${describe(numFmt)}.`, `Use a number format code of 1 to ${String(max)} characters, such as "#,##0.00" or "yyyy-mm-dd".`);
    }
    if (numFmt !== 'General') out.numFmt = numFmt;
  }
  return out;
}

/** Merges `patch` into `base` part by part (font, fill, border, alignment, protection, numFmt). */
export function mergeStyles(base: NormalizedStyle, patch: NormalizedStyle): NormalizedStyle {
  const out: NormalizedStyle = { ...base };
  if (patch.font) out.font = { ...base.font, ...patch.font };
  if (patch.fill) out.fill = patch.fill;
  if (patch.border) out.border = { ...base.border, ...patch.border };
  if (patch.alignment) out.alignment = { ...base.alignment, ...patch.alignment };
  if (patch.protection) out.protection = { ...base.protection, ...patch.protection };
  if (patch.numFmt !== undefined) out.numFmt = patch.numFmt;
  return out;
}

/** A stable string key for a value made of plain objects, arrays and primitives. */
export function stableKey(value: unknown): string {
  if (typeof value !== 'object' || value === null) return JSON.stringify(value) ?? 'undefined';
  if (Array.isArray(value)) return `[${value.map(stableKey).join(',')}]`;
  const keys = Object.keys(value).sort();
  return `{${keys.map((k) => `${JSON.stringify(k)}:${stableKey((value as Record<string, unknown>)[k])}`).join(',')}}`;
}
