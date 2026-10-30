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
  | { readonly pattern: PatternType; readonly fgColor?: Color; readonly bgColor?: Color };

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
  };
  fill?: { pattern: PatternType; fgColor?: NormalizedColor; bgColor?: NormalizedColor };
  border?: { top?: BorderEdge; right?: BorderEdge; bottom?: BorderEdge; left?: BorderEdge };
  alignment?: {
    horizontal?: HorizontalAlignment;
    vertical?: VerticalAlignment;
    wrapText?: true;
    indent?: number;
    textRotation?: number;
  };
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
  const o = readObject(value, 'font', ['name', 'size', 'bold', 'italic', 'underline', 'strike', 'color']);
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
  for (const key of ['bold', 'italic', 'strike'] as const) {
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
  return out;
}

function normalizeFill(value: unknown): NonNullable<NormalizedStyle['fill']> {
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
  const o = readObject(value, 'border', EDGES);
  const out: NonNullable<NormalizedStyle['border']> = {};
  for (const edge of EDGES) {
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
  const o = readObject(value, 'alignment', ['horizontal', 'vertical', 'wrapText', 'indent', 'textRotation']);
  const out: NonNullable<NormalizedStyle['alignment']> = {};
  const h = o.get('horizontal');
  if (h !== undefined) out.horizontal = oneOf(h, 'alignment.horizontal', HORIZONTAL);
  const v = o.get('vertical');
  if (v !== undefined) out.vertical = oneOf(v, 'alignment.vertical', VERTICAL);
  const wrap = o.get('wrapText');
  if (wrap !== undefined && bool(wrap, 'alignment.wrapText')) out.wrapText = true;
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
  const o = readObject(value, 'style', ['font', 'fill', 'border', 'alignment', 'numFmt']);
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

/** Merges `patch` into `base` part by part (font, fill, border, alignment, numFmt). */
export function mergeStyles(base: NormalizedStyle, patch: NormalizedStyle): NormalizedStyle {
  const out: NormalizedStyle = { ...base };
  if (patch.font) out.font = { ...base.font, ...patch.font };
  if (patch.fill) out.fill = patch.fill;
  if (patch.border) out.border = { ...base.border, ...patch.border };
  if (patch.alignment) out.alignment = { ...base.alignment, ...patch.alignment };
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
