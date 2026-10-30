// Parses the number-format side of the Styles part (ECMA-376 Part 1 §12.3.20, §18.8; facts
// F-STY-01 … F-STY-09, decisions D-STY-1 … D-STY-6 in the facts file).
//
// - cellXfs/xf records are referenced from cells by zero-based index (`c@s`, §18.8.10).
// - xf@numFmtId is the id of a built-in (implied) or a declared numFmt (§18.8.30, §18.8.45);
//   numFmts/numFmt maps ids to format codes, matched by id, not by position.
// - The number format of a cell is the numFmtId of its cellXfs record alone. ECMA-376 says both
//   the cell style and the cell xf records are read (§18.8.10), but in Office only the cell xf
//   record defines a cell's formatting ([MS-OI29500] "Part 1 Section 18.8.10, cellXfs", note b;
//   fact F-STY-09, decision D-STY-2). applyNumberFormat and cellStyleXfs are therefore not used.
// - Ids from the file are kept in arrays and a Map, never as object keys. Invalid unsignedInt
//   values are reported as XLV_W_INVALID_VALUE and treated as absent.
import { XlvError } from '../errors';
import {
  BORDER_STYLES,
  EDGES,
  HORIZONTAL,
  PATTERNS,
  UNDERLINES,
  VERTICAL,
  type AlignmentStyle,
  type BorderEdge,
  type CellStyle,
  type Color,
  type FontStyle,
  type FillStyle,
} from '../style';
import { childNamed, childrenNamed, parseXmlTree, type XmlElement } from '../xml/tree';
import { SML_NS } from './namespaces';
import type { PartContext } from './shared-strings';
import { decodeXstring } from './xstring';

/** The number-format information of the Styles part. */
export interface Styles {
  /** Number of cell formats (cellXfs/xf entries). */
  readonly cellFormatCount: number;
  /** numFmtId of the cell format at this index, or 0 (General) when the index is out of range. */
  numFmtId(styleIndex: number): number;
  /** Format code of a custom number format declared in numFmts, if any. */
  customFormatCode(numFmtId: number): string | undefined;
  /**
   * The resolved, frozen style of the cell format at this index, with `numFmt` set to the given
   * code (the caller resolves number formats). Cached per index.
   */
  cellStyle(styleIndex: number, numFmt: string | null): CellStyle;
  /** Font 0, the workbook's default font (W2-12), if the part declares one. */
  defaultFont(): FontStyle | undefined;
  /** The differential format (dxfs/dxf) at a zero-based index, if there is one (SR-16). */
  dxf(index: number): CellStyle | undefined;
}

/** Font, fill, border and alignment ids of one cellXfs record. */
interface XfParts {
  readonly fontId: number;
  readonly fillId: number;
  readonly borderId: number;
  readonly alignment: AlignmentStyle | undefined;
}

const EMPTY_STYLE: CellStyle = Object.freeze({});

/** The number format id of General (§18.8.30, built-in id 0). */
const GENERAL = 0;
const MAX_UNSIGNED_INT = 0xffffffff;

class NumberFormats implements Styles {
  readonly cellFormatCount: number;
  private readonly resolved = new Map<number, CellStyle>();

  constructor(
    private readonly ids: readonly number[],
    private readonly codes: ReadonlyMap<number, string>,
    private readonly xfs: readonly XfParts[] = [],
    private readonly fonts: readonly (FontStyle | undefined)[] = [],
    private readonly fills: readonly (FillStyle | undefined)[] = [],
    private readonly borders: readonly (Readonly<Record<string, BorderEdge>> | undefined)[] = [],
    private readonly dxfs: readonly CellStyle[] = [],
  ) {
    this.cellFormatCount = ids.length;
  }

  defaultFont(): FontStyle | undefined {
    return this.fonts[0];
  }

  dxf(index: number): CellStyle | undefined {
    return Number.isInteger(index) && index >= 0 ? this.dxfs[index] : undefined;
  }

  cellStyle(styleIndex: number, numFmt: string | null): CellStyle {
    const cached = this.resolved.get(styleIndex);
    if (cached !== undefined) return cached;
    const xf = Number.isInteger(styleIndex) && styleIndex >= 0 ? this.xfs[styleIndex] : undefined;
    if (xf === undefined && numFmt === null) return EMPTY_STYLE;
    const style: { -readonly [K in keyof CellStyle]: CellStyle[K] } = {};
    if (xf !== undefined) {
      // Font 0 is the workbook's default font (W2-12); other fonts are reported in full.
      const font = xf.fontId === 0 ? undefined : this.fonts[xf.fontId];
      if (font !== undefined) style.font = font;
      const fill = this.fills[xf.fillId];
      if (fill !== undefined) style.fill = fill;
      const border = this.borders[xf.borderId];
      if (border !== undefined) style.border = border;
      if (xf.alignment !== undefined) style.alignment = xf.alignment;
    }
    if (numFmt !== null) style.numFmt = numFmt;
    const frozen = Object.freeze(style);
    if (xf !== undefined) this.resolved.set(styleIndex, frozen);
    return frozen;
  }

  numFmtId(styleIndex: number): number {
    if (!Number.isInteger(styleIndex) || styleIndex < 0 || styleIndex >= this.ids.length) return GENERAL;
    return this.ids[styleIndex] ?? GENERAL;
  }

  customFormatCode(numFmtId: number): string | undefined {
    return this.codes.get(numFmtId);
  }
}

/** Styles of a workbook without a Styles part: no cell formats, everything General. */
export const EMPTY_STYLES: Styles = new NumberFormats([], new Map());

/**
 * Parses the Styles part. Throws `XLV_E_PART_INVALID` if the root element is not `styleSheet`
 * in a SpreadsheetML namespace, and the XML parser's errors for malformed XML. Recoverable
 * problems are reported to `ctx.warnings` as `XLV_W_INVALID_VALUE`.
 */
export function parseStyles(xml: string, ctx: PartContext): Styles {
  const root = parseXmlTree(xml, ctx);
  if (root.local !== 'styleSheet' || (root.ns !== SML_NS.strict && root.ns !== SML_NS.transitional)) {
    throw new XlvError(
      'XLV_E_PART_INVALID',
      `The styles part has the root element <${root.local}> in namespace "${root.ns}"; expected <styleSheet> in the SpreadsheetML namespace.`,
      {
        location: { part: ctx.part },
        hint: 'The part the workbook names as its styles is not a SpreadsheetML style sheet. Re-save the file in a spreadsheet application; if it was generated by a program, report the problem to its authors.',
        details: { element: root.local, namespace: root.ns },
      },
    );
  }
  const ns = root.ns;
  const warn = (message: string): void => ctx.warnings.add('XLV_W_INVALID_VALUE', message, { part: ctx.part });

  /** The first child named `local`; further ones (the schema allows one) are reported and ignored. */
  const single = (local: string): XmlElement | undefined => {
    const found = childrenNamed(root, ns, local);
    if (found.length > 1) warn(`The styles part has ${String(found.length)} <${local}> elements; only the first is used.`);
    return found[0];
  };

  const codes = new Map<number, string>();
  const numFmts = single('numFmts');
  for (const el of numFmts === undefined ? [] : childrenNamed(numFmts, ns, 'numFmt')) {
    const rawId = el.attrs.get('numFmtId');
    const formatCode = el.attrs.get('formatCode');
    if (rawId === undefined || formatCode === undefined) {
      warn(`A numFmt element has no ${rawId === undefined ? 'numFmtId' : 'formatCode'} attribute; it is ignored.`);
      continue;
    }
    const id = parseUnsignedInt(rawId);
    if (id === null) {
      warn(`numFmt@numFmtId has the invalid value ${quote(rawId)}; the number format is ignored.`);
      continue;
    }
    if (codes.has(id)) {
      warn(`Number format ${String(id)} is declared more than once; the first declaration is used.`);
      continue;
    }
    codes.set(id, decodeXstring(formatCode));
  }

  const listOf = <T>(container: string, item: string, parse: (el: XmlElement) => T): T[] => {
    const el = single(container);
    return el === undefined ? [] : childrenNamed(el, ns, item).map(parse);
  };
  const colorOf = (el: XmlElement | undefined, what: string): Color | undefined => (el === undefined ? undefined : parseColor(el, what, warn));
  const fonts = listOf('fonts', 'font', (el) => parseFont(el, ns, colorOf, warn));
  const fills = listOf('fills', 'fill', (el) => parseFill(el, ns, colorOf, warn));
  const borders = listOf('borders', 'border', (el) => parseBorder(el, ns, colorOf, warn));

  const cellXfs = single('cellXfs');
  const xfElements = cellXfs === undefined ? [] : childrenNamed(cellXfs, ns, 'xf');
  const index = (xf: XmlElement, name: string): number => {
    const raw = xf.attrs.get(name);
    if (raw === undefined) return 0;
    const n = parseUnsignedInt(raw);
    if (n === null) warn(`cellXfs/xf@${name} has the invalid value ${quote(raw)}; using 0.`);
    return n ?? 0;
  };
  const xfs = xfElements.map(
    (xf): XfParts => ({
      fontId: index(xf, 'fontId'),
      fillId: index(xf, 'fillId'),
      borderId: index(xf, 'borderId'),
      alignment: parseAlignment(childNamed(xf, ns, 'alignment'), warn),
    }),
  );
  const ids = xfElements.map((xf): number => {
    const raw = xf.attrs.get('numFmtId');
    if (raw === undefined) return GENERAL;
    const id = parseUnsignedInt(raw);
    if (id === null) {
      warn(`cellXfs/xf@numFmtId has the invalid value ${quote(raw)}; using the General number format.`);
      return GENERAL;
    }
    return id;
  });

  const dxfs = listOf('dxfs', 'dxf', (el) => parseDxf(el, ns, colorOf, warn));
  return new NumberFormats(ids, codes, xfs, fonts, fills, borders, dxfs);
}

/**
 * Parses a differential format (CT_Dxf, SR-16): font, number format, fill and border; the
 * properties a dxf leaves out are not part of the result. A pattern fill without `patternType`
 * is read as solid, and a solid fill whose background equals its foreground as its foreground
 * only (implementation choice, SR-16: the standard does not say how a dxf's fill colours apply).
 */
function parseDxf(el: XmlElement, ns: string, colorOf: ColorOf, warn: (message: string) => void): CellStyle {
  const style: { -readonly [K in keyof CellStyle]: CellStyle[K] } = {};
  const fontEl = childNamed(el, ns, 'font');
  if (fontEl !== undefined) {
    const font = parseFont(fontEl, ns, colorOf, warn);
    if (Object.keys(font).length > 0) style.font = font;
  }
  const code = childNamed(el, ns, 'numFmt')?.attrs.get('formatCode');
  if (code !== undefined) style.numFmt = decodeXstring(code);
  const fillEl = childNamed(el, ns, 'fill');
  const pf = fillEl === undefined ? undefined : childNamed(fillEl, ns, 'patternFill');
  if (fillEl !== undefined && pf !== undefined && pf.attrs.get('patternType') === undefined) {
    const fg = colorOf(childNamed(pf, ns, 'fgColor'), 'fill');
    const bg = colorOf(childNamed(pf, ns, 'bgColor'), 'fill');
    if (fg !== undefined || bg !== undefined) style.fill = Object.freeze({ pattern: 'solid' as const, ...(fg === undefined ? {} : { fgColor: fg }), ...(bg === undefined ? {} : { bgColor: bg }) });
  } else if (fillEl !== undefined) {
    const fill = parseFill(fillEl, ns, colorOf, warn);
    if (fill !== undefined) {
      // parseFill always gives the pattern form.
      const f = fill as { readonly pattern: string; readonly fgColor?: Color; readonly bgColor?: Color };
      const same = f.pattern === 'solid' && f.fgColor !== undefined && f.bgColor !== undefined && JSON.stringify(f.fgColor) === JSON.stringify(f.bgColor);
      style.fill = same && f.fgColor !== undefined ? Object.freeze({ pattern: 'solid' as const, fgColor: f.fgColor }) : fill;
    }
  }
  const borderEl = childNamed(el, ns, 'border');
  const border = borderEl === undefined ? undefined : parseBorder(borderEl, ns, colorOf, warn);
  if (border !== undefined) style.border = border;
  return Object.freeze(style);
}

type ColorOf = (el: XmlElement | undefined, what: string) => Color | undefined;

/** Parses CT_Color (W2-16): rgb as ARGB, or theme (with tint), indexed or auto. */
export function parseColor(el: XmlElement, what: string, warn: (message: string) => void): Color | undefined {
  const rgb = el.attrs.get('rgb');
  if (rgb !== undefined) {
    if (/^[0-9A-Fa-f]{8}$/.test(rgb)) return rgb.toUpperCase();
    warn(`The ${what} colour ${quote(rgb)} is not an ARGB value; it is ignored.`);
    return undefined;
  }
  const theme = el.attrs.get('theme');
  if (theme !== undefined) {
    const n = parseUnsignedInt(theme);
    if (n === null) return undefined;
    const rawTint = el.attrs.get('tint');
    const tint = rawTint === undefined ? 0 : Number(rawTint.trim());
    return Number.isFinite(tint) && tint !== 0 && tint >= -1 && tint <= 1 ? Object.freeze({ theme: n, tint }) : Object.freeze({ theme: n });
  }
  const indexed = el.attrs.get('indexed');
  if (indexed !== undefined) {
    const n = parseUnsignedInt(indexed);
    return n === null ? undefined : Object.freeze({ indexed: n });
  }
  const auto = el.attrs.get('auto');
  if (auto !== undefined && parseBoolean(auto) === true) return Object.freeze({ auto: true as const });
  return undefined;
}

/** xsd:boolean: "true", "false", "1", "0" with surrounding white space. */
function parseBoolean(raw: string): boolean | null {
  const s = raw.trim();
  if (s === 'true' || s === '1') return true;
  if (s === 'false' || s === '0') return false;
  return null;
}

/** CT_BooleanProperty: `val` defaults to true (W2-11). */
function booleanProperty(el: XmlElement | undefined): boolean {
  if (el === undefined) return false;
  const val = el.attrs.get('val');
  return val === undefined ? true : parseBoolean(val) === true;
}

/**
 * Parses the run properties of a rich text run (CT_RPrElt, F-SST-08): the font properties of
 * CT_Font, with the font name in `rFont` instead of `name`. Unknown or invalid properties are left
 * out and reported through `warn`.
 */
export function parseRunFont(el: XmlElement, ns: string, warn: (message: string) => void): FontStyle {
  return parseFont(el, ns, (c, what) => (c === undefined ? undefined : parseColor(c, what, warn)), warn, 'rFont');
}

/** Parses CT_Font (W2-11). Unknown or invalid properties are left out. */
function parseFont(el: XmlElement, ns: string, colorOf: ColorOf, warn: (message: string) => void, nameElement = 'name'): FontStyle {
  const font: { -readonly [K in keyof FontStyle]: FontStyle[K] } = {};
  const name = childNamed(el, ns, nameElement)?.attrs.get('val');
  if (name !== undefined) font.name = decodeXstring(name);
  const sz = childNamed(el, ns, 'sz')?.attrs.get('val');
  if (sz !== undefined) {
    const n = Number(sz.trim());
    if (sz.trim() !== '' && Number.isFinite(n) && n > 0) font.size = n;
    else warn(`The font size ${quote(sz)} is not valid; it is ignored.`);
  }
  if (booleanProperty(childNamed(el, ns, 'b'))) font.bold = true;
  if (booleanProperty(childNamed(el, ns, 'i'))) font.italic = true;
  if (booleanProperty(childNamed(el, ns, 'strike'))) font.strike = true;
  const u = childNamed(el, ns, 'u');
  if (u !== undefined) {
    const val = u.attrs.get('val') ?? 'single';
    const found = UNDERLINES.find((x) => x === val);
    if (found !== undefined) font.underline = found;
    else if (val !== 'none') warn(`The underline style ${quote(val)} is unknown; it is ignored.`);
  }
  const color = colorOf(childNamed(el, ns, 'color'), 'font');
  if (color !== undefined) font.color = color;
  return Object.freeze(font);
}

/** Parses CT_Fill (W2-14); only pattern fills other than "none" give a fill. */
function parseFill(el: XmlElement, ns: string, colorOf: ColorOf, warn: (message: string) => void): FillStyle | undefined {
  const pf = childNamed(el, ns, 'patternFill');
  if (pf === undefined) return undefined;
  const type = pf.attrs.get('patternType');
  if (type === undefined || type === 'none') return undefined;
  const pattern = PATTERNS.find((p) => p === type);
  if (pattern === undefined) {
    warn(`The fill pattern ${quote(type)} is unknown; it is ignored.`);
    return undefined;
  }
  const fill: { pattern: typeof pattern; fgColor?: Color; bgColor?: Color } = { pattern };
  const fg = colorOf(childNamed(pf, ns, 'fgColor'), 'fill');
  if (fg !== undefined) fill.fgColor = fg;
  const bg = colorOf(childNamed(pf, ns, 'bgColor'), 'fill');
  if (bg !== undefined) fill.bgColor = bg;
  return Object.freeze(fill);
}

/** Parses CT_Border (W2-15): the left, right, top and bottom edges whose style is not "none". */
function parseBorder(
  el: XmlElement,
  ns: string,
  colorOf: ColorOf,
  warn: (message: string) => void,
): Readonly<Record<string, BorderEdge>> | undefined {
  const border: Record<string, BorderEdge> = {};
  let any = false;
  for (const edge of EDGES) {
    const e = childNamed(el, ns, edge);
    const raw = e?.attrs.get('style');
    if (e === undefined || raw === undefined || raw === 'none') continue;
    const style = BORDER_STYLES.find((b) => b === raw);
    if (style === undefined) {
      warn(`The border style ${quote(raw)} is unknown; it is ignored.`);
      continue;
    }
    const color = colorOf(childNamed(e, ns, 'color'), 'border');
    border[edge] = Object.freeze(color === undefined ? { style } : { style, color });
    any = true;
  }
  return any ? Object.freeze(border) : undefined;
}

/** Parses CT_CellAlignment (W2-17): the properties xlvane models, defaults left out. */
function parseAlignment(el: XmlElement | undefined, warn: (message: string) => void): AlignmentStyle | undefined {
  if (el === undefined) return undefined;
  const a: { -readonly [K in keyof AlignmentStyle]: AlignmentStyle[K] } = {};
  const h = el.attrs.get('horizontal');
  if (h !== undefined) {
    const found = HORIZONTAL.find((x) => x === h);
    if (found !== undefined) a.horizontal = found;
    else warn(`The horizontal alignment ${quote(h)} is unknown; it is ignored.`);
  }
  const v = el.attrs.get('vertical');
  if (v !== undefined) {
    const found = VERTICAL.find((x) => x === v);
    if (found !== undefined) a.vertical = found;
    else warn(`The vertical alignment ${quote(v)} is unknown; it is ignored.`);
  }
  const wrap = el.attrs.get('wrapText');
  if (wrap !== undefined && parseBoolean(wrap) === true) a.wrapText = true;
  const indent = el.attrs.get('indent');
  const indentN = indent === undefined ? null : parseUnsignedInt(indent);
  if (indentN !== null && indentN > 0) a.indent = indentN;
  const rotation = el.attrs.get('textRotation');
  const rotationN = rotation === undefined ? null : parseUnsignedInt(rotation);
  if (rotationN !== null && rotationN > 0) a.textRotation = rotationN;
  return Object.keys(a).length > 0 ? Object.freeze(a) : undefined;
}

/**
 * Parses an `xsd:unsignedInt` value: ASCII digits (leading zeros allowed) with optional
 * surrounding white space, at most 4294967295. Returns `null` for anything else. Linear time.
 */
function parseUnsignedInt(raw: string): number | null {
  const s = raw.trim();
  if (s.length === 0) return null;
  let value = 0;
  for (let i = 0; i < s.length; i++) {
    const d = s.charCodeAt(i) - 0x30;
    if (d < 0 || d > 9) return null;
    value = value * 10 + d;
    if (value > MAX_UNSIGNED_INT) return null;
  }
  return value;
}

/** Quotes a value from the file for a message, shortened to keep messages small. */
function quote(value: string): string {
  return JSON.stringify(value.length > 40 ? `${value.slice(0, 40)}…` : value);
}
