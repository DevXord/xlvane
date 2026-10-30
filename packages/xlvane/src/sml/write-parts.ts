// Serialises a workbook model to the parts of a Transitional SpreadsheetML package and zips them.
// Element order follows the schema sequences (docs/facts/sml-write.md W2-01 … W2-09); content
// types, relationship types and part names follow W2-30 … W2-33 and docs/facts/ecma-376.md E-03,
// E-05. Every string from the caller goes through escapeText/escapeAttr, and text that the
// schema types as ST_Xstring also through encodeXstring.
import { definedNameXml, type NormalizedDefinedName } from '../defined-names';
import type { NormalizedHyperlink } from '../hyperlinks';
import { printXml, sheetProtectionXml, workbookProtectionXml, type NormalizedWorkbookProtection, type SheetPrintData } from '../printing';
import { X14_CF_URI, X14_ID_URI, X14_NS, XM_NS, type CfValue, type DataValidation, type StoredCfRule, type StoredConditionalFormat } from '../sheet-rules';
import { COMMENTS_CONTENT_TYPE, VML_CONTENT_TYPE } from '../notes';
import {
  APP_PROPERTIES_CONTENT_TYPE,
  APP_PROPERTIES_PART,
  APP_PROPERTIES_REL,
  appPropertiesXml,
  CORE_PROPERTIES_CONTENT_TYPE,
  CORE_PROPERTIES_PART,
  CORE_PROPERTIES_REL,
  corePropertiesXml,
  type WorkbookProperties,
} from '../doc-properties';
import type { DateSystem, SheetState, SheetView } from '../model';
import { checkStyleCount, type NormalizedColor, type NormalizedStyle } from '../style';
import { escapeAttr, escapeText, encodeXstring } from '../xml/writer';
import { writeZip, type ZipWriteEntry } from '../zip/writer';
import { formatCellRef, type CellRange } from './cell-ref';
import { CONTENT_TYPES_NS, PACKAGE_REL_NS, REL_NS, SML_NS } from './namespaces';
import { BUILTIN_NUMBER_FORMATS } from './numfmt';
import type { StoredCell, StoredNote, StoredRun, StoredValue } from '../write';

/** One sheet, as handed over by `NewSheet`. */
export interface WriteSheetData {
  readonly name: string;
  readonly state: SheetState;
  readonly cells: ReadonlyMap<number, ReadonlyMap<number, StoredCell>>;
  /** Width, visibility and interned style id (0 = none) by column number. */
  readonly columns: ReadonlyMap<number, { readonly width: number | null; readonly hidden: boolean; readonly style: number }>;
  readonly rows: ReadonlyMap<number, { readonly height: number | null; readonly hidden: boolean }>;
  readonly merges: readonly CellRange[];
  readonly freeze: { readonly rows: number; readonly cols: number } | null;
  readonly autoFilter: CellRange | null;
  /** Colour of the sheet tab (W2-50); absent or `null` for none. */
  readonly tabColor?: NormalizedColor | null;
  /** Display settings (W2-51); absent for the defaults. */
  readonly view?: SheetView;
  /** Hyperlinks in the order they were added (W2-42). */
  readonly hyperlinks?: readonly NormalizedHyperlink[];
  /** Notes in row, then column order (W2-45). */
  readonly notes?: readonly StoredNote[];
  /** Conditional formats in the order they were added, with dxf indexes (SR-07 … SR-16). */
  readonly conditionalFormats?: readonly StoredConditionalFormat[];
  /** Data validations in the order they were added (SR-01 … SR-06). */
  readonly dataValidations?: readonly DataValidation[];
  /** Protection, page setup, headers and footers, page breaks (SR-20 … SR-28); absent for none. */
  readonly print?: SheetPrintData;
}

/** Everything needed to write a package. */
export interface WritePackageInput {
  readonly dateSystem: DateSystem;
  readonly sheets: readonly WriteSheetData[];
  /** Interned styles; index 0 is the default style. */
  readonly styles: readonly NormalizedStyle[];
  /** Document properties; `docProps/core.xml` is written only when one is set. */
  readonly properties?: WorkbookProperties;
  readonly definedNames?: readonly NormalizedDefinedName[];
  /** Differential formats of conditional formatting, by dxf index (SR-16). */
  readonly dxfs?: readonly NormalizedStyle[];
  /** `workbookProtection` (SR-22, D-C08). */
  readonly workbookProtection?: NormalizedWorkbookProtection | null;
}

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS = SML_NS.transitional;
const R_NS = REL_NS.transitional;
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const CT_BASE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.';

/** Default font of new workbooks (D-205). */
export const DEFAULT_FONT = { name: 'Calibri', size: 11 } as const;
/** First id used for number formats declared in the styles part (D-204). */
const FIRST_CUSTOM_NUMFMT_ID = 164;
/** Built-in ids reused when a style's format code equals theirs: the non-date ones (D-204). */
const REUSABLE_BUILTINS: ReadonlyMap<string, number> = new Map(
  [...BUILTIN_NUMBER_FORMATS].filter(([id]) => id !== 0 && !(id >= 14 && id <= 22) && !(id >= 45 && id <= 47)).map(([id, code]) => [code, id]),
);

const encoder = new TextEncoder();

/** Formats a finite number as an xsd:double lexical value that keeps every bit. */
export function formatNumber(n: number): string {
  // String(n) is the shortest text that reads back as the same double; only the exponent form
  // differs from xsd:double ("1e+21" → "1E21").
  return String(n).replace('e+', 'E').replace('e-', 'E-');
}

export function attr(name: string, value: string | number): string {
  return ` ${name}="${typeof value === 'number' ? formatNumber(value) : escapeAttr(value)}"`;
}

function rangeRef(r: CellRange): string {
  const a = formatCellRef(r.start.row, r.start.col);
  return r.start.row === r.end.row && r.start.col === r.end.col ? a : `${a}:${formatCellRef(r.end.row, r.end.col)}`;
}

function absoluteRangeRef(r: CellRange): string {
  const abs = (row: number, col: number): string => formatCellRef(row, col).replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');
  return `${abs(r.start.row, r.start.col)}:${abs(r.end.row, r.end.col)}`;
}

/** Quotes a sheet name for use in a formula: 'name' with apostrophes doubled (W2-08). */
function quoteSheetName(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

// ---------------------------------------------------------------------------------------------
// Styles

interface StyleTables {
  readonly xml: string;
  /** cellXfs index for each interned style id. */
  readonly xfOf: readonly number[];
}

export function colorXml(element: string, color: NormalizedColor): string {
  if (typeof color === 'string') return `<${element}${attr('rgb', color)}/>`;
  if ('theme' in color) return `<${element}${attr('theme', color.theme)}${color.tint === undefined ? '' : attr('tint', color.tint)}/>`;
  if ('indexed' in color) return `<${element}${attr('indexed', color.indexed)}/>`;
  return `<${element} auto="1"/>`;
}

class Table {
  readonly items: string[] = [];
  private readonly ids = new Map<string, number>();
  add(xml: string): number {
    let id = this.ids.get(xml);
    if (id === undefined) {
      id = this.items.length;
      this.items.push(xml);
      this.ids.set(xml, id);
    }
    return id;
  }
}

/**
 * A `font` record: the font completed with `defaults` (name and size). Children are written in the
 * order b, i, strike, u, sz, color, name: Office requires `i, strike, …, u, vertAlign, sz, color,
 * name, family, charset, scheme` ([MS-OI29500] §2.1.706 a, which omits `b`) and puts `b` first in
 * the run properties ([MS-OI29500] §2.1.674 a); CT_Font in sml.xsd allows any order (W2-11).
 */
export function fontXml(font: NonNullable<NormalizedStyle['font']>, defaults: { readonly name: string; readonly size: number } = DEFAULT_FONT): string {
  const f = { ...defaults, ...font };
  let xml = '<font>';
  if (f.bold) xml += '<b/>';
  if (f.italic) xml += '<i/>';
  if (f.strike) xml += '<strike/>';
  if (f.underline) xml += f.underline === 'single' ? '<u/>' : `<u${attr('val', f.underline)}/>`;
  xml += `<sz${attr('val', f.size)}/>`;
  if (f.color !== undefined) xml += colorXml('color', f.color);
  xml += `<name${attr('val', encodeXstring(f.name))}/>`;
  return `${xml}</font>`;
}

/** A `fill` record with a pattern fill. */
export function fillXml(fill: NonNullable<NormalizedStyle['fill']>): string {
  const { pattern, fgColor, bgColor } = fill;
  let inner = '';
  if (fgColor !== undefined) inner += colorXml('fgColor', fgColor);
  if (bgColor !== undefined) inner += colorXml('bgColor', bgColor);
  return `<fill><patternFill${attr('patternType', pattern)}${inner === '' ? '/>' : `>${inner}</patternFill>`}</fill>`;
}

/** A `border` record with the left, right, top and bottom edges in schema order (W2-15). */
export function borderXml(border: NonNullable<NormalizedStyle['border']>): string {
  let xml = '<border>';
  for (const edge of ['left', 'right', 'top', 'bottom'] as const) {
    const e = border[edge];
    xml += e === undefined ? `<${edge}/>` : `<${edge}${attr('style', e.style)}>${e.color === undefined ? '' : colorXml('color', e.color)}</${edge}>`;
  }
  return `${xml}<diagonal/></border>`;
}

/** The built-in id reused for a format code, if any (D-204). */
export function reusableBuiltinNumFmt(code: string): number | undefined {
  return REUSABLE_BUILTINS.get(code);
}

/**
 * Builds the styles part. With `identity`, cell format `i` is interned style `i` (no merging of
 * styles that give the same record), so a streaming writer can use style ids as `s` values
 * before the part is built.
 */
export function buildStyles(styles: readonly NormalizedStyle[], identity = false, dxfs: readonly NormalizedStyle[] = []): StyleTables {
  const fonts = new Table();
  const fills = new Table();
  const borders = new Table();
  const xfs = new Table();
  const numFmts = new Map<string, number>();

  // Fills 0 and 1 are reserved: none and gray125 (D-205, W2-14).
  fonts.add(fontXml({}));
  fills.add('<fill><patternFill patternType="none"/></fill>');
  fills.add('<fill><patternFill patternType="gray125"/></fill>');
  borders.add('<border><left/><right/><top/><bottom/><diagonal/></border>');

  const numFmtId = (code: string | number | undefined): number => {
    if (code === undefined) return 0;
    // A built-in id (checked by normalizeStyle) is implied and needs no numFmt element (E-41).
    if (typeof code === 'number') return code;
    const builtin = REUSABLE_BUILTINS.get(code);
    if (builtin !== undefined) return builtin;
    let id = numFmts.get(code);
    if (id === undefined) {
      id = FIRST_CUSTOM_NUMFMT_ID + numFmts.size;
      numFmts.set(code, id);
    }
    return id;
  };

  const xfOf = styles.map((style) => {
    const fontId = style.font ? fonts.add(fontXml(style.font)) : 0;
    let fillId = 0;
    if (style.fill) fillId = fills.add(fillXml(style.fill));
    let borderId = 0;
    if (style.border) borderId = borders.add(borderXml(style.border));
    const nf = numFmtId(style.numFmt);
    // Every cell xf has the Normal style xf (cellStyleXfs 0: numFmt, font, fill and border 0, no
    // alignment) as parent; where it differs, the matching applyX shall be true ([MS-OI29500]
    // §2.1.721 a–f), otherwise Office may replace the formatting when the Normal style changes.
    let xf = `<xf${attr('numFmtId', nf)}${attr('fontId', fontId)}${attr('fillId', fillId)}${attr('borderId', borderId)} xfId="0"`;
    if (nf !== 0) xf += ' applyNumberFormat="1"';
    if (fontId !== 0) xf += ' applyFont="1"';
    if (fillId !== 0) xf += ' applyFill="1"';
    if (borderId !== 0) xf += ' applyBorder="1"';
    const a = style.alignment;
    if (a) {
      let al = '<alignment';
      if (a.horizontal !== undefined) al += attr('horizontal', a.horizontal);
      if (a.vertical !== undefined) al += attr('vertical', a.vertical);
      if (a.textRotation !== undefined) al += attr('textRotation', a.textRotation);
      if (a.wrapText) al += ' wrapText="1"';
      if (a.indent !== undefined) al += attr('indent', a.indent);
      xf += ` applyAlignment="1">${al}/></xf>`;
    } else {
      xf += '/>';
    }
    if (identity) {
      xfs.items.push(xf);
      return xfs.items.length - 1;
    }
    return xfs.add(xf);
  });

  // Differential formats of conditional formatting (CT_Dxf, sml.xsd T3641; SR-16). Their number
  // formats get ids from the same table, so the numFmts part lists them too.
  const dxfXml = dxfs.map((d) => dxfRecordXml(d, numFmtId));
  checkStyleCount('customNumFmts', numFmts.size);
  checkStyleCount('fonts', fonts.items.length);
  checkStyleCount('fills', fills.items.length);
  checkStyleCount('borders', borders.items.length);
  checkStyleCount('cellXfs', xfs.items.length);
  let xml = `${XML_DECL}<styleSheet xmlns="${NS}">`;
  if (numFmts.size > 0) {
    xml += `<numFmts${attr('count', numFmts.size)}>`;
    for (const [code, id] of numFmts) xml += `<numFmt${attr('numFmtId', id)}${attr('formatCode', encodeXstring(code))}/>`;
    xml += '</numFmts>';
  }
  xml += `<fonts${attr('count', fonts.items.length)}>${fonts.items.join('')}</fonts>`;
  xml += `<fills${attr('count', fills.items.length)}>${fills.items.join('')}</fills>`;
  xml += `<borders${attr('count', borders.items.length)}>${borders.items.join('')}</borders>`;
  xml += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>';
  xml += `<cellXfs${attr('count', xfs.items.length)}>${xfs.items.join('')}</cellXfs>`;
  xml += '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>';
  if (dxfXml.length > 0) xml += `<dxfs${attr('count', dxfXml.length)}>${dxfXml.join('')}</dxfs>`;
  xml += '</styleSheet>';
  return { xml, xfOf };
}

// ---------------------------------------------------------------------------------------------
// Shared strings

/**
 * Where string cell values go: the shared string table. `add` takes the content of a string item
 * (see {@link stringItemXml}), so equal items, plain or rich, share one entry.
 */
export interface StringSink {
  add(item: string): number;
}

class SharedStrings implements StringSink {
  private readonly index = new Map<string, number>();
  readonly list: string[] = [];
  total = 0;
  add(item: string): number {
    this.total++;
    let i = this.index.get(item);
    if (i === undefined) {
      i = this.list.length;
      this.list.push(item);
      this.index.set(item, i);
    }
    return i;
  }
  xml(): string {
    let xml = `${XML_DECL}<sst xmlns="${NS}"${attr('count', this.total)}${attr('uniqueCount', this.list.length)}>`;
    for (const item of this.list) xml += `<si>${item}</si>`;
    return `${xml}</sst>`;
  }
}

/**
 * Run properties (`rPr`, CT_RPrElt, F-SST-08) of a rich text run, or '' for a run without font.
 * Only the given properties are written, in the order Office requires: b, i, strike, condense,
 * extend, outline, shadow, u, vertAlign, sz, color, rFont, family, charset, scheme
 * ([MS-OI29500] §2.1.674 a).
 */
export function runPropertiesXml(font: StoredRun['font']): string {
  if (font === undefined) return '';
  let xml = '';
  if (font.bold) xml += '<b/>';
  if (font.italic) xml += '<i/>';
  if (font.strike) xml += '<strike/>';
  if (font.underline) xml += font.underline === 'single' ? '<u/>' : `<u${attr('val', font.underline)}/>`;
  if (font.size !== undefined) xml += `<sz${attr('val', font.size)}/>`;
  if (font.color !== undefined) xml += colorXml('color', font.color);
  if (font.name !== undefined) xml += `<rFont${attr('val', encodeXstring(font.name))}/>`;
  return xml === '' ? '' : `<rPr>${xml}</rPr>`;
}

/**
 * The content of a string item (CT_Rst, F-SST-02): one `t` for plain text, or one `r` per rich
 * text run (`rPr` then `t`). Used for `si` in the shared string table and `is` in a cell.
 */
export function stringItemXml(v: Extract<StoredValue, { kind: 'string' }>): string {
  if (v.runs === undefined) return textElement(v.s);
  let xml = '';
  for (const run of v.runs) xml += `<r>${runPropertiesXml(run.font)}${textElement(run.text)}</r>`;
  return xml;
}

/** Never used for strings: inline strings are handled before valueXml is called. */
const INLINE_ONLY: StringSink = {
  add(): number {
    throw new Error('internal: inline strings are written by cellXml');
  },
};

/** A `t` element; `xml:space="preserve"` when white space at the ends or inside could be lost (W2-07). */
export function textElement(s: string): string {
  const preserve = /^[\s]|[\s]$|[\t\n\r]/.test(s) ? ' xml:space="preserve"' : '';
  return `<t${preserve}>${escapeText(encodeXstring(s))}</t>`;
}

// ---------------------------------------------------------------------------------------------
// Worksheets

function valueXml(v: Exclude<StoredValue, { kind: 'formula' }>, sst: StringSink, inFormula: boolean): { t: string; v: string } {
  switch (v.kind) {
    case 'number':
      return { t: '', v: formatNumber(v.n) };
    case 'string':
      // A formula's text result is stored in the cell as t="str"; plain text goes to the SST.
      return inFormula ? { t: 'str', v: escapeText(encodeXstring(v.s)) } : { t: 's', v: String(sst.add(stringItemXml(v))) };
    case 'boolean':
      return { t: 'b', v: v.b ? '1' : '0' };
    case 'error':
      return { t: 'e', v: escapeText(v.e) };
  }
}

/**
 * One `c` element. With `sst === null`, plain strings are written inline (`t="inlineStr"`,
 * E-22, E-29) instead of into the shared string table.
 */
export function cellXml(ref: string, cell: StoredCell, xfOf: readonly number[] | null, sst: StringSink | null): string {
  const xf = xfOf === null ? cell.style : (xfOf[cell.style] ?? 0);
  const s = xf === 0 ? '' : attr('s', xf);
  const v = cell.value;
  if (v === null) return `<c r="${ref}"${s}/>`;
  if (sst === null && v.kind === 'string') return `<c r="${ref}"${s} t="inlineStr"><is>${stringItemXml(v)}</is></c>`;
  if (v.kind === 'formula') {
    const f = `<f>${escapeText(encodeXstring(v.text))}</f>`;
    if (v.result === null) return `<c r="${ref}"${s}>${f}</c>`;
    const r = valueXml(v.result, sst ?? INLINE_ONLY, true);
    return `<c r="${ref}"${s}${r.t === '' ? '' : attr('t', r.t)}>${f}<v>${r.v}</v></c>`;
  }
  const r = valueXml(v, sst ?? INLINE_ONLY, false);
  return `<c r="${ref}"${s}${r.t === '' ? '' : attr('t', r.t)}><v>${r.v}</v></c>`;
}

/** Attributes of `sheetView` that differ from the CT_SheetView defaults (W2-51). */
function viewAttributes(view: SheetView | undefined): string {
  if (view === undefined) return '';
  let xml = '';
  if (!view.showGridLines) xml += ' showGridLines="0"';
  if (!view.showRowColHeaders) xml += ' showRowColHeaders="0"';
  if (view.rightToLeft) xml += ' rightToLeft="1"';
  if (view.layout !== 'normal') xml += attr('view', view.layout);
  if (view.zoom !== 100) xml += attr('zoomScale', view.zoom);
  return xml;
}

/**
 * The start of a worksheet part, up to and including `<sheetData>`; `dimension` is optional
 * (W2-01). `xfOf` maps interned style ids to cell format indexes (identity when omitted).
 */
export function sheetHead(
  sheet: Pick<WriteSheetData, 'freeze' | 'columns' | 'tabColor' | 'view' | 'print'>,
  selected: boolean,
  dimension: string | null,
  xfOf?: readonly number[],
): string {
  let xml = `${XML_DECL}<worksheet xmlns="${NS}" xmlns:r="${R_NS}">`;
  // sheetPr comes first, and tabColor first in it (CT_Worksheet, CT_SheetPr; W2-01, W2-50).
  // CT_SheetPr: tabColor, outlinePr, pageSetUpPr (sml.xsd T2306; SR-26).
  const fitToPage = sheet.print?.pageSetup.fitToPage === true;
  if (sheet.tabColor || fitToPage) xml += `<sheetPr>${sheet.tabColor ? colorXml('tabColor', sheet.tabColor) : ''}${fitToPage ? '<pageSetUpPr fitToPage="1"/>' : ''}</sheetPr>`;
  if (dimension !== null) xml += `<dimension${attr('ref', dimension)}/>`;
  xml += `<sheetViews><sheetView${selected ? ' tabSelected="1"' : ''}${viewAttributes(sheet.view)} workbookViewId="0"`;
  const split = sheet.view?.split ?? null;
  if (split !== null && !sheet.freeze) {
    // A split window: positions in 1/20 point, xSplit for the vertical bar; topLeftCell is
    // optional for state="split" ([MS-OI29500] §2.1.640 a, d; W2-51).
    const pane = split.x > 0 && split.y > 0 ? 'bottomRight' : split.y > 0 ? 'bottomLeft' : 'topRight';
    xml += '><pane';
    if (split.x > 0) xml += attr('xSplit', split.x * 20);
    if (split.y > 0) xml += attr('ySplit', split.y * 20);
    xml += `${attr('activePane', pane)} state="split"/></sheetView>`;
  } else if (sheet.freeze) {
    const { rows, cols } = sheet.freeze;
    const pane = rows > 0 && cols > 0 ? 'bottomRight' : rows > 0 ? 'bottomLeft' : 'topRight';
    const topLeft = formatCellRef(rows + 1, cols + 1);
    xml += '><pane';
    if (cols > 0) xml += attr('xSplit', cols);
    if (rows > 0) xml += attr('ySplit', rows);
    xml += `${attr('topLeftCell', topLeft)}${attr('activePane', pane)} state="frozen"/>`;
    xml += `<selection${attr('pane', pane)}${attr('activeCell', topLeft)}${attr('sqref', topLeft)}/></sheetView>`;
  } else {
    xml += '/>';
  }
  xml += '</sheetViews>';
  xml += '<sheetFormatPr defaultRowHeight="15"/>';
  if (sheet.columns.size > 0) {
    xml += '<cols>';
    for (const index of [...sheet.columns.keys()].sort((a, b) => a - b)) {
      const c = sheet.columns.get(index);
      if (!c) continue;
      xml += `<col${attr('min', index)}${attr('max', index)}`;
      if (c.width !== null) xml += `${attr('width', c.width)} customWidth="1"`;
      if (c.hidden) xml += ' hidden="1"';
      // CT_Col@style: index of a cell format, default 0 (sml.xsd L2238–2252, W2-03).
      const xf = c.style === 0 ? 0 : (xfOf?.[c.style] ?? c.style);
      if (xf !== 0) xml += attr('style', xf);
      xml += '/>';
    }
    xml += '</cols>';
  }
  return `${xml}<sheetData>`;
}

/** The start tag of a row (without `>` for rows that have cells: callers add `>` or `/>`). */
export function rowStart(r: number, info: { readonly height: number | null; readonly hidden: boolean } | undefined): string {
  let xml = `<row${attr('r', r)}`;
  if (info?.height != null) xml += `${attr('ht', info.height)} customHeight="1"`;
  if (info?.hidden) xml += ' hidden="1"';
  return xml;
}

/**
 * The end of a worksheet part, from `</sheetData>` (W2-01 order: autoFilter, mergeCells). `after`
 * is the XML of the elements that follow `mergeCells` in CT_Worksheet, already in schema order
 * (`hyperlinks`, …, `legacyDrawing`; W2-42).
 */
export function sheetTail(autoFilter: CellRange | null, merges: readonly CellRange[], after = '', protection = ''): string {
  // sheetProtection comes between sheetData and autoFilter (CT_Worksheet, sml.xsd T2179; SR-20).
  let xml = `</sheetData>${protection}`;
  if (autoFilter) xml += `<autoFilter${attr('ref', rangeRef(autoFilter))}/>`;
  if (merges.length > 0) {
    xml += `<mergeCells${attr('count', merges.length)}>`;
    for (const m of merges) xml += `<mergeCell${attr('ref', rangeRef(m))}/>`;
    xml += '</mergeCells>';
  }
  return `${xml}${after}</worksheet>`;
}

/**
 * The relationships of one worksheet part, collected while its XML is written. Each feature that
 * points from the sheet to another part or to an external resource (hyperlinks; comments and
 * their VML drawing) calls `add` and writes the returned id into the sheet's XML; `xml()` gives
 * the part `xl/worksheets/_rels/sheetN.xml.rels`, or `null` when nothing was added, in which case
 * no part is written. Relationship types are Transitional (Part 4 §13.3, W2-52), like every other
 * relationship the writer emits. The `.rels` extension is covered by the `Default` content type.
 */
export class SheetRels {
  private readonly items: string[] = [];

  /** Adds a relationship of type `REL_BASE + type` and returns its id (`rId1`, `rId2`, …). */
  add(type: string, target: string, external = false): string {
    const id = `rId${String(this.items.length + 1)}`;
    this.items.push(`<Relationship${attr('Id', id)}${attr('Type', REL_BASE + type)}${attr('Target', target)}${external ? ' TargetMode="External"' : ''}/>`);
    return id;
  }

  xml(): string | null {
    return this.items.length === 0 ? null : `${XML_DECL}<Relationships xmlns="${PACKAGE_REL_NS}">${this.items.join('')}</Relationships>`;
  }
}

/** Part name of the relationships part of worksheet `n` (1-based). */
export function sheetRelsPartName(n: number): string {
  return `xl/worksheets/_rels/sheet${String(n)}.xml.rels`;
}

// ---------------------------------------------------------------------------------------------
// Conditional formatting and data validation (SR-01 … SR-16)

/**
 * A differential format (CT_Dxf: font, numFmt, fill, alignment, border, protection; SR-16). The
 * font is written without the workbook defaults: only what the format changes. A solid fill with
 * one colour is written with that colour as both foreground and background, as in every dxf
 * example of ECMA-376 Part 1 (Annex L.2.7.3.8, L.2.7.4.3); the standard does not say which of
 * the two a spreadsheet application uses in a dxf (SR-16). `numFmtId` is required by the schema
 * and ignored by Office outside `numFmts` ([MS-OI29500] §2.1.712 c).
 */
function dxfRecordXml(style: NormalizedStyle, numFmtId: (code: string | number | undefined) => number): string {
  let xml = '<dxf>';
  const f = style.font;
  if (f) {
    xml += '<font>';
    if (f.bold) xml += '<b/>';
    if (f.italic) xml += '<i/>';
    if (f.strike) xml += '<strike/>';
    if (f.underline) xml += f.underline === 'single' ? '<u/>' : `<u${attr('val', f.underline)}/>`;
    if (f.size !== undefined) xml += `<sz${attr('val', f.size)}/>`;
    if (f.color !== undefined) xml += colorXml('color', f.color);
    if (f.name !== undefined) xml += `<name${attr('val', encodeXstring(f.name))}/>`;
    xml += '</font>';
  }
  if (style.numFmt !== undefined) {
    const id = numFmtId(style.numFmt);
    const code = typeof style.numFmt === 'number' ? (BUILTIN_NUMBER_FORMATS.get(style.numFmt) ?? 'General') : style.numFmt;
    xml += `<numFmt${attr('numFmtId', id)}${attr('formatCode', encodeXstring(code))}/>`;
  }
  if (style.fill) {
    const { pattern, fgColor, bgColor } = style.fill;
    const solidOne = pattern === 'solid' && (fgColor === undefined) !== (bgColor === undefined);
    const fg = solidOne ? (fgColor ?? bgColor) : fgColor;
    const bg = solidOne ? (bgColor ?? fgColor) : bgColor;
    xml += fillXml({ pattern, ...(fg === undefined ? {} : { fgColor: fg }), ...(bg === undefined ? {} : { bgColor: bg }) });
  }
  if (style.border) xml += borderXml(style.border);
  return `${xml}</dxf>`;
}

function cfValuesXml(values: readonly CfValue[]): string {
  // The main cfvo cannot hold autoMin / autoMax ([MS-XLSX] §2.7.9); a data bar that uses them
  // has the extension, whose value objects replace these (§2.6.30, SR-41).
  return values
    .map((v) => {
      const type = v.type === 'autoMin' ? 'min' : v.type === 'autoMax' ? 'max' : v.type;
      return `<cfvo${attr('type', type)}${v.value === null || type !== v.type ? '' : attr('val', encodeXstring(v.value))}${v.gte ? '' : ' gte="0"'}/>`;
    })
    .join('');
}

/** Value objects of the extension: the value in `xm:f`, `gte` only in icon sets ([MS-XLSX] §2.6.26). */
function x14CfValuesXml(values: readonly CfValue[], iconSet: boolean): string {
  return values
    .map((v) => {
      const open = `<x14:cfvo${attr('type', v.type)}${iconSet && !v.gte ? ' gte="0"' : ''}`;
      const noValue = v.value === null || v.type === 'min' || v.type === 'max' || v.type === 'autoMin' || v.type === 'autoMax';
      return noValue ? `${open}/>` : `${open}><xm:f>${escapeText(encodeXstring(v.value ?? ''))}</xm:f></x14:cfvo>`;
    })
    .join('');
}

/** Icon sets that exist only in the extension ([MS-XLSX] §2.7.10, SR-43). */
const X14_ICON_SETS = new Set(['3Stars', '3Triangles', '5Boxes', 'NoIcons']);

/** A rule that only the [MS-XLSX] extension can hold: an extension icon set or custom icons (SR-43). */
function extensionOnly(r: StoredCfRule): boolean {
  return r.iconSet !== null && (r.iconSet.custom !== null || X14_ICON_SETS.has(r.iconSet.iconSet));
}

/**
 * The `id` that links a main data bar rule to its extension (ST_Guid, [MS-XLSX] §2.4.28,
 * §2.6.27). [MS-XLSX] requires only that the main rule and the extension carry the same value;
 * xlvane derives it from the sheet number and the rule's position on the sheet, so the output is
 * the same for the same input (SR-41).
 */
export function ruleGuid(sheet: number, rule: number): string {
  const hex = (n: number, width: number): string => n.toString(16).toUpperCase().padStart(width, '0');
  return `{${hex(sheet, 8)}-0000-0000-0000-${hex(rule, 12)}}`;
}

/** `x14:dataBar` of a rule (§2.6.30); `fill` only for a rule with its own priority. */
function x14DataBarXml(d: NonNullable<StoredCfRule['dataBar']>, fill: boolean): string {
  let xml = '<x14:dataBar';
  if (d.minLength !== 10) xml += attr('minLength', d.minLength);
  if (d.maxLength !== 90) xml += attr('maxLength', d.maxLength);
  if (!d.showValue) xml += ' showValue="0"';
  if (d.border) xml += ' border="1"';
  if (!d.gradient) xml += ' gradient="0"';
  if (d.direction !== 'context') xml += attr('direction', d.direction);
  // The negative colours exist exactly when the "same as positive" flags are false (§2.6.30).
  if (d.negativeFillColor === null) xml += ' negativeBarColorSameAsPositive="1"';
  if (d.negativeBorderColor !== null) xml += ' negativeBarBorderColorSameAsPositive="0"';
  if (d.axisPosition !== 'automatic') xml += attr('axisPosition', d.axisPosition);
  xml += `>${x14CfValuesXml(d.cfvo, false)}`;
  if (fill) xml += colorXml('x14:fillColor', d.color);
  if (d.borderColor !== null) xml += colorXml('x14:borderColor', d.borderColor);
  if (d.negativeFillColor !== null) xml += colorXml('x14:negativeFillColor', d.negativeFillColor);
  if (d.negativeBorderColor !== null) xml += colorXml('x14:negativeBorderColor', d.negativeBorderColor);
  if (d.axisColor !== null) xml += colorXml('x14:axisColor', d.axisColor);
  return `${xml}</x14:dataBar>`;
}

/** `x14:iconSet` of a rule (§2.6.28, §2.6.36); `percent` is undefined there and not written. */
function x14IconSetXml(s: NonNullable<StoredCfRule['iconSet']>): string {
  let xml = '<x14:iconSet';
  if (s.iconSet !== '3TrafficLights1') xml += attr('iconSet', s.iconSet);
  if (!s.showValue) xml += ' showValue="0"';
  // reverse is ignored with custom icons (§2.6.28).
  if (s.reverse && s.custom === null) xml += ' reverse="1"';
  if (s.custom !== null) xml += ' custom="1"';
  xml += `>${x14CfValuesXml(s.cfvo, true)}`;
  for (const icon of s.custom ?? []) xml += `<x14:cfIcon${attr('iconSet', icon.iconSet)}${attr('iconId', icon.iconId)}/>`;
  return `${xml}</x14:iconSet>`;
}

/** The parts of the conditional formatting and validations of a sheet; see {@link sheetRulesXml}. */
export interface SheetRulesXml {
  /** `conditionalFormatting*` and `dataValidations`, written after `mergeCells`. */
  readonly main: string;
  /** The worksheet `extLst` with the [MS-XLSX] extensions, the last child of `worksheet`; or ''. */
  readonly ext: string;
}

/**
 * The `conditionalFormatting` elements and the `dataValidations` element of a sheet, in this
 * order (CT_Worksheet: … mergeCells, phoneticPr, conditionalFormatting*, dataValidations,
 * hyperlinks …, sml.xsd T2186–T2191; SR-01, SR-07), and the worksheet `extLst` with the
 * [MS-XLSX] extensions (SR-40 … SR-43): the extension of a data bar that uses extension
 * features, linked to its main rule by `id`, and rules only the extension can hold, with their
 * own priority. Attributes are written only where they differ from the schema defaults. `sheet`
 * is the 1-based sheet number, used for the rule ids.
 */
export function sheetRulesXml(sheet: Pick<WriteSheetData, 'conditionalFormats' | 'dataValidations'>, sheetNumber = 1): SheetRulesXml {
  let xml = '';
  let ext = '';
  let ruleNumber = 0;
  for (const cf of sheet.conditionalFormats ?? []) {
    let main = '';
    let extended = '';
    cf.rules.forEach((r, i) => {
      ruleNumber++;
      if (extensionOnly(r) && r.iconSet !== null) {
        extended += `<x14:cfRule type="iconSet"${attr('priority', r.priority)}${r.stopIfTrue ? ' stopIfTrue="1"' : ''}>${x14IconSetXml(r.iconSet)}</x14:cfRule>`;
        return;
      }
      const dxf = cf.dxfIds[i];
      main += `<cfRule${attr('type', r.type)}${dxf === null || dxf === undefined ? '' : attr('dxfId', dxf)}${attr('priority', r.priority)}`;
      if (r.stopIfTrue) main += ' stopIfTrue="1"';
      if (r.type === 'aboveAverage' && !r.aboveAverage) main += ' aboveAverage="0"';
      if (r.percent) main += ' percent="1"';
      if (r.bottom) main += ' bottom="1"';
      if (r.operator !== null) main += attr('operator', r.operator);
      if (r.text !== null) main += attr('text', r.text);
      if (r.timePeriod !== null) main += attr('timePeriod', r.timePeriod);
      if (r.rank !== null) main += attr('rank', r.rank);
      if (r.stdDev !== null) main += attr('stdDev', r.stdDev);
      if (r.equalAverage) main += ' equalAverage="1"';
      let inner = r.formulas.map((f) => `<formula>${escapeText(encodeXstring(f))}</formula>`).join('');
      if (r.colorScale) inner += `<colorScale>${cfValuesXml(r.colorScale.cfvo)}${r.colorScale.colors.map((c) => colorXml('color', c)).join('')}</colorScale>`;
      if (r.dataBar) {
        const d = r.dataBar;
        // A main rule with an extension of 0 and 100 has 10 and 90 ([MS-XLSX] §2.6.30, SR-41).
        const full = d.extended && d.minLength === 0 && d.maxLength === 100;
        const min = full ? 10 : d.minLength;
        const max = full ? 90 : d.maxLength;
        inner += `<dataBar${min === 10 ? '' : attr('minLength', min)}${max === 90 ? '' : attr('maxLength', max)}${d.showValue ? '' : ' showValue="0"'}>${cfValuesXml(d.cfvo)}${colorXml('color', d.color)}</dataBar>`;
        if (d.extended) {
          const id = ruleGuid(sheetNumber, ruleNumber);
          inner += `<extLst><ext${attr('uri', X14_ID_URI)}${attr('xmlns:x14', X14_NS)}><x14:id>${id}</x14:id></ext></extLst>`;
          extended += `<x14:cfRule type="dataBar"${attr('id', id)}>${x14DataBarXml(d, false)}</x14:cfRule>`;
        }
      }
      if (r.iconSet) {
        const s = r.iconSet;
        inner += `<iconSet${s.iconSet === '3TrafficLights1' ? '' : attr('iconSet', s.iconSet)}${s.showValue ? '' : ' showValue="0"'}${s.percent ? '' : ' percent="0"'}${s.reverse ? ' reverse="1"' : ''}>${cfValuesXml(s.cfvo)}</iconSet>`;
      }
      main += inner === '' ? '/>' : `>${inner}</cfRule>`;
    });
    // Part 1 requires at least one cfRule in conditionalFormatting (§18.3.1.18).
    if (main !== '') xml += `<conditionalFormatting${attr('sqref', cf.sqref)}>${main}</conditionalFormatting>`;
    if (extended !== '') ext += `<x14:conditionalFormatting>${extended}<xm:sqref>${cf.sqref}</xm:sqref></x14:conditionalFormatting>`;
  }
  const dvs = sheet.dataValidations ?? [];
  if (dvs.length > 0) {
    xml += `<dataValidations${attr('count', dvs.length)}>`;
    for (const d of dvs) {
      xml += `<dataValidation${d.type === 'none' ? '' : attr('type', d.type)}`;
      if (d.errorStyle !== 'stop') xml += attr('errorStyle', d.errorStyle);
      if (d.imeMode !== 'noControl') xml += attr('imeMode', d.imeMode);
      if (d.operator !== null && d.operator !== 'between') xml += attr('operator', d.operator);
      if (d.allowBlank) xml += ' allowBlank="1"';
      if (d.showDropDown) xml += ' showDropDown="1"';
      if (d.showInputMessage) xml += ' showInputMessage="1"';
      if (d.showErrorMessage) xml += ' showErrorMessage="1"';
      for (const k of ['errorTitle', 'error', 'promptTitle', 'prompt'] as const) {
        const v = d[k];
        if (v !== null) xml += attr(k, encodeXstring(v));
      }
      xml += attr('sqref', d.sqref);
      const f = (d.formula1 === null ? '' : `<formula1>${escapeText(encodeXstring(d.formula1))}</formula1>`) + (d.formula2 === null ? '' : `<formula2>${escapeText(encodeXstring(d.formula2))}</formula2>`);
      xml += f === '' ? '/>' : `>${f}</dataValidation>`;
    }
    xml += '</dataValidations>';
  }
  return {
    main: xml,
    ext: ext === '' ? '' : `<extLst><ext${attr('uri', X14_CF_URI)}${attr('xmlns:x14', X14_NS)}${attr('xmlns:xm', XM_NS)}><x14:conditionalFormattings>${ext}</x14:conditionalFormattings></ext></extLst>`,
  };
}

/**
 * The `hyperlinks` element (§18.3.1.47, §18.3.1.48; W2-42). An external target gets a
 * relationship with TargetMode="External" and `r:id`; a place in the workbook is written in
 * `location` only, without a relationship.
 */
export function hyperlinksXml(links: readonly NormalizedHyperlink[], rels: SheetRels): string {
  if (links.length === 0) return '';
  let xml = '<hyperlinks>';
  for (const l of links) {
    xml += `<hyperlink${attr('ref', l.ref)}`;
    if (l.url !== null) xml += attr('r:id', rels.add('hyperlink', l.url, true));
    if (l.location !== null) xml += attr('location', encodeXstring(l.location));
    if (l.tooltip !== null) xml += attr('tooltip', encodeXstring(l.tooltip));
    if (l.display !== null) xml += attr('display', encodeXstring(l.display));
    xml += '/>';
  }
  return `${xml}</hyperlinks>`;
}

// ---------------------------------------------------------------------------------------------
// Notes (comments): a Comments part with the text and a VML Drawing part with the boxes

/** The parts that hold the notes of one worksheet, and the `legacyDrawing` element for it. */
export interface NoteParts {
  /** `<legacyDrawing r:id="…"/>`, to be written after `hyperlinks` (CT_Worksheet, W2-48). */
  readonly legacyDrawing: string;
  readonly comments: { readonly name: string; readonly xml: string };
  readonly vml: { readonly name: string; readonly xml: string };
  /** Number of 1024-id blocks of shape ids the VML part uses. */
  readonly blocks: number;
}

/** Part name of the Comments part of worksheet `n` (1-based). */
export function commentsPartName(n: number): string {
  return `xl/comments${String(n)}.xml`;
}

/**
 * Builds the Comments part (W2-45, W2-46) and the VML Drawing part (W2-47) of worksheet `n` and
 * adds their relationships to `rels`; `null` when the sheet has no notes. Shape ids use the
 * 1024-id blocks from `block` on, written in `o:idmap`, as in the examples of Part 1 Annex
 * L.5.5.2 and Part 4 §19.4 (`data="1"` with ids from 1025).
 */
export function noteParts(notes: readonly StoredNote[], n: number, rels: SheetRels, block: number): NoteParts | null {
  if (notes.length === 0) return null;
  const commentsName = commentsPartName(n);
  const vmlName = `xl/drawings/vmlDrawing${String(n)}.vml`;
  rels.add('comments', `../comments${String(n)}.xml`);
  const vmlId = rels.add('vmlDrawing', `../drawings/vmlDrawing${String(n)}.vml`);
  const authors = new Map<string, number>();
  let list = '';
  for (const note of notes) {
    let id = authors.get(note.author);
    if (id === undefined) {
      id = authors.size;
      authors.set(note.author, id);
    }
    list += `<comment${attr('ref', formatCellRef(note.row, note.col))}${attr('authorId', id)}><text>${stringItemXml(note.value)}</text></comment>`;
  }
  let comments = `${XML_DECL}<comments xmlns="${NS}"><authors>`;
  for (const author of authors.keys()) comments += `<author>${escapeText(encodeXstring(author))}</author>`;
  comments += `</authors><commentList>${list}</commentList></comments>`;
  const blocks = Math.ceil(notes.length / 1024);
  return { legacyDrawing: `<legacyDrawing${attr('r:id', vmlId)}/>`, comments: { name: commentsName, xml: comments }, vml: { name: vmlName, xml: vmlXml(notes, block, blocks) }, blocks };
}

/** Points per default column (64 px) and row (20 px) and per pixel, for the box position. */
const COLUMN_PT = 48;
const ROW_PT = 15;
const PX_PT = 0.75;

/**
 * The VML Drawing part (Part 4 §8.2, §19.1.2.19, §19.4; W2-47): one hidden text box shape per
 * note, with the `x:ClientData` of the Part 4 §19.4 example minus `x:Visible` (so the note is
 * hidden, §19.4.2.65). The box lies right of the cell; its anchor follows the Part 1 Annex
 * L.5.5.2 example, and the CSS position is computed for default column widths and row heights.
 */
function vmlXml(notes: readonly StoredNote[], block: number, blocks: number): string {
  const ids = Array.from({ length: blocks }, (_, i) => String(block + i)).join(',');
  let xml =
    '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
    `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit"${attr('data', ids)}/></o:shapelayout>` +
    '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';
  notes.forEach((note, k) => {
    const row = note.row - 1;
    const col = note.col - 1;
    const top = Math.max(row - 1, 0);
    const style =
      `position:absolute;margin-left:${String((col + 1) * COLUMN_PT + 15 * PX_PT)}pt;margin-top:${String(top * ROW_PT + 10 * PX_PT)}pt;` +
      `width:${String(2 * COLUMN_PT)}pt;height:${String(4 * ROW_PT - 6 * PX_PT)}pt;z-index:${String(k + 1)};visibility:hidden`;
    xml +=
      `<v:shape${attr('id', `_x0000_s${String(block * 1024 + k + 1)}`)} type="#_x0000_t202"${attr('style', style)} fillcolor="#ffffe1" o:insetmode="auto">` +
      '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>' +
      '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>' +
      '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>' +
      `<x:Anchor>${[col + 1, 15, top, 10, col + 3, 15, top + 4, 4].join(', ')}</x:Anchor>` +
      `<x:AutoFill>False</x:AutoFill><x:Row>${String(row)}</x:Row><x:Column>${String(col)}</x:Column></x:ClientData></v:shape>`;
  });
  return `${xml}</xml>`;
}

function sheetXml(sheet: WriteSheetData, selected: boolean, xfOf: readonly number[], sst: SharedStrings, rels: SheetRels, after = '', sheetNumber = 1): string {
  const rowNumbers = new Set<number>([...sheet.cells.keys(), ...sheet.rows.keys()]);
  const sortedRows = [...rowNumbers].sort((a, b) => a - b);

  let minRow = Infinity;
  let maxRow = 0;
  let minCol = Infinity;
  let maxCol = 0;
  for (const [r, cells] of sheet.cells) {
    for (const c of cells.keys()) {
      if (r < minRow) minRow = r;
      if (r > maxRow) maxRow = r;
      if (c < minCol) minCol = c;
      if (c > maxCol) maxCol = c;
    }
  }
  const dimension = maxRow === 0 ? 'A1' : rangeRef({ start: { row: minRow, col: minCol }, end: { row: maxRow, col: maxCol } });

  let xml = sheetHead(sheet, selected, dimension, xfOf);
  for (const r of sortedRows) {
    const cells = sheet.cells.get(r);
    const start = rowStart(r, sheet.rows.get(r));
    if (!cells || cells.size === 0) {
      xml += `${start}/>`;
      continue;
    }
    xml += `${start}>`;
    for (const c of [...cells.keys()].sort((a, b) => a - b)) {
      const cell = cells.get(c);
      if (cell) xml += cellXml(formatCellRef(r, c), cell, xfOf, sst);
    }
    xml += '</row>';
  }
  const print = sheet.print;
  const rules = sheetRulesXml(sheet, sheetNumber);
  return xml + sheetTail(sheet.autoFilter, sheet.merges, rules.main + hyperlinksXml(sheet.hyperlinks ?? [], rels) + (print ? printXml(print) : '') + after + rules.ext, print ? sheetProtectionXml(print.protection) : '');
}

// ---------------------------------------------------------------------------------------------
// Package

/** Builds all parts and returns the zipped package. */
export function writePackage(input: WritePackageInput): Uint8Array {
  const { xml: stylesXml, xfOf } = buildStyles(input.styles, false, input.dxfs);
  const sst = new SharedStrings();
  const active = input.sheets.findIndex((s) => s.state === 'visible');
  const sheetRels = input.sheets.map(() => new SheetRels());
  const notes: (NoteParts | null)[] = [];
  let block = 1;
  const sheetParts = input.sheets.map((sheet, i) => {
    const rels = sheetRels[i]!;
    const parts = noteParts(sheet.notes ?? [], i + 1, rels, block);
    notes.push(parts);
    if (parts !== null) block += parts.blocks;
    return sheetXml(sheet, i === active, xfOf, sst, rels, parts?.legacyDrawing, i + 1);
  });

  const n = input.sheets.length;
  const workbook = workbookXml(input.sheets, input.dateSystem, input.definedNames, input.workbookProtection ?? null);
  const workbookRels = workbookRelsXml(n, true);
  const core = corePropertiesXml(input.properties ?? {});
  const app = appPropertiesXml(input.properties ?? {}, input.sheets.map((sheet) => sheet.name), nameTitles(input.definedNames ?? []));
  const rootRels = rootRelsXml(core !== null, app !== null);
  const noteSheets = notes.flatMap((p, i) => (p === null ? [] : [i + 1]));
  const types = contentTypesXml(n, true, core !== null, app !== null, noteSheets);

  const entries: ZipWriteEntry[] = [
    { name: '[Content_Types].xml', data: encoder.encode(types) },
    { name: '_rels/.rels', data: encoder.encode(rootRels) },
    { name: 'xl/workbook.xml', data: encoder.encode(workbook) },
    { name: 'xl/_rels/workbook.xml.rels', data: encoder.encode(workbookRels) },
    ...sheetParts.map((xml, i) => ({ name: `xl/worksheets/sheet${String(i + 1)}.xml`, data: encoder.encode(xml) })),
    { name: 'xl/styles.xml', data: encoder.encode(stylesXml) },
    { name: 'xl/sharedStrings.xml', data: encoder.encode(sst.xml()) },
  ];
  sheetRels.forEach((rels, i) => {
    const xml = rels.xml();
    if (xml !== null) entries.push({ name: sheetRelsPartName(i + 1), data: encoder.encode(xml) });
    const parts = notes[i];
    if (parts) entries.push({ name: parts.comments.name, data: encoder.encode(parts.comments.xml) }, { name: parts.vml.name, data: encoder.encode(parts.vml.xml) });
  });
  if (core !== null) entries.push({ name: CORE_PROPERTIES_PART, data: encoder.encode(core) });
  if (app !== null) entries.push({ name: APP_PROPERTIES_PART, data: encoder.encode(app) });
  return writeZip(entries);
}

/**
 * The defined names as listed in `TitlesOfParts` of `docProps/app.xml` (Part 1 §18.17.6.5): names
 * that are not hidden, a sheet-level name prefixed with its quoted sheet, e.g. `'Data'!Total`.
 * This follows the files written by Excel in the Apache POI corpus (W2-52).
 */
export function nameTitles(definedNames: readonly NormalizedDefinedName[]): string[] {
  // Built-in names are listed without "_xlnm.", e.g. 'Sheet 2'!Print_Titles (POI corpus, SR-31).
  return definedNames
    .filter((n) => !n.hidden)
    .map((n) => {
      const name = n.name.startsWith('_xlnm.') ? n.name.slice(6) : n.name;
      return n.sheet === null ? name : `${quoteSheetName(n.sheet)}!${name}`;
    });
}

/** The workbook part: date system, book view, sheets, and the autofilter and user defined names. */
export function workbookXml(
  sheets: readonly { readonly name: string; readonly state: SheetState; readonly autoFilter: CellRange | null }[],
  dateSystem: DateSystem,
  definedNames: readonly NormalizedDefinedName[] = [],
  protection: NormalizedWorkbookProtection | null = null,
): string {
  const active = sheets.findIndex((s) => s.state === 'visible');
  let workbook = `${XML_DECL}<workbook xmlns="${NS}" xmlns:r="${R_NS}">`;
  if (dateSystem === '1904') workbook += '<workbookPr date1904="1"/>';
  workbook += workbookProtectionXml(protection);
  // Always written, also by the streaming writer: Excel fails to print files without bookViews
  // ([MS-OI29500] §2.1.584 a).
  workbook += `<bookViews><workbookView${active > 0 ? `${attr('firstSheet', active)}${attr('activeTab', active)}` : ''}/></bookViews>`;
  workbook += '<sheets>';
  sheets.forEach((sheet, i) => {
    workbook += `<sheet${attr('name', encodeXstring(sheet.name))}${attr('sheetId', i + 1)}`;
    if (sheet.state !== 'visible') workbook += attr('state', sheet.state);
    workbook += `${attr('r:id', `rId${String(i + 1)}`)}/>`;
  });
  workbook += '</sheets>';
  const names = sheets
    .map((sheet, i) =>
      sheet.autoFilter
        ? `<definedName name="_xlnm._FilterDatabase"${attr('localSheetId', i)} hidden="1">${escapeText(
            encodeXstring(`${quoteSheetName(sheet.name)}!${absoluteRangeRef(sheet.autoFilter)}`),
          )}</definedName>`
        : '',
    )
    .join('') +
    definedNames
      .map((n) => {
        const id = n.sheet === null ? null : sheets.findIndex((s) => s.name === n.sheet);
        return definedNameXml(n, id === null || id < 0 ? null : id);
      })
      .join('');
  if (names !== '') workbook += `<definedNames>${names}</definedNames>`;
  return `${workbook}</workbook>`;
}

/** Relationships of the workbook part: sheets rId1…rIdN, then styles and (optionally) shared strings. */
export function workbookRelsXml(sheetCount: number, sharedStrings: boolean): string {
  let xml = `${XML_DECL}<Relationships xmlns="${PACKAGE_REL_NS}">`;
  for (let i = 1; i <= sheetCount; i++) {
    xml += `<Relationship Id="rId${String(i)}" Type="${REL_BASE}worksheet" Target="worksheets/sheet${String(i)}.xml"/>`;
  }
  xml += `<Relationship Id="rId${String(sheetCount + 1)}" Type="${REL_BASE}styles" Target="styles.xml"/>`;
  if (sharedStrings) xml += `<Relationship Id="rId${String(sheetCount + 2)}" Type="${REL_BASE}sharedStrings" Target="sharedStrings.xml"/>`;
  return `${xml}</Relationships>`;
}

/** The package relationships part, with the Core and Extended File Properties relationships when asked. */
export function rootRelsXml(coreProperties = false, appProperties = false): string {
  return (
    `${XML_DECL}<Relationships xmlns="${PACKAGE_REL_NS}">` +
    `<Relationship Id="rId1" Type="${REL_BASE}officeDocument" Target="xl/workbook.xml"/>` +
    (coreProperties ? `<Relationship Id="rId2" Type="${CORE_PROPERTIES_REL}" Target="${CORE_PROPERTIES_PART}"/>` : '') +
    (appProperties ? `<Relationship Id="rId3" Type="${APP_PROPERTIES_REL}" Target="${APP_PROPERTIES_PART}"/>` : '') +
    '</Relationships>'
  );
}

/**
 * The Media Types stream (W2-30). `noteSheets` lists the worksheets (1-based) that have notes:
 * each gets an Override for its Comments part, and the VML parts are covered by a `vml` Default.
 */
export function contentTypesXml(sheetCount: number, sharedStrings: boolean, coreProperties = false, appProperties = false, noteSheets: readonly number[] = []): string {
  let types = `${XML_DECL}<Types xmlns="${CONTENT_TYPES_NS}">`;
  types += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
  types += '<Default Extension="xml" ContentType="application/xml"/>';
  if (noteSheets.length > 0) types += `<Default Extension="vml"${attr('ContentType', VML_CONTENT_TYPE)}/>`;
  for (const i of noteSheets) types += `<Override${attr('PartName', `/${commentsPartName(i)}`)}${attr('ContentType', COMMENTS_CONTENT_TYPE)}/>`;
  types += `<Override PartName="/xl/workbook.xml" ContentType="${CT_BASE}sheet.main+xml"/>`;
  for (let i = 1; i <= sheetCount; i++) types += `<Override PartName="/xl/worksheets/sheet${String(i)}.xml" ContentType="${CT_BASE}worksheet+xml"/>`;
  types += `<Override PartName="/xl/styles.xml" ContentType="${CT_BASE}styles+xml"/>`;
  if (sharedStrings) types += `<Override PartName="/xl/sharedStrings.xml" ContentType="${CT_BASE}sharedStrings+xml"/>`;
  if (coreProperties) types += `<Override PartName="/${CORE_PROPERTIES_PART}" ContentType="${CORE_PROPERTIES_CONTENT_TYPE}"/>`;
  if (appProperties) types += `<Override PartName="/${APP_PROPERTIES_PART}" ContentType="${APP_PROPERTIES_CONTENT_TYPE}"/>`;
  return `${types}</Types>`;
}

/** A shared string table that grows as strings are added; `xml()` writes the part. */
export function createSharedStrings(): StringSink & { xml(): string } {
  return new SharedStrings();
}
