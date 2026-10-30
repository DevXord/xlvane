// Serialises a workbook model to the parts of a Transitional SpreadsheetML package and zips them.
// Element order follows the schema sequences (docs/facts/sml-write.md W2-01 … W2-09); content
// types, relationship types and part names follow W2-30 … W2-33 and docs/facts/ecma-376.md E-03,
// E-05. Every string from the caller goes through escapeText/escapeAttr, and text that the
// schema types as ST_Xstring also through encodeXstring.
import type { DateSystem, SheetState } from '../model';
import type { NormalizedColor, NormalizedStyle } from '../style';
import { escapeAttr, escapeText, encodeXstring } from '../xml/writer';
import { writeZip, type ZipWriteEntry } from '../zip/writer';
import { formatCellRef, type CellRange } from './cell-ref';
import { CONTENT_TYPES_NS, PACKAGE_REL_NS, REL_NS, SML_NS } from './namespaces';
import { BUILTIN_NUMBER_FORMATS } from './numfmt';
import type { StoredCell, StoredValue } from '../write';

/** One sheet, as handed over by `NewSheet`. */
export interface WriteSheetData {
  readonly name: string;
  readonly state: SheetState;
  readonly cells: ReadonlyMap<number, ReadonlyMap<number, StoredCell>>;
  readonly columns: ReadonlyMap<number, { readonly width: number | null; readonly hidden: boolean }>;
  readonly rows: ReadonlyMap<number, { readonly height: number | null; readonly hidden: boolean }>;
  readonly merges: readonly CellRange[];
  readonly freeze: { readonly rows: number; readonly cols: number } | null;
  readonly autoFilter: CellRange | null;
}

/** Everything needed to write a package. */
export interface WritePackageInput {
  readonly dateSystem: DateSystem;
  readonly sheets: readonly WriteSheetData[];
  /** Interned styles; index 0 is the default style. */
  readonly styles: readonly NormalizedStyle[];
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

function attr(name: string, value: string | number): string {
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

function colorXml(element: string, color: NormalizedColor): string {
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

function buildStyles(styles: readonly NormalizedStyle[]): StyleTables {
  const fonts = new Table();
  const fills = new Table();
  const borders = new Table();
  const xfs = new Table();
  const numFmts = new Map<string, number>();

  const fontXml = (font: NonNullable<NormalizedStyle['font']>): string => {
    const f = { ...DEFAULT_FONT, ...font };
    let xml = '<font>';
    if (f.bold) xml += '<b/>';
    if (f.italic) xml += '<i/>';
    if (f.strike) xml += '<strike/>';
    if (f.underline) xml += f.underline === 'single' ? '<u/>' : `<u${attr('val', f.underline)}/>`;
    xml += `<sz${attr('val', f.size)}/>`;
    if (f.color !== undefined) xml += colorXml('color', f.color);
    xml += `<name${attr('val', encodeXstring(f.name))}/>`;
    return `${xml}</font>`;
  };

  // Fills 0 and 1 are reserved: none and gray125 (D-205, W2-14).
  fonts.add(fontXml({}));
  fills.add('<fill><patternFill patternType="none"/></fill>');
  fills.add('<fill><patternFill patternType="gray125"/></fill>');
  borders.add('<border><left/><right/><top/><bottom/><diagonal/></border>');

  const numFmtId = (code: string | undefined): number => {
    if (code === undefined) return 0;
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
    if (style.fill) {
      const { pattern, fgColor, bgColor } = style.fill;
      let inner = '';
      if (fgColor !== undefined) inner += colorXml('fgColor', fgColor);
      if (bgColor !== undefined) inner += colorXml('bgColor', bgColor);
      fillId = fills.add(`<fill><patternFill${attr('patternType', pattern)}${inner === '' ? '/>' : `>${inner}</patternFill>`}</fill>`);
    }
    let borderId = 0;
    if (style.border) {
      let xml = '<border>';
      for (const edge of ['left', 'right', 'top', 'bottom'] as const) {
        const e = style.border[edge];
        xml += e === undefined ? `<${edge}/>` : `<${edge}${attr('style', e.style)}>${e.color === undefined ? '' : colorXml('color', e.color)}</${edge}>`;
      }
      borderId = borders.add(`${xml}<diagonal/></border>`);
    }
    const nf = numFmtId(style.numFmt);
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
    return xfs.add(xf);
  });

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
  xml += '</styleSheet>';
  return { xml, xfOf };
}

// ---------------------------------------------------------------------------------------------
// Shared strings

class SharedStrings {
  private readonly index = new Map<string, number>();
  readonly list: string[] = [];
  total = 0;
  add(s: string): number {
    this.total++;
    let i = this.index.get(s);
    if (i === undefined) {
      i = this.list.length;
      this.list.push(s);
      this.index.set(s, i);
    }
    return i;
  }
  xml(): string {
    let xml = `${XML_DECL}<sst xmlns="${NS}"${attr('count', this.total)}${attr('uniqueCount', this.list.length)}>`;
    for (const s of this.list) xml += `<si>${textElement(s)}</si>`;
    return `${xml}</sst>`;
  }
}

/** A `t` element; `xml:space="preserve"` when white space at the ends or inside could be lost (W2-07). */
function textElement(s: string): string {
  const preserve = /^[\s]|[\s]$|[\t\n\r]/.test(s) ? ' xml:space="preserve"' : '';
  return `<t${preserve}>${escapeText(encodeXstring(s))}</t>`;
}

// ---------------------------------------------------------------------------------------------
// Worksheets

function valueXml(v: Exclude<StoredValue, { kind: 'formula' }>, sst: SharedStrings, inFormula: boolean): { t: string; v: string } {
  switch (v.kind) {
    case 'number':
      return { t: '', v: formatNumber(v.n) };
    case 'string':
      // A formula's text result is stored in the cell as t="str"; plain text goes to the SST.
      return inFormula ? { t: 'str', v: escapeText(encodeXstring(v.s)) } : { t: 's', v: String(sst.add(v.s)) };
    case 'boolean':
      return { t: 'b', v: v.b ? '1' : '0' };
    case 'error':
      return { t: 'e', v: escapeText(v.e) };
  }
}

function cellXml(ref: string, cell: StoredCell, xfOf: readonly number[], sst: SharedStrings): string {
  const xf = xfOf[cell.style] ?? 0;
  const s = xf === 0 ? '' : attr('s', xf);
  const v = cell.value;
  if (v === null) return `<c r="${ref}"${s}/>`;
  if (v.kind === 'formula') {
    const f = `<f>${escapeText(encodeXstring(v.text))}</f>`;
    if (v.result === null) return `<c r="${ref}"${s}>${f}</c>`;
    const r = valueXml(v.result, sst, true);
    return `<c r="${ref}"${s}${r.t === '' ? '' : attr('t', r.t)}>${f}<v>${r.v}</v></c>`;
  }
  const r = valueXml(v, sst, false);
  return `<c r="${ref}"${s}${r.t === '' ? '' : attr('t', r.t)}><v>${r.v}</v></c>`;
}

function sheetXml(sheet: WriteSheetData, selected: boolean, xfOf: readonly number[], sst: SharedStrings): string {
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

  let xml = `${XML_DECL}<worksheet xmlns="${NS}" xmlns:r="${R_NS}">`;
  xml += `<dimension${attr('ref', dimension)}/>`;
  xml += `<sheetViews><sheetView${selected ? ' tabSelected="1"' : ''} workbookViewId="0"`;
  if (sheet.freeze) {
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
      xml += '/>';
    }
    xml += '</cols>';
  }

  xml += '<sheetData>';
  for (const r of sortedRows) {
    const info = sheet.rows.get(r);
    let rowXml = `<row${attr('r', r)}`;
    if (info?.height != null) rowXml += `${attr('ht', info.height)} customHeight="1"`;
    if (info?.hidden) rowXml += ' hidden="1"';
    const cells = sheet.cells.get(r);
    if (!cells || cells.size === 0) {
      xml += `${rowXml}/>`;
      continue;
    }
    xml += `${rowXml}>`;
    for (const c of [...cells.keys()].sort((a, b) => a - b)) {
      const cell = cells.get(c);
      if (cell) xml += cellXml(formatCellRef(r, c), cell, xfOf, sst);
    }
    xml += '</row>';
  }
  xml += '</sheetData>';

  if (sheet.autoFilter) xml += `<autoFilter${attr('ref', rangeRef(sheet.autoFilter))}/>`;
  if (sheet.merges.length > 0) {
    xml += `<mergeCells${attr('count', sheet.merges.length)}>`;
    for (const m of sheet.merges) xml += `<mergeCell${attr('ref', rangeRef(m))}/>`;
    xml += '</mergeCells>';
  }
  return `${xml}</worksheet>`;
}

// ---------------------------------------------------------------------------------------------
// Package

/** Builds all parts and returns the zipped package. */
export function writePackage(input: WritePackageInput): Uint8Array {
  const { xml: stylesXml, xfOf } = buildStyles(input.styles);
  const sst = new SharedStrings();
  const active = input.sheets.findIndex((s) => s.state === 'visible');
  const sheetParts = input.sheets.map((sheet, i) => sheetXml(sheet, i === active, xfOf, sst));

  let workbook = `${XML_DECL}<workbook xmlns="${NS}" xmlns:r="${R_NS}">`;
  if (input.dateSystem === '1904') workbook += '<workbookPr date1904="1"/>';
  workbook += `<bookViews><workbookView${active > 0 ? `${attr('firstSheet', active)}${attr('activeTab', active)}` : ''}/></bookViews>`;
  workbook += '<sheets>';
  input.sheets.forEach((sheet, i) => {
    workbook += `<sheet${attr('name', encodeXstring(sheet.name))}${attr('sheetId', i + 1)}`;
    if (sheet.state !== 'visible') workbook += attr('state', sheet.state);
    workbook += `${attr('r:id', `rId${String(i + 1)}`)}/>`;
  });
  workbook += '</sheets>';
  const names = input.sheets
    .map((sheet, i) =>
      sheet.autoFilter
        ? `<definedName name="_xlnm._FilterDatabase"${attr('localSheetId', i)} hidden="1">${escapeText(
            encodeXstring(`${quoteSheetName(sheet.name)}!${absoluteRangeRef(sheet.autoFilter)}`),
          )}</definedName>`
        : '',
    )
    .join('');
  if (names !== '') workbook += `<definedNames>${names}</definedNames>`;
  workbook += '</workbook>';

  const n = input.sheets.length;
  let workbookRels = `${XML_DECL}<Relationships xmlns="${PACKAGE_REL_NS}">`;
  for (let i = 1; i <= n; i++) {
    workbookRels += `<Relationship Id="rId${String(i)}" Type="${REL_BASE}worksheet" Target="worksheets/sheet${String(i)}.xml"/>`;
  }
  workbookRels += `<Relationship Id="rId${String(n + 1)}" Type="${REL_BASE}styles" Target="styles.xml"/>`;
  workbookRels += `<Relationship Id="rId${String(n + 2)}" Type="${REL_BASE}sharedStrings" Target="sharedStrings.xml"/>`;
  workbookRels += '</Relationships>';

  const rootRels =
    `${XML_DECL}<Relationships xmlns="${PACKAGE_REL_NS}">` +
    `<Relationship Id="rId1" Type="${REL_BASE}officeDocument" Target="xl/workbook.xml"/></Relationships>`;

  let types = `${XML_DECL}<Types xmlns="${CONTENT_TYPES_NS}">`;
  types += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
  types += '<Default Extension="xml" ContentType="application/xml"/>';
  types += `<Override PartName="/xl/workbook.xml" ContentType="${CT_BASE}sheet.main+xml"/>`;
  for (let i = 1; i <= n; i++) types += `<Override PartName="/xl/worksheets/sheet${String(i)}.xml" ContentType="${CT_BASE}worksheet+xml"/>`;
  types += `<Override PartName="/xl/styles.xml" ContentType="${CT_BASE}styles+xml"/>`;
  types += `<Override PartName="/xl/sharedStrings.xml" ContentType="${CT_BASE}sharedStrings+xml"/>`;
  types += '</Types>';

  const entries: ZipWriteEntry[] = [
    { name: '[Content_Types].xml', data: encoder.encode(types) },
    { name: '_rels/.rels', data: encoder.encode(rootRels) },
    { name: 'xl/workbook.xml', data: encoder.encode(workbook) },
    { name: 'xl/_rels/workbook.xml.rels', data: encoder.encode(workbookRels) },
    ...sheetParts.map((xml, i) => ({ name: `xl/worksheets/sheet${String(i + 1)}.xml`, data: encoder.encode(xml) })),
    { name: 'xl/styles.xml', data: encoder.encode(stylesXml) },
    { name: 'xl/sharedStrings.xml', data: encoder.encode(sst.xml()) },
  ];
  return writeZip(entries);
}
