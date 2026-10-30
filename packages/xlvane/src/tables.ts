// Tables ("list objects") of a worksheet: ECMA-376 Part 1 §18.5, §12.3.21, §18.3.1.94/.95 and
// [MS-OI29500] §2.1.679 … §2.1.684, §2.1.1094 (docs/facts/tables-drawings.md TD-01 … TD-12;
// docs/decisions.md D-D02). A table is a range of the sheet with a header row, data rows and an
// optional totals row; its cells hold the values, and a Table Definition part (`xl/tables/…`)
// holds the name, the columns, the autofilter and the style.
import { A1_NAME_RE, NAME_RE, R1C1_NAME_RE } from './defined-names';
import { XlvError } from './errors';
import { formatCellRef, formatRange, MAX_COLUMNS, MAX_ROWS, parseRange, type CellRange } from './sml/cell-ref';
import { SML_NS } from './sml/namespaces';
import type { PartContext } from './sml/shared-strings';
import { decodeXstring } from './sml/xstring';
import { childNamed, childrenNamed, parseXmlTree, type XmlElement } from './xml/tree';
import { encodeXstring, escapeAttr, escapeText, firstLoneSurrogate } from './xml/writer';

/**
 * How a column is summarised in the totals row (ST_TotalsRowFunction, Part 1 §18.18.83). All but
 * `none` and `custom` are written as `SUBTOTAL(n, Table[Column])` with the function numbers of
 * [MS-OI29500] §2.1.1094 (TD-06); `custom` uses `totalsRowFormula`.
 *
 * @example
 * ```ts
 * import type { TotalsRowFunction } from 'xlvane';
 *
 * const total: TotalsRowFunction = 'sum';
 * ```
 */
export type TotalsRowFunction = 'none' | 'sum' | 'min' | 'max' | 'average' | 'count' | 'countNums' | 'stdDev' | 'var' | 'custom';

/**
 * The table style and which of its parts are shown (CT_TableStyleInfo, Part 1 §18.5.1.5). `theme`
 * is one of the 60 built-in styles of Part 1 Annex G.1 (`TableStyleLight1` … `TableStyleLight21`,
 * `TableStyleMedium1` … `TableStyleMedium28`, `TableStyleDark1` … `TableStyleDark11`), or `null`
 * for no style name.
 *
 * @example
 * ```ts
 * import type { TableStyleOptions } from 'xlvane';
 *
 * const style: TableStyleOptions = { theme: 'TableStyleMedium9', showRowStripes: true };
 * ```
 */
export interface TableStyleOptions {
  /** Default `"TableStyleMedium2"`. */
  readonly theme?: string | null;
  /** Default `false`. */
  readonly showFirstColumn?: boolean;
  /** Default `false`. */
  readonly showLastColumn?: boolean;
  /** Default `false`. */
  readonly showRowStripes?: boolean;
  /** Default `false`. */
  readonly showColumnStripes?: boolean;
}

/**
 * A column of a table to be written. `name` is also the text of the column's header cell
 * ([MS-OI29500] §2.1.681 i).
 *
 * @example
 * ```ts
 * import type { TableColumnInput } from 'xlvane';
 *
 * const columns: TableColumnInput[] = [
 *   { name: 'Product', totalsRowLabel: 'Total' },
 *   { name: 'Units', totalsRowFunction: 'sum' },
 *   { name: 'Margin', calculatedColumnFormula: 'Sales[[#This Row],[Units]]*2' },
 * ];
 * ```
 */
export interface TableColumnInput {
  /** Column name, 1 to 255 characters, unique in the table ignoring case. */
  readonly name: string;
  /** Summary in the totals row. Default `"none"`. */
  readonly totalsRowFunction?: TotalsRowFunction;
  /** Text in the totals row; only with `totalsRowFunction: "none"`. At most 32767 characters. */
  readonly totalsRowLabel?: string;
  /** The formula of a `"custom"` total, without "=". */
  readonly totalsRowFormula?: string;
  /** A formula written into every data cell of the column that has no value, without "=". */
  readonly calculatedColumnFormula?: string;
  /** Show the filter button in the header cell. Default `true`. */
  readonly filterButton?: boolean;
}

/**
 * A table to be added with `NewSheet.addTable`.
 *
 * - `ref` is the top-left cell (the table then spans the columns and the given rows, at least one
 *   data row) or the whole range including the header and totals rows.
 * - `rows` are the data rows; a value `undefined` leaves the cell empty.
 *
 * @example
 * ```ts
 * import type { TableInput } from 'xlvane';
 *
 * const table: TableInput = {
 *   name: 'Sales',
 *   ref: 'B2',
 *   totalsRow: true,
 *   columns: [{ name: 'Region', totalsRowLabel: 'Total' }, { name: 'Revenue', totalsRowFunction: 'sum' }],
 *   rows: [['North', 1200], ['South', 980]],
 * };
 * ```
 */
export interface TableInput {
  /** Table name, for example `"Sales"`: letters, digits, `_`, `.`, `\`; no spaces. */
  readonly name: string;
  /** The name used in formulas and shown in applications (`displayName`); default `name`. Same rules. */
  readonly displayName?: string;
  /** Top-left cell, or the whole range of the table. */
  readonly ref: string;
  /** Write a header row. Default `true`. */
  readonly headerRow?: boolean;
  /** Write a totals row. Default `false`. */
  readonly totalsRow?: boolean;
  readonly style?: TableStyleOptions;
  readonly columns: readonly TableColumnInput[];
  /** Data rows, each an array of cell values by column. */
  readonly rows?: readonly (readonly unknown[])[];
}

/**
 * A column of a table read from a file (CT_TableColumn, Part 1 §18.5.1.3).
 *
 * @example
 * ```ts
 * import { readWorkbook, type TableDefinitionColumn } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const columns: readonly TableDefinitionColumn[] = readWorkbook(bytes).getSheet('Data').tables[0]?.columns ?? [];
 * console.log(columns.map((c) => `${c.name}: ${c.totalsRowFunction}`));
 * ```
 */
export interface TableDefinitionColumn {
  /** Column id, unique in the table. */
  readonly id: number;
  readonly name: string;
  readonly totalsRowFunction: TotalsRowFunction;
  readonly totalsRowLabel: string | null;
  readonly totalsRowFormula: string | null;
  readonly calculatedColumnFormula: string | null;
  /** `false` when the autofilter hides the column's filter button. */
  readonly filterButton: boolean;
}

/**
 * A table of a worksheet read from a file (CT_Table, Part 1 §18.5.1.2).
 *
 * @example
 * ```ts
 * import { readWorkbook, type TableDefinition } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * for (const table of readWorkbook(bytes).getSheet('Data').tables as readonly TableDefinition[]) {
 *   console.log(table.name, table.ref, table.columns.map((c) => c.name));
 * }
 * ```
 */
export interface TableDefinition {
  /** The `name` of the table; the display name when the file has none ([MS-OI29500] §2.1.680 c). */
  readonly name: string;
  /** The name used in formulas and shown in applications (`displayName`). */
  readonly displayName: string;
  /** The range of the table, header and totals rows included, for example `"B2:D10"`. */
  readonly ref: string;
  /** Header rows (0 or 1). */
  readonly headerRowCount: number;
  /** Totals rows (0 or 1). */
  readonly totalsRowCount: number;
  /** Whether the totals row has ever been shown (`totalsRowShown`). */
  readonly totalsRowShown: boolean;
  readonly tableType: 'worksheet' | 'xml' | 'queryTable';
  readonly comment: string | null;
  readonly columns: readonly TableDefinitionColumn[];
  /** The style, or `null` when the file names none. */
  readonly style: Required<TableStyleOptions> | null;
  /** The autofilter range of the table, or `null`. */
  readonly autoFilter: string | null;
  /** Part name of the Table Definition part. */
  readonly part: string;
}

/** A column checked for writing. */
export interface NormalizedTableColumn {
  readonly name: string;
  readonly totalsRowFunction: TotalsRowFunction;
  readonly totalsRowLabel: string | null;
  readonly totalsRowFormula: string | null;
  readonly calculatedColumnFormula: string | null;
  readonly filterButton: boolean;
}

/** A table checked for writing. */
export interface NormalizedTable {
  readonly name: string;
  readonly displayName: string;
  readonly range: CellRange;
  readonly headerRow: boolean;
  readonly totalsRow: boolean;
  readonly style: Required<TableStyleOptions>;
  readonly columns: readonly NormalizedTableColumn[];
  readonly rows: readonly (readonly unknown[])[];
}

/** Content type of a Table Definition part (Part 1 §12.3.21, TD-01). */
export const TABLE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';

/** SUBTOTAL function numbers by totals function ([MS-OI29500] §2.1.1094, TD-06). */
const SUBTOTAL: Readonly<Record<string, number>> = { average: 101, countNums: 102, count: 103, max: 104, min: 105, stdDev: 107, sum: 109, var: 110 };
const FUNCTIONS: readonly TotalsRowFunction[] = ['none', 'sum', 'min', 'max', 'average', 'count', 'countNums', 'stdDev', 'var', 'custom'];
/** Built-in table style names of Part 1 Annex G.1 / presetTableStyles.xml (TD-09). */
const THEME_RE = /^TableStyle(?:Light(?:[1-9]|1\d|2[01])|Medium(?:[1-9]|1\d|2[0-8])|Dark(?:[1-9]|1[01]))$/;
/** [MS-OI29500] §2.1.680 b, §2.1.681 a, j (TD-04). */
const MAX_NAME = 255;
const MAX_LABEL = 32767;
const MAX_FORMULA = 8192;
/** A sheet holds at most this many tables in the reader (an implementation limit). */
const MAX_TABLES_PER_SHEET = 4096;

function invalidTable(sheet: string, message: string, hint: string, details: Record<string, string | number> = {}): XlvError {
  return new XlvError('XLV_E_INVALID_TABLE', message, { location: { sheet }, hint, details });
}

/** Checks a table or display name (P1 §18.5.1.2, §18.17.2.5; [MS-OI29500] §2.1.680 b; TD-03). */
function checkTableName(value: unknown, what: string, sheet: string): string {
  const hint = 'Use letters, digits, "_" and ".", starting with a letter or "_", for example "Sales_2026"; no spaces.';
  if (typeof value !== 'string' || value.length < 1 || value.length > MAX_NAME) {
    throw invalidTable(sheet, `The table ${what} must be text of 1 to ${String(MAX_NAME)} characters.`, hint, { name: String(value).slice(0, 40) });
  }
  const upper = value.toUpperCase();
  if (!NAME_RE.test(value) || firstLoneSurrogate(value) !== -1 || upper === 'TRUE' || upper === 'FALSE' || A1_NAME_RE.test(value) || R1C1_NAME_RE.test(value)) {
    throw invalidTable(sheet, `"${value.slice(0, 40)}" is not a valid table ${what}.`, `${hint} It must not read as a cell reference such as "A1" or "R1C1".`, { name: value.slice(0, 40) });
  }
  return value;
}

function text(value: unknown, sheet: string, what: string, max: number): string {
  if (typeof value !== 'string' || value.length > max || firstLoneSurrogate(value) !== -1) {
    throw invalidTable(sheet, `The ${what} must be text of at most ${String(max)} characters.`, `Shorten the ${what}.`);
  }
  return value;
}

function formula(value: unknown, sheet: string, what: string): string {
  const f = typeof value === 'string' && value.startsWith('=') ? value.slice(1) : value;
  if (typeof f !== 'string' || f.trim() === '' || f.length > MAX_FORMULA || firstLoneSurrogate(f) !== -1) {
    throw invalidTable(sheet, `The ${what} must be a formula of at most ${String(MAX_FORMULA)} characters.`, 'Pass the formula text with or without "=", for example "SUM(Sales[Units])".');
  }
  return f;
}

function bool(o: Record<string, unknown>, key: string, def: boolean, sheet: string): boolean {
  const v = o[key];
  if (v === undefined) return def;
  if (typeof v !== 'boolean') throw invalidTable(sheet, `The table option "${key}" must be true or false.`, `Use ${key}: true or leave it out.`, { option: key });
  return v;
}

function known(o: object, keys: readonly string[], what: string, sheet: string): Record<string, unknown> {
  if (typeof o !== 'object' || o === null || Array.isArray(o)) throw invalidTable(sheet, `The ${what} must be an object.`, `For example { ${keys[0] ?? ''}: … }.`);
  for (const key of Object.keys(o)) {
    if (!keys.includes(key)) throw invalidTable(sheet, `Unknown ${what} option "${key.slice(0, 40)}".`, `Known options: ${keys.join(', ')}.`, { option: key.slice(0, 40) });
  }
  return o as Record<string, unknown>;
}

/**
 * Checks a table given to `NewSheet.addTable` (or the streaming writer, `streaming = true`: then
 * `ref` must be the whole range and `rows` must not be given). Names, columns and the range are
 * checked here; uniqueness and overlaps are checked by the caller.
 */
export function normalizeTable(input: unknown, sheet: string, streaming = false): NormalizedTable {
  const o = known(input as object, ['name', 'displayName', 'ref', 'headerRow', 'totalsRow', 'style', 'columns', 'rows'], 'table', sheet);
  const name = checkTableName(o['name'], 'name', sheet);
  const displayName = o['displayName'] === undefined ? name : checkTableName(o['displayName'], 'display name', sheet);
  const headerRow = bool(o, 'headerRow', true, sheet);
  const totalsRow = bool(o, 'totalsRow', false, sheet);

  const s = o['style'] === undefined ? {} : known(o['style'] as object, ['theme', 'showFirstColumn', 'showLastColumn', 'showRowStripes', 'showColumnStripes'], 'table style', sheet);
  const theme = s['theme'] === undefined ? 'TableStyleMedium2' : s['theme'];
  if (theme !== null && (typeof theme !== 'string' || !THEME_RE.test(theme))) {
    throw invalidTable(sheet, `"${String(theme).slice(0, 40)}" is not a built-in table style.`, 'Use TableStyleLight1 … TableStyleLight21, TableStyleMedium1 … TableStyleMedium28, TableStyleDark1 … TableStyleDark11, or null.', { option: 'theme' });
  }
  const style = {
    theme: theme as string | null,
    showFirstColumn: bool(s, 'showFirstColumn', false, sheet),
    showLastColumn: bool(s, 'showLastColumn', false, sheet),
    showRowStripes: bool(s, 'showRowStripes', false, sheet),
    showColumnStripes: bool(s, 'showColumnStripes', false, sheet),
  };

  const rawColumns = o['columns'];
  if (!Array.isArray(rawColumns) || rawColumns.length === 0 || rawColumns.length > MAX_COLUMNS) {
    throw invalidTable(sheet, `The table "${name}" needs a list of 1 to ${String(MAX_COLUMNS)} columns.`, "For example columns: [{ name: 'Region' }, { name: 'Revenue' }].", { name });
  }
  const seen = new Set<string>();
  const columns = rawColumns.map((raw: unknown, i) => {
    const c = known(raw as object, ['name', 'totalsRowFunction', 'totalsRowLabel', 'totalsRowFormula', 'calculatedColumnFormula', 'filterButton'], `column ${String(i + 1)} of table "${name}"`, sheet);
    const columnName = c['name'];
    if (typeof columnName !== 'string' || columnName.length < 1 || columnName.length > MAX_NAME || columnName.trim() === '' || firstLoneSurrogate(columnName) !== -1 || /[\r\n]/.test(columnName)) {
      throw invalidTable(sheet, `Column ${String(i + 1)} of table "${name}" needs a name of 1 to ${String(MAX_NAME)} characters on one line.`, 'Give every column a short, unique name, for example "Revenue".', { name, column: i + 1 });
    }
    const key = columnName.toLowerCase();
    if (seen.has(key)) throw invalidTable(sheet, `Table "${name}" has two columns named "${columnName.slice(0, 40)}".`, 'Column names must be unique in a table (case is ignored), for example "Revenue" and "Revenue_2".', { name, column: i + 1 });
    seen.add(key);
    const fn = c['totalsRowFunction'] ?? 'none';
    if (typeof fn !== 'string' || !FUNCTIONS.includes(fn as TotalsRowFunction)) {
      throw invalidTable(sheet, `"${String(fn).slice(0, 20)}" is not a totals function.`, `Use one of ${FUNCTIONS.join(', ')}.`, { name, column: i + 1 });
    }
    const label = c['totalsRowLabel'] === undefined ? null : text(c['totalsRowLabel'], sheet, `totals label of column "${columnName}"`, MAX_LABEL);
    if (label !== null && fn !== 'none') {
      throw invalidTable(sheet, `Column "${columnName}" of table "${name}" has both a totals label and the function "${fn}".`, 'A totals cell shows either a label (with totalsRowFunction "none") or a function, not both (ECMA-376 Part 1 §18.5.1.3).', { name, column: i + 1 });
    }
    const totalsFormula = c['totalsRowFormula'] === undefined ? null : formula(c['totalsRowFormula'], sheet, `totals formula of column "${columnName}"`);
    if ((fn === 'custom') !== (totalsFormula !== null)) {
      throw invalidTable(sheet, `Column "${columnName}" of table "${name}": totalsRowFormula goes with totalsRowFunction "custom" and only with it.`, "For example { totalsRowFunction: 'custom', totalsRowFormula: 'SUM(Sales[Units])*2' }.", { name, column: i + 1 });
    }
    const calculated = c['calculatedColumnFormula'] === undefined ? null : formula(c['calculatedColumnFormula'], sheet, `calculated formula of column "${columnName}"`);
    const filterButton = bool(c, 'filterButton', true, sheet);
    return { name: columnName, totalsRowFunction: fn as TotalsRowFunction, totalsRowLabel: label, totalsRowFormula: totalsFormula, calculatedColumnFormula: calculated, filterButton };
  });

  const rawRows = o['rows'] ?? [];
  if (!Array.isArray(rawRows) || rawRows.some((r) => !Array.isArray(r))) {
    throw invalidTable(sheet, `The rows of table "${name}" must be an array of arrays.`, "For example rows: [['North', 1200], ['South', 980]].", { name });
  }
  if (streaming && rawRows.length > 0) {
    throw invalidTable(sheet, `The streaming writer takes no rows for table "${name}".`, 'Stream the header, data and totals rows as ordinary rows of the sheet and give the whole range in ref.', { name });
  }
  for (const [i, r] of (rawRows as unknown[][]).entries()) {
    if (r.length > columns.length) throw invalidTable(sheet, `Row ${String(i + 1)} of table "${name}" has ${String(r.length)} values for ${String(columns.length)} columns.`, 'Give at most one value per column.', { name, row: i + 1 });
  }
  const rows = rawRows as readonly (readonly unknown[])[];

  const rawRef = o['ref'];
  const given = typeof rawRef === 'string' ? parseRange(rawRef) : null;
  if (given === null) throw invalidTable(sheet, `The ref of table "${name}" must be a cell such as "B2" or a range such as "B2:D10".`, 'Give the top-left cell, or the whole range including the header and totals rows.', { name });
  const single = !String(rawRef).includes(':');
  if (streaming && single) throw invalidTable(sheet, `The streaming writer needs the whole range of table "${name}", not only its top-left cell.`, 'For example ref: "A1:C101" for a header row and 100 data rows.', { name });
  const extra = (headerRow ? 1 : 0) + (totalsRow ? 1 : 0);
  let range: CellRange;
  if (single) {
    const height = extra + Math.max(rows.length, 1);
    range = { start: given.start, end: { row: given.start.row + height - 1, col: given.start.col + columns.length - 1 } };
    if (range.end.row > MAX_ROWS || range.end.col > MAX_COLUMNS) {
      throw invalidTable(sheet, `Table "${name}" at ${String(rawRef)} would reach beyond the last row or column of the sheet.`, 'Start the table higher or further left, or use fewer rows or columns.', { name });
    }
  } else {
    range = given;
    const width = range.end.col - range.start.col + 1;
    if (width !== columns.length) {
      throw invalidTable(sheet, `The range ${formatRange(range)} of table "${name}" is ${String(width)} columns wide, but ${String(columns.length)} columns are given.`, 'Give one column for each column of the range, or only the top-left cell.', { name });
    }
    const dataRows = range.end.row - range.start.row + 1 - extra;
    if (dataRows < 1) throw invalidTable(sheet, `The range ${formatRange(range)} of table "${name}" leaves no data row.`, 'A table needs at least one data row between its header and totals rows.', { name });
    if (rows.length > dataRows) throw invalidTable(sheet, `Table "${name}" has ${String(rows.length)} rows, but its range ${formatRange(range)} holds ${String(dataRows)}.`, 'Make the range larger, or give only the top-left cell.', { name });
  }
  return { name, displayName, range, headerRow, totalsRow, style, columns, rows };
}

/** True when two ranges share a cell. */
export function rangesOverlap(a: CellRange, b: CellRange): boolean {
  return a.start.row <= b.end.row && b.start.row <= a.end.row && a.start.col <= b.end.col && b.start.col <= a.end.col;
}

/**
 * A structured reference to a whole column of a table, as Excel writes it in totals rows (TD-07):
 * `Sales[Units]`; names with a leading or trailing space or with `'`, `[`, `]`, `#` get double
 * brackets, and those four characters are escaped with `'` ([MS-OI29500] §3.2.3.1.1).
 */
export function columnReference(table: string, column: string): string {
  const escaped = column.replace(/['[\]#]/g, "'$&");
  return escaped !== column || column.trim() !== column ? `${table}[[${escaped}]]` : `${table}[${column}]`;
}

/** The formula of a column's totals cell, or `null` when it has none. */
export function totalsFormula(table: string, column: NormalizedTableColumn): string | null {
  if (column.totalsRowFunction === 'custom') return column.totalsRowFormula;
  const n = SUBTOTAL[column.totalsRowFunction];
  return n === undefined ? null : `SUBTOTAL(${String(n)},${columnReference(table, column.name)})`;
}

const a = (name: string, value: string | number): string => ` ${name}="${typeof value === 'number' ? String(value) : escapeAttr(value)}"`;

/** The Table Definition part of a table (CT_Table; element order of sml.xsd T3946; TD-02). */
export function tableXml(t: NormalizedTable, id: number): string {
  // Always a range, also for one cell (ST_Ref of the Part 1 examples).
  const ref = `${formatCellRef(t.range.start.row, t.range.start.col)}:${formatCellRef(t.range.end.row, t.range.end.col)}`;
  let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="${SML_NS.transitional}"${a('id', id)}${a('name', encodeXstring(t.name))}${a('displayName', encodeXstring(t.displayName))}${a('ref', ref)}`;
  if (!t.headerRow) xml += a('headerRowCount', 0);
  // totalsRowShown is "whether the totals row has ever been shown"; Excel writes 0 for tables
  // that never had one, and true is required with a totals row ([MS-OI29500] §2.1.680 f, TD-05).
  xml += t.totalsRow ? a('totalsRowCount', 1) : a('totalsRowShown', 0);
  xml += '>';
  if (t.headerRow) {
    // The autofilter covers the header and data rows, not the totals row, as in the tables Excel
    // wrote in the POI corpus (TD-08); Part 1 has no rule.
    const filterRange: CellRange = { start: t.range.start, end: { row: t.range.end.row - (t.totalsRow ? 1 : 0), col: t.range.end.col } };
    const hidden = t.columns.map((c, i) => (c.filterButton ? '' : `<filterColumn${a('colId', i)} hiddenButton="1"/>`)).join('');
    xml += `<autoFilter${a('ref', formatRange(filterRange))}${hidden === '' ? '/>' : `>${hidden}</autoFilter>`}`;
  }
  xml += `<tableColumns${a('count', t.columns.length)}>`;
  t.columns.forEach((c, i) => {
    xml += `<tableColumn${a('id', i + 1)}${a('name', encodeXstring(c.name))}`;
    if (t.totalsRow && c.totalsRowFunction !== 'none') xml += a('totalsRowFunction', c.totalsRowFunction);
    if (t.totalsRow && c.totalsRowLabel !== null) xml += a('totalsRowLabel', encodeXstring(c.totalsRowLabel));
    let children = '';
    if (c.calculatedColumnFormula !== null) children += `<calculatedColumnFormula>${escapeText(encodeXstring(c.calculatedColumnFormula))}</calculatedColumnFormula>`;
    if (t.totalsRow && c.totalsRowFormula !== null) children += `<totalsRowFormula>${escapeText(encodeXstring(c.totalsRowFormula))}</totalsRowFormula>`;
    xml += children === '' ? '/>' : `>${children}</tableColumn>`;
  });
  xml += '</tableColumns>';
  const s = t.style;
  xml += `<tableStyleInfo${s.theme === null ? '' : a('name', s.theme)}${a('showFirstColumn', s.showFirstColumn ? 1 : 0)}${a('showLastColumn', s.showLastColumn ? 1 : 0)}${a('showRowStripes', s.showRowStripes ? 1 : 0)}${a('showColumnStripes', s.showColumnStripes ? 1 : 0)}/>`;
  return `${xml}</table>`;
}

/** Part name of the n-th table of a workbook (1-based), as Excel names them. */
export function tablePartName(n: number): string {
  return `xl/tables/table${String(n)}.xml`;
}

function boolAttr(el: XmlElement, name: string, def: boolean): boolean {
  const v = el.attrs.get(name)?.trim();
  return v === undefined ? def : v === '1' || v === 'true';
}

function uintAttr(el: XmlElement, name: string, def: number): number {
  const v = el.attrs.get(name)?.trim();
  return v !== undefined && /^\d{1,10}$/.test(v) ? Number(v) : def;
}

/**
 * Reads a Table Definition part. Returns `null` with a warning when the part is not a table or
 * its range is not valid; unknown attributes and elements are ignored.
 */
export function parseTable(xml: string, ctx: PartContext & { readonly sheet: string }): TableDefinition | null {
  const warn = (message: string): void => ctx.warnings.add('XLV_W_INVALID_VALUE', message, { part: ctx.part, sheet: ctx.sheet });
  let root: XmlElement;
  try {
    root = parseXmlTree(xml, ctx);
  } catch (error) {
    if (error instanceof XlvError && error.code === 'XLV_E_XML_MALFORMED') {
      warn('A table part of the sheet is not well-formed XML; the table is ignored.');
      return null;
    }
    throw error;
  }
  const ns = root.ns;
  if (root.local !== 'table' || (ns !== SML_NS.strict && ns !== SML_NS.transitional)) {
    warn(`A table part of the sheet has the root element <${root.local}>; it is ignored.`);
    return null;
  }
  const range = parseRange((root.attrs.get('ref') ?? '').trim().toUpperCase());
  if (range === null) {
    warn(`A table of the sheet has the invalid range "${(root.attrs.get('ref') ?? '').slice(0, 40)}"; it is ignored.`);
    return null;
  }
  const displayName = decodeXstring(root.attrs.get('displayName') ?? '');
  const objectName = decodeXstring(root.attrs.get('name') ?? '');
  const name = objectName !== '' ? objectName : displayName;
  const autoFilterEl = childNamed(root, ns, 'autoFilter');
  const filterRef = autoFilterEl?.attrs.get('ref')?.trim().toUpperCase();
  const filterRange = filterRef === undefined ? null : parseRange(filterRef);
  const hidden = new Set<number>();
  if (autoFilterEl) for (const fc of childrenNamed(autoFilterEl, ns, 'filterColumn')) if (boolAttr(fc, 'hiddenButton', false) || !boolAttr(fc, 'showButton', true)) hidden.add(uintAttr(fc, 'colId', -1));
  const columnsEl = childNamed(root, ns, 'tableColumns');
  const columns = (columnsEl === undefined ? [] : childrenNamed(columnsEl, ns, 'tableColumn')).slice(0, MAX_COLUMNS).map((c, i): TableDefinitionColumn => {
    const fn = c.attrs.get('totalsRowFunction')?.trim() ?? 'none';
    const label = c.attrs.get('totalsRowLabel');
    const calc = childNamed(c, ns, 'calculatedColumnFormula');
    const total = childNamed(c, ns, 'totalsRowFormula');
    return Object.freeze({
      id: uintAttr(c, 'id', i + 1),
      name: decodeXstring(c.attrs.get('name') ?? ''),
      totalsRowFunction: FUNCTIONS.includes(fn as TotalsRowFunction) ? (fn as TotalsRowFunction) : 'none',
      totalsRowLabel: label === undefined ? null : decodeXstring(label),
      totalsRowFormula: total === undefined ? null : decodeXstring(total.text),
      calculatedColumnFormula: calc === undefined ? null : decodeXstring(calc.text),
      filterButton: !hidden.has(i),
    });
  });
  const styleEl = childNamed(root, ns, 'tableStyleInfo');
  const type = root.attrs.get('tableType')?.trim();
  const comment = root.attrs.get('comment');
  return Object.freeze({
    name,
    displayName: displayName !== '' ? displayName : name,
    ref: formatRange(range),
    headerRowCount: Math.min(uintAttr(root, 'headerRowCount', 1), 1),
    totalsRowCount: Math.min(uintAttr(root, 'totalsRowCount', 0), 1),
    totalsRowShown: boolAttr(root, 'totalsRowShown', true),
    tableType: type === 'xml' || type === 'queryTable' ? type : 'worksheet',
    comment: comment === undefined ? null : decodeXstring(comment),
    columns,
    style:
      styleEl === undefined
        ? null
        : Object.freeze({
            theme: styleEl.attrs.has('name') ? decodeXstring(styleEl.attrs.get('name') ?? '') : null,
            showFirstColumn: boolAttr(styleEl, 'showFirstColumn', false),
            showLastColumn: boolAttr(styleEl, 'showLastColumn', false),
            showRowStripes: boolAttr(styleEl, 'showRowStripes', false),
            showColumnStripes: boolAttr(styleEl, 'showColumnStripes', false),
          }),
    autoFilter: filterRange === null ? null : formatRange(filterRange),
    part: ctx.part,
  });
}

export { MAX_TABLES_PER_SHEET };
