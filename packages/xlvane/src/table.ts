// Task-oriented table functions: a sheet as a list of row objects and back (docs/decisions.md
// D-303 … D-305). Built on the read model (read.ts) and the writer (write.ts).
import { XlvError } from './errors';
import { detectHeader } from './inspect';
import type { ReadLimits } from './limits';
import type { Cell, Sheet, Workbook } from './model';
import { columnLetters, formatCellRef, MAX_COLUMNS, MAX_ROWS, parseRange, type CellRange } from './sml/cell-ref';
import type { StandardSchemaV1 } from './standard-schema';
import type { CellStyle } from './style';
import { booleanOption, oneOfOption, openInput, readOptions } from './util/options';
import { createWorkbook, writeWorkbook, type NewWorkbook } from './write';
import type { DateSystem } from './model';

/**
 * A value in a table row read by {@link readTable}. Error cells give their code, e.g. `"#N/A"`.
 *
 * @example
 * ```ts
 * import type { TableValue } from 'xlvane';
 *
 * const show = (value: TableValue): string => (value instanceof Date ? value.toISOString().slice(0, 10) : String(value ?? ''));
 * ```
 */
export type TableValue = string | number | boolean | Date | null;

/**
 * A row read by {@link readTable} without a schema: header text → value.
 *
 * @example
 * ```ts
 * import { readTable, type TableRow } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const rows: TableRow[] = readTable(bytes).rows;
 * console.log(rows[0]?.['Customer']);
 * ```
 */
export type TableRow = Record<string, TableValue>;

/**
 * Options for {@link readTable} and {@link readTableAsync}.
 *
 * @example
 * ```ts
 * import type { ReadTableOptions } from 'xlvane';
 *
 * const options: ReadTableOptions = { sheet: 'Sales', header: 'auto', dates: 'iso' };
 * ```
 */
export interface ReadTableOptions<S extends StandardSchemaV1 | undefined = undefined> {
  /** Sheet name (exact). Default: the first visible worksheet. */
  readonly sheet?: string;
  /**
   * Where the header row is: `"auto"` (default) detects it like `inspect()`; a number is the
   * 1-based row; `false` means no header, and keys are the column letters.
   */
  readonly header?: 'auto' | number | false;
  /** Only read this range, for example `"A3:F200"`. Default: the used range of the sheet. */
  readonly range?: string;
  /** A Standard Schema (Zod, Valibot, ArkType, …) that every row is validated with. */
  readonly schema?: S;
  /** `"date"` (default) returns `Date` objects; `"iso"` returns ISO 8601 strings. */
  readonly dates?: 'date' | 'iso';
  /** Leave out rows without any value. Default `true`. */
  readonly skipEmptyRows?: boolean;
  /** Read limits for byte input; see `ReadLimits`. */
  readonly limits?: Partial<ReadLimits>;
}

/**
 * A column of a table.
 *
 * @example
 * ```ts
 * import { readTable } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * for (const column of readTable(bytes).columns) console.log(column.letter, column.key);
 * ```
 */
export interface TableColumn {
  /** Property name in the row objects: the trimmed header text, made unique with `_2`, `_3`, …. */
  readonly key: string;
  /** Header text as in the sheet, or `null` without a header or for an empty header cell. */
  readonly header: string | null;
  /** 1-based column number. */
  readonly col: number;
  /** Column letters, for example `"C"`. */
  readonly letter: string;
}

/**
 * A validation problem found by the schema in one row.
 *
 * @example
 * ```ts
 * import { readTable, type TableIssue } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const describe = (issue: TableIssue): string => `${issue.cell ?? `row ${String(issue.row)}`}: ${issue.message}`;
 * console.log(readTable(bytes).issues.map(describe));
 * ```
 */
export interface TableIssue {
  /** 1-based sheet row of the invalid row. */
  readonly row: number;
  /** Cell address when the issue belongs to one column, for example `"B5"`; otherwise `null`. */
  readonly cell: string | null;
  /** Column key the issue belongs to, or `null`. */
  readonly column: string | null;
  /** The schema library's message. */
  readonly message: string;
  /** The issue path from the schema, with keys as strings or numbers. */
  readonly path: readonly (string | number)[];
}

/**
 * The result of {@link readTable}.
 *
 * @example
 * ```ts
 * import { readTable, type TableResult, type TableRow } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const table: TableResult<TableRow> = readTable(bytes);
 * console.log(table.sheet, table.headerRow, table.rows.length, table.issues.length);
 * ```
 */
export interface TableResult<Row> {
  /** Name of the sheet that was read. */
  readonly sheet: string;
  /** Valid rows (all rows when there is no schema). */
  readonly rows: Row[];
  /** 1-based sheet row of each entry of `rows`. */
  readonly rowNumbers: number[];
  readonly columns: TableColumn[];
  /** 1-based header row, or `null`. */
  readonly headerRow: number | null;
  /** The range read, from the header row to the last data row, or `null` for an empty sheet. */
  readonly range: string | null;
  /** Schema validation problems, with cell addresses. Empty without a schema. */
  readonly issues: TableIssue[];
}

const READ_OPTIONS = ['sheet', 'header', 'range', 'schema', 'dates', 'skipEmptyRows', 'limits'];
const HEADER_SCAN_ROWS = 20;

interface PreparedTable {
  readonly sheet: string;
  readonly rows: Record<string, TableValue>[];
  readonly rowNumbers: number[];
  readonly columns: TableColumn[];
  readonly headerRow: number | null;
  readonly range: string | null;
  readonly schema: StandardSchemaV1 | undefined;
}

/** The sheet to read: by name, or the first visible worksheet (or the first sheet). */
export function pickSheet(workbook: Workbook, name: unknown): Sheet {
  if (name !== undefined) {
    if (typeof name !== 'string') {
      throw new XlvError('XLV_E_INVALID_OPTION', 'The option "sheet" must be a sheet name.', {
        hint: `Available sheets: ${workbook.sheets.map((s) => s.name).join(', ')}.`,
        details: { option: 'sheet' },
      });
    }
    return workbook.getSheet(name);
  }
  const sheet = workbook.sheets.find((s) => s.kind === 'worksheet' && s.state === 'visible') ?? workbook.sheets[0];
  if (sheet === undefined) {
    throw new XlvError('XLV_E_SHEET_NOT_FOUND', 'The workbook has no sheets.', { hint: 'Pass a workbook that contains at least one worksheet.' });
  }
  return sheet;
}

function checkSchema(schema: unknown): StandardSchemaV1 | undefined {
  if (schema === undefined) return undefined;
  const props = typeof schema === 'object' || typeof schema === 'function' ? (schema as { '~standard'?: unknown } | null)?.['~standard'] : undefined;
  if (
    typeof props !== 'object' ||
    props === null ||
    (props as { version?: unknown }).version !== 1 ||
    typeof (props as { validate?: unknown }).validate !== 'function'
  ) {
    throw new XlvError('XLV_E_SCHEMA_INVALID', 'The option "schema" is not a Standard Schema (version 1).', {
      hint: 'Pass a schema from a library that implements Standard Schema, such as Zod 3.24+, Valibot 1.0+ or ArkType 2.0+ (for example z.object({ ... })).',
      details: { option: 'schema' },
    });
  }
  return schema as StandardSchemaV1;
}

/** Sets an own enumerable data property, also for keys such as `__proto__`. */
function setOwn(target: Record<string, unknown>, key: string, value: unknown): void {
  Object.defineProperty(target, key, { value, enumerable: true, writable: true, configurable: true });
}

function cellValue(cell: Cell | undefined, dates: 'date' | 'iso'): TableValue {
  if (cell === undefined) return null;
  const v = cell.value;
  if (v instanceof Date) return dates === 'iso' ? v.toISOString() : v;
  return v;
}

function headerText(cell: Cell | undefined): string | null {
  const v = cell?.value;
  if (v === undefined || v === null) return null;
  if (v instanceof Date) return v.toISOString();
  return String(v);
}

function prepare(input: unknown, options: unknown): PreparedTable {
  const o = readOptions(options, READ_OPTIONS, 'readTable options');
  const workbook = openInput(input, o.get('limits'));
  const sheet = pickSheet(workbook, o.get('sheet'));
  const schema = checkSchema(o.get('schema'));
  const dates = oneOfOption('dates', o.get('dates'), ['date', 'iso'], 'date');
  const skipEmpty = booleanOption('skipEmptyRows', o.get('skipEmptyRows'), true);
  const headerOption = o.get('header') ?? 'auto';
  if (headerOption !== 'auto' && headerOption !== false && !(typeof headerOption === 'number' && Number.isInteger(headerOption) && headerOption >= 1 && headerOption <= MAX_ROWS)) {
    throw new XlvError('XLV_E_INVALID_OPTION', 'The option "header" must be "auto", a row number or false.', {
      hint: 'Use header: "auto" to detect the header row, header: 3 for row 3, or header: false for no header.',
      details: { option: 'header' },
    });
  }

  // Cells by row, limited to the requested range.
  let range: CellRange | null = null;
  const rawRange = o.get('range');
  if (rawRange !== undefined) {
    range = typeof rawRange === 'string' ? parseRange(rawRange) : null;
    if (range === null) {
      throw new XlvError('XLV_E_INVALID_CELL_REF', `The range ${JSON.stringify(String(rawRange).slice(0, 40))} is not valid.`, {
        location: { sheet: sheet.name },
        hint: 'Pass an A1-style range such as "A1:D100".',
        details: { ref: String(rawRange).slice(0, 40) },
      });
    }
  }
  const byRow = new Map<number, Map<number, Cell>>();
  let minRow = Infinity;
  let maxRow = 0;
  let minCol = Infinity;
  let maxCol = 0;
  for (const cell of sheet.cells()) {
    if (range && (cell.row < range.start.row || cell.row > range.end.row || cell.col < range.start.col || cell.col > range.end.col)) continue;
    let row = byRow.get(cell.row);
    if (!row) {
      row = new Map();
      byRow.set(cell.row, row);
    }
    row.set(cell.col, cell);
    if (cell.row < minRow) minRow = cell.row;
    if (cell.row > maxRow) maxRow = cell.row;
    if (cell.col < minCol) minCol = cell.col;
    if (cell.col > maxCol) maxCol = cell.col;
  }
  if (range) {
    minRow = range.start.row;
    maxRow = range.end.row;
    minCol = range.start.col;
    maxCol = range.end.col;
  }

  if (maxRow === 0 || maxCol === 0) {
    if (typeof headerOption === 'number') throw headerNotFound(sheet.name, headerOption);
    return { sheet: sheet.name, rows: [], rowNumbers: [], columns: [], headerRow: null, range: null, schema };
  }

  let headerRow: number | null = null;
  if (headerOption === 'auto') {
    const ordered = [...byRow.keys()].sort((a, b) => a - b).map((r) => [...(byRow.get(r)?.values() ?? [])].sort((a, b) => a.col - b.col));
    headerRow = detectHeader(ordered, maxCol - minCol + 1, HEADER_SCAN_ROWS)?.row ?? null;
  } else if (typeof headerOption === 'number') {
    if (headerOption < minRow || headerOption > maxRow || (byRow.get(headerOption)?.size ?? 0) === 0) throw headerNotFound(sheet.name, headerOption);
    headerRow = headerOption;
  }

  const columns: TableColumn[] = [];
  const used = new Set<string>();
  const headerCells = headerRow === null ? undefined : byRow.get(headerRow);
  for (let col = minCol; col <= maxCol; col++) {
    const letter = columnLetters(col);
    const header = headerRow === null ? null : headerText(headerCells?.get(col));
    const base = header === null || header.trim() === '' ? letter : header.trim();
    let key = base;
    for (let n = 2; used.has(key); n++) key = `${base}_${String(n)}`;
    used.add(key);
    columns.push({ key, header: header === null || header.trim() === '' ? null : header, col, letter });
  }

  const rows: Record<string, TableValue>[] = [];
  const rowNumbers: number[] = [];
  const first = headerRow === null ? minRow : headerRow + 1;
  const numbers = skipEmpty ? [...byRow.keys()].filter((r) => r >= first && r <= maxRow).sort((a, b) => a - b) : rangeOf(first, maxRow);
  for (const r of numbers) {
    const cells = byRow.get(r);
    if (skipEmpty && (cells === undefined || [...cells.values()].every((c) => c.value === null))) continue;
    const row: Record<string, TableValue> = {};
    for (const column of columns) setOwn(row, column.key, cellValue(cells?.get(column.col), dates));
    rows.push(row);
    rowNumbers.push(r);
  }
  const lastRow = rowNumbers.length > 0 ? Math.max(rowNumbers[rowNumbers.length - 1] ?? 0, headerRow ?? 0) : (headerRow ?? maxRow);
  const startRow = headerRow ?? (rowNumbers[0] ?? minRow);
  return {
    sheet: sheet.name,
    rows,
    rowNumbers,
    columns,
    headerRow,
    range: `${formatCellRef(startRow, minCol)}:${formatCellRef(Math.max(lastRow, startRow), maxCol)}`,
    schema,
  };
}

function rangeOf(from: number, to: number): number[] {
  const out: number[] = [];
  for (let r = from; r <= to; r++) out.push(r);
  return out;
}

function headerNotFound(sheet: string, row: number): XlvError {
  return new XlvError('XLV_E_HEADER_NOT_FOUND', `Row ${String(row)} of sheet "${sheet}" has no cells to use as a header.`, {
    location: { sheet },
    hint: 'Pass the row that holds the column names, header: "auto" to detect it, or header: false to use column letters as keys.',
    details: { row },
  });
}

function isThenable(value: unknown): value is PromiseLike<unknown> {
  return typeof value === 'object' && value !== null && typeof (value as { then?: unknown }).then === 'function';
}

/** Converts a schema issue to a {@link TableIssue}, mapping the first path key to a column. */
function toIssue(issue: StandardSchemaV1.Issue, row: number, columns: readonly TableColumn[]): TableIssue {
  const path = (issue.path ?? []).map((segment) => {
    const key = typeof segment === 'object' && segment !== null ? segment.key : segment;
    return typeof key === 'number' ? key : String(key);
  });
  const first = path[0];
  const column = typeof first === 'string' ? columns.find((c) => c.key === first) : undefined;
  return {
    row,
    cell: column === undefined ? null : formatCellRef(row, column.col),
    column: column === undefined ? null : column.key,
    message: typeof issue.message === 'string' ? issue.message : String(issue.message),
    path,
  };
}

function collect(
  prepared: PreparedTable,
  results: readonly StandardSchemaV1.Result<unknown>[],
): { rows: unknown[]; rowNumbers: number[]; issues: TableIssue[] } {
  const rows: unknown[] = [];
  const rowNumbers: number[] = [];
  const issues: TableIssue[] = [];
  results.forEach((result, i) => {
    const row = prepared.rowNumbers[i] ?? 0;
    if (result.issues) {
      for (const issue of result.issues) issues.push(toIssue(issue, row, prepared.columns));
    } else {
      rows.push(result.value);
      rowNumbers.push(row);
    }
  });
  return { rows, rowNumbers, issues };
}

/**
 * Reads a sheet as a list of row objects keyed by the header text. The header row is detected
 * automatically unless `header` says otherwise. With `schema` (any Standard Schema: Zod,
 * Valibot, ArkType, …), each row is validated: `rows` holds the valid rows with the schema's
 * output type and `issues` lists the problems with their cell addresses. Throws
 * `XLV_E_ASYNC_SCHEMA` for schemas that validate asynchronously; use {@link readTableAsync}.
 *
 * @example
 * ```ts
 * import { readTable } from 'xlvane';
 * import { z } from 'zod';
 *
 * declare const bytes: Uint8Array;
 * const { rows, issues } = readTable(bytes, {
 *   sheet: 'Orders',
 *   schema: z.object({ Customer: z.string(), Total: z.number() }),
 * });
 * for (const row of rows) console.log(row.Customer, row.Total.toFixed(2));
 * for (const issue of issues) console.log(`${issue.cell ?? `row ${String(issue.row)}`}: ${issue.message}`);
 * ```
 */
export function readTable<S extends StandardSchemaV1>(
  input: Uint8Array | ArrayBuffer | Workbook,
  options: ReadTableOptions<S> & { readonly schema: S },
): TableResult<StandardSchemaV1.InferOutput<S>>;
export function readTable(input: Uint8Array | ArrayBuffer | Workbook, options?: ReadTableOptions): TableResult<TableRow>;
export function readTable(input: Uint8Array | ArrayBuffer | Workbook, options?: ReadTableOptions<StandardSchemaV1 | undefined>): TableResult<unknown> {
  const prepared = prepare(input, options);
  const base = { sheet: prepared.sheet, columns: prepared.columns, headerRow: prepared.headerRow, range: prepared.range };
  const schema = prepared.schema;
  if (schema === undefined) return { ...base, rows: prepared.rows, rowNumbers: prepared.rowNumbers, issues: [] };
  const results = prepared.rows.map((row) => {
    const result = schema['~standard'].validate(row);
    if (isThenable(result)) {
      // Avoid an unhandled rejection from the promise we are not going to await.
      void Promise.resolve(result).catch(() => undefined);
      throw new XlvError('XLV_E_ASYNC_SCHEMA', 'The schema validates asynchronously, so readTable cannot use it.', {
        location: { sheet: prepared.sheet },
        hint: 'Use await readTableAsync(input, options) with this schema, or a schema without asynchronous checks.',
      });
    }
    return result;
  });
  return { ...base, ...collect(prepared, results) };
}

/**
 * Like {@link readTable}, but awaits schemas that validate asynchronously (for example Zod
 * schemas with async refinements).
 *
 * @example
 * ```ts
 * import { readTableAsync } from 'xlvane';
 * import { z } from 'zod';
 *
 * declare const bytes: Uint8Array;
 * const schema = z.object({ Email: z.string().refine(async (email: string) => email.includes('@')) });
 * const { rows } = await readTableAsync(bytes, { schema });
 * console.log(rows.map((row) => row.Email));
 * ```
 */
export async function readTableAsync<S extends StandardSchemaV1>(
  input: Uint8Array | ArrayBuffer | Workbook,
  options: ReadTableOptions<S> & { readonly schema: S },
): Promise<TableResult<StandardSchemaV1.InferOutput<S>>>;
export async function readTableAsync(input: Uint8Array | ArrayBuffer | Workbook, options?: ReadTableOptions): Promise<TableResult<TableRow>>;
export async function readTableAsync(
  input: Uint8Array | ArrayBuffer | Workbook,
  options?: ReadTableOptions<StandardSchemaV1 | undefined>,
): Promise<TableResult<unknown>> {
  const prepared = prepare(input, options);
  const base = { sheet: prepared.sheet, columns: prepared.columns, headerRow: prepared.headerRow, range: prepared.range };
  const schema = prepared.schema;
  if (schema === undefined) return { ...base, rows: prepared.rows, rowNumbers: prepared.rowNumbers, issues: [] };
  const results: StandardSchemaV1.Result<unknown>[] = [];
  for (const row of prepared.rows) results.push(await schema['~standard'].validate(row));
  return { ...base, ...collect(prepared, results) };
}

/**
 * A JSON-safe table, as returned by {@link toJSON}.
 *
 * @example
 * ```ts
 * import { toJSON, type JsonTable } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const table: JsonTable = toJSON(bytes);
 * console.log(JSON.stringify(table.rows));
 * ```
 */
export interface JsonTable {
  readonly sheet: string;
  readonly headerRow: number | null;
  readonly range: string | null;
  /** Column keys in sheet order. */
  readonly columns: string[];
  /** Rows with numbers, strings, booleans, `null`, and dates as ISO 8601 strings. */
  readonly rows: Record<string, string | number | boolean | null>[];
}

/**
 * Options for {@link toJSON}: the options of {@link readTable} without `schema` and `dates`.
 *
 * @example
 * ```ts
 * import type { ToJSONOptions } from 'xlvane';
 *
 * const options: ToJSONOptions = { sheet: 'Data', header: 1 };
 * ```
 */
export type ToJSONOptions = Omit<ReadTableOptions, 'schema' | 'dates'>;

/**
 * Reads a sheet as JSON-safe data: like {@link readTable} without a schema, with dates as ISO
 * 8601 strings. `JSON.stringify` of the result never loses information.
 *
 * @example
 * ```ts
 * import { toJSON } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * console.log(JSON.stringify(toJSON(bytes, { sheet: 'Data' }).rows, null, 2));
 * ```
 */
export function toJSON(input: Uint8Array | ArrayBuffer | Workbook, options?: ToJSONOptions): JsonTable {
  const o = readOptions(options, ['sheet', 'header', 'range', 'skipEmptyRows', 'limits'], 'toJSON options');
  const table = readTable(input, { ...Object.fromEntries(o), dates: 'iso' } as ReadTableOptions);
  return {
    sheet: table.sheet,
    headerRow: table.headerRow,
    range: table.range,
    columns: table.columns.map((c) => c.key),
    rows: table.rows as Record<string, string | number | boolean | null>[],
  };
}

// ---------------------------------------------------------------------------------------------
// Writing

/**
 * A column of {@link writeTable}: a key of the row objects, or an object with display options.
 *
 * @example
 * ```ts
 * import type { TableColumnSpec } from 'xlvane';
 *
 * const columns: TableColumnSpec[] = ['name', { key: 'total', header: 'Total (EUR)', numFmt: '#,##0.00', width: 14 }];
 * ```
 */
export type TableColumnSpec =
  | string
  | {
      /** Property of the row objects (for rows given as arrays: ignored; columns are by position). */
      readonly key?: string;
      /** Header text. Default: the key. */
      readonly header?: string;
      /** Number format for the data cells of the column. */
      readonly numFmt?: string;
      /** Column width in characters (0–255). */
      readonly width?: number;
    };

/**
 * Options for {@link writeTable} and {@link fromJSON}.
 *
 * @example
 * ```ts
 * import type { WriteTableOptions } from 'xlvane';
 *
 * const options: WriteTableOptions = { sheet: 'Orders', preset: 'report', columns: ['id', 'customer', 'total'] };
 * ```
 */
export interface WriteTableOptions {
  /** Columns in order. Default: every key found in the rows, in order of first appearance. */
  readonly columns?: readonly TableColumnSpec[];
  /** Sheet name. Default `"Sheet1"`. */
  readonly sheet?: string;
  /**
   * `"report"` (default): bold header with a fill and a bottom border, frozen header row,
   * autofilter and column widths fitted to the content. `"plain"`: values only.
   */
  readonly preset?: 'report' | 'plain';
  /** Date system of the new workbook. Default `"1900"`. */
  readonly dateSystem?: DateSystem;
}

const WRITE_OPTIONS = ['columns', 'sheet', 'preset', 'dateSystem'];
/** Header style of the `report` preset (D-304). */
const REPORT_HEADER: CellStyle = { font: { bold: true }, fill: { color: 'D9E1F2' }, border: { bottom: { style: 'thin' } } };
const MIN_WIDTH = 8;
const MAX_FITTED_WIDTH = 60;

interface ResolvedColumn {
  readonly key: string;
  readonly header: string;
  readonly numFmt: string | undefined;
  readonly width: number | undefined;
}

function tableError(message: string, hint: string, details: Record<string, string | number> = {}): XlvError {
  return new XlvError('XLV_E_TABLE_INVALID', message, { hint, details });
}

function resolveColumns(spec: unknown, rows: readonly unknown[], arrays: boolean): ResolvedColumn[] {
  if (spec === undefined) {
    if (arrays) {
      const width = rows.reduce<number>((n, r) => Math.max(n, (r as unknown[]).length), 0);
      return Array.from({ length: width }, (_, i) => ({ key: String(i), header: '', numFmt: undefined, width: undefined }));
    }
    const keys = new Set<string>();
    for (const row of rows) for (const key of Object.keys(row as object)) keys.add(key);
    return [...keys].map((key) => ({ key, header: key, numFmt: undefined, width: undefined }));
  }
  if (!Array.isArray(spec)) {
    throw new XlvError('XLV_E_INVALID_OPTION', 'The option "columns" must be an array.', {
      hint: 'Pass column keys or objects, for example columns: ["name", { key: "total", numFmt: "#,##0.00" }].',
      details: { option: 'columns' },
    });
  }
  return spec.map((entry: unknown, i): ResolvedColumn => {
    if (typeof entry === 'string') return { key: arrays ? String(i) : entry, header: entry, numFmt: undefined, width: undefined };
    const o = readOptions(entry, ['key', 'header', 'numFmt', 'width'], `columns[${String(i)}]`);
    const key = o.get('key');
    const header = o.get('header');
    const numFmt = o.get('numFmt');
    const width = o.get('width');
    if (!arrays && typeof key !== 'string') {
      throw new XlvError('XLV_E_INVALID_OPTION', `columns[${String(i)}] has no "key".`, {
        hint: 'Give each column the property name it shows, for example { key: "total", header: "Total" }.',
        details: { option: 'columns' },
      });
    }
    if (header !== undefined && typeof header !== 'string') throw new XlvError('XLV_E_INVALID_OPTION', `columns[${String(i)}].header must be text.`, { hint: 'Use a string such as "Total".', details: { option: 'columns' } });
    if (numFmt !== undefined && typeof numFmt !== 'string') throw new XlvError('XLV_E_INVALID_OPTION', `columns[${String(i)}].numFmt must be text.`, { hint: 'Use a format code such as "#,##0.00".', details: { option: 'columns' } });
    if (width !== undefined && typeof width !== 'number') throw new XlvError('XLV_E_INVALID_OPTION', `columns[${String(i)}].width must be a number.`, { hint: 'Use a width in characters, for example 20.', details: { option: 'columns' } });
    const k = arrays ? String(i) : (key as string);
    return { key: k, header: (header as string | undefined) ?? (typeof key === 'string' ? key : ''), numFmt: numFmt as string | undefined, width: width as number | undefined };
  });
}

function displayLength(value: unknown): number {
  if (typeof value === 'string') return Math.min(value.length, 200);
  if (typeof value === 'number') return String(value).length;
  if (typeof value === 'boolean') return 5;
  if (value instanceof Date) return value.getTime() % 86_400_000 === 0 ? 10 : 19;
  return 0;
}

/** Adds a table as a new sheet of `workbook`; shared by writeTable, fromJSON and fromCSV. */
export function addTableSheet(workbook: NewWorkbook, rows: unknown, options: Map<string, unknown>): void {
  if (!Array.isArray(rows)) {
    throw tableError(`Expected an array of rows, got ${rows === null ? 'null' : typeof rows}.`, 'Pass an array of objects, such as [{ name: "Alice", total: 10 }], or an array of arrays.');
  }
  if (rows.length > MAX_ROWS - 1) {
    throw new XlvError('XLV_E_LIMIT_EXCEEDED', `${String(rows.length)} rows do not fit in one sheet.`, {
      hint: `A sheet has ${String(MAX_ROWS)} rows, one of them for the header. Split the data over several sheets.`,
      details: { limit: 'rows', value: MAX_ROWS - 1 },
    });
  }
  const arrays = rows.length > 0 && Array.isArray(rows[0]);
  rows.forEach((row: unknown, i) => {
    const ok = arrays ? Array.isArray(row) : typeof row === 'object' && row !== null && !Array.isArray(row) && !(row instanceof Date);
    if (!ok) {
      throw tableError(
        `Row ${String(i + 1)} is ${row === null ? 'null' : Array.isArray(row) ? 'an array' : typeof row}, but the first row is ${arrays ? 'an array' : 'an object'}.`,
        'Give every row the same shape: all objects (keys become columns) or all arrays (columns by position).',
        { row: i + 2 },
      );
    }
  });
  const columns = resolveColumns(options.get('columns'), rows, arrays);
  if (columns.length > MAX_COLUMNS) {
    throw new XlvError('XLV_E_LIMIT_EXCEEDED', `${String(columns.length)} columns do not fit in one sheet.`, {
      hint: `A sheet has ${String(MAX_COLUMNS)} columns. Pass fewer columns with the "columns" option.`,
      details: { limit: 'columns', value: MAX_COLUMNS },
    });
  }
  const preset = oneOfOption('preset', options.get('preset'), ['report', 'plain'], 'report');
  const sheetName = options.get('sheet') ?? 'Sheet1';
  const sheet = workbook.addSheet(sheetName as string);
  const hasHeader = !arrays || columns.some((c) => c.header !== '');
  const offset = hasHeader ? 1 : 0;
  const widths = columns.map((c) => displayLength(c.header));

  columns.forEach((column, c) => {
    if (hasHeader && column.header !== '') sheet.set(formatCellRef(1, c + 1), column.header, preset === 'report' ? REPORT_HEADER : undefined);
  });
  rows.forEach((row: unknown, i) => {
    const r = i + 1 + offset;
    columns.forEach((column, c) => {
      const record = row as Record<string, unknown>;
      const raw = arrays ? (row as unknown[])[c] : Object.prototype.hasOwnProperty.call(record, column.key) ? record[column.key] : undefined;
      const value = raw === undefined ? null : raw;
      if (!(value === null || typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean' || value instanceof Date)) {
        throw tableError(
          `Row ${String(i + 1)}, column "${column.key}" holds ${Array.isArray(value) ? 'an array' : typeof value}, which is not a cell value.`,
          'Cell values are numbers, strings, booleans, Dates or null. Convert nested data first, for example with JSON.stringify.',
          { row: r, key: column.key },
        );
      }
      if (value === null && column.numFmt === undefined) return;
      sheet.set(formatCellRef(r, c + 1), value, column.numFmt === undefined ? undefined : { numFmt: column.numFmt });
      widths[c] = Math.max(widths[c] ?? 0, displayLength(value));
    });
  });
  columns.forEach((column, c) => {
    if (column.width !== undefined) sheet.column(c + 1, { width: column.width });
    else if (preset === 'report') sheet.column(c + 1, { width: Math.min(MAX_FITTED_WIDTH, Math.max(MIN_WIDTH, (widths[c] ?? 0) + 2)) });
  });
  if (preset === 'report' && hasHeader && columns.length > 0) {
    sheet.freeze({ rows: 1 });
    sheet.autoFilter(`A1:${formatCellRef(Math.max(1, rows.length + offset), columns.length)}`);
  }
}

/**
 * Writes rows as a formatted XLSX table and returns the file's bytes. Rows are objects (keys
 * become columns) or arrays (columns by position, headers from `columns`). The default
 * `"report"` preset makes a readable report: bold header with a fill and a bottom border,
 * frozen header row, autofilter, and column widths fitted to the content. Dates get a date
 * format.
 *
 * @example
 * ```ts
 * import { writeTable } from 'xlvane';
 *
 * const bytes = writeTable(
 *   [
 *     { customer: 'Alice', total: 120.5, date: new Date(Date.UTC(2024, 0, 31)) },
 *     { customer: 'Bob', total: 80, date: new Date(Date.UTC(2024, 1, 2)) },
 *   ],
 *   { sheet: 'Orders', columns: ['date', 'customer', { key: 'total', header: 'Total', numFmt: '#,##0.00' }] },
 * );
 * ```
 */
export function writeTable(rows: readonly object[] | readonly (readonly unknown[])[], options?: WriteTableOptions): Uint8Array {
  const o = readOptions(options, WRITE_OPTIONS, 'writeTable options');
  const dateSystem = o.get('dateSystem');
  const workbook = createWorkbook(dateSystem === undefined ? undefined : { dateSystem: dateSystem as DateSystem });
  addTableSheet(workbook, rows, o);
  return writeWorkbook(workbook);
}

/**
 * Writes JSON data as a formatted XLSX table: the same as {@link writeTable}, but also accepts
 * JSON text. Throws `XLV_E_JSON_INVALID` for text that is not valid JSON and
 * `XLV_E_TABLE_INVALID` when the data is not an array of objects or arrays. JSON keys such as
 * `__proto__` become ordinary columns.
 *
 * @example
 * ```ts
 * import { fromJSON } from 'xlvane';
 *
 * const bytes = fromJSON('[{"sku":"A-1","stock":12},{"sku":"B-7","stock":0}]', { sheet: 'Stock' });
 * ```
 */
export function fromJSON(input: string | readonly object[], options?: WriteTableOptions): Uint8Array {
  let rows: unknown = input;
  if (typeof input === 'string') {
    try {
      rows = JSON.parse(input);
    } catch (error) {
      throw new XlvError('XLV_E_JSON_INVALID', `The text is not valid JSON: ${error instanceof Error ? error.message.slice(0, 200) : 'parse error'}.`, {
        hint: 'Check the JSON text (the message gives the position of the problem), or pass the parsed array instead.',
        cause: error,
      });
    }
  }
  if (!Array.isArray(rows)) {
    throw tableError('The JSON data is not an array.', 'Pass an array of row objects, for example [{"name":"Alice"}]. For an object that holds the rows, pass that property, e.g. data.rows.');
  }
  return writeTable(rows as object[], options);
}
