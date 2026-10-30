// Building new workbooks in memory and writing them as XLSX (docs/decisions.md D-201 … D-210).
// Runtime-neutral: the result is a Uint8Array; `xlvane/node` adds `writeWorkbookFile`.
import { XlvError } from './errors';
import type { DateSystem, SheetState } from './model';
import { MAX_COLUMNS, MAX_ROWS, columnNumber, formatCellRef, parseCellRef, parseRange, type CellRange } from './sml/cell-ref';
import { dateToSerial } from './sml/dates';
import { writePackage, type WriteSheetData } from './sml/write-parts';
import { mergeStyles, normalizeStyle, stableKey, type CellStyle, type NormalizedStyle } from './style';
import { suggest } from './util/suggest';
import { firstLoneSurrogate } from './xml/writer';

/**
 * Error values a cell can hold (ECMA-376 Part 1 §18.17.3).
 *
 * @example
 * ```ts
 * import { createWorkbook, type ErrorValue } from 'xlvane';
 *
 * const missing: ErrorValue = '#N/A';
 * createWorkbook().addSheet('Data').set('A1', { error: missing });
 * ```
 */
export type ErrorValue = '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A' | '#GETTING_DATA';

/**
 * A value that a formula cell stores as its cached result.
 *
 * @example
 * ```ts
 * import { createWorkbook, type FormulaResult } from 'xlvane';
 *
 * const cached: FormulaResult = 42;
 * createWorkbook().addSheet('Data').set('A1', 40).set('A2', { formula: 'A1+2', result: cached });
 * ```
 */
export type FormulaResult = number | string | boolean | Date | { readonly error: ErrorValue };

/**
 * A value for {@link NewSheet.set}: a number, a string, a boolean, a `Date` (its UTC fields are
 * the spreadsheet's date and time), `null` for no value, `{ error: "#N/A" }` for an error
 * value, or `{ formula, result }` for a formula with an optional cached result.
 *
 * @example
 * ```ts
 * import { createWorkbook, type CellInput } from 'xlvane';
 *
 * const row: CellInput[] = ['Tea', 12.5, true, new Date(Date.UTC(2024, 0, 31)), null, { formula: 'B1*2' }];
 * const sheet = createWorkbook().addSheet('Data');
 * row.forEach((value, i) => sheet.set(`${String.fromCharCode(65 + i)}1`, value));
 * ```
 */
export type CellInput =
  | number
  | string
  | boolean
  | Date
  | null
  | { readonly error: ErrorValue }
  | { readonly formula: string; readonly result?: FormulaResult };

/**
 * Options for {@link createWorkbook}.
 *
 * @example
 * ```ts
 * import { createWorkbook, type CreateWorkbookOptions } from 'xlvane';
 *
 * const options: CreateWorkbookOptions = { dateSystem: '1904' };
 * const workbook = createWorkbook(options);
 * console.log(workbook.dateSystem); // "1904"
 * ```
 */
export interface CreateWorkbookOptions {
  /** Date system for serial dates. Default `"1900"`, the system Excel uses on Windows. */
  readonly dateSystem?: DateSystem;
}

/**
 * Options for {@link NewWorkbook.addSheet}.
 *
 * @example
 * ```ts
 * import { createWorkbook, type AddSheetOptions } from 'xlvane';
 *
 * const options: AddSheetOptions = { state: 'hidden' };
 * createWorkbook().addSheet('Lookup', options);
 * ```
 */
export interface AddSheetOptions {
  /** Visibility. Default `"visible"`. */
  readonly state?: SheetState;
}

/**
 * Options for {@link NewSheet.column}.
 *
 * @example
 * ```ts
 * import { createWorkbook, type ColumnOptions } from 'xlvane';
 *
 * const wide: ColumnOptions = { width: 40 };
 * createWorkbook().addSheet('Data').column('A', wide);
 * ```
 */
export interface ColumnOptions {
  /**
   * Column width as stored in the file (`col@width`), 0 to 255 ([MS-OI29500] §2.1.598 c). The
   * unit is characters of the maximum digit width of the Normal style's font, including 5 pixels
   * of padding, truncated to 1/256 (ECMA-376 Part 1 §18.3.1.13). This is not the number Excel
   * shows in its column width box: with Calibri 11, a column Excel shows as 8 characters wide is
   * stored as 8.7109375. The value is written unchanged.
   */
  readonly width?: number;
  readonly hidden?: boolean;
}

/**
 * Options for {@link NewSheet.row}.
 *
 * @example
 * ```ts
 * import { createWorkbook, type RowOptions } from 'xlvane';
 *
 * const tall: RowOptions = { height: 30 };
 * createWorkbook().addSheet('Data').row(1, tall);
 * ```
 */
export interface RowOptions {
  /** Height in points, 0 to 409. */
  readonly height?: number;
  readonly hidden?: boolean;
}

/**
 * Options for {@link NewSheet.freeze}.
 *
 * @example
 * ```ts
 * import { createWorkbook, type FreezeOptions } from 'xlvane';
 *
 * const header: FreezeOptions = { rows: 1, cols: 1 };
 * createWorkbook().addSheet('Data').freeze(header);
 * ```
 */
export interface FreezeOptions {
  /** Number of rows at the top that stay visible when scrolling. Default 0. */
  readonly rows?: number;
  /** Number of columns at the left that stay visible when scrolling. Default 0. */
  readonly cols?: number;
}

/**
 * Error values accepted in cells (ECMA-376 Part 1 §18.17.3, fact E-33).
 *
 * @example
 * ```ts
 * import { ERROR_VALUES } from 'xlvane';
 *
 * console.log(ERROR_VALUES.includes('#DIV/0!')); // true
 * ```
 */
export const ERROR_VALUES: readonly ErrorValue[] = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'];
/** Most characters a cell can contain (Excel specifications and limits, W-04). */
const MAX_CELL_TEXT = 32_767;
/** Widest column, in characters (Excel specifications and limits, W-04). */
/** Widest column Office accepts ([MS-OI29500] §2.1.598 c). */
const MAX_COLUMN_WIDTH = 255;
/** Tallest row, in points (Excel specifications and limits, W-04). */
const MAX_ROW_HEIGHT = 409;
/** Most cells one call to {@link NewSheet.style} may touch. Implementation cap (D-206). */
export const MAX_STYLE_RANGE_CELLS = 1_048_576;
/** Longest sheet name ([MS-OI29500] §2.1.578 a; docs/facts/sml-write.md W2-20). */
const MAX_SHEET_NAME = 31;
/** Characters not allowed in sheet names (ECMA-376 Part 1 §18.17 sheet-name-special; [MS-OI29500] §2.1.578 a, W2-20). */
const SHEET_NAME_FORBIDDEN = /[\\/?*:[\]]/;
/** Formula text limit (Excel specifications and limits: 8,192 characters; to be verified, W2-21). */
const MAX_FORMULA_LENGTH = 8192;
export const DATE_FORMAT = 'yyyy-mm-dd';
export const DATE_TIME_FORMAT = 'yyyy-mm-dd hh:mm:ss';
const DAY_MS = 86_400_000;

/** A cell value after validation, as the writer stores it. */
export type StoredValue =
  | { readonly kind: 'number'; readonly n: number }
  | { readonly kind: 'string'; readonly s: string }
  | { readonly kind: 'boolean'; readonly b: boolean }
  | { readonly kind: 'error'; readonly e: ErrorValue }
  | {
      readonly kind: 'formula';
      readonly text: string;
      readonly result: Exclude<StoredValue, { kind: 'formula' }> | null;
    };

/** A stored cell: its value (or `null`) and the id of its interned style (0 = default). */
export interface StoredCell {
  value: StoredValue | null;
  style: number;
}

function invalidValue(sheet: string, cell: string, message: string, hint: string): XlvError {
  return new XlvError('XLV_E_INVALID_VALUE', message, { location: { sheet, cell }, hint });
}

function invalidOption(message: string, hint: string, details?: Record<string, string | readonly string[]>): XlvError {
  return new XlvError('XLV_E_INVALID_OPTION', message, { hint, ...(details ? { details } : {}) });
}

/** Reads the own enumerable properties of an options object, rejecting unknown ones. */
function readOptions(value: unknown, known: readonly string[], what: string): Map<string, unknown> {
  const out = new Map<string, unknown>();
  if (value === undefined) return out;
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw invalidOption(`${what} must be an object.`, `Pass an object with some of these properties: ${known.join(', ')}.`);
  }
  for (const key of Object.keys(value)) {
    if (!known.includes(key)) {
      const close = suggest(key, known);
      throw invalidOption(
        `Unknown option "${key.slice(0, 40)}" in ${what}.`,
        `${close.length > 0 ? `Did you mean "${close[0] ?? ''}"? ` : ''}Known options: ${known.join(', ')}.`,
        { option: key.slice(0, 40), known },
      );
    }
    const v = (value as Record<string, unknown>)[key];
    if (v !== undefined) out.set(key, v);
  }
  return out;
}

function invalidRef(ref: unknown, sheet: string, what = 'cell reference', example = '"B7"'): XlvError {
  const shown = String(typeof ref === 'symbol' ? 'symbol' : ref).slice(0, 40);
  return new XlvError('XLV_E_INVALID_CELL_REF', `"${shown}" is not a valid ${what}.`, {
    location: { sheet },
    hint: `Pass an A1-style ${what} such as ${example} (column letters A to XFD, rows 1 to 1048576).`,
    details: { ref: shown },
  });
}

/** Validates a sheet name; throws `XLV_E_INVALID_SHEET_NAME` with the reason. */
function checkSheetName(name: unknown): string {
  const fail = (reason: string): never => {
    throw new XlvError('XLV_E_INVALID_SHEET_NAME', `Invalid sheet name ${JSON.stringify(String(name).slice(0, 40))}: ${reason}`, {
      hint: `Use 1 to ${String(MAX_SHEET_NAME)} characters without \\ / ? * : [ ] and without an apostrophe at the start or end.`,
      details: { name: String(name).slice(0, 40) },
    });
  };
  if (typeof name !== 'string') return fail('it must be a string.');
  if (name.length === 0) return fail('it is empty.');
  if (name.length > MAX_SHEET_NAME) return fail(`it has ${String(name.length)} characters, more than ${String(MAX_SHEET_NAME)}.`);
  const bad = SHEET_NAME_FORBIDDEN.exec(name);
  if (bad) return fail(`it contains "${bad[0]}".`);
  if (name.startsWith("'") || name.endsWith("'")) return fail('it starts or ends with an apostrophe.');
  if (firstLoneSurrogate(name) !== -1) return fail('it contains an unpaired UTF-16 surrogate.');
  // Characters that XML cannot represent; sheet names are attribute values, not ST_Xstring.
  // This includes U+0003 END OF TEXT, which Office also excludes from sheet names
  // ([MS-OI29500] §2.1.578 a, §2.1.863 a).
  if (/[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]/.test(name)) return fail('it contains a control character.');
  return name;
}

/**
 * A sheet of a workbook created with {@link createWorkbook}. Methods return the sheet, so calls
 * can be chained.
 *
 * @example
 * ```ts
 * import { createWorkbook } from 'xlvane';
 *
 * const sheet = createWorkbook().addSheet('Sales');
 * sheet.set('A1', 'Product', { font: { bold: true } }).set('B1', 'Amount', { font: { bold: true } });
 * sheet.set('A2', 'Tea').set('B2', 12.5, { numFmt: '#,##0.00' });
 * sheet.column('A', { width: 20 }).freeze({ rows: 1 }).autoFilter('A1:B2');
 * ```
 */
export class NewSheet {
  /** Visibility of the sheet. */
  readonly state: SheetState;
  /** @internal */
  readonly cellMap: Map<number, Map<number, StoredCell>> = new Map();
  /** @internal */
  readonly columnMap: Map<number, { width: number | null; hidden: boolean }> = new Map();
  /** @internal */
  readonly rowMap: Map<number, { height: number | null; hidden: boolean }> = new Map();
  /** @internal */
  readonly mergeList: CellRange[] = [];
  /** @internal */
  frozen: { rows: number; cols: number } | null = null;
  /** @internal */
  filter: CellRange | null = null;
  private count = 0;

  /** @internal Use {@link NewWorkbook.addSheet}. */
  constructor(
    /** Name of the sheet. */
    readonly name: string,
    state: SheetState,
    private readonly book: NewWorkbook,
  ) {
    this.state = state;
  }

  /** Number of cells that have a value or a style. */
  get cellCount(): number {
    return this.count;
  }

  private address(ref: unknown): { row: number; col: number; ref: string } {
    const address = typeof ref === 'string' ? parseCellRef(ref) : null;
    if (address === null) throw invalidRef(ref, this.name);
    return { ...address, ref: formatCellRef(address.row, address.col) };
  }

  private range(ref: unknown, what = 'range', example = '"A1:C3"'): CellRange {
    const range = typeof ref === 'string' ? parseRange(ref) : null;
    if (range === null) throw invalidRef(ref, this.name, what, example);
    return range;
  }

  private getCell(row: number, col: number): StoredCell | undefined {
    return this.cellMap.get(row)?.get(col);
  }

  private putCell(row: number, col: number, cell: StoredCell | null): void {
    let cells = this.cellMap.get(row);
    if (cell === null) {
      if (cells?.delete(col) === true) {
        this.count--;
        if (cells.size === 0) this.cellMap.delete(row);
      }
      return;
    }
    if (!cells) {
      cells = new Map();
      this.cellMap.set(row, cells);
    }
    if (!cells.has(col)) this.count++;
    cells.set(col, cell);
  }

  /**
   * Sets the value of a cell, and its style when `style` is given (otherwise the cell keeps its
   * current style). Dates get a date number format (`yyyy-mm-dd`, or `yyyy-mm-dd hh:mm:ss` when
   * they have a time) unless the style sets one. Setting `null` without a style removes an
   * unstyled cell.
   *
   * Throws `XLV_E_INVALID_CELL_REF` for a bad reference, `XLV_E_INVALID_VALUE` for values a
   * spreadsheet cannot hold (NaN, infinities, text over 32 767 characters, dates outside the
   * date system) and `XLV_E_INVALID_STYLE` for a bad style.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * const sheet = createWorkbook().addSheet('Data');
   * sheet.set('A1', 'Total').set('B1', { formula: 'SUM(B2:B10)', result: 0 });
   * sheet.set('C1', new Date(Date.UTC(2024, 0, 31)), { font: { italic: true } });
   * sheet.set('D1', { error: '#N/A' });
   * ```
   */
  set(ref: string, value: CellInput, style?: CellStyle): this {
    const at = this.address(ref);
    const normalized = style === undefined ? null : normalizeStyle(style);
    const { stored, isDate, hasTime } = this.convert(value, at.ref);
    let base = normalized ?? this.book.styleOf(this.getCell(at.row, at.col)?.style ?? 0);
    if (isDate && base.numFmt === undefined) base = { ...base, numFmt: hasTime ? DATE_TIME_FORMAT : DATE_FORMAT };
    const styleId = this.book.intern(base);
    this.putCell(at.row, at.col, stored === null && styleId === 0 ? null : { value: stored, style: styleId });
    return this;
  }

  private convert(value: unknown, cell: string): ConvertedValue {
    return convertCellInput(value, { sheet: this.name, dateSystem: this.book.dateSystem }, cell);
  }

  /**
   * Applies a style to every cell of a range, merging it with each cell's current style: given
   * font, border and alignment properties replace the cell's ones, a given fill or number format
   * replaces the cell's. Cells without a value are created as empty styled cells. Throws
   * `XLV_E_LIMIT_EXCEEDED` for ranges of more than 1 048 576 cells.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * const sheet = createWorkbook().addSheet('Data').set('A1', 'Name').set('B1', 'Amount');
   * sheet.style('A1:B1', { font: { bold: true }, fill: { color: 'D9E1F2' } });
   * ```
   */
  style(range: string, style: CellStyle): this {
    const r = this.range(range);
    const patch = normalizeStyle(style);
    const cells = (r.end.row - r.start.row + 1) * (r.end.col - r.start.col + 1);
    if (cells > MAX_STYLE_RANGE_CELLS) {
      throw new XlvError('XLV_E_LIMIT_EXCEEDED', `The range ${range} has ${String(cells)} cells, more than ${String(MAX_STYLE_RANGE_CELLS)} can be styled at once.`, {
        location: { sheet: this.name },
        hint: 'Style only the cells that hold data, for example the header row and the used range.',
        details: { limit: 'styleRangeCells', value: MAX_STYLE_RANGE_CELLS },
      });
    }
    const cache = new Map<number, number>();
    for (let row = r.start.row; row <= r.end.row; row++) {
      for (let col = r.start.col; col <= r.end.col; col++) {
        const cell = this.getCell(row, col);
        const old = cell?.style ?? 0;
        let id = cache.get(old);
        if (id === undefined) {
          id = this.book.intern(mergeStyles(this.book.styleOf(old), patch));
          cache.set(old, id);
        }
        if (cell) cell.style = id;
        else if (id !== 0) this.putCell(row, col, { value: null, style: id });
      }
    }
    return this;
  }

  /**
   * Sets the width or visibility of a column, given by letters (`"C"`) or number (3).
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * createWorkbook().addSheet('Data').column('A', { width: 30 }).column(2, { hidden: true });
   * ```
   */
  column(col: string | number, options: ColumnOptions): this {
    const index = typeof col === 'number' ? col : typeof col === 'string' ? (columnNumber(col.toUpperCase()) ?? 0) : 0;
    if (!Number.isInteger(index) || index < 1 || index > MAX_COLUMNS) throw invalidRef(col, this.name, 'column', '"C" or 3');
    const o = readOptions(options, ['width', 'hidden'], 'column options');
    const current = this.columnMap.get(index) ?? { width: null, hidden: false };
    const width = o.get('width');
    if (width !== undefined) {
      if (typeof width !== 'number' || !(width >= 0 && width <= MAX_COLUMN_WIDTH)) {
        throw invalidValue(this.name, '', `Column width ${String(width)} is not valid.`, `Use a width from 0 to ${String(MAX_COLUMN_WIDTH)} characters.`);
      }
      current.width = width;
    }
    const hidden = o.get('hidden');
    if (hidden !== undefined) {
      if (typeof hidden !== 'boolean') throw invalidOption('Column option "hidden" must be true or false.', 'Use { hidden: true }.');
      current.hidden = hidden;
    }
    this.columnMap.set(index, current);
    return this;
  }

  /**
   * Sets the height or visibility of a row (1-based).
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * createWorkbook().addSheet('Data').row(1, { height: 30 }).row(5, { hidden: true });
   * ```
   */
  row(row: number, options: RowOptions): this {
    if (typeof row !== 'number' || !Number.isInteger(row) || row < 1 || row > MAX_ROWS) throw invalidRef(row, this.name, 'row number', '1');
    const o = readOptions(options, ['height', 'hidden'], 'row options');
    const current = this.rowMap.get(row) ?? { height: null, hidden: false };
    const height = o.get('height');
    if (height !== undefined) {
      if (typeof height !== 'number' || !(height >= 0 && height <= MAX_ROW_HEIGHT)) {
        throw invalidValue(this.name, '', `Row height ${String(height)} is not valid.`, `Use a height from 0 to ${String(MAX_ROW_HEIGHT)} points.`);
      }
      current.height = height;
    }
    const hidden = o.get('hidden');
    if (hidden !== undefined) {
      if (typeof hidden !== 'boolean') throw invalidOption('Row option "hidden" must be true or false.', 'Use { hidden: true }.');
      current.hidden = hidden;
    }
    this.rowMap.set(row, current);
    return this;
  }

  /**
   * Merges a range of at least two cells. Throws `XLV_E_OVERLAPPING_MERGE` if it overlaps a
   * range merged before.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * createWorkbook().addSheet('Report').set('A1', 'Quarterly report').merge('A1:D1');
   * ```
   */
  merge(range: string): this {
    const r = this.range(range);
    if (r.start.row === r.end.row && r.start.col === r.end.col) throw invalidRef(range, this.name, 'range of two or more cells', '"A1:C1"');
    for (const m of this.mergeList) {
      if (r.start.row <= m.end.row && m.start.row <= r.end.row && r.start.col <= m.end.col && m.start.col <= r.end.col) {
        const other = `${formatCellRef(m.start.row, m.start.col)}:${formatCellRef(m.end.row, m.end.col)}`;
        throw new XlvError('XLV_E_OVERLAPPING_MERGE', `The range ${range} overlaps the merged range ${other}.`, {
          location: { sheet: this.name },
          hint: 'Merged ranges must not overlap. Choose ranges that share no cells.',
          details: { range: String(range), other },
        });
      }
    }
    this.mergeList.push(r);
    return this;
  }

  /**
   * Freezes the top rows and left columns so that they stay visible when scrolling. Pass
   * `{ rows: 0, cols: 0 }` to remove the frozen panes.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * createWorkbook().addSheet('Data').freeze({ rows: 1 }); // keep the header row visible
   * ```
   */
  freeze(options: FreezeOptions): this {
    const o = readOptions(options, ['rows', 'cols'], 'freeze options');
    const count = (key: string, max: number): number => {
      const v = o.get(key) ?? 0;
      if (typeof v !== 'number' || !Number.isInteger(v) || v < 0 || v >= max) {
        throw invalidOption(`freeze option "${key}" must be a whole number from 0 to ${String(max - 1)}.`, `For example { ${key}: 1 }.`);
      }
      return v;
    };
    const rows = count('rows', MAX_ROWS);
    const cols = count('cols', MAX_COLUMNS);
    this.frozen = rows === 0 && cols === 0 ? null : { rows, cols };
    return this;
  }

  /**
   * Adds filter buttons to the header row of a range, or removes them with `null`.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * createWorkbook().addSheet('Data').set('A1', 'Name').set('B1', 'Amount').autoFilter('A1:B100');
   * ```
   */
  autoFilter(range: string | null): this {
    this.filter = range === null ? null : this.range(range);
    return this;
  }

  /** @internal */
  data(): WriteSheetData {
    return {
      name: this.name,
      state: this.state,
      cells: this.cellMap,
      columns: this.columnMap,
      rows: this.rowMap,
      merges: this.mergeList,
      freeze: this.frozen,
      autoFilter: this.filter,
    };
  }
}

/** Where a value is converted: names the sheet in errors and gives the date system. */
export interface ConvertContext {
  readonly sheet: string;
  readonly dateSystem: DateSystem;
}

/** A converted cell value; `isDate`/`hasTime` choose the default date format. */
export interface ConvertedValue {
  readonly stored: StoredValue | null;
  readonly isDate: boolean;
  readonly hasTime: boolean;
}

/**
 * Validates and converts a {@link CellInput} as `NewSheet.set` does (shared with the streaming
 * writer). Throws `XLV_E_INVALID_VALUE` with the sheet and cell.
 */
export function convertCellInput(value: unknown, ctx: ConvertContext, cell: string): ConvertedValue {
  if (value === null || value === undefined) return { stored: null, isDate: false, hasTime: false };
  if (typeof value === 'object' && !(value instanceof Date)) {
    const keys = Object.keys(value);
    if (keys.includes('formula')) {
      const o = readFormula(value, ctx.sheet, cell);
      if (o.result === undefined) return { stored: { kind: 'formula', text: o.text, result: null }, isDate: false, hasTime: false };
      const r = scalarValue(ctx, o.result, cell, 'formula result');
      return { stored: { kind: 'formula', text: o.text, result: r.stored }, isDate: r.isDate, hasTime: r.hasTime };
    }
  }
  return scalarValue(ctx, value, cell, 'value');
}

function scalarValue(
ctx: ConvertContext,
value: unknown,
cell: string,
what: string,
): { stored: Exclude<StoredValue, { kind: 'formula' }>; isDate: boolean; hasTime: boolean } {
  const plain = (stored: Exclude<StoredValue, { kind: 'formula' }>) => ({ stored, isDate: false, hasTime: false });
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      throw invalidValue(ctx.sheet, cell, `The ${what} for ${cell} is ${String(value)}.`, 'Spreadsheet cells hold finite numbers only. Write the value as text or leave the cell empty.');
    }
    return plain({ kind: 'number', n: Object.is(value, -0) ? 0 : value });
  }
  if (typeof value === 'string') {
    if (value.length > MAX_CELL_TEXT) {
      throw invalidValue(ctx.sheet, cell, `The ${what} for ${cell} has ${String(value.length)} characters.`, `A cell can contain at most ${String(MAX_CELL_TEXT)} characters. Split the text over several cells.`);
    }
    if (firstLoneSurrogate(value) !== -1) {
      throw invalidValue(ctx.sheet, cell, `The ${what} for ${cell} contains an unpaired UTF-16 surrogate.`, 'The text is not valid Unicode; it may have been cut in the middle of a character such as an emoji. Fix the text before writing it.');
    }
    return plain({ kind: 'string', s: value });
  }
  if (typeof value === 'boolean') return plain({ kind: 'boolean', b: value });
  if (value instanceof Date) {
    const time = value.getTime();
    let serial: number;
    try {
      serial = dateToSerial(value, ctx.dateSystem);
    } catch {
      serial = Number.NaN;
    }
    if (!(serial >= 0)) {
      throw invalidValue(
        ctx.sheet,
        cell,
        Number.isNaN(time) ? `The ${what} for ${cell} is an invalid Date.` : `The date ${value.toISOString()} for ${cell} is outside the ${ctx.dateSystem} date system.`,
        ctx.dateSystem === '1904'
          ? 'The 1904 date system covers 1904-01-01 to 9999-12-31. Write earlier dates as text.'
          : 'The 1900 date system covers 1899-12-31 (time-only values) to 9999-12-31. Write earlier dates as text.',
      );
    }
    return { stored: { kind: 'number', n: serial }, isDate: true, hasTime: time % DAY_MS !== 0 };
  }
  if (typeof value === 'object' && value !== null && Object.keys(value).includes('error')) {
    const keys = Object.keys(value);
    const e = (value as { error: unknown }).error;
    if (keys.length === 1 && typeof e === 'string' && (ERROR_VALUES as readonly string[]).includes(e)) {
      return plain({ kind: 'error', e: e as ErrorValue });
    }
    throw invalidValue(ctx.sheet, cell, `The error value for ${cell} is not valid.`, `Use { error: code } with one of: ${ERROR_VALUES.join(', ')}.`);
  }
  throw invalidValue(
    ctx.sheet,
    cell,
    `The ${what} for ${cell} has an unsupported type (${typeof value === 'object' ? 'object' : typeof value}).`,
    'Use a number, string, boolean, Date, null, { error: "#N/A" } or { formula: "A1*2", result: 2 }.',
  );
}

function readFormula(value: object, sheet: string, cell: string): { text: string; result: unknown } {
  const keys = Object.keys(value);
  for (const key of keys) {
    if (key !== 'formula' && key !== 'result') {
      throw invalidValue(sheet, cell, `Unknown property "${key.slice(0, 40)}" in the formula for ${cell}.`, 'Use { formula: "A1*2" } or { formula: "A1*2", result: 2 }.');
    }
  }
  const raw = (value as { formula: unknown }).formula;
  const text = typeof raw === 'string' && raw.startsWith('=') ? raw.slice(1) : raw;
  if (typeof text !== 'string' || text.trim().length === 0 || text.length > MAX_FORMULA_LENGTH || firstLoneSurrogate(text) !== -1) {
    throw invalidValue(sheet, cell, `The formula for ${cell} is empty, too long or not valid text.`, `Pass the formula text, with or without "=", for example { formula: "SUM(A1:A10)" }; at most ${String(MAX_FORMULA_LENGTH)} characters.`);
  }
  return { text, result: keys.includes('result') ? (value as { result: unknown }).result : undefined };
}

/**
 * A workbook being built in memory, created by {@link createWorkbook} and written with
 * {@link writeWorkbook}.
 *
 * @example
 * ```ts
 * import { createWorkbook, writeWorkbook } from 'xlvane';
 *
 * const workbook = createWorkbook();
 * workbook.addSheet('Summary').set('A1', 'Hello');
 * const bytes = writeWorkbook(workbook);
 * ```
 */
export class NewWorkbook {
  /** Date system used to store dates. */
  readonly dateSystem: DateSystem;
  private readonly list: NewSheet[] = [];
  private readonly styleIds = new Map<string, number>();
  private readonly styleList: NormalizedStyle[] = [{}];

  /** @internal Use {@link createWorkbook}. */
  constructor(dateSystem: DateSystem) {
    this.dateSystem = dateSystem;
    this.styleIds.set(stableKey({}), 0);
  }

  /** Sheets in workbook order. */
  get sheets(): readonly NewSheet[] {
    return [...this.list];
  }

  /**
   * Adds a sheet at the end. Throws `XLV_E_INVALID_SHEET_NAME` for names Excel does not accept
   * and `XLV_E_DUPLICATE_SHEET_NAME` if a sheet with the same name (ignoring case) exists.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * const workbook = createWorkbook();
   * workbook.addSheet('Data');
   * workbook.addSheet('Lookup', { state: 'hidden' });
   * ```
   */
  addSheet(name: string, options?: AddSheetOptions): NewSheet {
    const o = readOptions(options, ['state'], 'sheet options');
    const state = o.get('state') ?? 'visible';
    if (state !== 'visible' && state !== 'hidden' && state !== 'veryHidden') {
      throw invalidOption(`Sheet state ${JSON.stringify(String(state).slice(0, 40))} is not valid.`, 'Use "visible", "hidden" or "veryHidden".');
    }
    const checked = checkSheetName(name);
    const lower = checked.toLowerCase();
    const existing = this.list.find((s) => s.name.toLowerCase() === lower);
    if (existing) {
      throw new XlvError('XLV_E_DUPLICATE_SHEET_NAME', `The workbook already has a sheet named "${existing.name}".`, {
        hint: 'Sheet names must be unique, ignoring upper and lower case. Choose another name.',
        details: { name: checked, existing: existing.name },
      });
    }
    const sheet = new NewSheet(checked, state, this);
    this.list.push(sheet);
    return sheet;
  }

  /**
   * Returns the sheet with this name (exact match); throws `XLV_E_SHEET_NOT_FOUND` otherwise.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * const workbook = createWorkbook();
   * workbook.addSheet('Data');
   * workbook.getSheet('Data').set('A1', 1);
   * ```
   */
  getSheet(name: string): NewSheet {
    const found = this.list.find((s) => s.name === name);
    if (found) return found;
    const names = this.list.map((s) => s.name);
    const close = suggest(String(name), names);
    throw new XlvError('XLV_E_SHEET_NOT_FOUND', `The workbook has no sheet named "${String(name).slice(0, 100)}".`, {
      hint: `${close.length > 0 ? `Did you mean "${close[0] ?? ''}"? ` : ''}Available sheets: ${names.length > 0 ? names.join(', ') : '(none)'}.`,
      details: { name: String(name).slice(0, 100), available: names },
    });
  }

  /** @internal Returns the id of a normalised style, adding it if new. */
  intern(style: NormalizedStyle): number {
    const key = stableKey(style);
    let id = this.styleIds.get(key);
    if (id === undefined) {
      id = this.styleList.length;
      this.styleList.push(style);
      this.styleIds.set(key, id);
    }
    return id;
  }

  /** @internal */
  styleOf(id: number): NormalizedStyle {
    return this.styleList[id] ?? {};
  }

  /** @internal */
  styles(): readonly NormalizedStyle[] {
    return this.styleList;
  }
}

/**
 * Creates an empty workbook. Add sheets with {@link NewWorkbook.addSheet} and write it with
 * {@link writeWorkbook}.
 *
 * @example
 * ```ts
 * import { createWorkbook, writeWorkbook } from 'xlvane';
 *
 * const workbook = createWorkbook();
 * const sheet = workbook.addSheet('Sales');
 * sheet.set('A1', 'Region', { font: { bold: true } }).set('B1', 'Revenue', { font: { bold: true } });
 * sheet.set('A2', 'North').set('B2', 1250.5, { numFmt: '#,##0.00' });
 * const bytes = writeWorkbook(workbook); // Uint8Array with the .xlsx file
 * ```
 */
export function createWorkbook(options?: CreateWorkbookOptions): NewWorkbook {
  const o = readOptions(options, ['dateSystem'], 'workbook options');
  const dateSystem = o.get('dateSystem') ?? '1900';
  if (dateSystem !== '1900' && dateSystem !== '1904') {
    throw invalidOption(`Date system ${JSON.stringify(String(dateSystem).slice(0, 20))} is not valid.`, 'Use "1900" (the default) or "1904".');
  }
  return new NewWorkbook(dateSystem);
}

/**
 * Writes a workbook as an XLSX file (Office Open XML, Transitional) and returns its bytes. The
 * output is deterministic: the same workbook always gives the same bytes. Throws
 * `XLV_E_INVALID_INPUT` for a workbook without sheets or without a visible sheet.
 *
 * @example
 * ```ts
 * import { createWorkbook, writeWorkbook, readWorkbook } from 'xlvane';
 *
 * const workbook = createWorkbook();
 * workbook.addSheet('Data').set('A1', 42);
 * const bytes = writeWorkbook(workbook);
 * console.log(readWorkbook(bytes).getSheet('Data').cell('A1')?.value); // 42
 * ```
 */
export function writeWorkbook(workbook: NewWorkbook): Uint8Array {
  if (!(workbook instanceof NewWorkbook)) {
    throw new XlvError('XLV_E_INVALID_INPUT', 'writeWorkbook expects a workbook created with createWorkbook().', {
      hint: 'Create the workbook with createWorkbook(), add sheets with addSheet() and pass it to writeWorkbook().',
    });
  }
  const sheets = workbook.sheets;
  if (sheets.length === 0) {
    throw new XlvError('XLV_E_INVALID_INPUT', 'The workbook has no sheets.', {
      hint: 'Add at least one sheet with workbook.addSheet("Sheet1") before writing.',
    });
  }
  if (!sheets.some((s) => s.state === 'visible')) {
    throw new XlvError('XLV_E_INVALID_INPUT', 'All sheets of the workbook are hidden.', {
      hint: 'A workbook needs at least one visible sheet. Add one or create a sheet with { state: "visible" }.',
    });
  }
  return writePackage({
    dateSystem: workbook.dateSystem,
    sheets: sheets.map((s) => s.data()),
    styles: workbook.styles(),
  });
}
