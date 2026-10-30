// Building new workbooks in memory and writing them as XLSX (docs/decisions.md D-201 … D-210).
// Runtime-neutral: the result is a Uint8Array; `xlvane/node` adds `writeWorkbookFile`.
import { checkDefinedName, checkUniqueName, type DefinedName, type DefinedNameOptions, type NormalizedDefinedName } from './defined-names';
import { normalizeProperties, type WorkbookProperties } from './doc-properties';
import { XlvError } from './errors';
import { normalizeHyperlink, type HyperlinkTarget, type NormalizedHyperlink } from './hyperlinks';
import type { DateSystem, RichTextRun, SheetLayout, SheetState, SheetView, SplitPanes } from './model';
import { checkImage, metricsOf, normalizeImage, resolveImage, type ImageAnchorInput, type ImageFormat, type ImageOptions, type StoredImage } from './images';
import { MAX_AUTHOR_LENGTH, type NoteInput } from './notes';
import { MAX_DATA_VALIDATIONS, normalizeConditionalFormat, normalizeDataValidation, type ConditionalFormatRuleInput, type DataValidation, type DataValidationRule, type StoredConditionalFormat } from './sheet-rules';
import {
  EMPTY_PRINT_DATA,
  normalizeHeaderFooter,
  normalizePageBreaks,
  normalizePageSetup,
  normalizePrintArea,
  normalizePrintTitles,
  normalizeProtection,
  normalizeWorkbookProtection,
  printNames,
  type HeaderFooter,
  type PageBreaks,
  type PageSetupOptions,
  type PrintTitles,
  type SheetPrintData,
  type NormalizedWorkbookProtection,
  type SheetProtectionOptions,
  type WorkbookProtectionOptions,
} from './printing';
import { MAX_COLUMNS, MAX_ROWS, columnNumber, formatCellRef, formatRange, parseCellRef, parseRange, type CellRange } from './sml/cell-ref';
import { dateToSerial } from './sml/dates';
import { writePackage, type WriteSheetData } from './sml/write-parts';
import { mergeStyles, normalizeColor, normalizeStyle, stableKey, type CellStyle, type Color, type NormalizedColor, type NormalizedStyle } from './style';
import { normalizeTable, rangesOverlap, totalsFormula, type NormalizedTable, type TableInput } from './tables';
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
 * value, `{ formula, result }` for a formula with an optional cached result, or
 * `{ richText: [{ text, font? }, …] }` for text whose parts have their own fonts.
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
  | { readonly formula: string; readonly result?: FormulaResult }
  | { readonly richText: readonly RichTextRun[] };

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
  /** Document properties (title, creator, dates …), written to `docProps/core.xml` when given. */
  readonly properties?: WorkbookProperties;
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
 * const wide: ColumnOptions = { width: 40, style: { numFmt: '#,##0.00' } };
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
  /**
   * Default style of the column, written as `col@style` (the column's cell format; [MS-OI29500]
   * §2.1.598 d). xlvane does not apply it to the column's cells: style cells with `set` or
   * `style`. Applications such as Excel use it for cells added to the column later. `{}` removes
   * it.
   */
  readonly style?: CellStyle;
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
 * Options for {@link NewSheet.view}: how the sheet is displayed when opened (`sheetView`,
 * ECMA-376 Part 1 §18.3.1.87; W2-51). Settings not given keep their current value.
 *
 * @example
 * ```ts
 * import { createWorkbook, type SheetViewOptions } from 'xlvane';
 *
 * const view: SheetViewOptions = { zoom: 80, showGridLines: false };
 * createWorkbook().addSheet('Report').view(view);
 * ```
 */
export interface SheetViewOptions {
  /** Zoom in percent, a whole number from 10 to 400. Default 100. */
  readonly zoom?: number;
  /** Default true. */
  readonly showGridLines?: boolean;
  /** Show row numbers and column letters. Default true. */
  readonly showRowColHeaders?: boolean;
  /** Column A at the right and text right to left. Default false. */
  readonly rightToLeft?: boolean;
  /** Default "normal". */
  readonly layout?: SheetLayout;
  /**
   * Splits the window (without freezing) at `x` points from the left and `y` points from the
   * top; a missing one is 0, but not both. `null` removes the split. A split replaces frozen
   * panes, and {@link NewSheet.freeze} replaces a split.
   */
  readonly split?: Partial<SplitPanes> | null;
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
/** The view of a new sheet: the defaults of CT_SheetView (W2-51). */
const DEFAULT_VIEW: SheetView = { zoom: 100, showGridLines: true, showRowColHeaders: true, rightToLeft: false, layout: 'normal', split: null };
/** Largest split position accepted, in points: an implementation bound, not a documented limit. */
const MAX_SPLIT_POINTS = 1_000_000;
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
  | { readonly kind: 'string'; readonly s: string; readonly runs?: readonly StoredRun[] }
  | { readonly kind: 'boolean'; readonly b: boolean }
  | { readonly kind: 'error'; readonly e: ErrorValue }
  | {
      readonly kind: 'formula';
      readonly text: string;
      readonly result: Exclude<StoredValue, { kind: 'formula' }> | null;
    };

/** A rich text run after validation: its text and its font with only known keys. */
export interface StoredRun {
  readonly text: string;
  readonly font?: NonNullable<NormalizedStyle['font']>;
}

/** Most runs in one rich text value ([MS-OI29500] §2.1.675, F-SST-09). */
const MAX_RICH_TEXT_RUNS = 32_767;

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
  readonly columnMap: Map<number, { width: number | null; hidden: boolean; style: number }> = new Map();
  /** @internal */
  readonly rowMap: Map<number, { height: number | null; hidden: boolean }> = new Map();
  /** @internal */
  readonly mergeList: CellRange[] = [];
  /** @internal */
  frozen: { rows: number; cols: number } | null = null;
  /** @internal */
  filter: CellRange | null = null;
  /** @internal */
  tab: NormalizedColor | null = null;
  /** @internal */
  viewSettings: SheetView = DEFAULT_VIEW;
  /** @internal Hyperlinks by normalised range, in the order they were added. */
  readonly linkMap: Map<string, NormalizedHyperlink> = new Map();
  /** @internal Notes by cell reference. */
  readonly noteMap: Map<string, StoredNote> = new Map();
  /** @internal Pictures in the order they were added (TD-35). */
  readonly imageList: StoredImage[] = [];
  /** @internal Background image (index of `NewWorkbook.addImage`), or `null`. */
  backgroundImage: number | null = null;
  /** @internal Data validations in the order they were added (SR-01). */
  readonly validationList: DataValidation[] = [];
  /** @internal Conditional formats in the order they were added (SR-07). */
  readonly conditionalList: StoredConditionalFormat[] = [];
  /** Priorities of the conditional formatting rules of the sheet (unique, SR-10). */
  private readonly priorities = new Set<number>();
  /** @internal Protection, page setup, headers and footers, page breaks, print area and titles. */
  print: SheetPrintData = EMPTY_PRINT_DATA;
  /** @internal Tables in the order they were added (TD-01). */
  readonly tableList: NormalizedTable[] = [];
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
   * Sets the width, visibility or default style of a column, given by letters (`"C"`) or number
   * (3). Options not given keep their current value.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * createWorkbook().addSheet('Data').column('A', { width: 30 }).column(2, { hidden: true, style: { font: { bold: true } } });
   * ```
   */
  column(col: string | number, options: ColumnOptions): this {
    const index = typeof col === 'number' ? col : typeof col === 'string' ? (columnNumber(col.toUpperCase()) ?? 0) : 0;
    if (!Number.isInteger(index) || index < 1 || index > MAX_COLUMNS) throw invalidRef(col, this.name, 'column', '"C" or 3');
    const o = readOptions(options, ['width', 'hidden', 'style'], 'column options');
    const current = this.columnMap.get(index) ?? { width: null, hidden: false, style: 0 };
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
    const style = o.get('style');
    if (style !== undefined) current.style = this.book.intern(normalizeStyle(style));
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
    for (const t of this.tableList) {
      if (r.start.row <= t.range.end.row && t.range.start.row <= r.end.row && r.start.col <= t.range.end.col && t.range.start.col <= r.end.col) {
        throw new XlvError('XLV_E_INVALID_TABLE', `The range ${range} overlaps table "${t.name}".`, { location: { sheet: this.name }, hint: 'Tables cannot contain merged cells; merge cells outside the table.', details: { name: t.name } });
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
    if (this.frozen !== null && this.viewSettings.split !== null) this.viewSettings = { ...this.viewSettings, split: null };
    return this;
  }

  /**
   * Sets the colour of the sheet tab (`sheetPr/tabColor`, W2-50), or removes it with `null`.
   * Colours are given as for cell styles.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * createWorkbook().addSheet('Summary').tabColor('00B050');
   * ```
   */
  tabColor(color: Color | null): this {
    this.tab = color === null ? null : normalizeColor(color, 'tabColor');
    return this;
  }

  /**
   * Sets how the sheet is displayed when opened: zoom, gridlines, headings, direction, layout
   * and a split window (W2-51). Settings not given keep their value.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * createWorkbook().addSheet('Report').view({ zoom: 120, showGridLines: false });
   * ```
   */
  view(options: SheetViewOptions): this {
    const o = readOptions(options, ['zoom', 'showGridLines', 'showRowColHeaders', 'rightToLeft', 'layout', 'split'], 'view options');
    const next: { -readonly [K in keyof SheetView]: SheetView[K] } = { ...this.viewSettings };
    const zoom = o.get('zoom');
    if (zoom !== undefined) {
      if (typeof zoom !== 'number' || !Number.isInteger(zoom) || zoom < 10 || zoom > 400) {
        throw invalidOption('The view option "zoom" must be a whole number from 10 to 400 (percent).', 'For example { zoom: 80 }.');
      }
      next.zoom = zoom;
    }
    for (const key of ['showGridLines', 'showRowColHeaders', 'rightToLeft'] as const) {
      const v = o.get(key);
      if (v === undefined) continue;
      if (typeof v !== 'boolean') throw invalidOption(`The view option "${key}" must be true or false.`, `For example { ${key}: false }.`);
      next[key] = v;
    }
    const layout = o.get('layout');
    if (layout !== undefined) {
      if (layout !== 'normal' && layout !== 'pageBreakPreview' && layout !== 'pageLayout') {
        throw invalidOption(`The view option "layout" is ${JSON.stringify(String(layout).slice(0, 30))}.`, 'Use "normal", "pageBreakPreview" or "pageLayout".');
      }
      next.layout = layout;
    }
    let unfreeze = false;
    if (o.has('split')) {
      const split = o.get('split');
      if (split === null) next.split = null;
      else {
        const p = readOptions(split, ['x', 'y'], 'view option "split"');
        const at = (key: string): number => {
          const v = p.get(key) ?? 0;
          if (typeof v !== 'number' || !Number.isFinite(v) || v < 0 || v > MAX_SPLIT_POINTS) {
            throw invalidOption(`The split position "${key}" must be a number of points from 0 to ${String(MAX_SPLIT_POINTS)}.`, 'For example { split: { x: 100, y: 50 } }.');
          }
          return v;
        };
        const x = at('x');
        const y = at('y');
        // Office does not support a split pane at 0, 0 ([MS-OI29500] §2.1.640 e).
        if (x === 0 && y === 0) throw invalidOption('A split needs "x" or "y" greater than 0.', 'For example { split: { y: 50 } }, or { split: null } to remove it.');
        next.split = { x, y };
        unfreeze = true;
      }
    }
    this.viewSettings = next;
    if (unfreeze) this.frozen = null;
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
    const filter = range === null ? null : this.range(range);
    const table = filter === null ? undefined : this.tableList.find((t) => rangesOverlap(t.range, filter));
    if (filter !== null && table !== undefined) {
      throw new XlvError('XLV_E_INVALID_OPTION', `The filter ${formatRange(filter)} overlaps table "${table.name}" (${formatRange(table.range)}).`, {
        location: { sheet: this.name },
        hint: 'A table has its own filter, and the sheet filter must not share cells with a table ([MS-OI29500] §2.1.680 h); filter a range outside the table.',
        details: { option: 'autoFilter', table: table.name },
      });
    }
    this.filter = filter;
    return this;
  }

  /**
   * Attaches a hyperlink to a cell or range, or removes it with `null`. `{ url }` links to a web
   * page, `mailto:` address or file; `{ location }` to a place in the workbook such as
   * `"'Sales 2026'!A1"` or a defined name. The cell text is set separately with `set`; a link
   * added again for the same range replaces the earlier one. Throws `XLV_E_INVALID_HYPERLINK`
   * (see `HyperlinkTarget` for the limits). xlvane never opens or checks the target.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * const sheet = createWorkbook().addSheet('Links');
   * sheet.set('A1', 'Example').link('A1', { url: 'https://example.com', tooltip: 'Open example.com' });
   * sheet.set('A2', 'Back to the top').link('A2', { location: 'Links!A1' });
   * ```
   */
  link(range: string, target: HyperlinkTarget | null): this {
    const r = this.range(range, 'cell or range', '"A1" or "A1:B2"');
    const ref = r.start.row === r.end.row && r.start.col === r.end.col ? formatCellRef(r.start.row, r.start.col) : `${formatCellRef(r.start.row, r.start.col)}:${formatCellRef(r.end.row, r.end.col)}`;
    this.linkMap.delete(ref);
    if (target !== null) this.linkMap.set(ref, { ref, ...normalizeHyperlink(target, { sheet: this.name, cell: ref }) });
    return this;
  }

  /**
   * Attaches a note (comment) to a cell, or removes it with `null`. The note is plain text or
   * `{ text | runs, author? }` (see `NoteInput` for the limits); the cell does not need a value.
   * Notes are hidden until the pointer rests on the cell, as in spreadsheet applications. Throws
   * `XLV_E_INVALID_CELL_REF` or `XLV_E_INVALID_VALUE`.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * const sheet = createWorkbook().addSheet('Budget');
   * sheet.set('B2', 1250).note('B2', 'Estimate from March');
   * sheet.note('C2', { runs: [{ text: 'Ann: ', font: { bold: true } }, { text: 'check the total' }], author: 'Ann' });
   * ```
   */
  note(ref: string, note: NoteInput | null): this {
    const at = this.address(ref);
    this.noteMap.delete(at.ref);
    if (note !== null) this.noteMap.set(at.ref, { row: at.row, col: at.col, ...normalizeNote(note, { sheet: this.name, dateSystem: this.book.dateSystem }, at.ref) });
    return this;
  }

  /**
   * Protects the sheet (`sheetProtection`, SR-20), or removes the protection with `null`. Each
   * flag is `true` when the action is locked and `false` when it stays allowed; flags not given
   * keep the schema defaults (formatting, inserting, deleting, sorting, autofilter and pivot
   * tables locked; selecting cells, objects and scenarios allowed). A password is not supported
   * yet and throws `XLV_E_NOT_SUPPORTED`.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * createWorkbook().addSheet('Form').set('A1', 'Name').protect({ sort: false, autoFilter: false });
   * ```
   */
  protect(options: SheetProtectionOptions | null = {}): this {
    this.print = { ...this.print, protection: options === null ? null : normalizeProtection(options, this.name) };
    return this;
  }

  /**
   * Sets page setup for printing: margins, orientation, paper size, scale or fit to pages, page
   * order, print options and more (SR-23 … SR-26). Settings not given keep their value. Throws
   * `XLV_E_INVALID_OPTION` for values outside the limits of `PageSetupOptions`.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * createWorkbook().addSheet('Report').pageSetup({
   *   orientation: 'landscape',
   *   paperSize: 9,
   *   margins: { left: 0.5, right: 0.5 },
   *   fitToPage: true,
   *   fitToWidth: 1,
   *   fitToHeight: 0,
   *   printOptions: { gridLines: true },
   * });
   * ```
   */
  pageSetup(options: PageSetupOptions): this {
    this.print = { ...this.print, pageSetup: normalizePageSetup(options, this.print.pageSetup), pageSetupSet: true };
    return this;
  }

  /**
   * Sets printed headers and footers (SR-27), or removes them with `null`. Texts use the
   * formatting codes of `HeaderFooter`; settings not given keep their value.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * createWorkbook().addSheet('Report').headerFooter({ oddHeader: '&C&A', oddFooter: '&RPage &P of &N' });
   * ```
   */
  headerFooter(options: HeaderFooter | null): this {
    this.print = { ...this.print, headerFooter: options === null ? null : normalizeHeaderFooter(options, this.print.headerFooter) };
    return this;
  }

  /**
   * Sets the manual page breaks (SR-28): the rows and columns that start a new printed page.
   * Replaces earlier breaks; `{}` removes them.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * createWorkbook().addSheet('Report').pageBreaks({ rows: [41, 81] });
   * ```
   */
  pageBreaks(breaks: PageBreaks): this {
    const { rows, cols } = normalizePageBreaks(breaks);
    this.print = { ...this.print, rowBreaks: rows, colBreaks: cols };
    return this;
  }

  /**
   * Sets the print area, one range or several, or removes it with `null`. It is stored as the
   * defined name `_xlnm.Print_Area` of the sheet (SR-30).
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * createWorkbook().addSheet('Report').printArea('A1:H40');
   * ```
   */
  printArea(ranges: string | readonly string[] | null): this {
    this.print = { ...this.print, printArea: normalizePrintArea(ranges) };
    return this;
  }

  /**
   * Sets rows and columns repeated on every printed page, or removes them with `null`. They are
   * stored as the defined name `_xlnm.Print_Titles` of the sheet (SR-30).
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * createWorkbook().addSheet('Report').printTitles({ rows: '1:2', columns: 'A' });
   * ```
   */
  printTitles(titles: PrintTitles | null): this {
    this.print = { ...this.print, printTitles: normalizePrintTitles(titles) };
    return this;
  }

  /**
   * Adds a data validation to a cell or ranges (`"B2:B50"`, or several separated by spaces), or
   * with `null` removes the validations added for exactly that range. Each call writes one
   * `dataValidation` element. See `DataValidationRule` for the rules and the limits; throws
   * `XLV_E_INVALID_OPTION` or `XLV_E_INVALID_CELL_REF`.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * const sheet = createWorkbook().addSheet('Orders');
   * sheet.validate('C2:C500', { type: 'list', formula1: '"Open,Shipped,Paid"', showErrorMessage: true, error: 'Pick a status from the list.' });
   * sheet.validate('D2:D500', { type: 'whole', operator: 'greaterThan', formula1: 0, showInputMessage: true, prompt: 'Quantity, at least 1' });
   * ```
   */
  validate(sqref: string, rule: DataValidationRule | null): this {
    const where = { sheet: this.name, dateSystem: this.book.dateSystem };
    if (rule === null) {
      const range = normalizeDataValidation(sqref, { type: 'none' }, where).sqref;
      for (let i = this.validationList.length - 1; i >= 0; i--) if (this.validationList[i]?.sqref === range) this.validationList.splice(i, 1);
      return this;
    }
    const dv = normalizeDataValidation(sqref, rule, where);
    if (this.validationList.length >= MAX_DATA_VALIDATIONS) {
      throw new XlvError('XLV_E_LIMIT_EXCEEDED', `The sheet "${this.name}" already has ${String(MAX_DATA_VALIDATIONS)} data validations, the most Office accepts.`, {
        location: { sheet: this.name },
        hint: 'Give one validation a list of ranges ("A1:A9 C1:C9") instead of adding one per range.',
        details: { limit: 'dataValidations', value: MAX_DATA_VALIDATIONS },
      });
    }
    this.validationList.push(dv);
    return this;
  }

  /**
   * Adds conditional formatting to a cell or ranges: one or more rules, each with the format
   * applied where it is true. Rules get unique priorities on the sheet (1 is applied first). See
   * `ConditionalFormatRuleInput` for the rule types; throws `XLV_E_INVALID_OPTION`,
   * `XLV_E_INVALID_STYLE` or `XLV_E_INVALID_CELL_REF`.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * const sheet = createWorkbook().addSheet('Sales');
   * sheet.conditionalFormat('B2:B100', [
   *   { type: 'cellIs', operator: 'lessThan', formulas: [0], style: { font: { color: '9C0006' }, fill: { pattern: 'solid', fgColor: 'FFC7CE' } } },
   *   { type: 'dataBar', dataBar: { cfvo: [{ type: 'min' }, { type: 'max' }], color: '638EC6' } },
   * ]);
   * ```
   */
  conditionalFormat(sqref: string, rules: readonly ConditionalFormatRuleInput[]): this {
    const cf = normalizeConditionalFormat(sqref, rules, { sheet: this.name, dateSystem: this.book.dateSystem }, this.priorities);
    this.conditionalList.push({ sqref: cf.sqref, rules: cf.rules, dxfIds: cf.rules.map((r) => (r.style === null ? null : this.book.internDxf(r.style))) });
    return this;
  }


  /**
   * Adds a table (Part 1 §18.5): writes the header row (the column names), the data rows and the
   * totals row into the cells, and the table itself. `ref` is the top-left cell or the whole
   * range. Throws `XLV_E_INVALID_TABLE` for invalid or duplicate names (unique in the workbook,
   * also against defined names), duplicate column names, or a range that overlaps another table
   * or a merged range; see `TableInput`.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * createWorkbook().addSheet('Data').addTable({
   *   name: 'Sales',
   *   ref: 'A1',
   *   totalsRow: true,
   *   style: { theme: 'TableStyleMedium9', showRowStripes: true },
   *   columns: [{ name: 'Region', totalsRowLabel: 'Total' }, { name: 'Revenue', totalsRowFunction: 'sum' }],
   *   rows: [['North', 1200], ['South', 980]],
   * });
   * ```
   */
  addTable(table: TableInput): this {
    const t = normalizeTable(table, this.name);
    this.book.checkTable(t, this);
    const { start, end } = t.range;
    let row = start.row;
    if (t.headerRow) {
      t.columns.forEach((c, i) => this.set(formatCellRef(row, start.col + i), c.name));
      row++;
    }
    const dataEnd = end.row - (t.totalsRow ? 1 : 0);
    for (let r = row; r <= dataEnd; r++) {
      const values = t.rows[r - row] ?? [];
      t.columns.forEach((c, i) => {
        const value = values[i];
        const ref = formatCellRef(r, start.col + i);
        if (value !== undefined) this.set(ref, value as CellInput);
        else if (c.calculatedColumnFormula !== null) this.set(ref, { formula: c.calculatedColumnFormula });
      });
    }
    if (t.totalsRow) {
      t.columns.forEach((c, i) => {
        const ref = formatCellRef(end.row, start.col + i);
        const f = totalsFormula(t.displayName, c);
        if (f !== null) this.set(ref, { formula: f });
        else if (c.totalsRowLabel !== null) this.set(ref, c.totalsRowLabel);
      });
    }
    this.tableList.push(t);
    return this;
  }

  /**
   * Places a picture on the sheet. `image` is the number `NewWorkbook.addImage` returned; the same
   * image can be placed many times and is stored once. `anchor` is a range such as `"B2:D6"`,
   * two corners `{ from, to }` (with `editAs`), or a corner and a size in pixels
   * (`{ from, size }`), see `ImageAnchorInput`. Fractional positions and pixel sizes use the
   * sheet's column widths and row heights at the time the workbook is written, at 96 pixels per
   * inch (TD-33).
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * declare const png: Uint8Array;
   * const workbook = createWorkbook();
   * const logo = workbook.addImage(png);
   * workbook.addSheet('Report').addImage(logo, 'B2:D6', { description: 'Company logo' });
   * ```
   */
  addImage(image: number, anchor: ImageAnchorInput, options?: ImageOptions): this {
    this.imageList.push(normalizeImage(image, this.book.imageCount, anchor, options));
    return this;
  }

  /**
   * Sets the picture shown behind the cells of the sheet (`picture` element, Part 1 §18.3.1.67),
   * or removes it with `null`. `image` is the number `NewWorkbook.addImage` returned.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * declare const png: Uint8Array;
   * const workbook = createWorkbook();
   * workbook.addSheet('Cover').background(workbook.addImage(png));
   * ```
   */
  background(image: number | null): this {
    if (image !== null) normalizeImage(image, this.book.imageCount, 'A1', undefined);
    this.backgroundImage = image;
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
      tabColor: this.tab,
      view: this.viewSettings,
      hyperlinks: [...this.linkMap.values()],
      notes: sortedNotes(this.noteMap),
      conditionalFormats: this.conditionalList,
      dataValidations: this.validationList,
      print: this.print,
      tables: this.tableList,
      images: this.imageList.map((i) => resolveImage(i, metricsOf(this.columnMap, this.rowMap))),
      background: this.backgroundImage,
    };
  }
}

/** A validated note: its cell, its text (plain or rich) and its author (`""` when not given). */
export interface StoredNote {
  readonly row: number;
  readonly col: number;
  readonly value: Extract<StoredValue, { kind: 'string' }>;
  readonly author: string;
}

/** The notes of a sheet in row, then column order (deterministic output). */
export function sortedNotes(notes: ReadonlyMap<string, StoredNote>): StoredNote[] {
  return [...notes.values()].sort((a, b) => a.row - b.row || a.col - b.col);
}

/**
 * Validates a {@link NoteInput}: text as a cell string, runs as `{ richText }`, the author up to
 * 65 535 characters ([MS-OI29500] §2.1.689 a). Throws `XLV_E_INVALID_VALUE` naming the cell.
 */
function normalizeNote(note: unknown, ctx: ConvertContext, cell: string): Omit<StoredNote, 'row' | 'col'> {
  const usage = 'Pass the text, or { text: "…", author: "Ann" }, or { runs: [{ text: "…", font: { bold: true } }] }.';
  let input: unknown = note;
  let author = '';
  if (typeof note === 'object' && note !== null && !Array.isArray(note)) {
    const keys = Object.keys(note);
    const o = note as { text?: unknown; runs?: unknown; author?: unknown };
    if (keys.some((k) => k !== 'text' && k !== 'runs' && k !== 'author') || (o.text === undefined) === (o.runs === undefined)) {
      throw invalidValue(ctx.sheet, cell, `The note for ${cell} must have exactly one of "text" and "runs", and optionally "author".`, usage);
    }
    if (o.author !== undefined) {
      if (typeof o.author !== 'string' || o.author.length > MAX_AUTHOR_LENGTH || firstLoneSurrogate(o.author) !== -1) {
        throw invalidValue(ctx.sheet, cell, `The author of the note for ${cell} must be text of at most ${String(MAX_AUTHOR_LENGTH)} characters.`, 'Pass the author as a string, for example author: "Ann".');
      }
      author = o.author;
    }
    input = o.text !== undefined ? o.text : { richText: o.runs };
    if (o.text !== undefined && typeof o.text !== 'string') throw invalidValue(ctx.sheet, cell, `The text of the note for ${cell} must be a string.`, usage);
  } else if (typeof note !== 'string') {
    throw invalidValue(ctx.sheet, cell, `The note for ${cell} must be text or an object.`, usage);
  }
  let stored: StoredValue | null;
  try {
    stored = convertCellInput(input, ctx, cell).stored;
  } catch (error) {
    if (!(error instanceof XlvError)) throw error;
    throw new XlvError(error.code, `The note for ${cell}: ${error.message}`, { location: { sheet: ctx.sheet, cell }, hint: error.hint, details: error.details });
  }
  return { value: stored?.kind === 'string' ? stored : { kind: 'string', s: '' }, author };
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
    if (keys.includes('richText')) return { stored: readRichText(value, ctx.sheet, cell), isDate: false, hasTime: false };
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
    'Use a number, string, boolean, Date, null, { error: "#N/A" }, { formula: "A1*2", result: 2 } or { richText: [{ text: "a" }] }.',
  );
}

/**
 * Validates `{ richText: [{ text, font? }, …] }` (F-SST-08, F-SST-09): at least one run, at most
 * 32,767 runs and 32,767 characters in all, fonts as in cell styles. Runs with the same font next
 * to each other are kept apart, as given.
 */
function readRichText(value: object, sheet: string, cell: string): Extract<StoredValue, { kind: 'string' }> {
  const usage = 'Use { richText: [{ text: "Total: ", font: { bold: true } }, { text: "42" }] }.';
  const keys = Object.keys(value);
  const runsIn = (value as { richText: unknown }).richText;
  if (keys.length !== 1 || !Array.isArray(runsIn) || runsIn.length === 0) {
    throw invalidValue(sheet, cell, `The rich text for ${cell} must be an object with only "richText", a non-empty array of runs.`, usage);
  }
  if (runsIn.length > MAX_RICH_TEXT_RUNS) {
    throw invalidValue(sheet, cell, `The rich text for ${cell} has ${String(runsIn.length)} runs.`, `A rich text value can have at most ${String(MAX_RICH_TEXT_RUNS)} runs. Join runs that have the same font.`);
  }
  const runs: StoredRun[] = [];
  let text = '';
  for (let i = 0; i < runsIn.length; i++) {
    const run: unknown = runsIn[i];
    const runKeys = typeof run === 'object' && run !== null && !Array.isArray(run) ? Object.keys(run) : null;
    const t = runKeys === null ? undefined : (run as { text: unknown }).text;
    if (runKeys === null || runKeys.some((k) => k !== 'text' && k !== 'font') || typeof t !== 'string') {
      throw invalidValue(sheet, cell, `Run ${String(i)} of the rich text for ${cell} is not { text, font? } with a string text.`, usage);
    }
    if (firstLoneSurrogate(t) !== -1) {
      throw invalidValue(sheet, cell, `Run ${String(i)} of the rich text for ${cell} contains an unpaired UTF-16 surrogate.`, 'The text is not valid Unicode; it may have been cut in the middle of a character such as an emoji. Fix the text before writing it.');
    }
    text += t;
    const fontIn = runKeys.includes('font') ? (run as { font: unknown }).font : undefined;
    let font: StoredRun['font'];
    if (fontIn !== undefined) {
      try {
        font = normalizeStyle({ font: fontIn }).font;
      } catch (error) {
        if (!(error instanceof XlvError)) throw error;
        throw new XlvError(error.code, `Run ${String(i)} of the rich text for ${cell}: ${error.message}`, {
          location: { sheet, cell },
          hint: error.hint,
          details: { ...error.details, run: i },
        });
      }
    }
    runs.push(font === undefined ? { text: t } : { text: t, font });
  }
  if (text.length > MAX_CELL_TEXT) {
    throw invalidValue(sheet, cell, `The rich text for ${cell} has ${String(text.length)} characters.`, `A cell can contain at most ${String(MAX_CELL_TEXT)} characters. Split the text over several cells.`);
  }
  return { kind: 'string', s: text, runs };
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
  /** Document properties given to {@link createWorkbook}. */
  readonly properties: WorkbookProperties;
  private readonly list: NewSheet[] = [];
  private readonly names: NormalizedDefinedName[] = [];
  private readonly styleIds = new Map<string, number>();
  private readonly styleList: NormalizedStyle[] = [{}];
  /** @internal `workbookProtection`, set by {@link NewWorkbook.protect}. */
  workbookProtection: NormalizedWorkbookProtection | null = null;

  /** @internal Use {@link createWorkbook}. */
  constructor(dateSystem: DateSystem, properties: WorkbookProperties = {}) {
    this.dateSystem = dateSystem;
    this.properties = properties;
    this.styleIds.set(stableKey({}), 0);
  }

  /**
   * Protects the workbook (`workbookProtection`, SR-22), by default its structure: sheets cannot
   * be added, deleted, moved or renamed in a spreadsheet application without the password.
   * `null` removes the protection. A password string is hashed with `hashPassword` defaults
   * (D-C08); it is not encryption, and the content stays readable.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * const workbook = createWorkbook().protect({ password: 'secret' });
   * workbook.addSheet('Data').set('A1', 1);
   * ```
   */
  protect(options: WorkbookProtectionOptions | null = {}): this {
    this.workbookProtection = options === null ? null : normalizeWorkbookProtection(options);
    return this;
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

  /**
   * Defines a name for a cell, range, constant or formula. `refersTo` is formula text with or
   * without `=`, for example `"Data!$A$2:$A$100"`; it is stored as given. A name is local to a
   * sheet when `options.sheet` is set. Throws `XLV_E_INVALID_DEFINED_NAME` for a name Excel does
   * not accept or a duplicate in the same scope (names are not case-sensitive), and
   * `XLV_E_SHEET_NOT_FOUND` for an unknown sheet.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * const workbook = createWorkbook();
   * workbook.addSheet('Data').set('A1', 0.19);
   * workbook.defineName('VatRate', 'Data!$A$1');
   * workbook.defineName('Rows', 'Data!$A$2:$A$100', { sheet: 'Data', hidden: true });
   * ```
   */
  defineName(name: string, refersTo: string, options?: DefinedNameOptions): this {
    const n = checkDefinedName(name, refersTo, options, this.list.map((s) => s.name));
    checkUniqueName(n, this.names);
    const table = this.list.flatMap((s) => s.tableList).find((t) => t.name.toLowerCase() === n.name.toLowerCase() || t.displayName.toLowerCase() === n.name.toLowerCase());
    if (table !== undefined) {
      throw new XlvError('XLV_E_INVALID_DEFINED_NAME', `"${n.name}" is already the name of a table.`, { hint: 'Table names and defined names share one list in a workbook (ECMA-376 Part 1 §18.5.1.2); choose another name.', details: { name: n.name } });
    }
    this.names.push(n);
    return this;
  }

  /**
   * @internal Checks that a table's names are unique in the workbook (tables and defined names,
   * ignoring case; Part 1 §18.5.1.2) and that it overlaps no table or merged range of its sheet.
   */
  checkTable(t: NormalizedTable, sheet: NewSheet): void {
    const where = { sheet: sheet.name };
    const names = new Set([t.name.toLowerCase(), t.displayName.toLowerCase()]);
    for (const other of this.list.flatMap((s) => s.tableList)) {
      if (names.has(other.name.toLowerCase()) || names.has(other.displayName.toLowerCase())) {
        throw new XlvError('XLV_E_INVALID_TABLE', `The workbook already has a table named "${other.displayName}".`, { location: where, hint: 'Table names are unique in a workbook, ignoring case; choose another name, for example "Sales_2".', details: { name: t.name } });
      }
    }
    for (const n of this.names) {
      if (names.has(n.name.toLowerCase())) {
        throw new XlvError('XLV_E_INVALID_TABLE', `"${n.name}" is already a defined name of the workbook.`, { location: where, hint: 'Table names and defined names share one list in a workbook (ECMA-376 Part 1 §18.5.1.2); choose another table name.', details: { name: t.name } });
      }
    }
    const ref = (r: CellRange): string => `${formatCellRef(r.start.row, r.start.col)}:${formatCellRef(r.end.row, r.end.col)}`;
    for (const other of sheet.tableList) {
      if (rangesOverlap(t.range, other.range)) {
        throw new XlvError('XLV_E_INVALID_TABLE', `Table "${t.name}" (${ref(t.range)}) overlaps table "${other.name}" (${ref(other.range)}).`, { location: where, hint: 'Tables must not share cells ([MS-OI29500] §2.1.680 h); move one of them.', details: { name: t.name } });
      }
    }
    if (sheet.filter !== null && rangesOverlap(t.range, sheet.filter)) {
      throw new XlvError('XLV_E_INVALID_TABLE', `Table "${t.name}" (${ref(t.range)}) overlaps the sheet filter ${ref(sheet.filter)}.`, { location: where, hint: 'A table has its own filter, and the sheet filter must not share cells with a table ([MS-OI29500] §2.1.680 h); remove the sheet filter (autoFilter(null)) or move the table.', details: { name: t.name } });
    }
    for (const m of sheet.mergeList) {
      if (rangesOverlap(t.range, m)) {
        throw new XlvError('XLV_E_INVALID_TABLE', `Table "${t.name}" (${ref(t.range)}) overlaps the merged range ${ref(m)}.`, { location: where, hint: 'Tables cannot contain merged cells; unmerge the range or move the table.', details: { name: t.name } });
      }
    }
  }

  /** Defined names in the order they were added. */
  get definedNames(): readonly DefinedName[] {
    return [...this.names];
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

  private readonly dxfIds = new Map<string, number>();
  private readonly dxfList: NormalizedStyle[] = [];

  /** @internal Index of a differential format (conditional formatting), deduplicated (SR-16). */
  internDxf(style: NormalizedStyle): number {
    const key = stableKey(style);
    let id = this.dxfIds.get(key);
    if (id === undefined) {
      id = this.dxfList.length;
      this.dxfList.push(style);
      this.dxfIds.set(key, id);
    }
    return id;
  }

  /** @internal */
  dxfStyles(): readonly NormalizedStyle[] {
    return this.dxfList;
  }

  private readonly imageData: { readonly bytes: Uint8Array; readonly format: ImageFormat }[] = [];

  /**
   * Adds an image (the bytes of a PNG, JPEG or GIF file, recognised from its first bytes) and
   * returns its number for `NewSheet.addImage` and `NewSheet.background`. Other formats throw
   * `XLV_E_INVALID_IMAGE`. Identical images are stored once in the file.
   *
   * @example
   * ```ts
   * import { createWorkbook } from 'xlvane';
   *
   * declare const jpeg: Uint8Array;
   * const workbook = createWorkbook();
   * const photo = workbook.addImage(jpeg);
   * workbook.addSheet('Team').addImage(photo, { from: 'A1', size: { width: 120, height: 160 } });
   * ```
   */
  addImage(bytes: Uint8Array): number {
    this.imageData.push(checkImage(bytes));
    return this.imageData.length - 1;
  }

  /** @internal */
  get imageCount(): number {
    return this.imageData.length;
  }

  /** @internal */
  imageList(): readonly { readonly bytes: Uint8Array; readonly format: ImageFormat }[] {
    return this.imageData;
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
  const o = readOptions(options, ['dateSystem', 'properties'], 'workbook options');
  const dateSystem = o.get('dateSystem') ?? '1900';
  if (dateSystem !== '1900' && dateSystem !== '1904') {
    throw invalidOption(`Date system ${JSON.stringify(String(dateSystem).slice(0, 20))} is not valid.`, 'Use "1900" (the default) or "1904".');
  }
  return new NewWorkbook(dateSystem, normalizeProperties(o.get('properties')));
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
    properties: workbook.properties,
    definedNames: [...workbook.definedNames, ...sheets.flatMap((s) => printNames(s.name, s.print))],
    dxfs: workbook.dxfStyles(),
    workbookProtection: workbook.workbookProtection,
    images: workbook.imageList(),
  });
}
