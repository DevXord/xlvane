// Public types of the read model. A workbook read by `readWorkbook` is immutable.
import type { XlvWarning } from './errors';
import type { CellStyle } from './style';

/** The kind of value a cell holds after reading. */
export type CellType = 'number' | 'string' | 'boolean' | 'error' | 'date' | 'empty';

/**
 * The value of a cell: a number, a string, a boolean, a `Date` (for date-formatted numbers and
 * ISO 8601 date cells), an error code such as `"#DIV/0!"` (when `type` is `"error"`), or `null`
 * for an empty cell.
 */
export type CellValue = number | string | boolean | Date | null;

/** How a formula is stored in the file (ECMA-376 Part 1, ST_CellFormulaType). */
export type FormulaKind = 'normal' | 'shared' | 'array' | 'dataTable';

/** A formula as stored in the file. xlvane does not evaluate formulas. */
export interface Formula {
  readonly kind: FormulaKind;
  /**
   * The formula text without the leading "=". `null` for cells that share a formula written
   * in another cell (see `sharedIndex`); expanding shared formulas is not implemented yet.
   */
  readonly text: string | null;
  /** The range the formula applies to, for array, data-table and shared master formulas. */
  readonly ref: string | null;
  /** Shared formula group index (`si`), for shared formulas. */
  readonly sharedIndex: number | null;
}

/** One cell read from a worksheet. */
export interface Cell {
  /** A1-style reference, for example `"B7"`. */
  readonly ref: string;
  /** 1-based row number. */
  readonly row: number;
  /** 1-based column number. */
  readonly col: number;
  readonly type: CellType;
  /** For formula cells, the cached result stored in the file. */
  readonly value: CellValue;
  readonly formula: Formula | null;
  /** Index into the workbook's cell formats (`s` attribute); 0 when absent. */
  readonly styleIndex: number;
  /** Number format code applied to the cell, or `null` for the General format. */
  readonly numFmt: string | null;
  /**
   * The cell's formatting: font, fill, border, alignment and number format. Parts that equal the
   * workbook defaults are left out, so an unformatted cell has `{}`. Colours are 8-digit ARGB
   * strings or theme/indexed objects. The object is frozen and shared by cells with the same
   * formatting.
   */
  readonly style: CellStyle;
}

/** A `col` definition of a worksheet: width and visibility of columns `min` to `max` (1-based). */
export interface ColumnInfo {
  readonly min: number;
  readonly max: number;
  /** Width in characters, or `null` if the file does not set one. */
  readonly width: number | null;
  readonly hidden: boolean;
}

/** Height and visibility of a row. */
export interface RowInfo {
  /** Height in points, or `null` if the file does not set one. */
  readonly height: number | null;
  readonly hidden: boolean;
}

/** Rows at the top and columns at the left that stay visible when scrolling. */
export interface FrozenPanes {
  readonly rows: number;
  readonly cols: number;
}

/** The kind of sheet, from the relationship that points at its part. */
export type SheetKind = 'worksheet' | 'chartsheet' | 'dialogsheet' | 'macrosheet';

/** Visibility of a sheet (ECMA-376 Part 1, ST_SheetState). */
export type SheetState = 'visible' | 'hidden' | 'veryHidden';

/** Which date system the workbook uses for serial date values. */
export type DateSystem = '1900' | '1904';

/** Conformance class of the package. */
export type Conformance = 'transitional' | 'strict';

/** File format, derived from the content type of the workbook part. */
export type WorkbookFormat = 'xlsx' | 'xlsm' | 'xltx' | 'xltm' | 'unknown';

/** A sheet of a workbook read by `readWorkbook`. */
export interface Sheet {
  readonly name: string;
  /** 0-based position in the workbook. */
  readonly index: number;
  readonly kind: SheetKind;
  readonly state: SheetState;
  /** Part name of the sheet, for example `/xl/worksheets/sheet1.xml`; `null` if it is missing. */
  readonly part: string | null;
  /** The used range declared in the file (`dimension` element), or `null`. */
  readonly dimension: string | null;
  /** Merged ranges, for example `["A1:C1"]`. */
  readonly merges: readonly string[];
  /** Number of cells read from the sheet. */
  readonly cellCount: number;
  /** Column definitions in file order. Columns without one have the default width. */
  readonly columns: readonly ColumnInfo[];
  /** Frozen rows and columns of the sheet's first view, or `null`. */
  readonly freeze: FrozenPanes | null;
  /** The autofilter range, for example `"A1:D100"`, or `null`. */
  readonly autoFilter: string | null;
  /** Height and visibility of a row (1-based), or `undefined` if the file sets neither. */
  rowInfo(row: number): RowInfo | undefined;
  /** Returns the cell at an A1-style reference, or `undefined` if the file has none there. */
  cell(ref: string): Cell | undefined;
  /** Rows in ascending order; each row lists its cells in ascending column order. */
  rows(): IterableIterator<readonly Cell[]>;
  /** All cells, row by row. */
  cells(): IterableIterator<Cell>;
}

/** A workbook read by `readWorkbook`. */
export interface Workbook {
  readonly sheets: readonly Sheet[];
  readonly dateSystem: DateSystem;
  readonly conformance: Conformance;
  readonly format: WorkbookFormat;
  /** `true` if the package contains a VBA project or Excel 4.0 macro sheets. Macros are never run. */
  readonly hasMacros: boolean;
  /** Number of external link parts (links to other workbooks). They are never followed. */
  readonly externalLinkCount: number;
  /** Warnings found while reading. */
  readonly warnings: readonly XlvWarning[];
  /** Returns the sheet with this name; throws `XLV_E_SHEET_NOT_FOUND` if there is none. */
  getSheet(name: string): Sheet;
}
