// Public types of the read model. A workbook read by `readWorkbook` is immutable.
import type { DefinedName } from './defined-names';
import type { WorkbookProperties } from './doc-properties';
import type { XlvWarning } from './errors';
import type { Hyperlink } from './hyperlinks';
import type { SheetImage, WorkbookImage } from './images';
import type { HeaderFooter, PageSetup, SheetProtection, WorkbookProtection } from './printing';
import type { Note } from './notes';
import type { ConditionalFormat, DataValidation } from './sheet-rules';
import type { TableDefinition } from './tables';
import type { CellStyle, Color, FontStyle } from './style';

/**
 * The kind of value a cell holds after reading.
 *
 * @example
 * ```ts
 * import { readWorkbook, type CellType } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const type: CellType | undefined = readWorkbook(bytes).sheets[0]?.cell('A1')?.type;
 * console.log(type === 'date' ? 'a date' : type);
 * ```
 */
export type CellType = 'number' | 'string' | 'boolean' | 'error' | 'date' | 'empty';

/**
 * The value of a cell: a number, a string, a boolean, a `Date` (for date-formatted numbers and
 * ISO 8601 date cells), an error code such as `"#DIV/0!"` (when `type` is `"error"`), or `null`
 * for an empty cell.
 *
 * @example
 * ```ts
 * import { readWorkbook, type CellValue } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const value: CellValue | undefined = readWorkbook(bytes).sheets[0]?.cell('B2')?.value;
 * if (value instanceof Date) console.log(value.toISOString());
 * ```
 */
export type CellValue = number | string | boolean | Date | null;

/**
 * One run of rich text: a piece of text with its own font (ECMA-376 Part 1 §18.4.4 `r`, F-SST-08).
 * A run without `font` uses the cell's font. Cells read from a file carry their runs in
 * `Cell.richText`; the writers accept them as `{ richText: [...] }`.
 *
 * @example
 * ```ts
 * import { createWorkbook, type RichTextRun } from 'xlvane';
 *
 * const runs: RichTextRun[] = [{ text: 'Total: ', font: { bold: true } }, { text: '1,250 EUR' }];
 * createWorkbook().addSheet('Report').set('A1', { richText: runs });
 * ```
 */
export interface RichTextRun {
  readonly text: string;
  readonly font?: FontStyle;
}

/**
 * How a formula is stored in the file (ECMA-376 Part 1, ST_CellFormulaType).
 *
 * @example
 * ```ts
 * import { readWorkbook, type FormulaKind } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const kind: FormulaKind | undefined = readWorkbook(bytes).sheets[0]?.cell('C1')?.formula?.kind;
 * console.log(kind === 'array' ? 'array formula' : kind);
 * ```
 */
export type FormulaKind = 'normal' | 'shared' | 'array' | 'dataTable';

/**
 * A formula as stored in the file. xlvane does not evaluate formulas.
 *
 * @example
 * ```ts
 * import { readWorkbook, type Formula } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const formula: Formula | null | undefined = readWorkbook(bytes).sheets[0]?.cell('C1')?.formula;
 * console.log(formula?.text); // e.g. "SUM(A1:B1)"
 * ```
 */
export interface Formula {
  readonly kind: FormulaKind;
  /**
   * The formula text without the leading "=". `null` for cells that share a formula written
   * in another cell (see `sharedIndex`) when xlvane could not translate the shared formula to this
   * cell (`XLV_W_SHARED_FORMULA_NOT_EXPANDED`); otherwise the translated text.
   */
  readonly text: string | null;
  /** The range the formula applies to, for array, data-table and shared master formulas. */
  readonly ref: string | null;
  /** Shared formula group index (`si`), for shared formulas. */
  readonly sharedIndex: number | null;
}

/**
 * One cell read from a worksheet.
 *
 * @example
 * ```ts
 * import { readWorkbook, type Cell } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const cells: Cell[] = [...(readWorkbook(bytes).sheets[0]?.cells() ?? [])];
 * for (const cell of cells) console.log(cell.ref, cell.type, cell.value, cell.numFmt);
 * ```
 */
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
  /**
   * For text stored as rich text (runs with their own fonts), the runs; `value` is then their
   * text joined. Absent for plain text and other values.
   */
  readonly richText?: readonly RichTextRun[];
  /**
   * The hyperlink whose range covers this cell, if any (the first one in file order). Not set
   * by the streaming reader, because the links follow the rows in the file (D-A08).
   */
  readonly hyperlink?: Hyperlink;
  /**
   * The note (comment) of this cell, if any. Notes of cells without a value are listed only in
   * `Sheet.notes`. The streaming reader sets it on the cells it returns.
   */
  readonly note?: Note;
}

/**
 * A `col` definition of a worksheet: width and visibility of columns `min` to `max` (1-based).
 *
 * @example
 * ```ts
 * import { readWorkbook, type ColumnInfo } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const columns: readonly ColumnInfo[] = readWorkbook(bytes).sheets[0]?.columns ?? [];
 * console.log(columns.filter((column) => column.hidden).map((column) => [column.min, column.max]));
 * ```
 */
export interface ColumnInfo {
  readonly min: number;
  readonly max: number;
  /** Width in characters, or `null` if the file does not set one. */
  readonly width: number | null;
  readonly hidden: boolean;
  /**
   * The column's default style (`col@style`), resolved like `Cell.style`; present only when the
   * file gives the column a style. Applications use it for cells added to the column later;
   * cells already in the file carry their own style.
   */
  readonly style?: CellStyle;
  /** Outline level 1 to 7; absent for 0. */
  readonly outlineLevel?: number;
  /** Present and `true` when the column's outline group is collapsed. */
  readonly collapsed?: boolean;
}

/**
 * Height and visibility of a row.
 *
 * @example
 * ```ts
 * import { readWorkbook, type RowInfo } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const first: RowInfo | undefined = readWorkbook(bytes).sheets[0]?.rowInfo(1);
 * console.log(first?.height, first?.hidden);
 * ```
 */
export interface RowInfo {
  /** Height in points, or `null` if the file does not set one. */
  readonly height: number | null;
  readonly hidden: boolean;
  /** Outline level 1 to 7; absent for 0. */
  readonly outlineLevel?: number;
  /** Present and `true` when the row's outline group is collapsed. */
  readonly collapsed?: boolean;
}

/**
 * Rows at the top and columns at the left that stay visible when scrolling.
 *
 * @example
 * ```ts
 * import { readWorkbook, type FrozenPanes } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const frozen: FrozenPanes | null | undefined = readWorkbook(bytes).sheets[0]?.freeze;
 * console.log(frozen ? `${String(frozen.rows)} rows, ${String(frozen.cols)} columns` : 'none');
 * ```
 */
export interface FrozenPanes {
  readonly rows: number;
  readonly cols: number;
}

/**
 * How a sheet is displayed: normal, with page breaks, or as printed pages (ST_SheetViewType,
 * ECMA-376 Part 1 §18.18.69).
 *
 * @example
 * ```ts
 * import type { SheetLayout } from 'xlvane';
 *
 * const layout: SheetLayout = 'pageLayout';
 * console.log(layout);
 * ```
 */
export type SheetLayout = 'normal' | 'pageBreakPreview' | 'pageLayout';

/**
 * A split window (not frozen): the distance of the vertical split bar from the left and of the
 * horizontal split bar from the top, in points; 0 means no split in that direction. Stored as
 * `pane@xSplit` and `pane@ySplit` in 1/20 point (ECMA-376 Part 1 §18.3.1.66, W2-51).
 *
 * @example
 * ```ts
 * import type { SplitPanes } from 'xlvane';
 *
 * const split: SplitPanes = { x: 120, y: 60 };
 * console.log(split.x);
 * ```
 */
export interface SplitPanes {
  readonly x: number;
  readonly y: number;
}

/**
 * Display settings of a sheet's view (`sheetView`, ECMA-376 Part 1 §18.3.1.87, W2-51). Frozen
 * panes are in {@link Sheet.freeze}.
 *
 * @example
 * ```ts
 * import { readWorkbook, type SheetView } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const view: SheetView | undefined = readWorkbook(bytes).sheets[0]?.view;
 * console.log(view?.zoom, view?.showGridLines);
 * ```
 */
export interface SheetView {
  /** Zoom in percent, 10 to 400; 100 when the file does not set it. */
  readonly zoom: number;
  readonly showGridLines: boolean;
  /** Whether row numbers and column letters are shown. */
  readonly showRowColHeaders: boolean;
  /** Column A at the right, text right to left. */
  readonly rightToLeft: boolean;
  readonly layout: SheetLayout;
  /** A split window, or `null`. */
  readonly split: SplitPanes | null;
  /** The top-left visible cell (`sheetView@topLeftCell`), or `null` when the file does not set one. */
  readonly topLeftCell: string | null;
  /** The active cell of the first selection, or `null`. */
  readonly activeCell: string | null;
  /** Zoom of the normal view in percent, or 0 for automatic (the default). */
  readonly zoomScaleNormal: number;
  /** Whether the ruler is shown in page layout view (default true). */
  readonly showRuler: boolean;
}

/**
 * A workbook window read from `bookViews/workbookView` (sml.xsd CT_BookView L4137; W2-66).
 *
 * @example
 * ```ts
 * import { readWorkbook, type WorkbookView } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const view: WorkbookView | undefined = readWorkbook(bytes).views[0];
 * console.log(view?.activeTab);
 * ```
 */
export interface WorkbookView {
  readonly x: number | null;
  readonly y: number | null;
  readonly width: number | null;
  readonly height: number | null;
  /** `firstSheet`, or `null` when the file does not set it (default 0). */
  readonly firstSheet: number | null;
  /** `activeTab`, or `null` when the file does not set it (default 0). */
  readonly activeTab: number | null;
  readonly visibility: 'visible' | 'hidden' | 'veryHidden';
}

/**
 * Sheet-wide formatting read from `sheetFormatPr` and `sheetPr/outlinePr` (W2-65).
 *
 * @example
 * ```ts
 * import { readWorkbook, type SheetProperties } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const props: SheetProperties | undefined = readWorkbook(bytes).sheets[0]?.properties;
 * console.log(props?.defaultRowHeight, props?.outlineLevelRow);
 * ```
 */
export interface SheetProperties {
  /** Default row height in points (15 when the file has no `sheetFormatPr`). */
  readonly defaultRowHeight: number;
  /** Highest outline level of the rows (`outlineLevelRow`). */
  readonly outlineLevelRow: number;
  /** Highest outline level of the columns (`outlineLevelCol`). */
  readonly outlineLevelCol: number;
  readonly summaryBelow: boolean;
  readonly summaryRight: boolean;
}

/**
 * The kind of sheet, from the relationship that points at its part.
 *
 * @example
 * ```ts
 * import { readWorkbook, type SheetKind } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const kinds: SheetKind[] = readWorkbook(bytes).sheets.map((sheet) => sheet.kind);
 * console.log(kinds.filter((kind) => kind === 'worksheet').length);
 * ```
 */
export type SheetKind = 'worksheet' | 'chartsheet' | 'dialogsheet' | 'macrosheet';

/**
 * Visibility of a sheet (ECMA-376 Part 1, ST_SheetState).
 *
 * @example
 * ```ts
 * import { readWorkbook, type SheetState } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const states: SheetState[] = readWorkbook(bytes).sheets.map((sheet) => sheet.state);
 * console.log(states.includes('veryHidden'));
 * ```
 */
export type SheetState = 'visible' | 'hidden' | 'veryHidden';

/**
 * Which date system the workbook uses for serial date values.
 *
 * @example
 * ```ts
 * import { readWorkbook, type DateSystem } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const system: DateSystem = readWorkbook(bytes).dateSystem;
 * console.log(system); // "1900" or "1904"
 * ```
 */
export type DateSystem = '1900' | '1904';

/**
 * Conformance class of the package.
 *
 * @example
 * ```ts
 * import { readWorkbook, type Conformance } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const conformance: Conformance = readWorkbook(bytes).conformance;
 * console.log(conformance); // "transitional" or "strict"
 * ```
 */
export type Conformance = 'transitional' | 'strict';

/**
 * File format, derived from the content type of the workbook part.
 *
 * @example
 * ```ts
 * import { readWorkbook, type WorkbookFormat } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const format: WorkbookFormat = readWorkbook(bytes).format;
 * console.log(format === 'xlsm' ? 'macro-enabled' : format);
 * ```
 */
export type WorkbookFormat = 'xlsx' | 'xlsm' | 'xltx' | 'xltm' | 'unknown';

/**
 * A sheet of a workbook read by `readWorkbook`.
 *
 * @example
 * ```ts
 * import { readWorkbook, type Sheet } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const sheet: Sheet = readWorkbook(bytes).getSheet('Sales');
 * for (const row of sheet.rows()) console.log(row.map((cell) => cell.value));
 * ```
 */
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
  /** Hyperlinks of the sheet in file order (see `Hyperlink`); their targets are never opened. */
  readonly hyperlinks: readonly Hyperlink[];
  /** Notes (comments) of the sheet, in file order (see `Note`). */
  readonly notes: readonly Note[];
  /** Pictures of the sheet's drawing, in file order (see `SheetImage`); charts and shapes are not read. */
  readonly images: readonly SheetImage[];
  /** The background image of the sheet (`picture` element), or `null`. */
  readonly backgroundImage: WorkbookImage | null;
  /** Colour of the sheet tab (`sheetPr/tabColor`), or `null`. */
  readonly tabColor: Color | null;
  /** Display settings of the sheet's first view. */
  readonly view: SheetView;
  /** Sheet-wide formatting: default row height, outline levels and summary positions (W2-65). */
  readonly properties: SheetProperties;
  /** Data validations in file order (see `DataValidation`). */
  readonly dataValidations: readonly DataValidation[];
  /** Conditional formatting in file order (see `ConditionalFormat`); x14 extensions are not read. */
  readonly conditionalFormats: readonly ConditionalFormat[];
  /** `sheetProtection` of the sheet, or `null` when it has none (see `SheetProtection`). */
  readonly protection: SheetProtection | null;
  /** Page setup for printing, with the schema defaults where the file sets nothing. */
  readonly pageSetup: PageSetup;
  /** Printed headers and footers, or `null` when the sheet has none. */
  readonly headerFooter: Required<HeaderFooter> | null;
  /** Manual page breaks: rows and columns (numbers) that start a new printed page. */
  readonly pageBreaks: { readonly rows: readonly number[]; readonly columns: readonly number[] };
  /**
   * The print area as ranges such as `["A1:H40"]` (from `_xlnm.Print_Area`), or `null` when the
   * sheet has none or its text has another form; the raw text is in `Workbook.definedNames`.
   */
  readonly printArea: readonly string[] | null;
  /** Rows (`"1:2"`) and columns (`"A:B"`) repeated on each printed page, or `null`. */
  readonly printTitles: { readonly rows: string | null; readonly columns: string | null } | null;
  /** Tables of the sheet (see `TableDefinition`), in the order of the sheet's relationships. */
  readonly tables: readonly TableDefinition[];
  /** Height and visibility of a row (1-based), or `undefined` if the file sets neither. */
  rowInfo(row: number): RowInfo | undefined;
  /**
   * Returns the cell at an A1-style reference, or `undefined` if the file has none there. Cells
   * with a value or formula only; see `styledCells` for cells that hold only a style.
   */
  cell(ref: string): Cell | undefined;
  /** Rows in ascending order; each row lists its cells in ascending column order. */
  rows(): IterableIterator<readonly Cell[]>;
  /** All cells with a value or formula, row by row. */
  cells(): IterableIterator<Cell>;
  /**
   * Cells that hold only a style (no value, no formula), row by row and in column order: their
   * `type` is `"empty"` and `value` is `null`. They are kept apart from `cells()`, so tables,
   * CSV, `inspect` and `diff` see only cells with data (D-E05). They count against
   * `limits.maxCells` but not against `cellCount`.
   *
   * @example
   * ```ts
   * import { readWorkbook } from 'xlvane';
   *
   * declare const bytes: Uint8Array;
   * for (const cell of readWorkbook(bytes).sheets[0]?.styledCells() ?? []) console.log(cell.ref, cell.style);
   * ```
   */
  styledCells(): IterableIterator<Cell>;
}

/**
 * A workbook read by `readWorkbook`.
 *
 * @example
 * ```ts
 * import { readWorkbook, type Workbook } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const workbook: Workbook = readWorkbook(bytes);
 * console.log(workbook.sheets.map((sheet) => sheet.name), workbook.hasMacros);
 * ```
 */
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
  /** Document properties from `docProps/core.xml`; empty when the file has none. */
  readonly properties: WorkbookProperties;
  /** Defined names of the workbook part, in file order (built-in `_xlnm.` names included). */
  readonly definedNames: readonly DefinedName[];
  /** `workbookProtection` of the workbook, or `null` (see `WorkbookProtection`). */
  readonly protection: WorkbookProtection | null;
  /** Workbook windows (`bookViews/workbookView`, W2-66), in file order. */
  readonly views: readonly WorkbookView[];
  /** Calculation settings (`calcPr`, W2-67), or `null` when the file has none. */
  readonly calculation: { readonly fullCalcOnLoad: boolean } | null;
  /** The images used by pictures and sheet backgrounds, each once (see `WorkbookImage`). */
  readonly images: readonly WorkbookImage[];
  /** Returns the sheet with this name; throws `XLV_E_SHEET_NOT_FOUND` if there is none. */
  getSheet(name: string): Sheet;
}
