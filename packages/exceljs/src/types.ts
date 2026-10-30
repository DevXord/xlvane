// Types of the exceljs-compatible API. Shapes follow exceljs 4.4.0 `index.d.ts` (Font L217,
// Color L236, Border L248, Borders L258, Alignment L280, Style L295, CellValue L384, FillPattern
// L167, WorksheetViewFrozen L84); docs/exceljs-compat.md lists what is supported.

/**
 * A colour: `argb` such as `"FFFF0000"`, or a theme index (with optional tint), or an indexed colour.
 *
 * @example
 * ```ts
 * import type { Color } from '@xlvane/exceljs';
 *
 * const red: Partial<Color> = { argb: 'FFFF0000' };
 * ```
 */
export interface Color {
  argb: string;
  theme: number;
  tint: number;
  indexed: number;
}

/**
 * Font of a cell. `family`, `scheme` and `charset` are accepted and ignored when writing.
 *
 * @example
 * ```ts
 * import type { Font } from '@xlvane/exceljs';
 *
 * const font: Partial<Font> = { name: 'Arial', size: 12, bold: true, color: { argb: 'FF1F4E79' } };
 * ```
 */
export interface Font {
  name: string;
  size: number;
  family: number;
  scheme: 'minor' | 'major' | 'none';
  charset: number;
  color: Partial<Color>;
  bold: boolean;
  italic: boolean;
  underline: boolean | 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
  vertAlign: 'superscript' | 'subscript';
  strike: boolean;
  outline: boolean;
}

/**
 * Fill patterns.
 *
 * @example
 * ```ts
 * import type { FillPatterns } from '@xlvane/exceljs';
 *
 * const pattern: FillPatterns = 'solid';
 * ```
 */
export type FillPatterns =
  | 'none'
  | 'solid'
  | 'darkVertical'
  | 'darkHorizontal'
  | 'darkGrid'
  | 'darkTrellis'
  | 'darkDown'
  | 'darkUp'
  | 'lightVertical'
  | 'lightHorizontal'
  | 'lightGrid'
  | 'lightTrellis'
  | 'lightDown'
  | 'lightUp'
  | 'darkGray'
  | 'mediumGray'
  | 'lightGray'
  | 'gray125'
  | 'gray0625';

/**
 * A pattern fill. Gradient fills are not supported (`XLV_E_NOT_SUPPORTED`).
 *
 * @example
 * ```ts
 * import type { Fill } from '@xlvane/exceljs';
 *
 * const fill: Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
 * ```
 */
export interface Fill {
  type: 'pattern' | 'gradient';
  pattern?: FillPatterns;
  fgColor?: Partial<Color>;
  bgColor?: Partial<Color>;
  [key: string]: unknown;
}

/**
 * Border line styles.
 *
 * @example
 * ```ts
 * import type { BorderStyle } from '@xlvane/exceljs';
 *
 * const style: BorderStyle = 'thin';
 * ```
 */
export type BorderStyle =
  | 'thin'
  | 'dotted'
  | 'hair'
  | 'medium'
  | 'double'
  | 'thick'
  | 'dashed'
  | 'dashDot'
  | 'dashDotDot'
  | 'slantDashDot'
  | 'mediumDashed'
  | 'mediumDashDotDot'
  | 'mediumDashDot';

/**
 * One edge of a border.
 *
 * @example
 * ```ts
 * import type { Border } from '@xlvane/exceljs';
 *
 * const edge: Partial<Border> = { style: 'thin', color: { argb: 'FF000000' } };
 * ```
 */
export interface Border {
  style: BorderStyle;
  color: Partial<Color>;
}

/**
 * Borders of a cell. `diagonal` is not supported (`XLV_E_NOT_SUPPORTED` when it has a style).
 *
 * @example
 * ```ts
 * import type { Borders } from '@xlvane/exceljs';
 *
 * const border: Partial<Borders> = { top: { style: 'thin' }, bottom: { style: 'double' } };
 * ```
 */
export interface Borders {
  top: Partial<Border>;
  left: Partial<Border>;
  bottom: Partial<Border>;
  right: Partial<Border>;
  diagonal: Partial<Border> & { up?: boolean; down?: boolean };
}

/**
 * Alignment of a cell. `shrinkToFit` and `readingOrder` are not supported.
 *
 * @example
 * ```ts
 * import type { Alignment } from '@xlvane/exceljs';
 *
 * const alignment: Partial<Alignment> = { horizontal: 'center', vertical: 'middle', wrapText: true };
 * ```
 */
export interface Alignment {
  horizontal: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
  vertical: 'top' | 'middle' | 'bottom' | 'distributed' | 'justify';
  wrapText: boolean;
  shrinkToFit: boolean;
  indent: number;
  readingOrder: 'rtl' | 'ltr';
  textRotation: number | 'vertical';
}

/**
 * The style of a cell, row or column.
 *
 * @example
 * ```ts
 * import type { Style } from '@xlvane/exceljs';
 *
 * const style: Partial<Style> = { numFmt: '#,##0.00', font: { bold: true } };
 * ```
 */
export interface Style {
  numFmt?: string | undefined;
  font?: Partial<Font> | undefined;
  fill?: Fill | undefined;
  border?: Partial<Borders> | undefined;
  alignment?: Partial<Alignment> | undefined;
  protection?: unknown;
}

/**
 * An error value of a cell.
 *
 * @example
 * ```ts
 * import type { CellErrorValue } from '@xlvane/exceljs';
 *
 * const na: CellErrorValue = { error: '#N/A' };
 * ```
 */
export interface CellErrorValue {
  error: '#N/A' | '#REF!' | '#NAME?' | '#DIV/0!' | '#NULL!' | '#VALUE!' | '#NUM!';
}

/**
 * A formula with an optional cached result. xlvane does not calculate formulas.
 *
 * @example
 * ```ts
 * import type { CellFormulaValue } from '@xlvane/exceljs';
 *
 * const total: CellFormulaValue = { formula: 'SUM(B2:B10)', result: 120 };
 * ```
 */
export interface CellFormulaValue {
  formula: string;
  result?: number | string | boolean | Date | CellErrorValue;
  date1904?: boolean;
}

/**
 * A cell value, including rich text (`{ richText: [{ text, font? }] }`) and hyperlinks
 * (`{ text, hyperlink, tooltip? }`; a `hyperlink` starting with `#` is a place in the workbook).
 * Shared-string values are not supported (`XLV_E_NOT_SUPPORTED` when set); shared formulas are
 * written as ordinary formulas.
 *
 * @example
 * ```ts
 * import type { CellValue } from '@xlvane/exceljs';
 *
 * const values: CellValue[] = [
 *   1, 'text', true, new Date(), null, { formula: 'A1*2' }, { error: '#N/A' },
 *   { richText: [{ font: { bold: true }, text: 'Bold' }, { text: ' plain' }] },
 *   { text: 'Docs', hyperlink: 'https://example.com', tooltip: 'Open the docs' },
 *   { text: 'Totals', hyperlink: "#'Sheet 2'!A1" },
 * ];
 * ```
 */
export type CellValue = null | undefined | number | string | boolean | Date | CellErrorValue | CellFormulaValue | Record<string, unknown>;

/**
 * A column definition for `worksheet.columns`.
 *
 * @example
 * ```ts
 * import type { ColumnDefn } from '@xlvane/exceljs';
 *
 * const column: Partial<ColumnDefn> = { header: 'Total', key: 'total', width: 12, style: { numFmt: '#,##0.00' } };
 * ```
 */
export interface ColumnDefn {
  header?: string | string[] | undefined;
  key?: string | undefined;
  width?: number | undefined;
  style?: Style | undefined;
  hidden?: boolean | undefined;
}

/**
 * A worksheet view. The first view of a worksheet is written: frozen panes (`xSplit`/`ySplit`
 * in columns and rows) or a split window (`state: "split"`, `xSplit`/`ySplit` in 1/20 point),
 * `zoomScale`, `showGridLines`, `showRowColHeaders`, `rightToLeft` and `style`. `topLeftCell`,
 * `activeCell`, `activePane`, `showRuler` and `zoomScaleNormal` are not written.
 *
 * @example
 * ```ts
 * import type { WorksheetView } from '@xlvane/exceljs';
 *
 * const view: WorksheetView = { state: 'frozen', xSplit: 1, ySplit: 1 };
 * ```
 */
export interface WorksheetView {
  state?: 'normal' | 'frozen' | 'split';
  xSplit?: number;
  ySplit?: number;
  topLeftCell?: string | undefined;
  activeCell?: string;
  showGridLines?: boolean;
  zoomScale?: number;
  [key: string]: unknown;
}

/**
 * Options for `workbook.addWorksheet()`, as in exceljs (index.d.ts `AddWorksheetOptions`).
 *
 * @example
 * ```ts
 * import type { WorksheetOptions } from '@xlvane/exceljs';
 *
 * const options: WorksheetOptions = { state: 'hidden', views: [{ state: 'frozen', ySplit: 1 }] };
 * ```
 */
export interface WorksheetOptions {
  state?: 'visible' | 'hidden' | 'veryHidden';
  views?: WorksheetView[];
  properties?: Record<string, unknown>;
  pageSetup?: Partial<PageSetup>;
  headerFooter?: Partial<HeaderFooter>;
}

/**
 * A cell note (`cell.note = string | Comment`), as in exceljs (index.d.ts `Comment`, lib/doc/note.js).
 * `texts` are rich text runs. `margins`, `protection` and `editAs` are kept on the object but not
 * written: xlvane writes every note box in the form of the ECMA-376 examples (D-A14).
 *
 * @example
 * ```ts
 * import type { Comment } from '@xlvane/exceljs';
 *
 * const note: Comment = { texts: [{ font: { bold: true }, text: 'Check: ' }, { text: 'estimate only' }] };
 * ```
 */
export interface Comment {
  texts?: { text: string; font?: Partial<Font> }[];
  margins?: { insetmode?: 'auto' | 'custom'; inset?: number[] };
  protection?: { locked?: 'True' | 'False'; lockText?: 'True' | 'False' };
  editAs?: 'twoCells' | 'oneCells' | 'absolute';
}

/**
 * A data validation (`cell.dataValidation`, `worksheet.dataValidations`), as in exceljs
 * (index.d.ts `DataValidation`). `'any'` is a validation that only shows its input message.
 * `formulae` are formulas (`'"Yes,No"'` for a literal list, `'$D$1:$D$9'` for a range), numbers or
 * dates; reading gives numbers for whole, decimal and textLength validations and dates for date
 * validations, as exceljs does.
 *
 * @example
 * ```ts
 * import type { DataValidation } from '@xlvane/exceljs';
 *
 * const rule: DataValidation = { type: 'whole', operator: 'between', allowBlank: true, formulae: [1, 10], showErrorMessage: true, error: '1 to 10' };
 * ```
 */
export interface DataValidation {
  type: 'list' | 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'custom' | 'any';
  formulae?: unknown[];
  allowBlank?: boolean;
  operator?: 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual';
  error?: string;
  errorTitle?: string;
  errorStyle?: 'stop' | 'warning' | 'information';
  prompt?: string;
  promptTitle?: string;
  showErrorMessage?: boolean;
  showInputMessage?: boolean;
}

/**
 * A conditional formatting rule, as in exceljs (index.d.ts `ConditionalFormattingRule`): the
 * rule type with its own properties (`formulae`, `operator`, `rank`, `cfvo`, `color`, `iconSet`,
 * `text`, `timePeriod`, …) and the `style` applied where the rule is true.
 *
 * @example
 * ```ts
 * import type { ConditionalFormattingRule } from '@xlvane/exceljs';
 *
 * const rule: ConditionalFormattingRule = { type: 'cellIs', operator: 'greaterThan', formulae: [100], style: { font: { bold: true } } };
 * ```
 */
export interface ConditionalFormattingRule {
  type: 'expression' | 'cellIs' | 'top10' | 'aboveAverage' | 'colorScale' | 'iconSet' | 'containsText' | 'timePeriod' | 'dataBar';
  priority?: number;
  style?: Style;
  [key: string]: unknown;
}

/**
 * The argument of `worksheet.addConditionalFormatting()`: the range and its rules.
 *
 * @example
 * ```ts
 * import type { ConditionalFormattingOptions } from '@xlvane/exceljs';
 *
 * const cf: ConditionalFormattingOptions = { ref: 'B2:B20', rules: [{ type: 'top10', rank: 3, style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } }] };
 * ```
 */
export interface ConditionalFormattingOptions {
  ref: string;
  rules: ConditionalFormattingRule[];
}

/**
 * Page setup of a worksheet (`worksheet.pageSetup`), as in exceljs (index.d.ts `PageSetup`).
 * `printArea` (several areas joined by `&&`), `printTitlesRow` (`'1:2'`) and `printTitlesColumn`
 * (`'A:B'`) are stored as the built-in names Excel uses.
 *
 * @example
 * ```ts
 * import type { PageSetup } from '@xlvane/exceljs';
 *
 * const setup: Partial<PageSetup> = { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, printArea: 'A1:G40' };
 * ```
 */
export interface PageSetup {
  margins?: { left: number; right: number; top: number; bottom: number; header: number; footer: number } | undefined;
  orientation?: 'portrait' | 'landscape' | undefined;
  horizontalDpi?: number | undefined;
  verticalDpi?: number | undefined;
  fitToPage?: boolean | undefined;
  pageOrder?: 'downThenOver' | 'overThenDown' | undefined;
  blackAndWhite?: boolean | undefined;
  draft?: boolean | undefined;
  cellComments?: 'atEnd' | 'asDisplayed' | 'asDisplyed' | 'None' | undefined;
  errors?: 'dash' | 'blank' | 'NA' | 'displayed' | undefined;
  scale?: number | undefined;
  fitToWidth?: number | undefined;
  fitToHeight?: number | undefined;
  paperSize?: number | undefined;
  showRowColHeaders?: boolean | undefined;
  showGridLines?: boolean | undefined;
  firstPageNumber?: number | undefined;
  useFirstPageNumber?: boolean | undefined;
  usePrinterDefaults?: boolean | undefined;
  copies?: number | undefined;
  horizontalCentered?: boolean | undefined;
  verticalCentered?: boolean | undefined;
  rowBreaks?: unknown | undefined;
  colBreaks?: unknown | undefined;
  printArea?: string | undefined;
  printTitlesRow?: string | undefined;
  printTitlesColumn?: string | undefined;
}

/**
 * Headers and footers of a worksheet (`worksheet.headerFooter`), as in exceljs (index.d.ts
 * `HeaderFooter`). The texts use Excel's codes (`&C` centre section, `&P` page number, …).
 *
 * @example
 * ```ts
 * import type { HeaderFooter } from '@xlvane/exceljs';
 *
 * const hf: HeaderFooter = { oddHeader: '&CQuarterly report', oddFooter: '&RPage &P of &N' };
 * ```
 */
export interface HeaderFooter {
  differentFirst?: boolean;
  differentOddEven?: boolean;
  oddHeader?: string | null;
  oddFooter?: string | null;
  evenHeader?: string | null;
  evenFooter?: string | null;
  firstHeader?: string | null;
  firstFooter?: string | null;
}

/**
 * A manual row break (`worksheet.rowBreaks`), as `row.addPageBreak()` stores it in exceljs:
 * the page ends after row `id`.
 *
 * @example
 * ```ts
 * import type { RowBreak } from '@xlvane/exceljs';
 *
 * const pageEndsAfterRow20: RowBreak = { id: 20, max: 16383, man: 1 };
 * ```
 */
export interface RowBreak {
  id: number;
  max?: number;
  min?: number;
  man?: number;
}

/**
 * Options of `worksheet.protect()`, as in exceljs (index.d.ts `WorksheetProtection`): `true`
 * allows users to do the action on the protected sheet.
 *
 * @example
 * ```ts
 * import type { WorksheetProtection } from '@xlvane/exceljs';
 *
 * const options: Partial<WorksheetProtection> = { formatCells: true, sort: true, autoFilter: true };
 * ```
 */
export interface WorksheetProtection {
  objects: boolean;
  scenarios: boolean;
  selectLockedCells: boolean;
  selectUnlockedCells: boolean;
  formatCells: boolean;
  formatColumns: boolean;
  formatRows: boolean;
  insertColumns: boolean;
  insertRows: boolean;
  insertHyperlinks: boolean;
  deleteColumns: boolean;
  deleteRows: boolean;
  sort: boolean;
  autoFilter: boolean;
  pivotTables: boolean;
  spinCount: number;
}
