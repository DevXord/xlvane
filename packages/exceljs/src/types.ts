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
 * Options for `workbook.addWorksheet()`. `pageSetup` and `headerFooter` are not supported.
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
  pageSetup?: unknown;
  headerFooter?: unknown;
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
