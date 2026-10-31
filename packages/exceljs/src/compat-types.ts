// exceljs 4.4.0 type names that code written for exceljs uses (index.d.ts line numbers in each
// TSDoc), so that such code type-checks against @xlvane/exceljs (docs/decisions.md D-R01).
// Shapes follow exceljs (MIT); where a value is a Node.js Buffer in exceljs it is a Uint8Array here.
import type { CsvIO } from './csv';
import type { Worksheet } from './model';
import type { ValueType } from './model';
import type { Border, CellErrorValue, CellValue, Color, FillPatterns, Font, HeaderFooter, PageSetup, RowBreak, Style, WorksheetView } from './types';
import type { XlsxIO } from './workbook';

/**
 * The bytes exceljs returns from `writeBuffer()` (a Node.js Buffer in Node.js, a `Uint8Array` elsewhere); exceljs `Buffer` (index.d.ts L1).
 *
 * @example
 * ```ts
 * import type { Buffer } from '@xlvane/exceljs';
 * const value: Buffer | undefined = new Uint8Array(0);
 * console.log(value);
 * ```
 */
export type Buffer = Uint8Array;

/**
 * Paper sizes of `pageSetup.paperSize`, as exceljs `PaperSize` (index.d.ts L18).
 *
 * @example
 * ```ts
 * import { PaperSize } from '@xlvane/exceljs';
 *
 * const size: PaperSize = PaperSize.A4;
 * console.log(size); // 9
 * ```
 */
export const PaperSize: {
  readonly Legal: 5;
  readonly Executive: 7;
  readonly A4: 9;
  readonly A5: 11;
  readonly B5: 13;
  readonly Envelope_10: 20;
  readonly Envelope_DL: 27;
  readonly Envelope_C5: 28;
  readonly Envelope_B5: 34;
  readonly Envelope_Monarch: 37;
  readonly Double_Japan_Postcard_Rotated: 82;
  readonly K16_197x273_mm: 119;
} = Object.freeze({
  Legal: 5,
  Executive: 7,
  A4: 9,
  A5: 11,
  B5: 13,
  Envelope_10: 20,
  Envelope_DL: 27,
  Envelope_C5: 28,
  Envelope_B5: 34,
  Envelope_Monarch: 37,
  Double_Japan_Postcard_Rotated: 82,
  K16_197x273_mm: 119,
} as const);
/** A value of {@link PaperSize}. */
export type PaperSize = (typeof PaperSize)[keyof typeof PaperSize];

/**
 * One colour stop of a gradient fill (exceljs `GradientStop`, index.d.ts L174).
 *
 * @example
 * ```ts
 * import type { GradientStop } from '@xlvane/exceljs';
 * const value: GradientStop | undefined = { position: 0, color: { argb: 'FFFFFFFF' } };
 * console.log(value);
 * ```
 */
export interface GradientStop {
  position: number;
  color: Partial<Color>;
}

/**
 * A linear gradient fill (exceljs `FillGradientAngle`, index.d.ts L179).
 *
 * @example
 * ```ts
 * import type { FillGradientAngle } from '@xlvane/exceljs';
 * const value: FillGradientAngle | undefined = { type: 'gradient', gradient: 'angle', degree: 90, stops: [{ position: 0, color: { argb: 'FFFFFFFF' } }, { position: 1, color: { argb: 'FF0070C0' } }] };
 * console.log(value);
 * ```
 */
export interface FillGradientAngle {
  type: 'gradient';
  gradient: 'angle';
  degree: number;
  stops: GradientStop[];
}

/**
 * A path gradient fill (exceljs `FillGradientPath`, index.d.ts L197).
 *
 * @example
 * ```ts
 * import type { FillGradientPath } from '@xlvane/exceljs';
 * const value: FillGradientPath | undefined = { type: 'gradient', gradient: 'path', center: { left: 0.5, top: 0.5 }, stops: [] };
 * console.log(value);
 * ```
 */
export interface FillGradientPath {
  type: 'gradient';
  gradient: 'path';
  center: { left: number; top: number };
  stops: GradientStop[];
}

/**
 * A diagonal border (exceljs `BorderDiagonal`, index.d.ts L253).
 *
 * @example
 * ```ts
 * import type { BorderDiagonal } from '@xlvane/exceljs';
 * const value: BorderDiagonal | undefined = { style: 'thin', color: { argb: 'FF000000' }, up: true, down: false };
 * console.log(value);
 * ```
 */
export interface BorderDiagonal extends Border {
  up: boolean;
  down: boolean;
}

/**
 * Page margins in inches (exceljs `Margins`, index.d.ts L266).
 *
 * @example
 * ```ts
 * import type { Margins } from '@xlvane/exceljs';
 * const value: Margins | undefined = { top: 0.75, left: 0.7, bottom: 0.75, right: 0.7, header: 0.3, footer: 0.3 };
 * console.log(value);
 * ```
 */
export interface Margins {
  top: number;
  left: number;
  bottom: number;
  right: number;
  header: number;
  footer: number;
}

/**
 * Operators of a data validation (exceljs `DataValidationOperator`, index.d.ts L304).
 *
 * @example
 * ```ts
 * import type { DataValidationOperator } from '@xlvane/exceljs';
 * const value: DataValidationOperator | undefined = 'between';
 * console.log(value);
 * ```
 */
export type DataValidationOperator = 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual';

/**
 * One run of rich text (exceljs `RichText`, index.d.ts L336).
 *
 * @example
 * ```ts
 * import type { RichText } from '@xlvane/exceljs';
 * const value: RichText | undefined = { text: 'Bold', font: { bold: true } };
 * console.log(value);
 * ```
 */
export interface RichText {
  text: string;
  font?: Partial<Font>;
}

/**
 * A rich-text cell value (exceljs `CellRichTextValue`, index.d.ts L341).
 *
 * @example
 * ```ts
 * import type { CellRichTextValue } from '@xlvane/exceljs';
 * const value: CellRichTextValue | undefined = { richText: [{ text: 'a' }, { text: 'b', font: { italic: true } }] };
 * console.log(value);
 * ```
 */
export interface CellRichTextValue {
  richText: RichText[];
}

/**
 * A hyperlink cell value (exceljs `CellHyperlinkValue`, index.d.ts L345).
 *
 * @example
 * ```ts
 * import type { CellHyperlinkValue } from '@xlvane/exceljs';
 * const value: CellHyperlinkValue | undefined = { text: 'Site', hyperlink: 'https://example.com' };
 * console.log(value);
 * ```
 */
export interface CellHyperlinkValue {
  text: string;
  hyperlink: string;
  tooltip?: string;
}

/**
 * A decoded cell address (exceljs `Address`, index.d.ts L1622).
 *
 * @example
 * ```ts
 * import type { Address } from '@xlvane/exceljs';
 * const value: Address | undefined = undefined;
 * console.log(value);
 * ```
 */
export interface Address {
  sheetName?: string;
  address: string;
  col: string;
  row: string;
  $col$row: string;
}

/**
 * The data of one cell as exceljs `CellModel` (index.d.ts L410). `cell.model` returns xlvane's own snapshot with these fields where they apply.
 *
 * @example
 * ```ts
 * import type { CellModel } from '@xlvane/exceljs';
 * const value: CellModel | undefined = undefined;
 * console.log(value);
 * ```
 */
export interface CellModel {
  address: string;
  style: Partial<Style>;
  type: ValueType;
  text?: string;
  hyperlink?: string;
  value?: CellValue;
  master?: string;
  formula?: string;
  sharedFormula?: string;
  result?: unknown;
  comment?: unknown;
}

/**
 * The data of one row (exceljs `RowModel`, index.d.ts L511).
 *
 * @example
 * ```ts
 * import type { RowModel } from '@xlvane/exceljs';
 * const value: RowModel | undefined = undefined;
 * console.log(value);
 * ```
 */
export interface RowModel {
  cells: CellModel[];
  number: number;
  min: number;
  max: number;
  height: number;
  style: Partial<Style>;
  hidden: boolean;
  outlineLevel: number;
  collapsed: boolean;
}

/**
 * The autofilter of a worksheet (exceljs `AutoFilter`, index.d.ts L864).
 *
 * @example
 * ```ts
 * import type { AutoFilter } from '@xlvane/exceljs';
 * const value: AutoFilter | undefined = 'A1:C1';
 * console.log(value);
 * ```
 */
export type AutoFilter = string | { from: string | { row: number; column: number }; to: string | { row: number; column: number } };

/**
 * An image for `workbook.addImage()` (exceljs `Image`, index.d.ts L887). `filename` works only in Node.js, Bun and Deno.
 *
 * @example
 * ```ts
 * import type { Image } from '@xlvane/exceljs';
 * const value: Image | undefined = { extension: 'png', base64: 'iVBORw0KGgo=' };
 * console.log(value);
 * ```
 */
export interface Image {
  extension: 'jpeg' | 'png' | 'gif';
  base64?: string;
  filename?: string;
  buffer?: Uint8Array;
}

/**
 * Two corners of a picture (exceljs `ImageRange`, index.d.ts L915).
 *
 * @example
 * ```ts
 * import type { ImageRange } from '@xlvane/exceljs';
 * const value: ImageRange | undefined = { tl: { col: 0, row: 0 }, br: { col: 2, row: 3 } };
 * console.log(value);
 * ```
 */
export interface ImageRange {
  tl: IAnchorPosition;
  br: IAnchorPosition;
}

/**
 * A picture corner given by column and row, counted from 0 (the `tl`/`br` positions of exceljs `ImageRange`).
 *
 * @example
 * ```ts
 * import type { IAnchorPosition } from '@xlvane/exceljs';
 * const value: IAnchorPosition | undefined = { col: 1, row: 2 };
 * console.log(value);
 * ```
 */
export interface IAnchorPosition {
  col: number;
  row: number;
}

/**
 * A picture corner plus a size in pixels (exceljs `ImagePosition`, index.d.ts L920).
 *
 * @example
 * ```ts
 * import type { ImagePosition } from '@xlvane/exceljs';
 * const value: ImagePosition | undefined = { tl: { col: 0, row: 0 }, ext: { width: 120, height: 60 } };
 * console.log(value);
 * ```
 */
export interface ImagePosition {
  tl: { col: number; row: number };
  ext: { width: number; height: number };
}

/**
 * The link of a picture (exceljs `ImageHyperlinkValue`, index.d.ts L925).
 *
 * @example
 * ```ts
 * import type { ImageHyperlinkValue } from '@xlvane/exceljs';
 * const value: ImageHyperlinkValue | undefined = { hyperlink: 'https://example.com' };
 * console.log(value);
 * ```
 */
export interface ImageHyperlinkValue {
  hyperlink: string;
  tooltip?: string;
}

/**
 * A rectangle of cells, 1-based (exceljs `Location`, index.d.ts L1630).
 *
 * @example
 * ```ts
 * import type { Location } from '@xlvane/exceljs';
 * const value: Location | undefined = { top: 1, left: 1, bottom: 3, right: 2 };
 * console.log(value);
 * ```
 */
export interface Location {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

/**
 * A range of cells as exceljs `Range` (index.d.ts L930), the type of `worksheet.dimensions`.
 *
 * @example
 * ```ts
 * import type { Range } from '@xlvane/exceljs';
 * const value: Range | undefined = undefined;
 * console.log(value);
 * ```
 */
export interface Range extends Location {
  readonly tl: string;
  readonly br: string;
  readonly range: string;
  readonly shortRange: string;
  readonly $range: string;
  readonly count: number;
  readonly model: Location & { sheetName?: string };
  toString(): string;
}

/**
 * A picture of a workbook (exceljs `Media`, index.d.ts L1615).
 *
 * @example
 * ```ts
 * import type { Media } from '@xlvane/exceljs';
 * const value: Media | undefined = undefined;
 * console.log(value);
 * ```
 */
export interface Media {
  type: string;
  name: string;
  extension: string;
  buffer: Uint8Array;
}

/**
 * Visibility of a worksheet (exceljs `WorksheetState`, index.d.ts L997).
 *
 * @example
 * ```ts
 * import type { WorksheetState } from '@xlvane/exceljs';
 * const value: WorksheetState | undefined = 'hidden';
 * console.log(value);
 * ```
 */
export type WorksheetState = 'visible' | 'hidden' | 'veryHidden';

/**
 * Operators of a `cellIs` conditional formatting rule (exceljs index.d.ts L999).
 *
 * @example
 * ```ts
 * import type { CellIsOperators } from '@xlvane/exceljs';
 * const value: CellIsOperators | undefined = 'between';
 * console.log(value);
 * ```
 */
export type CellIsOperators = 'equal' | 'greaterThan' | 'lessThan' | 'between';

/**
 * Operators of a `containsText` rule (exceljs index.d.ts L1001).
 *
 * @example
 * ```ts
 * import type { ContainsTextOperators } from '@xlvane/exceljs';
 * const value: ContainsTextOperators | undefined = 'containsText';
 * console.log(value);
 * ```
 */
export type ContainsTextOperators = 'containsText' | 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors';

/**
 * Periods of a `timePeriod` rule (exceljs index.d.ts L1003).
 *
 * @example
 * ```ts
 * import type { TimePeriodTypes } from '@xlvane/exceljs';
 * const value: TimePeriodTypes | undefined = 'today';
 * console.log(value);
 * ```
 */
export type TimePeriodTypes = 'lastWeek' | 'thisWeek' | 'nextWeek' | 'yesterday' | 'today' | 'tomorrow' | 'last7Days' | 'lastMonth' | 'thisMonth' | 'nextMonth';

/**
 * Icon sets of an `iconSet` rule (exceljs index.d.ts L1006).
 *
 * @example
 * ```ts
 * import type { IconSetTypes } from '@xlvane/exceljs';
 * const value: IconSetTypes | undefined = '3Arrows';
 * console.log(value);
 * ```
 */
export type IconSetTypes =
  | '5Arrows'
  | '5ArrowsGray'
  | '5Boxes'
  | '5Quarters'
  | '5Rating'
  | '4Arrows'
  | '4ArrowsGray'
  | '4Rating'
  | '4RedToBlack'
  | '4TrafficLights'
  | 'NoIcons'
  | '3Arrows'
  | '3ArrowsGray'
  | '3Flags'
  | '3Signs'
  | '3Stars'
  | '3Symbols'
  | '3Symbols2'
  | '3TrafficLights1'
  | '3TrafficLights2'
  | '3Triangles';

/**
 * Value types of a conditional formatting value object (exceljs index.d.ts L1010).
 *
 * @example
 * ```ts
 * import type { CfvoTypes } from '@xlvane/exceljs';
 * const value: CfvoTypes | undefined = 'min';
 * console.log(value);
 * ```
 */
export type CfvoTypes = 'percentile' | 'percent' | 'num' | 'min' | 'max' | 'formula' | 'autoMin' | 'autoMax';

/**
 * A conditional formatting value object (exceljs `Cvfo`, index.d.ts L1012).
 *
 * @example
 * ```ts
 * import type { Cvfo } from '@xlvane/exceljs';
 * const value: Cvfo | undefined = { type: 'percentile', value: 50 };
 * console.log(value);
 * ```
 */
export interface Cvfo {
  type: CfvoTypes;
  value?: number;
}

/**
 * A `expression` conditional formatting rule (exceljs `ExpressionRuleType`, index.d.ts L1020).
 *
 * @example
 * ```ts
 * import type { ExpressionRuleType } from '@xlvane/exceljs';
 * const value: ExpressionRuleType | undefined = { type: 'expression' } as ExpressionRuleType;
 * console.log(value);
 * ```
 */
export interface ExpressionRuleType {
  type: 'expression';
  priority?: number;
  style?: Partial<Style>;
  formulae?: unknown[];
}

/**
 * A `cellIs` conditional formatting rule (exceljs `CellIsRuleType`, index.d.ts L1025).
 *
 * @example
 * ```ts
 * import type { CellIsRuleType } from '@xlvane/exceljs';
 * const value: CellIsRuleType | undefined = { type: 'cellIs' } as CellIsRuleType;
 * console.log(value);
 * ```
 */
export interface CellIsRuleType {
  type: 'cellIs';
  priority?: number;
  style?: Partial<Style>;
  formulae?: unknown[];
  operator?: CellIsOperators;
}

/**
 * A `aboveAverage` conditional formatting rule (exceljs `AboveAverageRuleType`, index.d.ts L1038).
 *
 * @example
 * ```ts
 * import type { AboveAverageRuleType } from '@xlvane/exceljs';
 * const value: AboveAverageRuleType | undefined = { type: 'aboveAverage' } as AboveAverageRuleType;
 * console.log(value);
 * ```
 */
export interface AboveAverageRuleType {
  type: 'aboveAverage';
  priority?: number;
  style?: Partial<Style>;
  aboveAverage: boolean;
}

/**
 * A `colorScale` conditional formatting rule (exceljs `ColorScaleRuleType`, index.d.ts L1043).
 *
 * @example
 * ```ts
 * import type { ColorScaleRuleType } from '@xlvane/exceljs';
 * const value: ColorScaleRuleType | undefined = { type: 'colorScale' } as ColorScaleRuleType;
 * console.log(value);
 * ```
 */
export interface ColorScaleRuleType {
  type: 'colorScale';
  priority?: number;
  style?: Partial<Style>;
  cfvo?: Cvfo[];
  color?: Partial<Color>[];
}

/**
 * A `iconSet` conditional formatting rule (exceljs `IconSetRuleType`, index.d.ts L1049).
 *
 * @example
 * ```ts
 * import type { IconSetRuleType } from '@xlvane/exceljs';
 * const value: IconSetRuleType | undefined = { type: 'iconSet' } as IconSetRuleType;
 * console.log(value);
 * ```
 */
export interface IconSetRuleType {
  type: 'iconSet';
  priority?: number;
  style?: Partial<Style>;
  showValue?: boolean;
  reverse?: boolean;
  custom?: boolean;
  iconSet?: IconSetTypes;
  cfvo?: Cvfo[];
}

/**
 * A `containsText` conditional formatting rule (exceljs `ContainsTextRuleType`, index.d.ts L1058).
 *
 * @example
 * ```ts
 * import type { ContainsTextRuleType } from '@xlvane/exceljs';
 * const value: ContainsTextRuleType | undefined = { type: 'containsText' } as ContainsTextRuleType;
 * console.log(value);
 * ```
 */
export interface ContainsTextRuleType {
  type: 'containsText';
  priority?: number;
  style?: Partial<Style>;
  operator?: ContainsTextOperators;
  text?: string;
}

/**
 * A `timePeriod` conditional formatting rule (exceljs `TimePeriodRuleType`, index.d.ts L1064).
 *
 * @example
 * ```ts
 * import type { TimePeriodRuleType } from '@xlvane/exceljs';
 * const value: TimePeriodRuleType | undefined = { type: 'timePeriod' } as TimePeriodRuleType;
 * console.log(value);
 * ```
 */
export interface TimePeriodRuleType {
  type: 'timePeriod';
  priority?: number;
  style?: Partial<Style>;
  timePeriod?: TimePeriodTypes;
}

/**
 * A `dataBar` conditional formatting rule (exceljs `DataBarRuleType`, index.d.ts L1069).
 *
 * @example
 * ```ts
 * import type { DataBarRuleType } from '@xlvane/exceljs';
 * const value: DataBarRuleType | undefined = { type: 'dataBar' } as DataBarRuleType;
 * console.log(value);
 * ```
 */
export interface DataBarRuleType {
  type: 'dataBar';
  priority?: number;
  style?: Partial<Style>;
  gradient?: boolean;
  minLength?: number;
  maxLength?: number;
  showValue?: boolean;
  border?: boolean;
  negativeBarColorSameAsPositive?: boolean;
  negativeBarBorderColorSameAsPositive?: boolean;
  axisPosition?: 'auto' | 'middle' | 'none';
  direction?: 'context' | 'leftToRight' | 'rightToLeft';
  cfvo?: Cvfo[];
}

/**
 * The values of a row given to `addRow()` (exceljs `RowValues`, index.d.ts L1087).
 *
 * @example
 * ```ts
 * import type { RowValues } from '@xlvane/exceljs';
 * const value: RowValues | undefined = [1, 'two', new Date(0)];
 * console.log(value);
 * ```
 */
export type RowValues = CellValue[] | { [key: string]: CellValue } | undefined | null;

/**
 * `workbook.calcProperties` (exceljs `CalculationProperties`, index.d.ts L1385).
 *
 * @example
 * ```ts
 * import type { CalculationProperties } from '@xlvane/exceljs';
 * const value: CalculationProperties | undefined = { fullCalcOnLoad: true };
 * console.log(value);
 * ```
 */
export interface CalculationProperties {
  fullCalcOnLoad: boolean;
}

/**
 * `workbook.properties` (exceljs `WorkbookProperties`, index.d.ts L1440).
 *
 * @example
 * ```ts
 * import type { WorkbookProperties } from '@xlvane/exceljs';
 * const value: WorkbookProperties | undefined = { date1904: false };
 * console.log(value);
 * ```
 */
export interface WorkbookProperties {
  date1904: boolean;
}

/**
 * The `zip` option of `writeBuffer()` (exceljs `JSZipGeneratorOptions`, index.d.ts L1447); xlvane uses `compressionOptions.level`.
 *
 * @example
 * ```ts
 * import type { JSZipGeneratorOptions } from '@xlvane/exceljs';
 * const value: JSZipGeneratorOptions | undefined = { compression: 'DEFLATE', compressionOptions: { level: 9 } };
 * console.log(value);
 * ```
 */
export interface JSZipGeneratorOptions {
  compression: 'STORE' | 'DEFLATE';
  compressionOptions: null | { level: number };
}

/**
 * `workbook.xlsx` (exceljs `Xlsx`, index.d.ts L1474).
 *
 * @example
 * ```ts
 * import type { Xlsx } from '@xlvane/exceljs';
 * const value: Xlsx | undefined = undefined;
 * console.log(value);
 * ```
 */
export type Xlsx = XlsxIO;

/**
 * `workbook.csv` (exceljs `Csv`, index.d.ts L1583).
 *
 * @example
 * ```ts
 * import type { Csv } from '@xlvane/exceljs';
 * const value: Csv | undefined = undefined;
 * console.log(value);
 * ```
 */
export type Csv = CsvIO;

/**
 * fast-csv parser options accepted by `csv.read()` (exceljs index.d.ts L1512).
 *
 * @example
 * ```ts
 * import type { FastCsvParserOptionsArgs } from '@xlvane/exceljs';
 * const value: FastCsvParserOptionsArgs | undefined = { delimiter: ';' } as FastCsvParserOptionsArgs;
 * console.log(value);
 * ```
 */
export interface FastCsvParserOptionsArgs {
  objectMode: boolean;
  delimiter: string;
  quote: string | null;
  escape: string;
  headers: boolean | ((headers: (string | undefined | null)[]) => (string | undefined | null)[]) | (string | undefined | null)[];
  renameHeaders: boolean;
  ignoreEmpty: boolean;
  comment: string;
  strictColumnHandling: boolean;
  discardUnmappedColumns: boolean;
  trim: boolean;
  ltrim: boolean;
  rtrim: boolean;
  encoding: string;
  maxRows: number;
  skipLines: number;
  skipRows: number;
}

/**
 * fast-csv formatter options accepted by `csv.write()` (exceljs index.d.ts L1550).
 *
 * @example
 * ```ts
 * import type { FastCsvFormatterOptionsArgs } from '@xlvane/exceljs';
 * const value: FastCsvFormatterOptionsArgs | undefined = { delimiter: ';' } as FastCsvFormatterOptionsArgs;
 * console.log(value);
 * ```
 */
export interface FastCsvFormatterOptionsArgs {
  objectMode: boolean;
  delimiter: string;
  rowDelimiter: string;
  quote: string | boolean;
  escape: string;
  quoteColumns: boolean | boolean[] | { [column: string]: boolean };
  quoteHeaders: boolean | boolean[] | { [column: string]: boolean };
  headers: null | boolean | string[];
  includeEndRowDelimiter: boolean;
  writeBOM: boolean;
  transform: (row: unknown) => unknown;
  alwaysWriteHeaders: boolean;
}

/**
 * The data of a worksheet as exceljs `WorksheetModel` (index.d.ts L984). `worksheet.model` returns xlvane's own snapshot.
 *
 * @example
 * ```ts
 * import type { WorksheetModel } from '@xlvane/exceljs';
 * const value: WorksheetModel | undefined = undefined;
 * console.log(value);
 * ```
 */
export interface WorksheetModel {
  id: number;
  name: string;
  properties: Partial<WorksheetProperties>;
  pageSetup: Partial<PageSetup>;
  headerFooter: Partial<HeaderFooter>;
  rowBreaks: RowBreak[];
  views: Partial<WorksheetView>[];
  autoFilter: AutoFilter;
  media: Media[];
  merges: string[];
}

/**
 * The data of a workbook as exceljs `WorkbookModel` (index.d.ts L1682). `workbook.model` returns xlvane's own snapshot.
 *
 * @example
 * ```ts
 * import type { WorkbookModel } from '@xlvane/exceljs';
 * const value: WorkbookModel | undefined = undefined;
 * console.log(value);
 * ```
 */
export interface WorkbookModel {
  creator: string;
  lastModifiedBy: string;
  lastPrinted: Date;
  created: Date;
  modified: Date;
  properties: WorkbookProperties;
  worksheets: Worksheet[];
  sheets: WorksheetModel[];
  definedNames: { name: string; ranges: string[] }[];
  views: Record<string, unknown>[];
  company: string;
  manager: string;
  title: string;
  subject: string;
  keywords: string;
  category: string;
  description: string;
  language: string;
  revision: Date;
  contentStatus: string;
  themes: string[];
  media: Media[];
}

/**
 * Settings shared by every worksheet view (exceljs `WorksheetViewCommon`, index.d.ts L33).
 *
 * @example
 * ```ts
 * import type { WorksheetViewCommon } from '@xlvane/exceljs';
 *
 * const value: WorksheetViewCommon | undefined = undefined;
 * console.log(value);
 * ```
 */
export interface WorksheetViewCommon {
  rightToLeft: boolean;
  activeCell: string;
  showRuler: boolean;
  showRowColHeaders: boolean;
  showGridLines: boolean;
  zoomScale: number;
  zoomScaleNormal: number;
}

/**
 * A normal worksheet view (exceljs `WorksheetViewNormal`, index.d.ts L72).
 *
 * @example
 * ```ts
 * import type { WorksheetViewNormal } from '@xlvane/exceljs';
 *
 * const value: WorksheetViewNormal | undefined = { state: 'normal', style: 'pageLayout' };
 * console.log(value);
 * ```
 */
export interface WorksheetViewNormal {
  state: 'normal';
  style: 'pageBreakPreview' | 'pageLayout';
}

/**
 * A view with frozen rows and columns (exceljs `WorksheetViewFrozen`, index.d.ts L84).
 *
 * @example
 * ```ts
 * import type { WorksheetViewFrozen } from '@xlvane/exceljs';
 *
 * const value: WorksheetViewFrozen | undefined = { state: 'frozen', ySplit: 1 };
 * console.log(value);
 * ```
 */
export interface WorksheetViewFrozen {
  state: 'frozen';
  style?: 'pageBreakPreview';
  xSplit?: number;
  ySplit?: number;
  topLeftCell?: string;
}

/**
 * A view split into panes (exceljs `WorksheetViewSplit`, index.d.ts L113).
 *
 * @example
 * ```ts
 * import type { WorksheetViewSplit } from '@xlvane/exceljs';
 *
 * const value: WorksheetViewSplit | undefined = { state: 'split', xSplit: 2000, ySplit: 3000 };
 * console.log(value);
 * ```
 */
export interface WorksheetViewSplit {
  state: 'split';
  style?: 'pageBreakPreview' | 'pageLayout';
  xSplit?: number;
  ySplit?: number;
  topLeftCell?: string;
  activePane?: 'topLeft' | 'topRight' | 'bottomLeft' | 'bottomRight';
}

/**
 * A window of the workbook (exceljs `WorkbookView`, index.d.ts L151).
 *
 * @example
 * ```ts
 * import type { WorkbookView } from '@xlvane/exceljs';
 *
 * const value: WorkbookView | undefined = { x: 0, y: 0, width: 10000, height: 20000, firstSheet: 0, activeTab: 0, visibility: 'visible' };
 * console.log(value);
 * ```
 */
export interface WorkbookView {
  x: number;
  y: number;
  width: number;
  height: number;
  firstSheet: number;
  activeTab: number;
  visibility: string;
}

/**
 * A pattern fill (exceljs `FillPattern`, index.d.ts L167).
 *
 * @example
 * ```ts
 * import type { FillPattern } from '@xlvane/exceljs';
 *
 * const value: FillPattern | undefined = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
 * console.log(value);
 * ```
 */
export interface FillPattern {
  type: 'pattern';
  pattern: FillPatterns;
  fgColor?: Partial<Color>;
  bgColor?: Partial<Color>;
}

/**
 * A cell that shares the formula of a master cell (exceljs `CellSharedFormulaValue`, index.d.ts L357).
 *
 * @example
 * ```ts
 * import type { CellSharedFormulaValue } from '@xlvane/exceljs';
 *
 * const value: CellSharedFormulaValue | undefined = { sharedFormula: 'A1', result: 2 };
 * console.log(value);
 * ```
 */
export interface CellSharedFormulaValue {
  sharedFormula: string;
  readonly formula?: string;
  result?: number | string | boolean | Date | CellErrorValue;
  date1904?: boolean;
}

/**
 * A picture anchor (exceljs `IAnchor`, index.d.ts L893).
 *
 * @example
 * ```ts
 * import type { IAnchor } from '@xlvane/exceljs';
 *
 * const value: IAnchor | undefined = undefined;
 * console.log(value);
 * ```
 */
export interface IAnchor {
  col: number;
  row: number;
  nativeCol: number;
  nativeRow: number;
  nativeColOff: number;
  nativeRowOff: number;
}

/**
 * Fields shared by conditional formatting rules (exceljs `ConditionalFormattingBaseRule`, index.d.ts L1015).
 *
 * @example
 * ```ts
 * import type { ConditionalFormattingBaseRule } from '@xlvane/exceljs';
 *
 * const value: ConditionalFormattingBaseRule | undefined = { priority: 1 };
 * console.log(value);
 * ```
 */
export interface ConditionalFormattingBaseRule {
  priority: number;
  style?: Partial<Style>;
}

/**
 * A `top10` conditional formatting rule (exceljs `Top10RuleType`, index.d.ts L1031).
 *
 * @example
 * ```ts
 * import type { Top10RuleType } from '@xlvane/exceljs';
 *
 * const value: Top10RuleType | undefined = { type: 'top10', rank: 5, percent: false, bottom: false };
 * console.log(value);
 * ```
 */
export interface Top10RuleType {
  type: 'top10';
  priority?: number;
  style?: Partial<Style>;
  rank: number;
  percent: boolean;
  bottom: boolean;
}

/**
 * `worksheet.properties` (exceljs `WorksheetProperties`, index.d.ts L1392).
 *
 * @example
 * ```ts
 * import type { WorksheetProperties } from '@xlvane/exceljs';
 *
 * const value: WorksheetProperties | undefined = { tabColor: { argb: 'FFFF0000' } } as WorksheetProperties;
 * console.log(value);
 * ```
 */
export interface WorksheetProperties {
  tabColor: Partial<Color>;
  outlineLevelCol: number;
  outlineLevelRow: number;
  outlineProperties: { summaryBelow: boolean; summaryRight: boolean };
  defaultRowHeight: number;
  defaultColWidth?: number;
  dyDescent: number;
  showGridLines: boolean;
}

/**
 * Options of `addWorksheet()` (exceljs `AddWorksheetOptions`, index.d.ts L1431).
 *
 * @example
 * ```ts
 * import type { AddWorksheetOptions } from '@xlvane/exceljs';
 *
 * const value: AddWorksheetOptions | undefined = { state: 'hidden' } as Partial<AddWorksheetOptions> as AddWorksheetOptions;
 * console.log(value);
 * ```
 */
export interface AddWorksheetOptions {
  properties: Partial<WorksheetProperties>;
  pageSetup: Partial<PageSetup>;
  headerFooter: Partial<HeaderFooter>;
  views: Array<Partial<WorksheetView>>;
  state: WorksheetState;
}

/**
 * One defined name and its ranges (exceljs `DefinedNamesRanges`, index.d.ts L1654).
 *
 * @example
 * ```ts
 * import type { DefinedNamesRanges } from '@xlvane/exceljs';
 *
 * const value: DefinedNamesRanges | undefined = { name: 'Totals', ranges: ['Sheet1!$A$1'] };
 * console.log(value);
 * ```
 */
export interface DefinedNamesRanges {
  name: string;
  ranges: string[];
}

/**
 * `workbook.definedNames.model` (exceljs `DefinedNamesModel`, index.d.ts L1659).
 *
 * @example
 * ```ts
 * import type { DefinedNamesModel } from '@xlvane/exceljs';
 *
 * const value: DefinedNamesModel | undefined = [];
 * console.log(value);
 * ```
 */
export type DefinedNamesModel = DefinedNamesRanges[];

/**
 * The style of a table (exceljs `TableStyleProperties`, index.d.ts L1783). `theme` is any table style name, such as `TableStyleMedium2`.
 *
 * @example
 * ```ts
 * import type { TableStyleProperties } from '@xlvane/exceljs';
 *
 * const value: TableStyleProperties | undefined = { theme: 'TableStyleMedium2', showRowStripes: true };
 * console.log(value);
 * ```
 */
export interface TableStyleProperties {
  theme?: string;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
}

/**
 * exceljs `config` (index.d.ts L1906): declared by exceljs for a custom Promise library; exceljs does
 * not export it at run time either.
 *
 * @example
 * ```ts
 * import type { config } from '@xlvane/exceljs';
 *
 * type SetValue = typeof config.setValue;
 * const kind: SetValue | undefined = undefined;
 * console.log(kind);
 * ```
 */
export declare namespace config {
  function setValue(key: 'promise', promise: unknown): void;
}
