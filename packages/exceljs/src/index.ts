// Entry point of @xlvane/exceljs: the commonly used exceljs 4.4.0 API on top of xlvane, so that
// `"exceljs": "npm:@xlvane/exceljs@^1"` switches a project with one line (docs/exceljs-compat.md).
//
// Module shape (docs/decisions.md D-501): the default export is an object with the same members
// as exceljs's CommonJS export (lib/exceljs.nodejs.js: Workbook, ModelContainer, stream.xlsx,
// and the enums), and the same object is exported under the string name "module.exports", which
// Node.js returns from require() (v22.12.0+, modules.md "Loading ECMAScript modules using
// require()"). Named exports serve `import { Workbook } from 'exceljs'`.
import { ErrorValue, FormulaType, notSupported, ValueType } from './model';
import { WorkbookReader, WorkbookWriter } from './stream';
import { Workbook } from './workbook';
import type * as CompatTypes from './compat-types';
import type * as CsvTypes from './csv';
import type * as NameTypes from './defined-names';
import type * as MediaTypes from './media';
import type * as ModelTypes from './model';
import type * as RuleTypes from './sheet-rules';
import type * as StreamTypes from './stream';
import type * as ApiTypes from './types';
import type * as WorkbookTypes from './workbook';

export { Cell, Column, ErrorValue, FormulaType, Row, ValueType, Worksheet } from './model';
export { WorkbookReader, WorkbookWriter, WorksheetReader, WorksheetWriter, type WorkbookReaderOptions, type WorkbookWriterOptions } from './stream';
export { CsvIO, type CsvReadOptions, type CsvWriteOptions } from './csv';
export { CellMatrix, DefinedNames, type DecodedLocation, type NameCell } from './defined-names';
export { Workbook, XlsxIO, type XlsxReadOptions, type XlsxWriteOptions } from './workbook';
export type {
  Alignment,
  Border,
  Borders,
  BorderStyle,
  CellErrorValue,
  CellFormulaValue,
  CellValue,
  Color,
  Comment,
  ColumnDefn,
  ConditionalFormattingOptions,
  ConditionalFormattingRule,
  DataValidation,
  Fill,
  FillPatterns,
  Font,
  HeaderFooter,
  PageSetup,
  Protection,
  RowBreak,
  Style,
  WorksheetOptions,
  WorksheetProtection,
  WorksheetView,
} from './types';
export { DataValidations } from './sheet-rules';
import { PaperSize } from './compat-types';
export { PaperSize, type config } from './compat-types';
export type {
  Buffer,
  GradientStop,
  FillGradientAngle,
  FillGradientPath,
  BorderDiagonal,
  Margins,
  DataValidationOperator,
  RichText,
  CellRichTextValue,
  CellHyperlinkValue,
  Address,
  CellModel,
  RowModel,
  AutoFilter,
  Image,
  ImageRange,
  IAnchorPosition,
  ImagePosition,
  ImageHyperlinkValue,
  Location,
  Range,
  Media,
  WorksheetState,
  CellIsOperators,
  ContainsTextOperators,
  TimePeriodTypes,
  IconSetTypes,
  CfvoTypes,
  Cvfo,
  ExpressionRuleType,
  CellIsRuleType,
  AboveAverageRuleType,
  ColorScaleRuleType,
  IconSetRuleType,
  ContainsTextRuleType,
  TimePeriodRuleType,
  DataBarRuleType,
  RowValues,
  CalculationProperties,
  WorkbookProperties,
  JSZipGeneratorOptions,
  Xlsx,
  Csv,
  FastCsvParserOptionsArgs,
  FastCsvFormatterOptionsArgs,
  WorksheetModel,
  WorkbookModel,
  WorksheetViewCommon,
  WorksheetViewNormal,
  WorksheetViewFrozen,
  WorksheetViewSplit,
  WorkbookView,
  FillPattern,
  CellSharedFormulaValue,
  IAnchor,
  ConditionalFormattingBaseRule,
  Top10RuleType,
  WorksheetProperties,
  AddWorksheetOptions,
  DefinedNamesRanges,
  DefinedNamesModel,
  TableStyleProperties,
} from './compat-types';
export { Anchor, Table, TableColumn, type ImageData, type ImageRangeInput, type TableColumnProperties, type TableProperties } from './media';

/**
 * exceljs `RelationshipType` (lib/doc/enums.js), kept for code that reads it.
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * console.log(ExcelJS.RelationshipType.Worksheet); // 2
 * ```
 */
export const RelationshipType: {
  readonly None: 0;
  readonly OfficeDocument: 1;
  readonly Worksheet: 2;
  readonly CalcChain: 3;
  readonly SharedStrings: 4;
  readonly Styles: 5;
  readonly Theme: 6;
  readonly Hyperlink: 7;
} = Object.freeze({ None: 0, OfficeDocument: 1, Worksheet: 2, CalcChain: 3, SharedStrings: 4, Styles: 5, Theme: 6, Hyperlink: 7 } as const);
/** A value of {@link RelationshipType} (usable as a type, as the exceljs enum). */
export type RelationshipType = (typeof RelationshipType)[keyof typeof RelationshipType];

/**
 * exceljs `DocumentType` (lib/doc/enums.js).
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * console.log(ExcelJS.DocumentType.Xlsx); // 1
 * ```
 */
export const DocumentType: { readonly Xlsx: 1 } = Object.freeze({ Xlsx: 1 } as const);
/** A value of {@link DocumentType} (usable as a type, as the exceljs enum). */
export type DocumentType = (typeof DocumentType)[keyof typeof DocumentType];

/**
 * exceljs `ReadingOrder` (lib/doc/enums.js).
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * console.log(ExcelJS.ReadingOrder.RightToLeft); // 2
 * ```
 */
export const ReadingOrder: { readonly LeftToRight: 1; readonly RightToLeft: 2 } = Object.freeze({ LeftToRight: 1, RightToLeft: 2 } as const);
/** A value of {@link ReadingOrder} (usable as a type, as the exceljs enum). */
export type ReadingOrder = (typeof ReadingOrder)[keyof typeof ReadingOrder];

/**
 * The streaming API, as `ExcelJS.stream`.
 *
 * @example
 * ```ts
 * import { stream } from '@xlvane/exceljs';
 *
 * const writer = new stream.xlsx.WorkbookWriter({ filename: 'out.xlsx' });
 * writer.addWorksheet('Data').addRow([1]).commit();
 * await writer.commit();
 * ```
 */
export const stream: { readonly xlsx: { readonly WorkbookWriter: typeof WorkbookWriter; readonly WorkbookReader: typeof WorkbookReader } } = Object.freeze({
  xlsx: Object.freeze({ WorkbookWriter, WorkbookReader }),
});
/** Types of the streaming API as exceljs declares them (`ExcelJS.stream.xlsx.WorkbookWriterOptions`). */
export declare namespace stream {
  namespace xlsx {
    type WorkbookWriter = XlvTypeOfWorkbookWriter;
    type WorkbookReader = XlvTypeOfWorkbookReader;
    type WorksheetWriter = XlvTypeOfWorksheetWriter;
    type WorksheetReader = XlvTypeOfWorksheetReader;
    type WorkbookWriterOptions = XlvTypeOfWorkbookWriterOptions;
    type WorkbookReaderOptions = XlvTypeOfWorkbookReaderOptions;
    /** The name exceljs's index.d.ts uses for the reader options. */
    type WorkbookStreamReaderOptions = XlvTypeOfWorkbookReaderOptions;
  }
}

/**
 * Kept for exceljs compatibility (`ExcelJS.ModelContainer`); not supported.
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * console.log(typeof ExcelJS.ModelContainer); // function
 * ```
 */
export class ModelContainer {
  constructor() {
    notSupported('ModelContainer');
  }
}

/**
 * The shape of the default export (and of `require('exceljs')`).
 *
 * @example
 * ```ts
 * import ExcelJS, { type ExcelJSModule } from '@xlvane/exceljs';
 *
 * const lib: ExcelJSModule = ExcelJS;
 * const workbook = new lib.Workbook();
 * ```
 */
export interface ExcelJSModule {
  readonly Workbook: typeof Workbook;
  readonly ModelContainer: typeof ModelContainer;
  readonly stream: typeof stream;
  readonly ValueType: typeof ValueType;
  readonly FormulaType: typeof FormulaType;
  readonly ErrorValue: typeof ErrorValue;
  readonly RelationshipType: typeof RelationshipType;
  readonly DocumentType: typeof DocumentType;
  readonly ReadingOrder: typeof ReadingOrder;
  readonly PaperSize: typeof PaperSize;
}

/**
 * The exceljs-compatible API as one object: `import ExcelJS from 'exceljs'` and
 * `const ExcelJS = require('exceljs')` both give it.
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const workbook = new ExcelJS.Workbook();
 * const sheet = workbook.addWorksheet('Sheet1');
 * sheet.addRow(['Hello', 'world']);
 * const buffer = await workbook.xlsx.writeBuffer();
 * console.log(buffer.length > 0);
 * ```
 */
const ExcelJS: ExcelJSModule = Object.freeze({ Workbook, ModelContainer, stream, ValueType, FormulaType, ErrorValue, RelationshipType, DocumentType, ReadingOrder, PaperSize });


type ModelContainerInstance = ModelContainer;

// Unexported aliases that the namespaces below refer to: a type alias inside a namespace cannot
// name the outer declaration of the same name directly.
type XlvTypeOfWorkbookWriter = StreamTypes.WorkbookWriter;
type XlvTypeOfWorkbookReader = StreamTypes.WorkbookReader;
type XlvTypeOfWorksheetWriter = StreamTypes.WorksheetWriter;
type XlvTypeOfWorksheetReader = StreamTypes.WorksheetReader;
type XlvTypeOfWorkbookWriterOptions = StreamTypes.WorkbookWriterOptions;
type XlvTypeOfWorkbookReaderOptions = StreamTypes.WorkbookReaderOptions;
type XlvTypeOfColor = ApiTypes.Color;
type XlvTypeOfFont = ApiTypes.Font;
type XlvTypeOfFillPatterns = ApiTypes.FillPatterns;
type XlvTypeOfFill = ApiTypes.Fill;
type XlvTypeOfBorderStyle = ApiTypes.BorderStyle;
type XlvTypeOfBorder = ApiTypes.Border;
type XlvTypeOfBorders = ApiTypes.Borders;
type XlvTypeOfAlignment = ApiTypes.Alignment;
type XlvTypeOfStyle = ApiTypes.Style;
type XlvTypeOfProtection = ApiTypes.Protection;
type XlvTypeOfCellErrorValue = ApiTypes.CellErrorValue;
type XlvTypeOfCellFormulaValue = ApiTypes.CellFormulaValue;
type XlvTypeOfCellValue = ApiTypes.CellValue;
type XlvTypeOfColumnDefn = ApiTypes.ColumnDefn;
type XlvTypeOfWorksheetView = ApiTypes.WorksheetView;
type XlvTypeOfWorksheetOptions = ApiTypes.WorksheetOptions;
type XlvTypeOfComment = ApiTypes.Comment;
type XlvTypeOfDataValidation = ApiTypes.DataValidation;
type XlvTypeOfConditionalFormattingRule = ApiTypes.ConditionalFormattingRule;
type XlvTypeOfConditionalFormattingOptions = ApiTypes.ConditionalFormattingOptions;
type XlvTypeOfPageSetup = ApiTypes.PageSetup;
type XlvTypeOfHeaderFooter = ApiTypes.HeaderFooter;
type XlvTypeOfRowBreak = ApiTypes.RowBreak;
type XlvTypeOfWorksheetProtection = ApiTypes.WorksheetProtection;
type XlvTypeOfBuffer = CompatTypes.Buffer;
type XlvTypeOfGradientStop = CompatTypes.GradientStop;
type XlvTypeOfFillGradientAngle = CompatTypes.FillGradientAngle;
type XlvTypeOfFillGradientPath = CompatTypes.FillGradientPath;
type XlvTypeOfBorderDiagonal = CompatTypes.BorderDiagonal;
type XlvTypeOfMargins = CompatTypes.Margins;
type XlvTypeOfDataValidationOperator = CompatTypes.DataValidationOperator;
type XlvTypeOfRichText = CompatTypes.RichText;
type XlvTypeOfCellRichTextValue = CompatTypes.CellRichTextValue;
type XlvTypeOfCellHyperlinkValue = CompatTypes.CellHyperlinkValue;
type XlvTypeOfAddress = CompatTypes.Address;
type XlvTypeOfCellModel = CompatTypes.CellModel;
type XlvTypeOfRowModel = CompatTypes.RowModel;
type XlvTypeOfAutoFilter = CompatTypes.AutoFilter;
type XlvTypeOfImage = CompatTypes.Image;
type XlvTypeOfImageRange = CompatTypes.ImageRange;
type XlvTypeOfIAnchorPosition = CompatTypes.IAnchorPosition;
type XlvTypeOfImagePosition = CompatTypes.ImagePosition;
type XlvTypeOfImageHyperlinkValue = CompatTypes.ImageHyperlinkValue;
type XlvTypeOfLocation = CompatTypes.Location;
type XlvTypeOfRange = CompatTypes.Range;
type XlvTypeOfMedia = CompatTypes.Media;
type XlvTypeOfWorksheetState = CompatTypes.WorksheetState;
type XlvTypeOfCellIsOperators = CompatTypes.CellIsOperators;
type XlvTypeOfContainsTextOperators = CompatTypes.ContainsTextOperators;
type XlvTypeOfTimePeriodTypes = CompatTypes.TimePeriodTypes;
type XlvTypeOfIconSetTypes = CompatTypes.IconSetTypes;
type XlvTypeOfCfvoTypes = CompatTypes.CfvoTypes;
type XlvTypeOfCvfo = CompatTypes.Cvfo;
type XlvTypeOfExpressionRuleType = CompatTypes.ExpressionRuleType;
type XlvTypeOfCellIsRuleType = CompatTypes.CellIsRuleType;
type XlvTypeOfAboveAverageRuleType = CompatTypes.AboveAverageRuleType;
type XlvTypeOfColorScaleRuleType = CompatTypes.ColorScaleRuleType;
type XlvTypeOfIconSetRuleType = CompatTypes.IconSetRuleType;
type XlvTypeOfContainsTextRuleType = CompatTypes.ContainsTextRuleType;
type XlvTypeOfTimePeriodRuleType = CompatTypes.TimePeriodRuleType;
type XlvTypeOfDataBarRuleType = CompatTypes.DataBarRuleType;
type XlvTypeOfRowValues = CompatTypes.RowValues;
type XlvTypeOfCalculationProperties = CompatTypes.CalculationProperties;
type XlvTypeOfWorkbookProperties = CompatTypes.WorkbookProperties;
type XlvTypeOfJSZipGeneratorOptions = CompatTypes.JSZipGeneratorOptions;
type XlvTypeOfXlsx = CompatTypes.Xlsx;
type XlvTypeOfCsv = CompatTypes.Csv;
type XlvTypeOfFastCsvParserOptionsArgs = CompatTypes.FastCsvParserOptionsArgs;
type XlvTypeOfFastCsvFormatterOptionsArgs = CompatTypes.FastCsvFormatterOptionsArgs;
type XlvTypeOfWorksheetModel = CompatTypes.WorksheetModel;
type XlvTypeOfWorkbookModel = CompatTypes.WorkbookModel;
type XlvTypeOfWorksheetViewCommon = CompatTypes.WorksheetViewCommon;
type XlvTypeOfWorksheetViewNormal = CompatTypes.WorksheetViewNormal;
type XlvTypeOfWorksheetViewFrozen = CompatTypes.WorksheetViewFrozen;
type XlvTypeOfWorksheetViewSplit = CompatTypes.WorksheetViewSplit;
type XlvTypeOfWorkbookView = CompatTypes.WorkbookView;
type XlvTypeOfFillPattern = CompatTypes.FillPattern;
type XlvTypeOfCellSharedFormulaValue = CompatTypes.CellSharedFormulaValue;
type XlvTypeOfIAnchor = CompatTypes.IAnchor;
type XlvTypeOfConditionalFormattingBaseRule = CompatTypes.ConditionalFormattingBaseRule;
type XlvTypeOfTop10RuleType = CompatTypes.Top10RuleType;
type XlvTypeOfWorksheetProperties = CompatTypes.WorksheetProperties;
type XlvTypeOfAddWorksheetOptions = CompatTypes.AddWorksheetOptions;
type XlvTypeOfDefinedNamesRanges = CompatTypes.DefinedNamesRanges;
type XlvTypeOfDefinedNamesModel = CompatTypes.DefinedNamesModel;
type XlvTypeOfTableStyleProperties = CompatTypes.TableStyleProperties;
type XlvTypeOfPaperSize = CompatTypes.PaperSize;
type XlvTypeOfCell = ModelTypes.Cell;
type XlvTypeOfColumn = ModelTypes.Column;
type XlvTypeOfRow = ModelTypes.Row;
type XlvTypeOfWorksheet = ModelTypes.Worksheet;
type XlvTypeOfValueType = ModelTypes.ValueType;
type XlvTypeOfFormulaType = ModelTypes.FormulaType;
type XlvTypeOfErrorValue = ModelTypes.ErrorValue;
type XlvTypeOfCsvIO = CsvTypes.CsvIO;
type XlvTypeOfCsvReadOptions = CsvTypes.CsvReadOptions;
type XlvTypeOfCsvWriteOptions = CsvTypes.CsvWriteOptions;
type XlvTypeOfCellMatrix = NameTypes.CellMatrix;
type XlvTypeOfDefinedNames = NameTypes.DefinedNames;
type XlvTypeOfDecodedLocation = NameTypes.DecodedLocation;
type XlvTypeOfNameCell = NameTypes.NameCell;
type XlvTypeOfWorkbook = WorkbookTypes.Workbook;
type XlvTypeOfXlsxIO = WorkbookTypes.XlsxIO;
type XlvTypeOfXlsxReadOptions = WorkbookTypes.XlsxReadOptions;
type XlvTypeOfXlsxWriteOptions = WorkbookTypes.XlsxWriteOptions;
type XlvTypeOfDataValidations = RuleTypes.DataValidations;
type XlvTypeOfAnchor = MediaTypes.Anchor;
type XlvTypeOfTable = MediaTypes.Table;
type XlvTypeOfTableColumn = MediaTypes.TableColumn;
type XlvTypeOfImageData = MediaTypes.ImageData;
type XlvTypeOfImageRangeInput = MediaTypes.ImageRangeInput;
type XlvTypeOfTableColumnProperties = MediaTypes.TableColumnProperties;
type XlvTypeOfTableProperties = MediaTypes.TableProperties;
type XlvTypeOfRelationshipType = (typeof RelationshipType)[keyof typeof RelationshipType];
type XlvTypeOfDocumentType = (typeof DocumentType)[keyof typeof DocumentType];
type XlvTypeOfReadingOrder = (typeof ReadingOrder)[keyof typeof ReadingOrder];
type XlvTypeOfModelContainer = ModelContainerInstance;

/**
 * The types of the exceljs-compatible API under the default export, so that exceljs code such as
 * `let sheet: ExcelJS.Worksheet` type-checks with `import ExcelJS from 'exceljs'` and with
 * `import ExcelJS = require('exceljs')` (D-R01).
 */
declare namespace ExcelJS {
  type Color = XlvTypeOfColor;
  type Font = XlvTypeOfFont;
  type FillPatterns = XlvTypeOfFillPatterns;
  type Fill = XlvTypeOfFill;
  type BorderStyle = XlvTypeOfBorderStyle;
  type Border = XlvTypeOfBorder;
  type Borders = XlvTypeOfBorders;
  type Alignment = XlvTypeOfAlignment;
  type Style = XlvTypeOfStyle;
  type Protection = XlvTypeOfProtection;
  type CellErrorValue = XlvTypeOfCellErrorValue;
  type CellFormulaValue = XlvTypeOfCellFormulaValue;
  type CellValue = XlvTypeOfCellValue;
  type ColumnDefn = XlvTypeOfColumnDefn;
  type WorksheetView = XlvTypeOfWorksheetView;
  type WorksheetOptions = XlvTypeOfWorksheetOptions;
  type Comment = XlvTypeOfComment;
  type DataValidation = XlvTypeOfDataValidation;
  type ConditionalFormattingRule = XlvTypeOfConditionalFormattingRule;
  type ConditionalFormattingOptions = XlvTypeOfConditionalFormattingOptions;
  type PageSetup = XlvTypeOfPageSetup;
  type HeaderFooter = XlvTypeOfHeaderFooter;
  type RowBreak = XlvTypeOfRowBreak;
  type WorksheetProtection = XlvTypeOfWorksheetProtection;
  type Buffer = XlvTypeOfBuffer;
  type GradientStop = XlvTypeOfGradientStop;
  type FillGradientAngle = XlvTypeOfFillGradientAngle;
  type FillGradientPath = XlvTypeOfFillGradientPath;
  type BorderDiagonal = XlvTypeOfBorderDiagonal;
  type Margins = XlvTypeOfMargins;
  type DataValidationOperator = XlvTypeOfDataValidationOperator;
  type RichText = XlvTypeOfRichText;
  type CellRichTextValue = XlvTypeOfCellRichTextValue;
  type CellHyperlinkValue = XlvTypeOfCellHyperlinkValue;
  type Address = XlvTypeOfAddress;
  type CellModel = XlvTypeOfCellModel;
  type RowModel = XlvTypeOfRowModel;
  type AutoFilter = XlvTypeOfAutoFilter;
  type Image = XlvTypeOfImage;
  type ImageRange = XlvTypeOfImageRange;
  type IAnchorPosition = XlvTypeOfIAnchorPosition;
  type ImagePosition = XlvTypeOfImagePosition;
  type ImageHyperlinkValue = XlvTypeOfImageHyperlinkValue;
  type Location = XlvTypeOfLocation;
  type Range = XlvTypeOfRange;
  type Media = XlvTypeOfMedia;
  type WorksheetState = XlvTypeOfWorksheetState;
  type CellIsOperators = XlvTypeOfCellIsOperators;
  type ContainsTextOperators = XlvTypeOfContainsTextOperators;
  type TimePeriodTypes = XlvTypeOfTimePeriodTypes;
  type IconSetTypes = XlvTypeOfIconSetTypes;
  type CfvoTypes = XlvTypeOfCfvoTypes;
  type Cvfo = XlvTypeOfCvfo;
  type ExpressionRuleType = XlvTypeOfExpressionRuleType;
  type CellIsRuleType = XlvTypeOfCellIsRuleType;
  type AboveAverageRuleType = XlvTypeOfAboveAverageRuleType;
  type ColorScaleRuleType = XlvTypeOfColorScaleRuleType;
  type IconSetRuleType = XlvTypeOfIconSetRuleType;
  type ContainsTextRuleType = XlvTypeOfContainsTextRuleType;
  type TimePeriodRuleType = XlvTypeOfTimePeriodRuleType;
  type DataBarRuleType = XlvTypeOfDataBarRuleType;
  type RowValues = XlvTypeOfRowValues;
  type CalculationProperties = XlvTypeOfCalculationProperties;
  type WorkbookProperties = XlvTypeOfWorkbookProperties;
  type JSZipGeneratorOptions = XlvTypeOfJSZipGeneratorOptions;
  type Xlsx = XlvTypeOfXlsx;
  type Csv = XlvTypeOfCsv;
  type FastCsvParserOptionsArgs = XlvTypeOfFastCsvParserOptionsArgs;
  type FastCsvFormatterOptionsArgs = XlvTypeOfFastCsvFormatterOptionsArgs;
  type WorksheetModel = XlvTypeOfWorksheetModel;
  type WorkbookModel = XlvTypeOfWorkbookModel;
  type WorksheetViewCommon = XlvTypeOfWorksheetViewCommon;
  type WorksheetViewNormal = XlvTypeOfWorksheetViewNormal;
  type WorksheetViewFrozen = XlvTypeOfWorksheetViewFrozen;
  type WorksheetViewSplit = XlvTypeOfWorksheetViewSplit;
  type WorkbookView = XlvTypeOfWorkbookView;
  type FillPattern = XlvTypeOfFillPattern;
  type CellSharedFormulaValue = XlvTypeOfCellSharedFormulaValue;
  type IAnchor = XlvTypeOfIAnchor;
  type ConditionalFormattingBaseRule = XlvTypeOfConditionalFormattingBaseRule;
  type Top10RuleType = XlvTypeOfTop10RuleType;
  type WorksheetProperties = XlvTypeOfWorksheetProperties;
  type AddWorksheetOptions = XlvTypeOfAddWorksheetOptions;
  type DefinedNamesRanges = XlvTypeOfDefinedNamesRanges;
  type DefinedNamesModel = XlvTypeOfDefinedNamesModel;
  type TableStyleProperties = XlvTypeOfTableStyleProperties;
  type PaperSize = XlvTypeOfPaperSize;
  type Cell = XlvTypeOfCell;
  type Column = XlvTypeOfColumn;
  type Row = XlvTypeOfRow;
  type Worksheet = XlvTypeOfWorksheet;
  type ValueType = XlvTypeOfValueType;
  type FormulaType = XlvTypeOfFormulaType;
  type ErrorValue = XlvTypeOfErrorValue;
  type WorkbookReader = XlvTypeOfWorkbookReader;
  type WorkbookWriter = XlvTypeOfWorkbookWriter;
  type WorksheetReader = XlvTypeOfWorksheetReader;
  type WorksheetWriter = XlvTypeOfWorksheetWriter;
  type WorkbookReaderOptions = XlvTypeOfWorkbookReaderOptions;
  type WorkbookWriterOptions = XlvTypeOfWorkbookWriterOptions;
  type CsvIO = XlvTypeOfCsvIO;
  type CsvReadOptions = XlvTypeOfCsvReadOptions;
  type CsvWriteOptions = XlvTypeOfCsvWriteOptions;
  type CellMatrix = XlvTypeOfCellMatrix;
  type DefinedNames = XlvTypeOfDefinedNames;
  type DecodedLocation = XlvTypeOfDecodedLocation;
  type NameCell = XlvTypeOfNameCell;
  type Workbook = XlvTypeOfWorkbook;
  type XlsxIO = XlvTypeOfXlsxIO;
  type XlsxReadOptions = XlvTypeOfXlsxReadOptions;
  type XlsxWriteOptions = XlvTypeOfXlsxWriteOptions;
  type DataValidations = XlvTypeOfDataValidations;
  type Anchor = XlvTypeOfAnchor;
  type Table = XlvTypeOfTable;
  type TableColumn = XlvTypeOfTableColumn;
  type ImageData = XlvTypeOfImageData;
  type ImageRangeInput = XlvTypeOfImageRangeInput;
  type TableColumnProperties = XlvTypeOfTableColumnProperties;
  type TableProperties = XlvTypeOfTableProperties;
  type RelationshipType = XlvTypeOfRelationshipType;
  type DocumentType = XlvTypeOfDocumentType;
  type ReadingOrder = XlvTypeOfReadingOrder;
  type ModelContainer = XlvTypeOfModelContainer;
  namespace stream {
    namespace xlsx {
      type WorkbookWriter = XlvTypeOfWorkbookWriter;
      type WorkbookReader = XlvTypeOfWorkbookReader;
      type WorksheetWriter = XlvTypeOfWorksheetWriter;
      type WorksheetReader = XlvTypeOfWorksheetReader;
      type WorkbookWriterOptions = XlvTypeOfWorkbookWriterOptions;
      type WorkbookReaderOptions = XlvTypeOfWorkbookReaderOptions;
      type WorkbookStreamReaderOptions = XlvTypeOfWorkbookReaderOptions;
    }
  }
}

export default ExcelJS;
export { ExcelJS as 'module.exports' };
