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
const ExcelJS: ExcelJSModule = Object.freeze({ Workbook, ModelContainer, stream, ValueType, FormulaType, ErrorValue, RelationshipType, DocumentType, ReadingOrder });

export default ExcelJS;
export { ExcelJS as 'module.exports' };
