// Public entry point of xlvane. Runtime-neutral: no Node.js APIs (file helpers are in
// `xlvane/node`).
export { readWorkbook, type ReadOptions } from './read';
export {
  inspect,
  type InspectColumn,
  type InspectOptions,
  type InspectReport,
  type InspectSheet,
  type InspectTypeCounts,
  type InspectValue,
} from './inspect';
export {
  ERROR_CODES,
  WARNING_CODES,
  XlvError,
  isXlvError,
  type XlvDetails,
  type XlvDetailValue,
  type XlvErrorCode,
  type XlvErrorInit,
  type XlvErrorJSON,
  type XlvLocation,
  type XlvWarning,
  type XlvWarningCode,
} from './errors';
export { DEFAULT_READ_LIMITS, type ReadLimits } from './limits';
export type {
  Cell,
  CellType,
  CellValue,
  Conformance,
  DateSystem,
  Formula,
  FormulaKind,
  Sheet,
  SheetKind,
  SheetState,
  Workbook,
  WorkbookFormat,
} from './model';
