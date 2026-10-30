/** Base URL of the error catalogue. Every code has a heading (and so an anchor) there. */
const DOCS_URL = 'https://github.com/DevXord/xlvane/blob/main/docs/errors.md';

/**
 * Every error code xlvane can throw. Codes are stable: a code is never reused for a different
 * condition. Each one is described in docs/errors.md.
 *
 * @example
 * ```ts
 * import { ERROR_CODES } from 'xlvane';
 * console.log(ERROR_CODES.includes('XLV_E_SHEET_NOT_FOUND')); // true
 * ```
 */
export const ERROR_CODES = [
  'XLV_E_INVALID_INPUT',
  'XLV_E_INVALID_OPTION',
  'XLV_E_LIMIT_EXCEEDED',
  'XLV_E_NOT_ZIP',
  'XLV_E_UNSUPPORTED_CONTAINER',
  'XLV_E_ZIP_CORRUPT',
  'XLV_E_ZIP_UNSUPPORTED',
  'XLV_E_XML_MALFORMED',
  'XLV_E_XML_DTD_FORBIDDEN',
  'XLV_E_XML_ENTITY_FORBIDDEN',
  'XLV_E_XML_ENCODING',
  'XLV_E_PART_MISSING',
  'XLV_E_PART_INVALID',
  'XLV_E_INVALID_CELL_REF',
  'XLV_E_SHEET_NOT_FOUND',
  'XLV_E_FILE_READ',
  'XLV_E_INVALID_VALUE',
  'XLV_E_INVALID_STYLE',
  'XLV_E_INVALID_SHEET_NAME',
  'XLV_E_DUPLICATE_SHEET_NAME',
  'XLV_E_OVERLAPPING_MERGE',
  'XLV_E_FILE_EXISTS',
  'XLV_E_FILE_WRITE',
  'XLV_E_HEADER_NOT_FOUND',
  'XLV_E_SCHEMA_INVALID',
  'XLV_E_ASYNC_SCHEMA',
  'XLV_E_TABLE_INVALID',
  'XLV_E_JSON_INVALID',
  'XLV_E_CSV_MALFORMED',
  'XLV_E_NOT_SUPPORTED',
  'XLV_E_USAGE',
  'XLV_E_NOT_INSTALLED',
  'XLV_E_INTERNAL',
  'XLV_E_PATH_NOT_ALLOWED',
] as const;

/**
 * Every warning code xlvane can report. Warnings never stop reading; they describe something
 * in the file that the caller may want to know about. Each one is described in docs/errors.md.
 *
 * @example
 * ```ts
 * import { WARNING_CODES } from 'xlvane';
 * console.log(WARNING_CODES.includes('XLV_W_MACROS_PRESENT')); // true
 * ```
 */
export const WARNING_CODES = [
  'XLV_W_DIMENSION_MISMATCH',
  'XLV_W_SHARED_FORMULA_NOT_EXPANDED',
  'XLV_W_FORMULA_NO_CACHED_VALUE',
  'XLV_W_VALUES_IN_MERGED_CELLS',
  'XLV_W_VERY_HIDDEN_SHEET',
  'XLV_W_MACROS_PRESENT',
  'XLV_W_EXTERNAL_LINKS',
  'XLV_W_DUPLICATE_CELL',
  'XLV_W_INVALID_VALUE',
  'XLV_W_SHEET_PART_MISSING',
  'XLV_W_PART_MISSING',
  'XLV_W_FORMULAS_NOT_RECALCULATED',
] as const;

/**
 * A stable error code, see {@link ERROR_CODES}.
 *
 * @example
 * ```ts
 * import { isXlvError, readWorkbook, type XlvErrorCode } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * try {
 *   readWorkbook(bytes);
 * } catch (error) {
 *   const code: XlvErrorCode | undefined = isXlvError(error) ? error.code : undefined;
 *   console.log(code === 'XLV_E_NOT_ZIP' ? 'not an xlsx file' : code);
 * }
 * ```
 */
export type XlvErrorCode = (typeof ERROR_CODES)[number];

/**
 * A stable warning code, see {@link WARNING_CODES}.
 *
 * @example
 * ```ts
 * import { readWorkbook, type XlvWarningCode } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const codes: XlvWarningCode[] = readWorkbook(bytes).warnings.map((warning) => warning.code);
 * console.log(codes.includes('XLV_W_MACROS_PRESENT'));
 * ```
 */
export type XlvWarningCode = (typeof WARNING_CODES)[number];

/**
 * Where in the input something happened. Only the fields that are known are present.
 *
 * @example
 * ```ts
 * import { isXlvError, readWorkbook, type XlvLocation } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * try {
 *   readWorkbook(bytes).getSheet('Sales');
 * } catch (error) {
 *   if (isXlvError(error)) {
 *     const where: XlvLocation = error.location;
 *     console.log(where.sheet, where.cell, where.part);
 *   }
 * }
 * ```
 */
export interface XlvLocation {
  /** Part inside the package, for example `/xl/worksheets/sheet1.xml`. */
  readonly part?: string;
  /** Sheet name as shown in the spreadsheet application. */
  readonly sheet?: string;
  /** Cell reference in A1 style, for example `B7`. */
  readonly cell?: string;
  /** 1-based line inside an XML part. */
  readonly line?: number;
  /** 1-based column inside an XML part. */
  readonly column?: number;
  /** Byte offset in the input file (ZIP structures). */
  readonly offset?: number;
}

/**
 * Value types allowed in {@link XlvError.details}; everything is JSON-safe.
 *
 * @example
 * ```ts
 * import { isXlvError, readWorkbook, type XlvDetailValue } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * try {
 *   readWorkbook(bytes).getSheet('Sales');
 * } catch (error) {
 *   if (isXlvError(error)) {
 *     const available: XlvDetailValue | undefined = error.details['available'];
 *     console.log(available);
 *   }
 * }
 * ```
 */
export type XlvDetailValue = string | number | boolean | null | readonly string[] | readonly number[];

/**
 * Machine-readable details of an error. Keys are chosen by xlvane, never taken from the file.
 *
 * @example
 * ```ts
 * import { isXlvError, readWorkbook, type XlvDetails } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * try {
 *   readWorkbook(bytes);
 * } catch (error) {
 *   if (isXlvError(error)) {
 *     const details: XlvDetails = error.details;
 *     console.log(Object.keys(details));
 *   }
 * }
 * ```
 */
export type XlvDetails = Readonly<Record<string, XlvDetailValue>>;

/**
 * The JSON form of an {@link XlvError}, as returned by {@link XlvError.toJSON}.
 *
 * @example
 * ```ts
 * import { isXlvError, readWorkbook, type XlvErrorJSON } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * try {
 *   readWorkbook(bytes);
 * } catch (error) {
 *   if (isXlvError(error)) {
 *     const json: XlvErrorJSON = error.toJSON();
 *     console.log(JSON.stringify(json));
 *   }
 * }
 * ```
 */
export interface XlvErrorJSON {
  readonly name: 'XlvError';
  readonly code: XlvErrorCode;
  readonly message: string;
  readonly location: XlvLocation;
  readonly hint: string;
  readonly docsUrl: string;
  readonly details: XlvDetails;
}

/**
 * Options for constructing an {@link XlvError}.
 *
 * @example
 * ```ts
 * import { XlvError, type XlvErrorInit } from 'xlvane';
 *
 * const init: XlvErrorInit = {
 *   location: { sheet: 'Sales', cell: 'B2' },
 *   hint: 'Use a number in B2.',
 *   details: { expected: 'number' },
 * };
 * const error = new XlvError('XLV_E_INVALID_VALUE', 'B2 is not a number.', init);
 * console.log(error.docsUrl);
 * ```
 */
export interface XlvErrorInit {
  /** Where the problem is. */
  readonly location?: XlvLocation;
  /** What to do next, phrased as an instruction. */
  readonly hint: string;
  /** Machine-readable details, for example the list of available sheets. */
  readonly details?: XlvDetails;
  /** The underlying error, if any. */
  readonly cause?: unknown;
}

/**
 * The only error type thrown by xlvane for problems with input files or options. It says what
 * happened (`message`), where (`location`), what to do next (`hint`) and links to the
 * description of the code (`docsUrl`). `toJSON()` returns the same information as data.
 *
 * @example
 * ```ts
 * import { readWorkbook, isXlvError } from 'xlvane';
 *
 * try {
 *   readWorkbook(new Uint8Array([1, 2, 3]));
 * } catch (error) {
 *   if (isXlvError(error)) console.log(error.code, error.hint); // "XLV_E_NOT_ZIP", "..."
 * }
 * ```
 */
export class XlvError extends Error {
  override readonly name: 'XlvError' = 'XlvError';
  /** Stable error code, see docs/errors.md. */
  readonly code: XlvErrorCode;
  /** Where the problem is. Empty when the problem is not tied to a place in the file. */
  readonly location: XlvLocation;
  /** What to do next. */
  readonly hint: string;
  /** Link to the description of the error code. */
  readonly docsUrl: string;
  /** Machine-readable details. */
  readonly details: XlvDetails;

  constructor(code: XlvErrorCode, message: string, init: XlvErrorInit) {
    super(message, init.cause === undefined ? undefined : { cause: init.cause });
    this.code = code;
    this.location = { ...init.location };
    this.hint = init.hint;
    this.docsUrl = `${DOCS_URL}#${code.toLowerCase()}`;
    this.details = { ...init.details };
  }

  /** Returns the error as plain JSON-safe data. */
  toJSON(): XlvErrorJSON {
    return {
      name: this.name,
      code: this.code,
      message: this.message,
      location: { ...this.location },
      hint: this.hint,
      docsUrl: this.docsUrl,
      details: { ...this.details },
    };
  }
}

/**
 * Returns `true` if `value` is an {@link XlvError}.
 *
 * @example
 * ```ts
 * import { isXlvError } from 'xlvane';
 * console.log(isXlvError(new Error('x'))); // false
 * ```
 */
export function isXlvError(value: unknown): value is XlvError {
  return value instanceof XlvError;
}

/**
 * A warning about the input file. Warnings of the same code are aggregated per sheet: `count`
 * says how often the condition occurred and `cells` lists up to ten example cells.
 *
 * @example
 * ```ts
 * import { readWorkbook, type XlvWarning } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const warnings: readonly XlvWarning[] = readWorkbook(bytes).warnings;
 * for (const warning of warnings) console.log(warning.code, warning.count, warning.cells ?? []);
 * ```
 */
export interface XlvWarning {
  readonly code: XlvWarningCode;
  readonly message: string;
  readonly location: XlvLocation;
  readonly count: number;
  readonly cells?: readonly string[];
}
