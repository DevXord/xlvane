// CSV export and import (docs/decisions.md D-306, D-307; facts docs/facts/csv.md).
// Quoting follows RFC 4180 (PROVISIONAL until the RFC text is checked, CSV-01 … CSV-04); the
// formula-injection rules follow OWASP "CSV Injection" (CSV-10). The parser is a single forward
// scan without regular expressions over untrusted text, so its run time is linear.
import { XlvError } from './errors';
import type { DateSystem, Workbook } from './model';
import { formatCellRef, MAX_COLUMNS, MAX_ROWS, parseRange } from './sml/cell-ref';
import { parseIsoDateTime } from './sml/dates';
import { addTableSheet, pickSheet } from './table';
import type { ReadLimits } from './limits';
import { booleanOption, oneOfOption, openInput, readOptions } from './util/options';
import { createWorkbook, writeWorkbook } from './write';

/**
 * Options for {@link toCSV}.
 *
 * @example
 * ```ts
 * import type { ToCSVOptions } from 'xlvane';
 *
 * const options: ToCSVOptions = { sheet: 'Data', delimiter: ';', bom: true };
 * ```
 */
export interface ToCSVOptions {
  /** Sheet name. Default: the first visible worksheet. */
  readonly sheet?: string;
  /** Only this range, for example `"A1:D100"`. Default: from A1 to the last used row and column. */
  readonly range?: string;
  /** Field separator: one character other than `"`, CR or LF. Default `","`. */
  readonly delimiter?: string;
  /** Line ending. Default `"\r\n"` (RFC 4180). */
  readonly newline?: '\r\n' | '\n';
  /** Start with a UTF-8 byte order mark, which helps Excel detect UTF-8. Default `false`. */
  readonly bom?: boolean;
  /**
   * `"escape"` (default): text cells that start with `=`, `+`, `-`, `@`, tab, CR, LF or their
   * full-width forms get a leading `'` and quotes, so spreadsheet applications show them as
   * text instead of running them as formulas (OWASP CSV Injection). `"off"`: write them as is.
   */
  readonly formulaInjection?: 'escape' | 'off';
  /** Read limits for byte input; see `ReadLimits`. */
  readonly limits?: Partial<ReadLimits>;
}

/** First characters that make spreadsheet applications treat a CSV field as a formula (CSV-10). */
const INJECTION_START = new Set(['=', '+', '-', '@', '\t', '\r', '\n', '＝', '＋', '－', '＠']);

function pad2(n: number): string {
  return n < 10 ? `0${String(n)}` : String(n);
}

/** ISO 8601: the date alone at midnight, otherwise date and time without a zone (D-306). */
function isoDate(d: Date): string {
  const date = `${String(d.getUTCFullYear()).padStart(4, '0')}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
  const ms = d.getUTCMilliseconds();
  if (d.getUTCHours() === 0 && d.getUTCMinutes() === 0 && d.getUTCSeconds() === 0 && ms === 0) return date;
  const time = `${pad2(d.getUTCHours())}:${pad2(d.getUTCMinutes())}:${pad2(d.getUTCSeconds())}`;
  return `${date}T${time}${ms === 0 ? '' : `.${String(ms).padStart(3, '0')}`}`;
}

/**
 * Converts a sheet to CSV text. Fields that contain the delimiter, a quote, CR or LF are
 * enclosed in quotes, and quotes are doubled (RFC 4180). Numbers are written in JavaScript's
 * shortest form, booleans as `TRUE`/`FALSE`, dates as ISO 8601, error cells as their code and
 * formula cells as their cached value. By default, text that a spreadsheet would run as a
 * formula is escaped (`formulaInjection: "escape"`).
 *
 * @example
 * ```ts
 * import { toCSV } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const csv = toCSV(bytes, { sheet: 'Orders' });
 * console.log(csv.split('\r\n')[0]); // the header line
 * ```
 */
export function toCSV(input: Uint8Array | ArrayBuffer | Workbook, options?: ToCSVOptions): string {
  const o = readOptions(options, ['sheet', 'range', 'delimiter', 'newline', 'bom', 'formulaInjection', 'limits'], 'toCSV options');
  const delimiter = o.get('delimiter') ?? ',';
  if (typeof delimiter !== 'string' || delimiter.length !== 1 || delimiter === '"' || delimiter === '\r' || delimiter === '\n') {
    throw new XlvError('XLV_E_INVALID_OPTION', 'The option "delimiter" must be one character other than a quote, CR or LF.', {
      hint: 'Use "," (default), ";" or "\\t".',
      details: { option: 'delimiter' },
    });
  }
  const newline = oneOfOption('newline', o.get('newline'), ['\r\n', '\n'], '\r\n');
  const bom = booleanOption('bom', o.get('bom'), false);
  const injection = oneOfOption('formulaInjection', o.get('formulaInjection'), ['escape', 'off'], 'escape');
  const workbook = openInput(input, o.get('limits'));
  const sheet = pickSheet(workbook, o.get('sheet'));

  let startRow = 1;
  let startCol = 1;
  let endRow = 0;
  let endCol = 0;
  const rawRange = o.get('range');
  if (rawRange !== undefined) {
    const range = typeof rawRange === 'string' ? parseRange(rawRange) : null;
    if (range === null) {
      throw new XlvError('XLV_E_INVALID_CELL_REF', `The range ${JSON.stringify(String(rawRange).slice(0, 40))} is not valid.`, {
        location: { sheet: sheet.name },
        hint: 'Pass an A1-style range such as "A1:D100".',
        details: { ref: String(rawRange).slice(0, 40) },
      });
    }
    ({ row: startRow, col: startCol } = range.start);
    ({ row: endRow, col: endCol } = range.end);
  } else {
    for (const cell of sheet.cells()) {
      if (cell.row > endRow) endRow = cell.row;
      if (cell.col > endCol) endCol = cell.col;
    }
  }

  const needsQuotes = (s: string): boolean => s.includes(delimiter) || s.includes('"') || s.includes('\r') || s.includes('\n');
  const field = (value: unknown): string => {
    if (value === null || value === undefined) return '';
    if (typeof value === 'number') return String(value);
    if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
    if (value instanceof Date) return isoDate(value);
    let s = String(value);
    let quote = needsQuotes(s);
    if (injection === 'escape' && s.length > 0 && INJECTION_START.has(s[0] ?? '')) {
      s = `'${s}`;
      quote = true;
    }
    return quote ? `"${s.replace(/"/g, '""')}"` : s;
  };

  const lines: string[] = [];
  for (let r = startRow; r <= endRow; r++) {
    const fields: string[] = [];
    for (let c = startCol; c <= endCol; c++) {
      const cell = sheet.cell(formatCellRef(r, c));
      // Error cells hold their code as a string; it is written as is (never escaped: it starts with "#").
      fields.push(field(cell?.value));
    }
    lines.push(fields.join(delimiter));
  }
  const body = lines.length === 0 ? '' : lines.join(newline) + newline;
  return (bom ? '\uFEFF' : '') + body;
}

/**
 * Options for {@link fromCSV}.
 *
 * @example
 * ```ts
 * import type { FromCSVOptions } from 'xlvane';
 *
 * const options: FromCSVOptions = { delimiter: ';', header: true, types: 'infer', preset: 'report' };
 * ```
 */
export interface FromCSVOptions {
  /** Field separator: one character other than `"`, CR or LF. Default `","`. */
  readonly delimiter?: string;
  /** The first line holds column names. Default `true`. */
  readonly header?: boolean;
  /**
   * `"infer"` (default): plain decimal numbers (up to 15 significant digits, no leading zeros),
   * `true`/`false` and ISO 8601 dates become numbers, booleans and dates; everything else stays
   * text. `"text"`: every field stays text.
   */
  readonly types?: 'infer' | 'text';
  /** Sheet name. Default `"Sheet1"`. */
  readonly sheet?: string;
  /** `"plain"` (default) or `"report"` (formatted header, frozen row, filter, widths). */
  readonly preset?: 'plain' | 'report';
  /** Date system of the new workbook. Default `"1900"`. */
  readonly dateSystem?: DateSystem;
  /** Largest input in characters. Default 256 Mi. */
  readonly maxChars?: number;
}

const DEFAULT_MAX_CHARS = 256 * 1024 * 1024;

function malformed(message: string, line: number, column: number, hint: string): XlvError {
  return new XlvError('XLV_E_CSV_MALFORMED', `${message} (line ${String(line)}, column ${String(column)}).`, {
    hint,
    details: { line, column },
  });
}

/**
 * Splits CSV text into rows of fields (RFC 4180 with CRLF, LF or CR line ends). Throws
 * `XLV_E_CSV_MALFORMED` for an unterminated quoted field or text after a closing quote.
 */
export function parseCsvRows(text: string, delimiter: string): string[][] {
  const rows: string[][] = [];
  if (text.length === 0) return rows;
  let row: string[] = [];
  let i = 0;
  let line = 1;
  let lineStart = 0;
  const n = text.length;
  const limitCheck = (): void => {
    if (row.length > MAX_COLUMNS) {
      throw new XlvError('XLV_E_LIMIT_EXCEEDED', `Line ${String(line)} has more than ${String(MAX_COLUMNS)} fields.`, {
        hint: `A sheet has ${String(MAX_COLUMNS)} columns. Check the delimiter option; the file may use another separator.`,
        details: { limit: 'columns', value: MAX_COLUMNS, line },
      });
    }
  };
  while (true) {
    // Start of a field.
    if (i < n && text[i] === '"') {
      const quoteLine = line;
      const quoteColumn = i - lineStart + 1;
      i++;
      let value = '';
      let segment = i;
      while (true) {
        const q = text.indexOf('"', i);
        if (q === -1) {
          throw malformed('A quoted field is not closed', quoteLine, quoteColumn, 'Add the closing quote, or double quotes that are part of the text ("" stands for one quote).');
        }
        // Track line numbers inside the quoted field.
        for (let k = i; k < q; k++) {
          const ch = text.charCodeAt(k);
          if (ch === 10 || (ch === 13 && text.charCodeAt(k + 1) !== 10)) {
            line++;
            lineStart = k + 1;
          }
        }
        if (text[q + 1] === '"') {
          value += text.slice(segment, q + 1);
          i = q + 2;
          segment = i;
          continue;
        }
        value += text.slice(segment, q);
        i = q + 1;
        break;
      }
      row.push(value);
      if (i < n && text[i] !== delimiter && text[i] !== '\r' && text[i] !== '\n') {
        throw malformed('Text follows a closing quote', line, i - lineStart + 1, 'Enclose the whole field in quotes, or remove the text after the closing quote.');
      }
    } else {
      let j = i;
      while (j < n) {
        const ch = text[j];
        if (ch === delimiter || ch === '\r' || ch === '\n') break;
        j++;
      }
      row.push(text.slice(i, j));
      i = j;
    }
    limitCheck();
    if (i >= n) {
      rows.push(row);
      break;
    }
    const ch = text[i];
    if (ch === delimiter) {
      i++;
      continue;
    }
    // Line end: CRLF, LF or CR.
    i += ch === '\r' && text[i + 1] === '\n' ? 2 : 1;
    rows.push(row);
    row = [];
    line++;
    lineStart = i;
    if (rows.length > MAX_ROWS) {
      throw new XlvError('XLV_E_LIMIT_EXCEEDED', `The CSV text has more than ${String(MAX_ROWS)} lines.`, {
        hint: `A sheet has ${String(MAX_ROWS)} rows. Split the file.`,
        details: { limit: 'rows', value: MAX_ROWS },
      });
    }
    if (i >= n) break;
  }
  return rows;
}

/** Plain decimal number: optional minus, no leading zeros, optional fraction and exponent (D-307). */
function inferNumber(s: string): number | null {
  let i = 0;
  const n = s.length;
  if (s[i] === '-') i++;
  const intStart = i;
  while (i < n && s.charCodeAt(i) >= 48 && s.charCodeAt(i) <= 57) i++;
  const intDigits = i - intStart;
  if (intDigits === 0 || (intDigits > 1 && s[intStart] === '0')) return null;
  let fracDigits = 0;
  if (s[i] === '.') {
    i++;
    const fracStart = i;
    while (i < n && s.charCodeAt(i) >= 48 && s.charCodeAt(i) <= 57) i++;
    fracDigits = i - fracStart;
    if (fracDigits === 0) return null;
  }
  if (s[i] === 'e' || s[i] === 'E') {
    i++;
    if (s[i] === '+' || s[i] === '-') i++;
    const expStart = i;
    while (i < n && s.charCodeAt(i) >= 48 && s.charCodeAt(i) <= 57) i++;
    if (i === expStart) return null;
  }
  if (i !== n) return null;
  // More than 15 significant digits would lose precision in a double: keep such values (IDs,
  // account numbers) as text.
  const digits = s.slice(intStart, intStart + intDigits + (fracDigits > 0 ? fracDigits + 1 : 0)).replace('.', '').replace(/^0+/, '');
  if (digits.length > 15) return null;
  const value = Number(s);
  return Number.isFinite(value) ? value : null;
}

/** The value of a CSV field with `types: "infer"` (also used by `xlvane edit --set`). */
export function inferValue(s: string, minDate: number): string | number | boolean | Date | null {
  if (s === '') return null;
  const number = inferNumber(s);
  if (number !== null) return number;
  const lower = s.toLowerCase();
  if (lower === 'true') return true;
  if (lower === 'false') return false;
  if (s.length >= 10 && s[4] === '-' && s[7] === '-') {
    const date = parseIsoDateTime(s);
    // Dates before the start of the date system stay text (they cannot be stored as dates).
    if (date !== null && date.getTime() >= minDate) return date;
  }
  return s;
}

/**
 * Builds an XLSX workbook from CSV text (or UTF-8 bytes) and returns the file's bytes. Fields are
 * never turned into formulas: a field such as `=HYPERLINK(...)` becomes text. With
 * `types: "infer"` (default), plain numbers, booleans and ISO 8601 dates are converted; numbers
 * with leading zeros or more than 15 significant digits stay text, so codes and IDs keep their
 * digits. Throws `XLV_E_CSV_MALFORMED` with the line and column of a broken quoted field.
 *
 * @example
 * ```ts
 * import { fromCSV } from 'xlvane';
 *
 * const bytes = fromCSV('sku,stock\r\nA-1,12\r\nB-7,0\r\n', { sheet: 'Stock', preset: 'report' });
 * ```
 */
export function fromCSV(input: string | Uint8Array, options?: FromCSVOptions): Uint8Array {
  const o = readOptions(options, ['delimiter', 'header', 'types', 'sheet', 'preset', 'dateSystem', 'maxChars'], 'fromCSV options');
  const delimiter = o.get('delimiter') ?? ',';
  if (typeof delimiter !== 'string' || delimiter.length !== 1 || delimiter === '"' || delimiter === '\r' || delimiter === '\n') {
    throw new XlvError('XLV_E_INVALID_OPTION', 'The option "delimiter" must be one character other than a quote, CR or LF.', {
      hint: 'Use "," (default), ";" or "\\t".',
      details: { option: 'delimiter' },
    });
  }
  const header = booleanOption('header', o.get('header'), true);
  const types = oneOfOption('types', o.get('types'), ['infer', 'text'], 'infer');
  const maxChars = o.get('maxChars') ?? DEFAULT_MAX_CHARS;
  if (typeof maxChars !== 'number' || !(maxChars > 0)) {
    throw new XlvError('XLV_E_INVALID_OPTION', 'The option "maxChars" must be a positive number.', { hint: 'For example { maxChars: 1e9 }, or Infinity.', details: { option: 'maxChars' } });
  }

  let text: string;
  if (typeof input === 'string') text = input;
  else if (input instanceof Uint8Array) {
    if (input.length > maxChars * 4) {
      throw new XlvError('XLV_E_LIMIT_EXCEEDED', 'The CSV data is larger than the limit.', { hint: 'Raise maxChars if you trust the input.', details: { limit: 'maxChars', value: maxChars } });
    }
    try {
      text = new TextDecoder('utf-8', { fatal: true }).decode(input);
    } catch (error) {
      throw new XlvError('XLV_E_CSV_MALFORMED', 'The CSV bytes are not valid UTF-8.', {
        hint: 'Convert the file to UTF-8 first, or decode it with the right encoding and pass the text.',
        details: { line: 0, column: 0 },
        cause: error,
      });
    }
  } else {
    throw new XlvError('XLV_E_INVALID_INPUT', 'fromCSV expects CSV text or UTF-8 bytes.', { hint: 'Pass a string, or a Uint8Array with UTF-8 text.' });
  }
  if (text.length > maxChars) {
    throw new XlvError('XLV_E_LIMIT_EXCEEDED', `The CSV text has ${String(text.length)} characters, more than the limit.`, {
      hint: 'Raise maxChars if you trust the input.',
      details: { limit: 'maxChars', value: maxChars },
    });
  }
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);

  const parsed = parseCsvRows(text, delimiter);
  const headerFields = header ? (parsed.shift() ?? []) : [];
  const width = parsed.reduce((w, r) => Math.max(w, r.length), headerFields.length);
  const dateSystem = o.get('dateSystem');
  const minDate = dateSystem === '1904' ? Date.UTC(1904, 0, 1) : Date.UTC(1900, 0, 1);
  const rows = parsed.map((fields) => fields.map((f) => (types === 'infer' ? inferValue(f, minDate) : f === '' ? null : f)));
  const columns = header ? Array.from({ length: width }, (_, i) => headerFields[i] ?? '') : undefined;
  if (!header && rows.length > 0 && (rows[0]?.length ?? 0) < width) {
    // The first row decides the width of positional rows in addTableSheet; pad it.
    const first = rows[0] ?? [];
    while (first.length < width) first.push(null);
  }
  const workbook = createWorkbook(dateSystem === undefined ? undefined : { dateSystem: dateSystem as DateSystem });
  const tableOptions = new Map<string, unknown>([
    ['sheet', o.get('sheet') ?? 'Sheet1'],
    ['preset', oneOfOption('preset', o.get('preset'), ['plain', 'report'], 'plain')],
  ]);
  if (columns) tableOptions.set('columns', columns);
  addTableSheet(workbook, rows, tableOptions);
  return writeWorkbook(workbook);
}
