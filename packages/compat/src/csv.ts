// exceljs `workbook.csv` (lib/csv/csv.js, exceljs 4.4.0), built on a small CSV reader and writer
// of its own instead of fast-csv (docs/facts/exceljs.md X4-23, X4-24; docs/decisions.md D-E03).
// Options and output follow exceljs and fast-csv 4.3; xlvane escapes formula-like text by default
// (OWASP CSV Injection, docs/facts/csv.md CSV-10) and fixes a few exceljs bugs, listed in
// docs/exceljs-compat.md.
import { DEFAULT_READ_LIMITS, XlvError } from 'xlvane';
import type { Worksheet } from './model';
import type { CellValue } from './types';
import type { Workbook } from './workbook';

/**
 * Options of `workbook.csv.read` and `readFile`, as exceljs `CsvReadOptions`.
 *
 * @example
 * ```ts
 * import type { CsvReadOptions } from '@xlvane/compat';
 *
 * const options: CsvReadOptions = { sheetName: 'Data', parserOptions: { delimiter: ';' } };
 * ```
 */
export interface CsvReadOptions {
  /** Formats tried, in order, for text that is not a number (dayjs tokens, strict, local time). */
  dateFormats?: string[];
  /** Turns each field into a cell value; replaces the default conversion. */
  map?(value: string, index: number): unknown;
  /** Name of the worksheet that is added; exceljs's default name `sheetN` otherwise. */
  sheetName?: string;
  /** fast-csv parser options: `delimiter`, `quote`, `escape`, `trim`, `ltrim`, `rtrim`, `ignoreEmpty`, `comment`, `skipLines`, `skipRows`, `maxRows`, `headers`, `renameHeaders`, `encoding`. */
  parserOptions?: Record<string, unknown>;
}

/**
 * Options of `workbook.csv.write`, `writeFile` and `writeBuffer`, as exceljs `CsvWriteOptions`,
 * plus `formulaInjection`.
 *
 * @example
 * ```ts
 * import type { CsvWriteOptions } from '@xlvane/compat';
 *
 * const options: CsvWriteOptions = { sheetName: 'Data', dateFormat: 'DD/MM/YYYY', formatterOptions: { delimiter: ';' } };
 * ```
 */
export interface CsvWriteOptions {
  /** dayjs format for dates; ISO 8601 with offset otherwise. */
  dateFormat?: string;
  /** Format dates in UTC instead of local time. */
  dateUTC?: boolean;
  sheetName?: string;
  sheetId?: number;
  /** Character set of `writeFile`: `utf8` (default), `utf16le` or `latin1`. */
  encoding?: string;
  /** Turns each cell value into a field; replaces the default conversion. */
  map?(value: unknown, index: number): unknown;
  /** Blank lines for missing rows (default `true`). */
  includeEmptyRows?: boolean;
  /** fast-csv formatter options: `delimiter`, `rowDelimiter`, `quote`, `escape`, `quoteColumns`, `headers`, `writeHeaders`, `includeEndRowDelimiter`, `writeBOM`. */
  formatterOptions?: Record<string, unknown>;
  /**
   * `"escape"` (default): text that starts with `=`, `+`, `-`, `@`, tab, CR, LF or another
   * control character (or the full-width forms of the signs), or with a byte order mark, is
   * written with a leading `'`, so spreadsheet applications show it as text (OWASP CSV
   * Injection); so is a `latin1` or `utf16le` field whose first bytes are a byte order mark. A
   * field holding a comma, `;`, tab or `|` is quoted whatever the delimiter. `"off"`: written
   * unchanged, as exceljs does. An xlvane option.
   */
  formulaInjection?: 'escape' | 'off';
}

// lib/csv/csv.js L13–23.
// A Map, so fields such as "constructor" or "__proto__" stay text (D-R01).
const SPECIAL_VALUES: ReadonlyMap<string, unknown> = new Map<string, unknown>(Object.entries({
  true: true,
  false: false,
  '#N/A': { error: '#N/A' },
  '#REF!': { error: '#REF!' },
  '#NAME?': { error: '#NAME?' },
  '#DIV/0!': { error: '#DIV/0!' },
  '#NULL!': { error: '#NULL!' },
  '#VALUE!': { error: '#VALUE!' },
  '#NUM!': { error: '#NUM!' },
}));
const DEFAULT_DATE_FORMATS = ['YYYY-MM-DD[T]HH:mm:ssZ', 'YYYY-MM-DD[T]HH:mm:ss', 'MM-DD-YYYY', 'YYYY-MM-DD'];
// docs/facts/csv.md CSV-10 (OWASP CSV Injection).
// The xlvane core's set (CSV-10, CSV-12); text starting with another C0 control is escaped too.
/** Signs that start a formula: refused as the first character of the delimiter while formulas are escaped (D-J01). */
const FORMULA_SIGNS = new Set(['=', '+', '-', '@', '\uff1d', '\uff0b', '\uff0d', '\uff20', '\ufe66', '\ufe62', '\ufe63', '\ufe6b']);
// U+FEFF and U+FFFE too: applications drop a byte order mark at the start of a file and run the
// formula that follows it (K2; LibreOffice confirmed).
const INJECTION_START = new Set(['=', '+', '-', '@', '\t', '\r', '\n', '\uff1d', '\uff0b', '\uff0d', '\uff20', '\ufe66', '\ufe62', '\ufe63', '\ufe6b', '\ufeff', '\ufffe']);
/** Characters a field is quoted for while formulas are escaped, whatever the delimiter: other applications may split on them (K2). */
const COMMON_DELIMITERS = ',;\t|';

function invalid(name: string, message: string, hint: string): XlvError {
  return new XlvError('XLV_E_INVALID_OPTION', message, { hint, details: { option: name } });
}

// ---- dates: the dayjs 1.11 format and strict customParseFormat tokens used by exceljs -------------

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
// dayjs esm/constant.js REGEX_FORMAT.
const FORMAT_TOKENS = /\[([^\]]+)]|YYYY|YY|M{1,4}|D{1,2}|d{1,4}|H{1,2}|h{1,2}|a|A|m{1,2}|s{1,2}|Z{1,2}|SSS/g;
const pad = (n: number, width = 2): string => String(Math.abs(n)).padStart(width, '0');

interface DateParts {
  y: number;
  M: number;
  d: number;
  wd: number;
  H: number;
  m: number;
  s: number;
  ms: number;
  offset: number;
}

function parts(date: Date, utc: boolean): DateParts {
  return utc
    ? { y: date.getUTCFullYear(), M: date.getUTCMonth(), d: date.getUTCDate(), wd: date.getUTCDay(), H: date.getUTCHours(), m: date.getUTCMinutes(), s: date.getUTCSeconds(), ms: date.getUTCMilliseconds(), offset: 0 }
    : { y: date.getFullYear(), M: date.getMonth(), d: date.getDate(), wd: date.getDay(), H: date.getHours(), m: date.getMinutes(), s: date.getSeconds(), ms: date.getMilliseconds(), offset: -date.getTimezoneOffset() };
}

function zone(offset: number, colon: boolean): string {
  return `${offset < 0 ? '-' : '+'}${pad(Math.floor(Math.abs(offset) / 60))}${colon ? ':' : ''}${pad(Math.abs(offset) % 60)}`;
}

/** A date formatted as dayjs `format()` formats it. */
export function formatDate(date: Date, format: string | undefined, utc: boolean): string {
  const p = parts(date, utc);
  // dayjs utc plugin: format() without a pattern ends in "Z" in UTC mode.
  // dayjs format(formatStr) uses `formatStr || default`, so '' gives the default too (D-E06).
  const pattern = format || (utc ? 'YYYY-MM-DDTHH:mm:ss[Z]' : 'YYYY-MM-DDTHH:mm:ssZ');
  const h12 = p.H % 12 || 12;
  return pattern.replace(FORMAT_TOKENS, (token, literal: string | undefined) => {
    if (literal !== undefined) return literal;
    switch (token) {
      case 'YY':
        return String(p.y).slice(-2);
      case 'YYYY':
        return pad(p.y, 4);
      case 'M':
        return String(p.M + 1);
      case 'MM':
        return pad(p.M + 1);
      case 'MMM':
        return MONTHS[p.M]!.slice(0, 3);
      case 'MMMM':
        return MONTHS[p.M]!;
      case 'D':
        return String(p.d);
      case 'DD':
        return pad(p.d);
      case 'd':
        return String(p.wd);
      case 'dd':
        return DAYS[p.wd]!.slice(0, 2);
      case 'ddd':
        return DAYS[p.wd]!.slice(0, 3);
      case 'dddd':
        return DAYS[p.wd]!;
      case 'H':
        return String(p.H);
      case 'HH':
        return pad(p.H);
      case 'h':
        return String(h12);
      case 'hh':
        return pad(h12);
      case 'a':
        return p.H < 12 ? 'am' : 'pm';
      case 'A':
        return p.H < 12 ? 'AM' : 'PM';
      case 'm':
        return String(p.m);
      case 'mm':
        return pad(p.m);
      case 's':
        return String(p.s);
      case 'ss':
        return pad(p.s);
      case 'SSS':
        return pad(p.ms, 3);
      case 'Z':
        return zone(p.offset, true);
      case 'ZZ':
        return zone(p.offset, false);
      default:
        return token;
    }
  });
}

// dayjs customParseFormat (1.11, as installed with exceljs 4.4.0): formattingTokens and the
// expression of each token. Characters that are no token are not in the list at all, as
// `String.match` with these tokens drops them (X4-23, D-E06).
const PARSE_TOKENS = /(\[[^[]*\])|([-_:/.,()\s]+)|(A|a|Q|YYYY|YY?|ww?|MM?M?M?|Do|DD?|hh?|HH?|mm?|ss?|S{1,3}|z|ZZ?)/g;
// The leftmost match of dayjs's /\d*[^-_:/,()\s\d]+/ never starts right after a digit (the match
// one position earlier would be found first), so the lookbehind gives the same matches without
// trying every position of a long run of digits again (quadratic before; D-J01).
const WORD = /(?<!\d)\d*[^-_:/,()\s\d]+/;
const ONE_OR_TWO = /\d\d?/;
const PARSE_EXPRESSIONS: Record<string, RegExp> = {
  A: WORD,
  a: WORD,
  Q: /\d/,
  S: /\d/,
  SS: /\d\d/,
  SSS: /\d{3}/,
  s: ONE_OR_TWO,
  ss: ONE_OR_TWO,
  m: ONE_OR_TWO,
  mm: ONE_OR_TWO,
  H: ONE_OR_TWO,
  h: ONE_OR_TWO,
  HH: ONE_OR_TWO,
  hh: ONE_OR_TWO,
  D: ONE_OR_TWO,
  DD: /\d\d/,
  Do: WORD,
  w: ONE_OR_TWO,
  ww: /\d\d/,
  M: ONE_OR_TWO,
  MM: /\d\d/,
  MMM: WORD,
  MMMM: WORD,
  Y: /[+-]?\d+/,
  YY: /\d\d/,
  YYYY: /\d{4}/,
  Z: /[+-]\d\d:?(\d\d)?|Z/,
  ZZ: /[+-]\d\d:?(\d\d)?|Z/,
};

/**
 * A date parsed from `text` with a dayjs format in strict mode, as exceljs's CSV reader parses
 * it (`dayjs(text, format, true)`; dayjs customParseFormat `makeParser`): literal parts only
 * advance the position, each token's expression is searched (not anchored) in the rest of the
 * input and its first match is then removed from the input; the result is local time unless the
 * text has an offset, and only valid when formatting it again gives the same text. `null`
 * otherwise.
 */
export function parseDate(text: string, format: string): Date | null {
  const tokens = format.match(PARSE_TOKENS) ?? [];
  let input = text;
  let start = 0;
  let year: number | undefined;
  let month: number | undefined;
  let day: number | undefined;
  let hours: number | undefined;
  let minutes: number | undefined;
  let seconds: number | undefined;
  let ms: number | undefined;
  let afternoon: boolean | undefined;
  let offset: number | undefined;
  for (const token of tokens) {
    const expression = PARSE_EXPRESSIONS[token];
    if (!expression) {
      start += token.replace(/^\[|\]$/g, '').length;
      continue;
    }
    const match = expression.exec(input.slice(start));
    // dayjs reads match[0] of a failed match, throws and gives an invalid date.
    if (!match) return null;
    const value = match[0];
    input = input.replace(value, '');
    const n = Number(value);
    switch (token) {
      case 'A':
      case 'a':
        afternoon = value === (token === 'a' ? 'pm' : 'PM');
        break;
      case 'Q':
        month = (n - 1) * 3 + 1;
        break;
      case 'S':
        ms = n * 100;
        break;
      case 'SS':
        ms = n * 10;
        break;
      case 'SSS':
        ms = n;
        break;
      case 's':
      case 'ss':
        seconds = n;
        break;
      case 'm':
      case 'mm':
        minutes = n;
        break;
      case 'H':
      case 'h':
      case 'HH':
      case 'hh':
        hours = n;
        break;
      case 'D':
      case 'DD':
        day = n;
        break;
      case 'Do':
        day = Number(/\d+/.exec(value)?.[0] ?? Number.NaN);
        break;
      case 'w':
      case 'ww':
        // dayjs needs its weekOfYear plugin for this, which exceljs does not load: invalid date.
        return null;
      case 'M':
      case 'MM':
        month = n;
        break;
      case 'MMM':
      case 'MMMM': {
        const names = token === 'MMM' ? MONTHS.map((m) => m.slice(0, 3)) : MONTHS;
        const index = names.indexOf(value);
        if (index < 0) return null;
        month = index + 1;
        break;
      }
      case 'Y':
      case 'YYYY':
        year = n;
        break;
      case 'YY':
        year = n + (n > 68 ? 1900 : 2000);
        break;
      default: {
        // Z, ZZ
        if (value === 'Z') offset = 0;
        else {
          const sign = value[0] === '-' ? -1 : 1;
          const digits = value.slice(1).replace(':', '');
          offset = sign * (Number(digits.slice(0, 2)) * 60 + Number(digits.slice(2) || '0'));
        }
      }
    }
  }
  if (afternoon !== undefined && hours !== undefined) {
    if (afternoon && hours < 12) hours += 12;
    else if (!afternoon && hours === 12) hours = 0;
  }
  // dayjs parseFormattedInput uses `||`, so 0 counts as missing.
  const now = new Date();
  const d = day || (!year && !month ? now.getDate() : 1);
  const y = year || now.getFullYear();
  const M = year && !month ? 0 : month !== undefined && month > 0 ? month - 1 : now.getMonth();
  const h = hours || 0;
  const mi = minutes || 0;
  const se = seconds || 0;
  const millis = ms || 0;
  const date = offset === undefined ? new Date(y, M, d, h, mi, se, millis) : new Date(Date.UTC(y, M, d, h, mi, se, millis) - offset * 60_000);
  if (Number.isNaN(date.getTime())) return null;
  // Strict mode: the date formatted with the same pattern (in local time) must give the input.
  if (formatDate(date, format, false) !== text) return null;
  return date;
}

// ---- reading ---------------------------------------------------------------------------------

interface ParserOptions {
  delimiter: string;
  quote: string;
  escape: string;
  trim: boolean;
  ltrim: boolean;
  rtrim: boolean;
  ignoreEmpty: boolean;
  comment: string | null;
  skipLines: number;
  skipRows: number;
  maxRows: number;
  skipHeader: boolean;
  encoding: string;
}

function parserOptions(raw: Record<string, unknown> | undefined): ParserOptions {
  const o = raw ?? {};
  const str = (name: string, fallback: string): string => {
    const v = o[name];
    if (v === undefined) return fallback;
    if (typeof v !== 'string') throw invalid(`parserOptions.${name}`, `parserOptions.${name} must be a string.`, `Give one character, for example "${fallback}".`);
    return v;
  };
  const delimiter = str('delimiter', ',');
  if (delimiter.length !== 1) throw invalid('parserOptions.delimiter', 'parserOptions.delimiter must be one character long.', 'Use one character, for example ";" or "\\t".');
  const quote = o['quote'] === null || o['quote'] === false ? '' : str('quote', '"');
  const escape = o['escape'] === null || o['escape'] === undefined ? quote : str('escape', quote);
  // fast-csv compares single characters; longer quote or escape strings would never match (D-E06).
  if (quote.length > 1) throw invalid('parserOptions.quote', 'parserOptions.quote must be one character long.', 'Use one character, for example "\\"" or "\'".');
  if (escape.length > 1) throw invalid('parserOptions.escape', 'parserOptions.escape must be one character long.', 'Use one character, for example "\\\\".');
  const count = (name: string): number => {
    const v = o[name] ?? 0;
    if (typeof v !== 'number' || !Number.isInteger(v) || v < 0) throw invalid(`parserOptions.${name}`, `parserOptions.${name} must be a whole number of at least 0.`, 'For example 1.');
    return v;
  };
  const headers = o['headers'];
  return {
    delimiter,
    quote,
    escape,
    trim: o['trim'] === true,
    ltrim: o['ltrim'] === true,
    rtrim: o['rtrim'] === true,
    ignoreEmpty: o['ignoreEmpty'] === true,
    comment: typeof o['comment'] === 'string' && o['comment'] !== '' ? o['comment'] : null,
    skipLines: count('skipLines'),
    skipRows: count('skipRows'),
    maxRows: count('maxRows'),
    // fast-csv does not emit the header line with `headers: true` or with `renameHeaders`.
    skipHeader: headers === true || (Array.isArray(headers) && o['renameHeaders'] === true),
    encoding: typeof o['encoding'] === 'string' ? o['encoding'] : 'utf8',
  };
}

function malformed(line: number, message: string): XlvError {
  return new XlvError('XLV_E_CSV_MALFORMED', `CSV line ${String(line)}: ${message}`, {
    location: { line },
    hint: 'Check the quotes on that line: a quoted field must end with a closing quote before the next delimiter or line break.',
    details: { line },
  });
}

/** Rows and columns of a worksheet (Part 1 §18.3.1.73, ST_CellRef up to XFD1048576). */
const MAX_SHEET_ROWS = 1_048_576;
const MAX_SHEET_COLUMNS = 16_384;

/** A parsed CSV row: its fields and the line it starts on. */
interface CsvRow {
  fields: string[];
  line: number;
}

function tooMany(what: 'rows' | 'columns', line: number): XlvError {
  const rows = what === 'rows';
  const max = rows ? MAX_SHEET_ROWS : MAX_SHEET_COLUMNS;
  return new XlvError('XLV_E_LIMIT_EXCEEDED', `CSV line ${String(line)}: ${rows ? `the CSV has more than ${String(max)} rows` : `the row has more than ${String(max)} fields`}, more than a worksheet can hold.`, {
    location: { line },
    hint: rows ? 'Split the CSV file into parts of at most 1048576 rows, or read only part of it with parserOptions.maxRows.' : 'Split the CSV file into parts of at most 16384 columns, or check the delimiter (parserOptions.delimiter).',
    details: { limit: rows ? 'maxRows' : 'maxColumns', value: max, line },
  });
}

/**
 * The rows of CSV text, split as fast-csv 4.3 does with the given parser options, one at a time:
 * the rows that `ignoreEmpty`, the header, `skipRows` and `maxRows` drop are left out. More than
 * 1 048 576 rows, or a row of more than 16 384 fields, throw `XLV_E_LIMIT_EXCEEDED` as soon as
 * they are reached, because a worksheet cannot hold them (K2).
 */
function* csvRows(text: string, o: ParserOptions): Generator<CsvRow> {
  let at = 0;
  let line = 1;
  const n = text.length;
  const atLineEnd = (i: number): boolean => i >= n || text[i] === '\n' || text[i] === '\r';
  const skipLineEnd = (i: number): number => {
    if (text[i] === '\r' && text[i + 1] === '\n') return i + 2;
    return i < n ? i + 1 : i;
  };
  const trim = (f: string): string => (o.trim ? f.trim() : o.ltrim ? f.trimStart() : o.rtrim ? f.trimEnd() : f);
  // Rows kept so far (after ignoreEmpty), for the header, skipRows and maxRows.
  let kept = 0;
  let emitted = 0;
  // skipLines: raw lines dropped before parsing.
  for (let k = 0; k < o.skipLines && at < n; k++) {
    while (!atLineEnd(at)) at++;
    at = skipLineEnd(at);
    line++;
  }
  while (at < n) {
    if (o.comment !== null && text.startsWith(o.comment, at)) {
      while (!atLineEnd(at)) at++;
      at = skipLineEnd(at);
      line++;
      continue;
    }
    const rowLine = line;
    const fields: string[] = [];
    // Fields past the 16 384 a worksheet holds are only counted, and whether one has text.
    let extra = 0;
    let extraText = false;
    const push = (field: string): void => {
      if (fields.length < MAX_SHEET_COLUMNS) fields.push(field);
      else {
        extra++;
        if (field.trim() !== '') extraText = true;
      }
    };
    for (;;) {
      // A field: optional quoted part (after leading white space), or plain text up to the delimiter.
      let start = at;
      while (start < n && (text[start] === ' ' || text[start] === '\t') && text[start] !== o.delimiter) start++;
      if (o.quote !== '' && text[start] === o.quote) {
        // The value is built from slices between escaped quotes (one character at a time before, K2).
        let value = '';
        let i = start + 1;
        let from = i;
        // A quote left open is reported on the line where it opens (D-E06).
        const opened = line;
        for (;;) {
          if (i >= n) throw malformed(opened, `missing closing ${o.quote} for the field that starts here.`);
          const ch = text[i]!;
          // An escaped quote: `""` by default, or the `escape` character followed by the quote.
          if (ch === o.escape && text[i + 1] === o.quote) {
            value += text.slice(from, i) + o.quote;
            i += 2;
            from = i;
            continue;
          }
          if (ch === o.quote) {
            value += text.slice(from, i);
            i++;
            break;
          }
          if (ch === '\n') line++;
          i++;
        }
        while (i < n && (text[i] === ' ' || text[i] === '\t') && text[i] !== o.delimiter) i++;
        if (i < n && text[i] !== o.delimiter && text[i] !== '\n' && text[i] !== '\r') throw malformed(line, `unexpected text after the closing ${o.quote}.`);
        push(value);
        at = i;
      } else {
        let i = at;
        while (i < n && text[i] !== o.delimiter && text[i] !== '\n' && text[i] !== '\r') i++;
        push(text.slice(at, i));
        at = i;
      }
      if (at < n && text[at] === o.delimiter) {
        at++;
        continue;
      }
      break;
    }
    at = skipLineEnd(at);
    line++;
    // A line of white space only is one empty field (fast-csv).
    const row = fields.length === 1 && fields[0]!.trim() === '' ? [''] : fields;
    if (o.ignoreEmpty && !extraText && row.every((f) => f.trim() === '')) continue;
    kept++;
    if (o.skipHeader && kept === 1) continue;
    if (kept - (o.skipHeader ? 1 : 0) <= o.skipRows) continue;
    // Rows past maxRows are still parsed, so that malformed text after them is reported as before.
    if (o.maxRows > 0 && emitted >= o.maxRows) continue;
    if (extra > 0) throw tooMany('columns', rowLine);
    if (++emitted > MAX_SHEET_ROWS) throw tooMany('rows', rowLine);
    yield { fields: row.map(trim), line: rowLine };
  }
}

/** Splits CSV text into rows of fields as fast-csv 4.3 does with the given parser options. */
export function parseCsv(text: string, o: ParserOptions): string[][] {
  return Array.from(csvRows(text, o), (r) => r.fields);
}

/**
 * exceljs's default conversion of a CSV field (lib/csv/csv.js L57–83), with two fixes: text of
 * white space only stays text (exceljs gives 0), and -Infinity stays text (it cannot be stored).
 */
export function csvValue(datum: string, dateFormats: readonly string[]): unknown {
  if (datum === '') return null;
  if (datum.trim() !== '') {
    const number = Number(datum);
    if (Number.isFinite(number)) return number;
  }
  for (const format of dateFormats) {
    const date = parseDate(datum, format);
    if (date) return date;
  }
  const special = SPECIAL_VALUES.get(datum);
  if (special !== undefined) return typeof special === 'object' ? { ...special } : special;
  return datum;
}

// ---- writing ---------------------------------------------------------------------------------

interface FormatterOptions {
  delimiter: string;
  rowDelimiter: string;
  quote: string;
  escapedQuote: string;
  quoteColumns: unknown;
  headers: string[] | null;
  includeEndRowDelimiter: boolean;
  writeBOM: boolean;
  escapeChars: RegExp;
  /** Formulas are escaped: with `quote: false`, a field that a reader splitting on a comma, ";", tab or "|" would turn into a formula is refused. */
  strictQuoting: boolean;
}

function formatterOptions(raw: Record<string, unknown> | undefined, escapeFormulas: boolean): FormatterOptions {
  const o = raw ?? {};
  const delimiter = typeof o['delimiter'] === 'string' ? o['delimiter'] : ',';
  const rowDelimiter = typeof o['rowDelimiter'] === 'string' ? o['rowDelimiter'] : '\n';
  const q = o['quote'];
  const quote = q === false ? '' : q === true || q === undefined ? '"' : String(q);
  const escape = typeof o['escape'] === 'string' ? o['escape'] : quote;
  const escapeRe = (s: string): string => s.replace(/[\\^$.*+?()[\]{}|-]/g, '\\$&');
  return {
    delimiter,
    rowDelimiter,
    quote,
    escapedQuote: `${escape}${quote}`,
    quoteColumns: o['quoteColumns'] ?? false,
    headers: Array.isArray(o['headers']) && o['writeHeaders'] !== false ? (o['headers'] as unknown[]).map(String) : null,
    includeEndRowDelimiter: o['includeEndRowDelimiter'] === true,
    writeBOM: o['writeBOM'] === true,
    // fast-csv FieldFormatter: the character class also contains "|" (it is written inside "[…]").
    // While formulas are escaped, a comma, ";", tab or "|" is quoted with any delimiter: an
    // application that splits on another one would otherwise see a new field there (K2).
    escapeChars: new RegExp(`[${escapeRe(delimiter)}${escapeRe(rowDelimiter)}|\\r|\\n${escapeFormulas ? escapeRe(COMMON_DELIMITERS) : ''}]`),
    strictQuoting: escapeFormulas,
  };
}

/**
 * Whether an application that splits `text` at a comma, ";", tab or "|" would get a field that
 * starts like a formula: one of those characters followed by a formula sign, a control character
 * or a byte order mark, after any spaces or tabs and an opening double quote (K2).
 */
function splitStartsFormula(text: string): boolean {
  for (let i = 0; i < text.length; i++) {
    if (!COMMON_DELIMITERS.includes(text[i]!)) continue;
    let j = i + 1;
    while (text[j] === ' ' || text[j] === '\t') j++;
    if (text[j] === '"') j++;
    while (text[j] === ' ') j++;
    const c = text[j];
    if (c !== undefined && (INJECTION_START.has(c) || c.charCodeAt(0) < 0x20)) return true;
  }
  return false;
}

/** One field as fast-csv FieldFormatter.format writes it. */
function formatField(value: unknown, index: number, f: FormatterOptions, forceQuote: boolean): string {
  const text = `${value === null || value === undefined ? '' : String(value)}`.replace(/\0/g, '');
  // Without a quote character a CR, LF or delimiter in a field would split it into other fields or
  // rows (D-R01).
  if (f.quote === '' && (/[\r\n]/.test(text) || (f.delimiter !== '' && text.includes(f.delimiter)))) {
    throw new XlvError('XLV_E_INVALID_OPTION', `Field ${String(index + 1)} contains a line break or the delimiter, which cannot be written with quote: false.`, {
      hint: 'Leave formatterOptions.quote at its default (") so such fields are quoted, or remove the line breaks from the value.',
      details: { option: 'quote', field: index + 1 },
    });
  }
  // While formulas are escaped, a comma, ";", tab or "|" followed by a formula sign would start a
  // formula for an application that splits on that character (K2); quoted otherwise.
  if (f.quote === '' && f.strictQuoting && splitStartsFormula(text)) {
    throw new XlvError('XLV_E_INVALID_OPTION', `Field ${String(index + 1)} has a comma, ";", tab or "|" followed by a formula sign, which an application splitting on that character reads as a formula; it cannot be written with quote: false.`, {
      hint: 'Leave formatterOptions.quote at its default (") so such fields are quoted, or pass formulaInjection: "off" for a trusted consumer.',
      details: { option: 'quote', field: index + 1 },
    });
  }
  if (f.quote !== '' && text.includes(f.quote)) return `${f.quote}${text.split(f.quote).join(f.escapedQuote)}${f.quote}`;
  const q = f.quoteColumns;
  const quoteThis = forceQuote || q === true || (Array.isArray(q) && q[index] === true);
  if (text.search(f.escapeChars) !== -1 || quoteThis) return `${f.quote}${text}${f.quote}`;
  return text;
}

/**
 * exceljs's default conversion of a cell value to a CSV field (lib/csv/csv.js L127–152), with
 * three fixes: a formula result of 0 or false is written (exceljs writes an empty field), a date
 * result is formatted like a date, and an error result gives its text.
 */
function defaultWriteMap(dateFormat: string | undefined, dateUTC: boolean) {
  const date = (d: Date): string => formatDate(d, dateFormat, dateUTC);
  return (value: unknown): unknown => {
    if (value === null || value === undefined || value === '' || value === 0 || value === false) return value;
    if (value instanceof Date) return date(value);
    if (typeof value === 'object') {
      const v = value as Record<string, unknown>;
      if (v['text'] || v['hyperlink']) return v['hyperlink'] || v['text'] || '';
      if (v['formula'] || v['sharedFormula'] || 'result' in v) {
        const result = v['result'];
        if (result === undefined || result === null) return '';
        if (result instanceof Date) return date(result);
        if (typeof result === 'object' && typeof (result as { error?: unknown }).error === 'string') return (result as { error: string }).error;
        return result;
      }
      if (typeof v['error'] === 'string') return v['error'];
      return JSON.stringify(value);
    }
    return value;
  };
}

/** The character sets `csv.writeFile` writes; `writeBuffer` and `write` write UTF-8. */
type OutputEncoding = 'utf8' | 'utf16le' | 'latin1';

/** The output encoding for the `encoding` option (an unsupported one is refused when encoding). */
function outputEncoding(encoding: string | undefined): OutputEncoding {
  const e = (encoding ?? '').toLowerCase();
  if (e === 'latin1' || e === 'binary') return 'latin1';
  if (e === 'utf16le' || e === 'utf-16le' || e === 'ucs2' || e === 'ucs-2') return 'utf16le';
  return 'utf8';
}

/** Leading UTF-8 sequences of U+FF1D, U+FF0B, U+FF0D, U+FF20, U+FE66, U+FE62, U+FE63, U+FE6B. */
const UTF8_FORMULA_SIGNS = [
  [0xef, 0xbc, 0x9d],
  [0xef, 0xbc, 0x8b],
  [0xef, 0xbc, 0x8d],
  [0xef, 0xbc, 0xa0],
  [0xef, 0xb9, 0xa6],
  [0xef, 0xb9, 0xa2],
  [0xef, 0xb9, 0xa3],
  [0xef, 0xb9, 0xab],
];

/**
 * Whether the first bytes of a field, as latin1 or UTF-16LE writes it, are a byte order mark
 * (EF BB BF, FF FE, FE FF), which makes a reader decode the rest of the file in that encoding
 * and drop the mark, or the UTF-8 form of a full-width or small formula sign, which a reader that
 * guesses UTF-8 shows as one (K2). `prefix` holds bytes written before the field (latin1 writes
 * the `writeBOM` mark U+FEFF as the single byte FF).
 */
function encodedStartIsUnsafe(text: string, encoding: OutputEncoding, prefix: readonly number[]): boolean {
  if (encoding === 'utf8') return false;
  const bytes = [...prefix];
  for (let i = 0; bytes.length < 3 && i < text.length; i++) {
    const c = text.charCodeAt(i);
    if (encoding === 'latin1') bytes.push(c & 0xff);
    else bytes.push(c & 0xff, c >> 8);
  }
  const starts = (lead: readonly number[]): boolean => lead.every((b, i) => bytes[i] === b);
  return starts([0xef, 0xbb, 0xbf]) || starts([0xff, 0xfe]) || starts([0xfe, 0xff]) || UTF8_FORMULA_SIGNS.some(starts);
}

/**
 * The CSV text of a worksheet, as exceljs `csv.write` gives it, for output in `encoding`.
 *
 * @internal
 */
export function worksheetCsv(sheet: Worksheet | undefined, options: CsvWriteOptions, encoding: OutputEncoding = 'utf8'): string {
  const injection = options.formulaInjection ?? 'escape';
  if (injection !== 'escape' && injection !== 'off') throw invalid('formulaInjection', 'The option "formulaInjection" must be "escape" or "off".', 'Use "escape" (default) or "off".');
  const f = formatterOptions(options.formatterOptions, injection === 'escape');
  const latin1 = encoding === 'latin1';
  if (latin1) {
    // latin1 keeps only the low byte of each character, so U+013D would become "=" and U+010A a
    // line feed after the fields were checked and quoted (D-J01).
    for (const [name, value] of [['delimiter', f.delimiter], ['rowDelimiter', f.rowDelimiter], ['quote', f.quote], ['escape', f.escapedQuote]] as const) {
      if (/[^\u0000-\u00ff]/.test(value)) throw invalid(name, `The option "${name}" has a character that latin1 cannot write.`, 'Use characters up to U+00FF with encoding "latin1", or write UTF-8.');
    }
  }
  if (injection === 'escape') {
    // Ways around the escape that a standard (RFC 4180) reader would split differently (D-J01):
    // formula signs or controls in the delimiter (a line with an empty first field starts with
    // it), a row delimiter other than a line break, and a quote other than '"', which such a
    // reader does not take as quoting, so line breaks and delimiters in a value split it.
    const off = 'or pass formulaInjection: "off" for a trusted consumer.';
    if (f.delimiter !== '\t' && [...f.delimiter].some((c) => INJECTION_START.has(c) || c.charCodeAt(0) < 0x20)) {
      throw invalid('delimiter', `The delimiter ${JSON.stringify(f.delimiter)} holds a formula sign or a control character, so a line could start with a formula.`, `Use "," (default), ";", "|" or "\\t", ${off}`);
    }
    if (!['\n', '\r\n', '\r'].includes(f.rowDelimiter)) {
      throw invalid('rowDelimiter', `The row delimiter ${JSON.stringify(f.rowDelimiter)} is not a line break, so a standard CSV reader would not split rows where it is written.`, `Use "\\n" (default), "\\r\\n" or "\\r", ${off}`);
    }
    if (f.quote !== '' && f.quote !== '"') {
      throw invalid('quote', `The quote ${JSON.stringify(f.quote)} is not '"', so a standard CSV reader would split a value that holds a line break or the delimiter.`, `Leave formatterOptions.quote at its default ("), or use quote: false, ${off}`);
    }
    if (f.quote !== '' && f.escapedQuote !== `${f.quote}${f.quote}`) {
      throw invalid('escape', 'An escape character other than the quote lets a standard CSV reader end a quoted field early, so a value could start a new field with a formula.', 'Leave formatterOptions.escape unset (the quote is doubled), or pass formulaInjection: "off" for a trusted consumer.');
    }
  }
  /** The text latin1 will write: characters above U+00FF as "?". */
  const narrow = (out: unknown): unknown => (latin1 && typeof out === 'string' ? out.replace(/[^\u0000-\u00ff]/g, '?') : out);
  /** A formatted field, refused if a standard reader would take its leading quote as quoting. */
  const field = (out: unknown, index: number, force: boolean): string => {
    const text = formatField(out, index, f, force);
    if (injection === 'escape' && f.quote !== '"' && text.startsWith('"')) {
      throw new XlvError('XLV_E_INVALID_OPTION', `Field ${String(index + 1)} starts with a double quote, which a standard CSV reader takes as quoting when formatterOptions.quote is not '"'.`, {
        hint: 'Leave formatterOptions.quote at its default ("), or pass formulaInjection: "off" for a trusted consumer.',
        details: { option: 'quote', field: index + 1 },
      });
    }
    return text;
  };
  const map = options.map ?? defaultWriteMap(options.dateFormat, options.dateUTC === true);
  const includeEmptyRows = options.includeEmptyRows === undefined || options.includeEmptyRows;
  const lines: string[] = [];
  // NUL characters are removed from fields (formatField), so the check looks past them: "\0=cmd"
  // is written as "=cmd" and must be escaped too (D-R01). Values other than numbers, booleans,
  // null and undefined (a String object, an array, an object with toString from `map`) are
  // written as their text, so they are turned into it first and checked like text, calling
  // toString once (K2). `atStart`: the field is the first one of the file.
  const escape = (raw: unknown, atStart: boolean): [unknown, boolean] => {
    const value = raw === null || raw === undefined || typeof raw === 'number' || typeof raw === 'boolean' || typeof raw === 'bigint' || typeof raw === 'string' ? raw : String(raw);
    const out = narrow(value);
    if (injection !== 'escape' || typeof out !== 'string') return [out, false];
    const text = out.replace(/\0/g, '');
    const first = text[0];
    if (first === undefined) return [out, false];
    // latin1 writes the writeBOM mark U+FEFF as the byte FF, before the first field.
    const unsafe = INJECTION_START.has(first) || first.charCodeAt(0) < 0x20 || encodedStartIsUnsafe(text, encoding, atStart && f.writeBOM && latin1 ? [0xff] : []);
    return unsafe ? [`'${out}`, f.quote !== ''] : [out, false];
  };
  if (f.headers) {
    lines.push(
      f.headers
        .map((h, i) => {
          const [out, force] = escape(h, i === 0);
          return field(out, i, force);
        })
        .join(f.delimiter),
    );
  }
  if (sheet) {
    // exceljs starts at 1 and so writes one blank line too few before the first row; xlvane writes one per missing row.
    let lastRow = 0;
    sheet.eachRow((row, rowNumber) => {
      if (includeEmptyRows) while (++lastRow < rowNumber) lines.push('');
      let values = row.values.slice(1);
      // With headers, fast-csv writes one field per header for an array row (RowFormatter
      // gatherColumns): longer rows are cut, shorter ones end in empty fields (D-E06).
      if (f.headers) values = f.headers.map((_h, i) => values[i]);
      const fields: string[] = [];
      for (let i = 0; i < values.length; i++) {
        // Array.prototype.map skips holes, and fast-csv writes them as empty, unquoted fields.
        if (!(i in values)) {
          fields.push('');
          continue;
        }
        const [out, force] = escape(map(values[i], i), lines.length === 0 && i === 0);
        fields.push(field(out, i, force));
      }
      lines.push(fields.join(f.delimiter));
      lastRow = rowNumber;
    });
  }
  let text = lines.join(f.rowDelimiter);
  if (f.includeEndRowDelimiter) text += f.rowDelimiter;
  return (f.writeBOM ? '\uFEFF' : '') + text;
}

// ---- bytes, streams and files ----------------------------------------------------------------

function encode(text: string, encoding: string | undefined): Uint8Array {
  const e = (encoding ?? 'utf8').toLowerCase();
  if (e === 'utf8' || e === 'utf-8') return new TextEncoder().encode(text);
  if (e === 'utf16le' || e === 'utf-16le' || e === 'ucs2' || e === 'ucs-2') {
    const out = new Uint8Array(text.length * 2);
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i);
      out[i * 2] = c & 0xff;
      out[i * 2 + 1] = c >> 8;
    }
    return out;
  }
  if (e === 'latin1' || e === 'binary') {
    const out = new Uint8Array(text.length);
    for (let i = 0; i < text.length; i++) out[i] = text.charCodeAt(i) & 0xff;
    return out;
  }
  throw invalid('encoding', `The encoding "${encoding ?? ''}" is not supported.`, 'Use "utf8" (default), "utf16le" or "latin1".');
}

function decoderFor(encoding: string): TextDecoder {
  const e = encoding.toLowerCase();
  const label = e === 'utf8' ? 'utf-8' : e === 'ucs2' || e === 'ucs-2' || e === 'utf16le' ? 'utf-16le' : e === 'binary' ? 'latin1' : e;
  try {
    return new TextDecoder(label);
  } catch {
    throw invalid('parserOptions.encoding', `The encoding "${encoding}" is not supported.`, 'Use "utf8" (default), "utf16le" or "latin1".');
  }
}

/** Text of a stream: a Node.js readable stream, a Web `ReadableStream` or any async iterable of bytes or strings. */
async function readText(stream: unknown, encoding: string): Promise<string> {
  const decoder = decoderFor(encoding);
  let text = '';
  let size = 0;
  // The input is collected before parsing, so it is limited while it arrives (D-J01).
  const count = (n: number): void => {
    size += n;
    if (size > DEFAULT_READ_LIMITS.maxInputBytes) {
      throw new XlvError('XLV_E_LIMIT_EXCEEDED', `The CSV input is larger than ${String(DEFAULT_READ_LIMITS.maxInputBytes)} bytes.`, {
        hint: 'Split the CSV file, or read it with the core fromCSV in parts.',
        details: { limit: 'maxInputBytes', value: DEFAULT_READ_LIMITS.maxInputBytes },
      });
    }
  };
  const add = (chunk: unknown): void => {
    if (typeof chunk === 'string') {
      count(chunk.length);
      text += chunk;
    } else if (chunk instanceof Uint8Array) {
      count(chunk.length);
      text += decoder.decode(chunk, { stream: true });
    } else if (chunk instanceof ArrayBuffer) {
      count(chunk.byteLength);
      text += decoder.decode(new Uint8Array(chunk), { stream: true });
    } else throw new XlvError('XLV_E_INVALID_INPUT', 'workbook.csv.read() got a stream chunk that is neither text nor bytes.', { hint: 'Pass a stream of strings, Buffers or Uint8Arrays.' });
  };
  const s = stream as { [Symbol.asyncIterator]?: unknown; [Symbol.iterator]?: unknown; getReader?: () => { read(): Promise<{ done: boolean; value?: unknown }>; releaseLock(): void } };
  if (typeof stream === 'string') add(stream);
  else if (s && (typeof s[Symbol.asyncIterator] === 'function' || (typeof s[Symbol.iterator] === 'function' && !(stream instanceof Uint8Array)))) {
    for await (const chunk of stream as AsyncIterable<unknown> | Iterable<unknown>) add(chunk);
  } else if (s && typeof s.getReader === 'function') {
    const reader = s.getReader();
    try {
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        add(value);
      }
    } finally {
      reader.releaseLock();
    }
  } else {
    throw new XlvError('XLV_E_INVALID_INPUT', 'workbook.csv.read() expects a readable stream.', {
      hint: 'Pass a Node.js readable stream (e.g. fs.createReadStream(path)), a Web ReadableStream or an async iterable of chunks; or use csv.readFile(path).',
    });
  }
  text += decoder.decode();
  // A byte order mark is not data (exceljs keeps it in the first field).
  return text.startsWith('\uFEFF') ? text.slice(1) : text;
}

/** A writable Node.js stream or Web `WritableStream`, as accepted by `csv.write()`. */
interface WritableLike {
  write?(chunk: Uint8Array): unknown;
  end?(callback?: () => void): unknown;
  getWriter?(): { write(chunk: Uint8Array): Promise<void>; close(): Promise<void>; releaseLock(): void };
}

async function writeTo(stream: unknown, bytes: Uint8Array): Promise<void> {
  const s = stream as WritableLike;
  if (s && typeof s.getWriter === 'function') {
    const writer = s.getWriter();
    try {
      await writer.write(bytes);
      await writer.close();
    } finally {
      writer.releaseLock();
    }
    return;
  }
  if (s && typeof s.write === 'function' && typeof s.end === 'function') {
    const end = s.end.bind(s);
    s.write(bytes);
    await new Promise<void>((resolve) => {
      end(() => resolve());
    });
    return;
  }
  throw new XlvError('XLV_E_INVALID_INPUT', 'workbook.csv.write() expects a writable stream.', {
    hint: 'Pass a Node.js writable stream (e.g. fs.createWriteStream(path)) or a Web WritableStream; or use csv.writeBuffer().',
  });
}

/** Node.js Buffer, when available, so `writeBuffer()` returns what exceljs returns there. */
interface BufferCtor {
  from(data: ArrayBuffer, offset: number, length: number): Uint8Array;
}

function asBuffer(bytes: Uint8Array): Uint8Array {
  const NodeBuffer = (globalThis as { Buffer?: BufferCtor }).Buffer;
  return NodeBuffer ? NodeBuffer.from(bytes.buffer as ArrayBuffer, bytes.byteOffset, bytes.byteLength) : bytes;
}

interface NodeFs {
  readFile(path: string): Promise<Uint8Array>;
  writeFile(path: string, data: Uint8Array): Promise<void>;
}

/**
 * node:fs/promises, loaded only when a file is read or written, through the package's own import
 * `#node-fs` (package.json "imports"): dist/node-fs.js (node:fs/promises) under the `node` condition (Node.js, Bun,
 * Deno), dist/no-node.js elsewhere, so browser and Workers bundles have no node: import (D-R07).
 */
async function nodeFs(method: string): Promise<NodeFs> {
  try {
    return (await import('#node-fs')) as unknown as NodeFs;
  } catch (cause) {
    throw new XlvError('XLV_E_NOT_SUPPORTED', `workbook.csv.${method}() needs a runtime with node:fs (Node.js, Bun or Deno).`, {
      hint: method === 'readFile' ? 'Read the file yourself and pass its contents as a stream to csv.read().' : 'Use csv.writeBuffer() and save the bytes yourself.',
      details: { feature: `csv.${method}` },
      cause,
    });
  }
}

/**
 * The `csv` property of a {@link Workbook}: CSV reading and writing as exceljs `workbook.csv`.
 * Reading adds a worksheet; writing takes one worksheet. Formula-like text is escaped on writing
 * unless `formulaInjection: "off"` is given.
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/compat';
 *
 * const workbook = new ExcelJS.Workbook();
 * workbook.addWorksheet('Data').addRow(['Tea', 3.5]);
 * const buffer = await workbook.csv.writeBuffer();
 * console.log(new TextDecoder().decode(buffer)); // Tea,3.5
 * ```
 */
export class CsvIO {
  /** @internal */
  constructor(private readonly workbook: Workbook) {}

  /** Reads CSV from a stream into a new worksheet and returns it. */
  async read(stream: unknown, options: CsvReadOptions = {}): Promise<Worksheet> {
    // exceljs adds the worksheet first (lib/csv/csv.js L49), so it exists even when parsing fails.
    const worksheet = this.workbook.addWorksheet(options.sheetName);
    const parser = parserOptions(options.parserOptions);
    const text = await readText(stream, parser.encoding);
    const formats = options.dateFormats ?? DEFAULT_DATE_FORMATS;
    const map = options.map ?? ((datum: string) => csvValue(datum, formats));
    // Each row is added as it is parsed, so the parsed rows are never all held at once (K2).
    for (const row of csvRows(text, parser)) worksheet.addRow(row.fields.map((datum, i) => map(datum, i)) as CellValue[]);
    return worksheet;
  }

  /** Reads a CSV file (Node.js, Bun, Deno) into a new worksheet. */
  async readFile(path: string, options: CsvReadOptions = {}): Promise<Worksheet> {
    const fs = await nodeFs('readFile');
    let bytes: Uint8Array;
    try {
      bytes = await fs.readFile(path);
    } catch (cause) {
      const code = typeof cause === 'object' && cause !== null && 'code' in cause ? String((cause as { code: unknown }).code) : 'UNKNOWN';
      throw new XlvError('XLV_E_FILE_READ', `Cannot read "${path}" (${code}).`, {
        hint: code === 'ENOENT' ? 'Check the path: the file does not exist.' : 'Check that the path points to a readable file.',
        details: { path, code },
        cause,
      });
    }
    return this.read([bytes], options);
  }

  /** Writes a worksheet as CSV to a stream and ends it. */
  async write(stream: unknown, options: CsvWriteOptions = {}): Promise<void> {
    await writeTo(stream, encode(this.text(options), 'utf8'));
  }

  /** Writes a worksheet to a CSV file (Node.js, Bun, Deno), replacing it. */
  async writeFile(path: string, options: CsvWriteOptions = {}): Promise<void> {
    const bytes = encode(this.text(options, outputEncoding(options.encoding)), options.encoding);
    const fs = await nodeFs('writeFile');
    try {
      await fs.writeFile(path, bytes);
    } catch (cause) {
      const code = typeof cause === 'object' && cause !== null && 'code' in cause ? String((cause as { code: unknown }).code) : 'UNKNOWN';
      throw new XlvError('XLV_E_FILE_WRITE', `Cannot write "${path}" (${code}).`, { hint: 'Check that the folder exists and can be written.', details: { path, code }, cause });
    }
  }

  /** The CSV of a worksheet as bytes (a Buffer in Node.js), UTF-8 encoded. */
  async writeBuffer(options: CsvWriteOptions = {}): Promise<Uint8Array> {
    return asBuffer(encode(this.text(options), 'utf8'));
  }

  /** Removed in exceljs 4 (it throws there too); use `read`. */
  createInputStream(): never {
    throw new Error(
      '`CSV#createInputStream` is deprecated. You should use `CSV#read` instead. This method will be removed in version 5.0. Please follow upgrade instruction: https://github.com/exceljs/exceljs/blob/master/UPGRADE-4.0.md',
    );
  }

  private text(options: CsvWriteOptions, encoding: OutputEncoding = 'utf8'): string {
    const key = options.sheetName || options.sheetId;
    const sheet = this.workbook.getWorksheet(key);
    if (key !== undefined && !sheet) {
      // exceljs writes an empty file here; xlvane names the missing sheet.
      throw new XlvError('XLV_E_SHEET_NOT_FOUND', `The workbook has no worksheet ${typeof key === 'string' ? `named "${key}"` : `with id ${String(key)}`}.`, {
        location: { sheet: String(key) },
        hint: `Available worksheets: ${this.workbook.worksheets.map((w) => `"${w.name}"`).join(', ') || 'none'}.`,
        details: { available: this.workbook.worksheets.map((w) => w.name) },
      });
    }
    return worksheetCsv(sheet, options, encoding);
  }
}
