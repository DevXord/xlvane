// A1-style cell references as stored in SpreadsheetML attributes (c@r, mergeCell@ref,
// dimension@ref). The input comes from untrusted files, so every parser is a hand-written linear
// scanner with hard length caps; no regular expressions.
//
// Syntax decisions (see the facts fragment for the sources):
// - Column letters are not case-sensitive (ECMA-376 Part 1 §18.17.2.3.1), so "b7" is accepted.
// - `$` markers are rejected: c@r is ST_CellRef and the refs are ST_Ref, while the spec has a
//   separate type, ST_RefA, for single cell references that may be absolute (§18.18.7,
//   §18.18.62, §18.18.63).
// - A row is a digit sequence whose value is non-zero (§18.17.2.3.1), so leading zeros are allowed.
// - Whole-row and whole-column references are rejected (dimension@ref, §18.3.1.35).

/** Number of rows in a worksheet: rows 1 to 1 048 576 (ECMA-376 Part 1 §18.17.5.1). */
export const MAX_ROWS: number = 1_048_576;

/** Number of columns in a worksheet: columns A to XFD (ECMA-376 Part 1 §18.17.5.1). */
export const MAX_COLUMNS: number = 16_384;

/** A cell position; both numbers are 1-based. */
export interface CellAddress {
  readonly row: number;
  readonly col: number;
}

/** A rectangular range given by its top-left and bottom-right corners. */
export interface CellRange {
  readonly start: CellAddress;
  readonly end: CellAddress;
}

/** "XFD" has three letters and "1048576" seven digits, so longer inputs are never valid. */
const MAX_LETTERS = 3;
const MAX_DIGITS = 7;
const MAX_CELL_REF_LENGTH = MAX_LETTERS + MAX_DIGITS;

const CODE_A = 0x41;
const CODE_Z = 0x5a;
const CODE_LOWER_A = 0x61;
const CODE_LOWER_Z = 0x7a;
const CODE_0 = 0x30;
const CODE_9 = 0x39;
const CODE_COLON = 0x3a;

/** Value 1–26 of an ASCII letter (either case), or 0 for any other UTF-16 code unit. */
function letterValue(code: number): number {
  if (code >= CODE_A && code <= CODE_Z) return code - CODE_A + 1;
  if (code >= CODE_LOWER_A && code <= CODE_LOWER_Z) return code - CODE_LOWER_A + 1;
  return 0;
}

function isRow(row: number): boolean {
  return Number.isInteger(row) && row >= 1 && row <= MAX_ROWS;
}

function isColumn(col: number): boolean {
  return Number.isInteger(col) && col >= 1 && col <= MAX_COLUMNS;
}

/**
 * Returns the column letters for a 1-based column number: 1 → "A", 27 → "AA", 16384 → "XFD".
 * Throws `RangeError` outside 1..MAX_COLUMNS.
 */
export function columnLetters(col: number): string {
  if (!isColumn(col)) throw new RangeError(`Column number must be a whole number from 1 to ${MAX_COLUMNS}, got ${col}.`);
  // Bijective base 26: A–Z, then AA–AZ, BA–BZ, …, ZA–ZZ, AAA… (ECMA-376 Part 1 §18.17.2.3.1).
  let letters = '';
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(CODE_A + rem) + letters;
    n = (n - 1 - rem) / 26;
  }
  return letters;
}

/**
 * Parses the letters at `ref[from..to)` as a column. Returns the 1-based column number, or 0 if
 * the slice is empty, longer than three letters, contains a non-letter or is beyond XFD.
 */
function scanColumn(ref: string, from: number, to: number): number {
  const length = to - from;
  if (length < 1 || length > MAX_LETTERS) return 0;
  let col = 0;
  for (let i = from; i < to; i++) {
    const value = letterValue(ref.charCodeAt(i));
    if (value === 0) return 0;
    col = col * 26 + value;
  }
  return col <= MAX_COLUMNS ? col : 0;
}

/**
 * Returns the 1-based column number for column letters ("A" → 1, "xfd" → 16384; case does not
 * matter), or `null` if `letters` is not 1–3 ASCII letters or names a column beyond XFD.
 */
export function columnNumber(letters: string): number | null {
  const col = scanColumn(letters, 0, letters.length);
  return col === 0 ? null : col;
}

/** Parses `ref[from..to)` as a single cell reference; returns `null` if it is not one. */
function scanCellRef(ref: string, from: number, to: number): CellAddress | null {
  const length = to - from;
  if (length < 2 || length > MAX_CELL_REF_LENGTH) return null;
  let i = from;
  while (i < to && letterValue(ref.charCodeAt(i)) !== 0) i++;
  const col = scanColumn(ref, from, i);
  if (col === 0) return null;
  const digitsStart = i;
  if (to - digitsStart < 1 || to - digitsStart > MAX_DIGITS) return null;
  let row = 0;
  for (; i < to; i++) {
    const code = ref.charCodeAt(i);
    if (code < CODE_0 || code > CODE_9) return null;
    row = row * 10 + (code - CODE_0);
  }
  // The row "shall be non-zero" (ECMA-376 Part 1 §18.17.2.3.1) and inside the grid (§18.17.5.1).
  if (row < 1 || row > MAX_ROWS) return null;
  return { row, col };
}

/**
 * Parses an A1-style cell reference such as `"B7"` into `{ row: 7, col: 2 }`. Lowercase letters
 * are accepted; `$` markers, ranges, whitespace and references outside the grid give `null`.
 */
export function parseCellRef(ref: string): CellAddress | null {
  return scanCellRef(ref, 0, ref.length);
}

/** Formats a 1-based row and column as an A1 reference. Throws `RangeError` outside the grid. */
export function formatCellRef(row: number, col: number): string {
  if (!isRow(row)) throw new RangeError(`Row number must be a whole number from 1 to ${MAX_ROWS}, got ${row}.`);
  return columnLetters(col) + String(row);
}

/**
 * Parses a range such as `"A1:C3"`, or a single cell such as `"B2"` (a one-cell range). The
 * corners are normalised so that `start` is the top-left and `end` the bottom-right cell.
 * Whole-row or whole-column references, `$` markers and malformed input give `null`.
 */
export function parseRange(ref: string): CellRange | null {
  const length = ref.length;
  if (length > MAX_CELL_REF_LENGTH * 2 + 1) return null;
  let colon = -1;
  for (let i = 0; i < length; i++) {
    if (ref.charCodeAt(i) === CODE_COLON) {
      if (colon !== -1) return null;
      colon = i;
    }
  }
  if (colon === -1) {
    const cell = scanCellRef(ref, 0, length);
    return cell === null ? null : { start: cell, end: cell };
  }
  const a = scanCellRef(ref, 0, colon);
  if (a === null) return null;
  const b = scanCellRef(ref, colon + 1, length);
  if (b === null) return null;
  return {
    start: { row: Math.min(a.row, b.row), col: Math.min(a.col, b.col) },
    end: { row: Math.max(a.row, b.row), col: Math.max(a.col, b.col) },
  };
}

/**
 * Formats a range as `"A1:C3"`, or as a single reference such as `"B2"` when both corners are the
 * same cell. The corners are written as given. Throws `RangeError` for corners outside the grid.
 */
export function formatRange(range: CellRange): string {
  const start = formatCellRef(range.start.row, range.start.col);
  if (range.start.row === range.end.row && range.start.col === range.end.col) return start;
  return `${start}:${formatCellRef(range.end.row, range.end.col)}`;
}
