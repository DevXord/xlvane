// Moves a formula written for one cell to another cell, as a spreadsheet application does when a
// shared formula is reused (ECMA-376 Part 1 §18.3.1.40, docs/facts/ecma-376.md E-28): relative
// parts of A1 references move by the row and column offset, `$`-marked parts stay
// (docs/facts/formulas.md F2-01 … F2-12).
//
// The tokenizer is a single forward pass with bounded lookahead and no regular expressions, so
// it runs in linear time on untrusted text. It is deliberately conservative: anything it cannot
// classify with confidence (R1C1-looking names, unbalanced quotes or brackets, unknown
// characters, references that would leave the grid) makes the whole formula "not shiftable",
// and the caller keeps the formula text unknown rather than guess.
import { columnLetters, MAX_COLUMNS, MAX_ROWS } from './cell-ref';

/** A cell reference: column and row, each relative or absolute. */
interface CellPart {
  readonly kind: 'cell';
  readonly col: number;
  readonly colAbs: boolean;
  readonly row: number;
  readonly rowAbs: boolean;
}

/** A whole-column range such as `A:C` or `$A:C`. */
interface ColumnsPart {
  readonly kind: 'cols';
  readonly c1: number;
  readonly c1Abs: boolean;
  readonly c2: number;
  readonly c2Abs: boolean;
}

/** A whole-row range such as `1:3` or `$1:3`. */
interface RowsPart {
  readonly kind: 'rows';
  readonly r1: number;
  readonly r1Abs: boolean;
  readonly r2: number;
  readonly r2Abs: boolean;
}

type Part = string | CellPart | ColumnsPart | RowsPart;

/** A formula split into literal text and references, ready to be moved to other cells. */
export interface CompiledFormula {
  /** `true` if every relative reference stays inside A1:XFD1048576 when moved by (dRow, dCol). */
  canShift(dRow: number, dCol: number): boolean;
  /** The formula moved by (dRow, dCol). Call only when {@link canShift} returned `true`. */
  render(dRow: number, dCol: number): string;
}

/** Error constants (§18.17.3, fact E-33). */
const ERROR_CONSTANTS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'];
/** Characters copied as operators or separators (F2-07). */
const OPERATORS = '+-*/^&=<>%,;(){}:@ \t\r\n';

function isLetter(c: number): boolean {
  return (c >= 0x41 && c <= 0x5a) || (c >= 0x61 && c <= 0x7a);
}

function isDigit(c: number): boolean {
  return c >= 0x30 && c <= 0x39;
}

/** Characters allowed inside a name or reference token (F2-05). */
function isIdentChar(c: number): boolean {
  return isLetter(c) || isDigit(c) || c === 0x5f || c === 0x2e || c === 0x5c || c === 0x3f || c === 0x24 || c >= 0x80;
}

function isIdentStart(c: number): boolean {
  return isLetter(c) || c === 0x5f || c === 0x5c || c === 0x24 || c >= 0x80;
}

/** Column number of 1–3 ASCII letters at s[from..to), or 0. */
function columnOf(s: string, from: number, to: number): number {
  if (to - from < 1 || to - from > 3) return 0;
  let col = 0;
  for (let i = from; i < to; i++) {
    const c = s.charCodeAt(i);
    if (!isLetter(c)) return 0;
    col = col * 26 + ((c | 0x20) - 0x60);
  }
  return col <= MAX_COLUMNS ? col : 0;
}

/** Row number of the decimal digits at s[from..to) (1 … MAX_ROWS), or 0. */
function rowOf(s: string, from: number, to: number): number {
  if (to - from < 1 || to - from > 7) return 0;
  let row = 0;
  for (let i = from; i < to; i++) {
    const c = s.charCodeAt(i);
    if (!isDigit(c)) return 0;
    row = row * 10 + (c - 0x30);
  }
  return row >= 1 && row <= MAX_ROWS ? row : 0;
}

/** Parses `$?letters$?digits` (the whole of `id`) as a cell reference. */
function cellRef(id: string): CellPart | null {
  let i = 0;
  const colAbs = id.charCodeAt(i) === 0x24;
  if (colAbs) i++;
  const colStart = i;
  while (i < id.length && isLetter(id.charCodeAt(i))) i++;
  const col = columnOf(id, colStart, i);
  if (col === 0) return null;
  const rowAbs = id.charCodeAt(i) === 0x24;
  if (rowAbs) i++;
  const rowStart = i;
  while (i < id.length && isDigit(id.charCodeAt(i))) i++;
  if (i !== id.length) return null;
  const row = rowOf(id, rowStart, i);
  return row === 0 ? null : { kind: 'cell', col, colAbs, row, rowAbs };
}

/** Parses `$?letters` (the whole of `id`) as a column; returns 0 if it is not one. */
function columnRef(id: string): { col: number; abs: boolean } | null {
  const abs = id.charCodeAt(0) === 0x24;
  const col = columnOf(id, abs ? 1 : 0, id.length);
  return col === 0 ? null : { col, abs };
}

/** Names that look like R1C1 references (`R`, `C`, `R1`, `C2`, `RC`, `R1C1`): ambiguous, F2-09. */
function looksR1C1(id: string): boolean {
  let i = 0;
  const upper = (k: number): number => id.charCodeAt(k) & ~0x20;
  if (upper(i) === 0x52) {
    i++;
    while (i < id.length && isDigit(id.charCodeAt(i))) i++;
    if (i === id.length) return true;
  }
  if (upper(i) !== 0x43) return false;
  i++;
  while (i < id.length && isDigit(id.charCodeAt(i))) i++;
  return i === id.length;
}

/**
 * Splits a formula into literal text and references. Returns `null` when the text contains
 * anything that cannot be classified with confidence.
 */
function tokenize(s: string): Part[] | null {
  const parts: Part[] = [];
  let literal = '';
  /** Whether the previous token can be followed by the spill operator `#` (F2-08). */
  let afterOperand = false;
  const flush = (): void => {
    if (literal !== '') {
      parts.push(literal);
      literal = '';
    }
  };
  const n = s.length;
  let i = 0;
  while (i < n) {
    const c = s.charCodeAt(i);
    // String literal: "...", with "" for a quote (F2-03).
    if (c === 0x22) {
      let j = i + 1;
      for (;;) {
        if (j >= n) return null;
        if (s.charCodeAt(j) === 0x22) {
          if (s.charCodeAt(j + 1) === 0x22) j += 2;
          else break;
        } else j++;
      }
      literal += s.slice(i, j + 1);
      i = j + 1;
      afterOperand = true;
      continue;
    }
    // Quoted sheet name: '...' with '' for an apostrophe, followed by "!" (F2-04).
    if (c === 0x27) {
      let j = i + 1;
      for (;;) {
        if (j >= n) return null;
        if (s.charCodeAt(j) === 0x27) {
          if (s.charCodeAt(j + 1) === 0x27) j += 2;
          else break;
        } else j++;
      }
      if (s.charCodeAt(j + 1) !== 0x21) return null;
      literal += s.slice(i, j + 2);
      i = j + 2;
      afterOperand = false;
      continue;
    }
    // Bracketed text: structured references and external workbook indexes, kept as they are.
    // Brackets nest; inside them an apostrophe escapes the next character (F2-06).
    if (c === 0x5b) {
      let depth = 0;
      let j = i;
      while (j < n) {
        const d = s.charCodeAt(j);
        if (d === 0x27) j++;
        else if (d === 0x5b) depth++;
        else if (d === 0x5d && --depth === 0) break;
        j++;
      }
      if (j >= n) return null;
      literal += s.slice(i, j + 1);
      i = j + 1;
      afterOperand = true;
      continue;
    }
    if (c === 0x5d) return null;
    if (c === 0x21) {
      // "!" after a bracketed workbook index, e.g. [1]!Name.
      literal += '!';
      i++;
      afterOperand = false;
      continue;
    }
    // Error constants, or the spill operator after an operand (F2-08).
    if (c === 0x23) {
      const constant = ERROR_CONSTANTS.find((e) => s.startsWith(e, i));
      if (constant !== undefined) {
        literal += constant;
        i += constant.length;
        afterOperand = true;
        continue;
      }
      if (!afterOperand) return null;
      literal += '#';
      i++;
      continue;
    }
    // Numbers, and whole-row ranges such as 1:3 or $1:$3 (F2-02, F2-10).
    if (isDigit(c) || c === 0x2e || (c === 0x24 && isDigit(s.charCodeAt(i + 1)))) {
      const rows = rowRange(s, i);
      if (rows !== null) {
        flush();
        parts.push(rows.part);
        i = rows.end;
        afterOperand = true;
        continue;
      }
      if (c === 0x24) return null;
      let j = i;
      while (j < n && isDigit(s.charCodeAt(j))) j++;
      if (s.charCodeAt(j) === 0x2e) {
        j++;
        while (j < n && isDigit(s.charCodeAt(j))) j++;
      }
      if (j === i + 1 && c === 0x2e) return null;
      const e = s.charCodeAt(j);
      if (e === 0x45 || e === 0x65) {
        j++;
        if (s.charCodeAt(j) === 0x2b || s.charCodeAt(j) === 0x2d) j++;
        const digits = j;
        while (j < n && isDigit(s.charCodeAt(j))) j++;
        if (j === digits) return null;
      }
      // A number followed by a letter, "$" or ":" is not valid text (1A, 1:A).
      if (j < n && (isIdentChar(s.charCodeAt(j)) || s.charCodeAt(j) === 0x3a)) return null;
      literal += s.slice(i, j);
      i = j;
      afterOperand = true;
      continue;
    }
    // Names, functions, sheet names and references.
    if (isIdentStart(c)) {
      let j = i;
      while (j < n && isIdentChar(s.charCodeAt(j))) j++;
      const id = s.slice(i, j);
      const next = s.charCodeAt(j);
      let k = j;
      while (k < n && (s.charCodeAt(k) === 0x20 || s.charCodeAt(k) === 0x0a || s.charCodeAt(k) === 0x0d)) k++;
      const cell = cellRef(id);
      const column = cell === null ? columnRef(id) : null;
      const refLike = cell !== null || looksR1C1(id);
      // A name containing "." whose first segment is a reference (A1.B2) is not valid text.
      const dot = id.indexOf('.');
      if (dot > 0 && cellRef(id.slice(0, dot)) !== null) return null;
      if (next === 0x28) {
        // A function call: a name directly followed by "(" is a function even when it also
        // reads as a reference, as LOG10( does (F2-05).
        literal += id;
        i = j;
        afterOperand = false;
        continue;
      }
      // "(" after white space: a reference there is not valid text.
      if (s.charCodeAt(k) === 0x28 && refLike) return null;
      if (next === 0x21 || next === 0x5b) {
        // A sheet name before "!" or a table name before "[": must not look like a reference.
        if (refLike) return null;
        literal += id;
        i = j;
        afterOperand = false;
        continue;
      }
      if (cell !== null) {
        flush();
        parts.push(cell);
        i = j;
        afterOperand = true;
        continue;
      }
      if (column !== null && next === 0x3a) {
        let m = j + 1;
        while (m < n && isIdentChar(s.charCodeAt(m))) m++;
        const second = columnRef(s.slice(j + 1, m));
        if (second !== null) {
          // A:C is a column range unless it is a 3-D sheet range such as Jan:Feb!A1 (F2-11).
          if (s.charCodeAt(m) === 0x21) return null;
          flush();
          parts.push({ kind: 'cols', c1: column.col, c1Abs: column.abs, c2: second.col, c2Abs: second.abs });
          i = m;
          afterOperand = true;
          continue;
        }
      }
      if (id.includes('$') || looksR1C1(id)) return null;
      // A defined name, TRUE/FALSE or another identifier: copied unchanged (F2-05).
      literal += id;
      i = j;
      afterOperand = true;
      continue;
    }
    if (OPERATORS.includes(s[i] ?? '')) {
      literal += s[i];
      i++;
      afterOperand = c === 0x29 || c === 0x7d || c === 0x25;
      continue;
    }
    return null;
  }
  flush();
  return parts;
}

/** Parses `$?digits:$?digits` at `i` as a whole-row range, or returns `null`. */
function rowRange(s: string, i: number): { part: RowsPart; end: number } | null {
  const one = (at: number): { row: number; abs: boolean; end: number } | null => {
    let j = at;
    const abs = s.charCodeAt(j) === 0x24;
    if (abs) j++;
    const start = j;
    while (j < s.length && isDigit(s.charCodeAt(j))) j++;
    const row = rowOf(s, start, j);
    return row === 0 ? null : { row, abs, end: j };
  };
  const a = one(i);
  if (a === null || s.charCodeAt(a.end) !== 0x3a) return null;
  const b = one(a.end + 1);
  if (b === null || (b.end < s.length && isIdentChar(s.charCodeAt(b.end)))) return null;
  return { part: { kind: 'rows', r1: a.row, r1Abs: a.abs, r2: b.row, r2Abs: b.abs }, end: b.end };
}

class Compiled implements CompiledFormula {
  // Smallest and largest relative row and column used, for the O(1) bounds check.
  private minRow = Infinity;
  private maxRow = -Infinity;
  private minCol = Infinity;
  private maxCol = -Infinity;

  constructor(private readonly parts: readonly Part[]) {
    const row = (r: number, abs: boolean): void => {
      if (abs) return;
      if (r < this.minRow) this.minRow = r;
      if (r > this.maxRow) this.maxRow = r;
    };
    const col = (c: number, abs: boolean): void => {
      if (abs) return;
      if (c < this.minCol) this.minCol = c;
      if (c > this.maxCol) this.maxCol = c;
    };
    for (const p of parts) {
      if (typeof p === 'string') continue;
      if (p.kind === 'cell') {
        row(p.row, p.rowAbs);
        col(p.col, p.colAbs);
      } else if (p.kind === 'cols') {
        col(p.c1, p.c1Abs);
        col(p.c2, p.c2Abs);
      } else {
        row(p.r1, p.r1Abs);
        row(p.r2, p.r2Abs);
      }
    }
  }

  canShift(dRow: number, dCol: number): boolean {
    if (!Number.isInteger(dRow) || !Number.isInteger(dCol)) return false;
    if (this.minRow !== Infinity && (this.minRow + dRow < 1 || this.maxRow + dRow > MAX_ROWS)) return false;
    if (this.minCol !== Infinity && (this.minCol + dCol < 1 || this.maxCol + dCol > MAX_COLUMNS)) return false;
    return true;
  }

  render(dRow: number, dCol: number): string {
    let out = '';
    const col = (c: number, abs: boolean): string => (abs ? '$' : '') + columnLetters(abs ? c : c + dCol);
    const row = (r: number, abs: boolean): string => (abs ? '$' : '') + String(abs ? r : r + dRow);
    for (const p of this.parts) {
      if (typeof p === 'string') out += p;
      else if (p.kind === 'cell') out += col(p.col, p.colAbs) + row(p.row, p.rowAbs);
      else if (p.kind === 'cols') out += `${col(p.c1, p.c1Abs)}:${col(p.c2, p.c2Abs)}`;
      else out += `${row(p.r1, p.r1Abs)}:${row(p.r2, p.r2Abs)}`;
    }
    return out;
  }
}

/**
 * Splits a formula (without the leading "=") into text and references once, so that it can be
 * moved to many cells cheaply. Returns `null` if the formula cannot be classified with
 * confidence; such a formula must not be moved.
 */
export function compileFormula(text: string): CompiledFormula | null {
  const parts = tokenize(text);
  return parts === null ? null : new Compiled(parts);
}

/**
 * Moves a formula by `dRow` rows and `dCol` columns: relative parts of A1 references change,
 * `$`-marked parts, strings, names, functions, numbers and structured references do not.
 * Returns `null` if the text cannot be classified with confidence or a reference would leave
 * A1:XFD1048576.
 */
export function shiftFormula(text: string, dRow: number, dCol: number): string | null {
  const compiled = compileFormula(text);
  return compiled !== null && compiled.canShift(dRow, dCol) ? compiled.render(dRow, dCol) : null;
}
