// Adjusts the references of a formula or a cell range when rows or columns of a sheet are
// inserted or deleted, as a spreadsheet application does (docs/facts/ref-adjust.md R-01 … R-14,
// docs/decisions.md D-B01). Built on the shared-formula tokenizer (sml/formula-shift.ts), so it
// accepts exactly what that tokenizer classifies with confidence and returns `null` otherwise.
import { XlvError } from '../errors';
import { columnLetters, MAX_COLUMNS, MAX_ROWS } from './cell-ref';
import { tokenize, type CellPart, type ColumnsPart, type Part, type PrefixPart, type RowsPart } from './formula-shift';

/**
 * Rows or columns inserted into, or deleted from, one sheet.
 *
 * @example
 * ```ts
 * import type { ReferenceChange } from 'xlvane';
 *
 * // Two rows inserted before row 5 of "Sales".
 * const change: ReferenceChange = { sheet: 'Sales', axis: 'row', at: 5, count: 2 };
 * ```
 */
export interface ReferenceChange {
  /** The sheet whose rows or columns change. Sheet names are compared ignoring case. */
  readonly sheet: string;
  /**
   * The sheet that holds the formula; references without a sheet prefix belong to it.
   * Default: `sheet`.
   */
  readonly formulaSheet?: string;
  /** `'row'` or `'col'`. */
  readonly axis: 'row' | 'col';
  /** The first row or column (1-based) of the change: insertion happens before it, deletion starts at it. */
  readonly at: number;
  /** How many rows or columns: positive to insert, negative to delete. */
  readonly count: number;
  /**
   * The sheet names of the workbook in order. Needed only to adjust 3-D references
   * (`Jan:Mar!A1`); without it a formula with a 3-D reference cannot be adjusted.
   */
  readonly sheets?: readonly string[];
}

const REF = '#REF!';

interface Change {
  readonly sheet: string;
  readonly formulaSheet: string;
  readonly row: boolean;
  readonly at: number;
  readonly count: number;
  readonly max: number;
  readonly sheets: readonly string[] | null;
}

function invalid(option: string, message: string, hint: string): XlvError {
  return new XlvError('XLV_E_INVALID_OPTION', message, { hint, details: { option } });
}

function check(change: ReferenceChange): Change {
  if (typeof change !== 'object' || change === null) throw invalid('change', 'The reference change must be an object.', "Pass { sheet: 'Sheet1', axis: 'row', at: 5, count: 2 }.");
  const { sheet, axis, at, count } = change;
  const formulaSheet = change.formulaSheet ?? sheet;
  if (typeof sheet !== 'string' || sheet === '') throw invalid('sheet', 'The option "sheet" must be a sheet name.', "For example { sheet: 'Sales' }.");
  if (typeof formulaSheet !== 'string' || formulaSheet === '') throw invalid('formulaSheet', 'The option "formulaSheet" must be a sheet name.', "For example { formulaSheet: 'Summary' }.");
  if (axis !== 'row' && axis !== 'col') throw invalid('axis', `The option "axis" must be "row" or "col", not ${JSON.stringify(axis)}.`, "Use { axis: 'row' } for rows or { axis: 'col' } for columns.");
  const max = axis === 'row' ? MAX_ROWS : MAX_COLUMNS;
  if (!Number.isInteger(at) || at < 1 || at > max) throw invalid('at', `The option "at" must be a whole number from 1 to ${String(max)}.`, 'For example { at: 5 } for row 5 or column E.');
  if (!Number.isInteger(count) || count === 0 || (count < 0 && at - count - 1 > max) || Math.abs(count) > max) {
    throw invalid('count', `The option "count" must be a whole number other than 0, and a deletion must end at ${String(max)} or before.`, 'Use a positive count to insert and a negative count to delete, for example { count: -2 }.');
  }
  let sheets: readonly string[] | null = null;
  if (change.sheets !== undefined) {
    if (!Array.isArray(change.sheets) || !change.sheets.every((s) => typeof s === 'string')) throw invalid('sheets', 'The option "sheets" must be an array of sheet names.', "For example { sheets: ['Jan', 'Feb', 'Mar'] }.");
    sheets = change.sheets.map((s) => s.toLowerCase());
  }
  return { sheet: sheet.toLowerCase(), formulaSheet: formulaSheet.toLowerCase(), row: axis === 'row', at, count, max, sheets };
}

/** New position of one coordinate, or `null` when it is deleted or pushed beyond the grid. */
function moveOne(v: number, c: Change): number | null {
  if (c.count > 0) {
    const out = v >= c.at ? v + c.count : v;
    return out > c.max ? null : out;
  }
  const end = c.at - c.count - 1;
  if (v < c.at) return v;
  if (v > end) return v + c.count;
  return null;
}

/**
 * New ends of a span v1..v2 (v1 ≤ v2), or `null` when it disappears (R-03 … R-12):
 * - insertion moves every end at or after `at`; a start pushed beyond the grid gives `null`, an end
 *   pushed beyond it stops at the last row or column;
 * - deletion cuts the deleted band out; a span entirely inside it gives `null`; a span longer than
 *   one row or column whose end is the last one keeps that end.
 */
function moveSpan(v1: number, v2: number, c: Change): [number, number] | null {
  if (c.count > 0) {
    const a = v1 >= c.at ? v1 + c.count : v1;
    if (a > c.max) return null;
    const b = v2 >= c.at ? v2 + c.count : v2;
    return [a, Math.min(b, c.max)];
  }
  const end = c.at - c.count - 1;
  if (v1 >= c.at && v2 <= end) return null;
  const a = v1 < c.at ? v1 : v1 > end ? v1 + c.count : c.at;
  const b = v2 === c.max && v1 < v2 ? v2 : v2 < c.at ? v2 : v2 > end ? v2 + c.count : c.at - 1;
  return [a, b];
}

/** Whether a reference with this prefix points into the changed sheet; `null` if unknown. */
function onChangedSheet(prefix: PrefixPart | null, c: Change): boolean | null {
  if (prefix === null) return c.formulaSheet === c.sheet;
  if (prefix.external) return false;
  const names = prefix.sheets.map((s) => s.toLowerCase());
  if (names.length === 1) return names[0] === c.sheet;
  // A 3-D reference covers every sheet between the two names in workbook order (R-13).
  if (c.sheets === null) return null;
  const i = c.sheets.indexOf(names[0] ?? '');
  const j = c.sheets.indexOf(names[1] ?? '');
  if (i < 0 || j < 0) return null;
  const k = c.sheets.indexOf(c.sheet);
  return k >= Math.min(i, j) && k <= Math.max(i, j) ? null : false;
}

function cellText(p: CellPart, col: number, row: number): string {
  return (p.colAbs ? '$' : '') + columnLetters(col) + (p.rowAbs ? '$' : '') + String(row);
}

/** Adjusts a single cell reference; returns its new text or `#REF!`. */
function adjustCell(p: CellPart, c: Change): string {
  const v = moveOne(c.row ? p.row : p.col, c);
  if (v === null) return REF;
  return c.row ? cellText(p, p.col, v) : cellText(p, v, p.row);
}

/** Adjusts a range of two cell references; returns its new text or `#REF!`. */
function adjustCellRange(first: CellPart, second: CellPart, c: Change): string | null {
  let a = first;
  let b = second;
  // A reversed range (B9:B2) is the same range as B2:B9; it is written in order (R-15).
  if (c.row ? a.row > b.row : a.col > b.col) {
    a = c.row ? { ...first, row: second.row, rowAbs: second.rowAbs } : { ...first, col: second.col, colAbs: second.colAbs };
    b = c.row ? { ...second, row: first.row, rowAbs: first.rowAbs } : { ...second, col: first.col, colAbs: first.colAbs };
  }
  const v1 = c.row ? a.row : a.col;
  const v2 = c.row ? b.row : b.col;
  const span = moveSpan(v1, v2, c);
  if (span === null) return REF;
  return c.row ? `${cellText(a, a.col, span[0])}:${cellText(b, b.col, span[1])}` : `${cellText(a, span[0], a.row)}:${cellText(b, span[1], b.row)}`;
}

function adjustLines(p: RowsPart | ColumnsPart, c: Change): string | null {
  const rows = p.kind === 'rows';
  // Whole rows do not change when columns change, and whole columns not when rows change.
  if (rows !== c.row) return render(p);
  let v1 = rows ? p.r1 : p.c1;
  let v2 = rows ? p.r2 : p.c2;
  let abs1 = rows ? p.r1Abs : p.c1Abs;
  let abs2 = rows ? p.r2Abs : p.c2Abs;
  // A reversed range (9:2) is written in order (R-15).
  if (v1 > v2) [v1, v2, abs1, abs2] = [v2, v1, abs2, abs1];
  const span = moveSpan(v1, v2, c);
  if (span === null) return REF;
  const show = (v: number, abs: boolean): string => (abs ? '$' : '') + (rows ? String(v) : columnLetters(v));
  return `${show(span[0], abs1)}:${show(span[1], abs2)}`;
}

function render(p: CellPart | RowsPart | ColumnsPart): string {
  if (p.kind === 'cell') return cellText(p, p.col, p.row);
  if (p.kind === 'rows') return `${p.r1Abs ? '$' : ''}${String(p.r1)}:${p.r2Abs ? '$' : ''}${String(p.r2)}`;
  return `${p.c1Abs ? '$' : ''}${columnLetters(p.c1)}:${p.c2Abs ? '$' : ''}${columnLetters(p.c2)}`;
}

function adjustParts(parts: readonly Part[], c: Change): string | null {
  let out = '';
  let i = 0;
  while (i < parts.length) {
    let p = parts[i] as Part;
    let prefix: PrefixPart | null = null;
    if (typeof p === 'string') {
      out += p;
      i++;
      continue;
    }
    if (p.kind === 'prefix') {
      prefix = p;
      const next = parts[i + 1];
      // A prefix followed by text qualifies a defined name or an error (`Sheet1!Total`,
      // `Sheet1!#REF!`), not cells; it is kept as written. Anything else is not understood.
      if (typeof next === 'string' && /^(?:#REF!|[A-Za-z_\\\u0080-\uFFFF])/.test(next)) {
        out += p.text;
        i++;
        continue;
      }
      if (next === undefined || typeof next === 'string' || next.kind === 'prefix') return null;
      p = next;
      i++;
    }
    // A range of two cells: A1:B5, written as cell ":" cell.
    const colon = parts[i + 1];
    const second = parts[i + 2];
    const range = p.kind === 'cell' && colon === ':' && second !== undefined && typeof second !== 'string' && second.kind === 'cell' ? second : null;
    if (p.kind === 'cell' && colon === ':' && second !== undefined && typeof second !== 'string' && second.kind === 'prefix') return null;
    const on = onChangedSheet(prefix, c);
    if (on === null) return null;
    const head = prefix?.text ?? '';
    let text: string | null;
    if (!on) text = range === null ? render(p) : `${render(p)}:${render(range)}`;
    else if (p.kind === 'cell') text = range === null ? adjustCell(p, c) : adjustCellRange(p, range, c);
    else text = adjustLines(p, c);
    if (text === null) return null;
    // A reference that no longer exists becomes #REF! without its sheet prefix (R-06).
    out += text === REF ? REF : head + text;
    i += range === null ? 1 : 3;
  }
  return out;
}

/**
 * Rewrites the references of a formula after rows or columns of a sheet are inserted or deleted,
 * as spreadsheet applications do: references at or after the change move, `$`-marked ones too;
 * a range grows when rows or columns are inserted inside it and shrinks when some of its rows or
 * columns are deleted; a reference whose cells are all deleted becomes `#REF!`. References to
 * other sheets, other workbooks, defined names, structured references and text in strings are not
 * changed. Returns `null` when the formula cannot be adjusted safely: text the tokenizer does not
 * classify with confidence (the same rules as shared formulas, docs/facts/formulas.md), or a 3-D
 * reference that may cover the changed sheet. The formula is given without the leading "=".
 *
 * @example
 * ```ts
 * import { adjustFormula } from 'xlvane';
 *
 * // Two rows inserted before row 5 of "Sales".
 * adjustFormula('SUM(B2:B10)*$B$12', { sheet: 'Sales', axis: 'row', at: 5, count: 2 }); // 'SUM(B2:B12)*$B$14'
 * // Rows 5 and 6 deleted; the formula is on another sheet.
 * adjustFormula('Sales!B5+Sales!B9', { sheet: 'Sales', formulaSheet: 'Summary', axis: 'row', at: 5, count: -2 }); // '#REF!+Sales!B7'
 * ```
 */
export function adjustFormula(text: string, change: ReferenceChange): string | null {
  const c = check(change);
  if (typeof text !== 'string') throw invalid('text', 'The formula must be a string.', "Pass the formula text without the leading '=', for example 'SUM(A1:A3)'.");
  const parts = tokenize(text);
  return parts === null ? null : adjustParts(parts, c);
}

/**
 * Rewrites a list of cell ranges of the changed sheet, such as the `sqref` of a conditional format
 * or the `ref` of a hyperlink (space-separated ranges like `A1:B5 D7`), after rows or columns are
 * inserted or deleted, with the rules of {@link adjustFormula}. Ranges whose cells are all deleted
 * are left out, so the result is `''` when nothing remains. Returns `null` if the text is not a
 * list of A1 references.
 *
 * @example
 * ```ts
 * import { adjustRange } from 'xlvane';
 *
 * adjustRange('A1:B5 D7', { sheet: 'Sales', axis: 'row', at: 3, count: 1 }); // 'A1:B6 D8'
 * adjustRange('A5:A6', { sheet: 'Sales', axis: 'row', at: 5, count: -2 }); // ''
 * ```
 */
export function adjustRange(sqref: string, change: ReferenceChange): string | null {
  const c = check(change);
  if (typeof sqref !== 'string') throw invalid('sqref', 'The ranges must be a string.', "For example 'A1:B5 D7'.");
  const out: string[] = [];
  for (const item of sqref.split(' ')) {
    if (item === '') continue;
    const parts = tokenize(item);
    if (parts === null) return null;
    const isCell = (p: Part | undefined): p is CellPart => p !== undefined && typeof p !== 'string' && p.kind === 'cell';
    let text: string;
    if (parts.length === 1 && isCell(parts[0])) text = adjustCell(parts[0], { ...c, formulaSheet: c.sheet });
    else if (parts.length === 3 && isCell(parts[0]) && parts[1] === ':' && isCell(parts[2])) {
      const t = adjustCellRange(parts[0], parts[2], c);
      if (t === null) return null;
      text = t;
    } else return null;
    if (text !== REF) out.push(text);
  }
  return out.join(' ');
}
