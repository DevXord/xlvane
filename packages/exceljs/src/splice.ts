// Helpers for row and column splicing (worksheet.spliceRows / spliceColumns, row.splice). xlvane
// behaves like a spreadsheet application here, not like exceljs 4.4.0, which moves only values,
// styles and heights (docs/facts/exceljs.md X4-15; docs/decisions.md D-B03).
import { adjustFormula, shiftFormula, XlvError } from 'xlvane';

/** A splice on one axis: `count` positions from `at` are removed, then `inserts` positions are inserted at `at`. */
export interface SpliceChange {
  readonly sheet: string;
  readonly axis: 'row' | 'col';
  readonly at: number;
  readonly count: number;
  readonly inserts: number;
}

/**
 * The new position of `n` after the splice, or `null` when it was removed. Positions before `at`
 * stay; removed positions are those from `at` to `at + count - 1`; later ones move by
 * `inserts - count`.
 */
export function movePosition(n: number, c: SpliceChange): number | null {
  if (n < c.at) return n;
  if (n < c.at + c.count) return null;
  return n - c.count + c.inserts;
}

/**
 * The new span of `[lo, hi]` after the splice, as a spreadsheet application adjusts a range: the
 * removed part is cut out (the span shrinks, and disappears when nothing is left), then inserted
 * positions at or above the first position move the span, inside it they extend it.
 */
export function moveSpan(lo: number, hi: number, c: SpliceChange): [number, number] | null {
  let a = lo;
  let b = hi;
  if (c.count > 0) {
    const end = c.at + c.count - 1;
    if (b >= c.at) {
      if (a > end) {
        a -= c.count;
        b -= c.count;
      } else {
        a = a < c.at ? a : c.at;
        b = b > end ? b - c.count : c.at - 1;
        if (b < a) return null;
      }
    }
  }
  if (c.inserts > 0) {
    if (a >= c.at) {
      a += c.inserts;
      b += c.inserts;
    } else if (b >= c.at) b += c.inserts;
  }
  return [a, b];
}

/** Throws XLV_E_INVALID_ARGUMENT for an argument exceljs would accept and corrupt data with. */
export function invalidArgument(method: string, name: string, value: unknown, hint: string): never {
  const shown = typeof value === 'string' ? `"${value.slice(0, 40)}"` : String(value).slice(0, 40);
  throw new XlvError('XLV_E_INVALID_ARGUMENT', `${method}: ${name} ${shown} is not valid.`, {
    hint,
    details: { method, argument: name, value: shown },
  });
}

/** Checks an integer argument in `[min, max]`. */
export function checkInteger(method: string, name: string, value: unknown, min: number, max: number): number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < min || value > max) {
    invalidArgument(method, name, value, `Use a whole number from ${String(min)} to ${String(max)}.`);
  }
  return value;
}

/** A deep copy of plain style data, so that moved or copied cells never share style objects. */
export function copyPlain<T>(value: T): T {
  if (Array.isArray(value)) return value.map((v) => copyPlain(v) as unknown) as T;
  if (value !== null && typeof value === 'object' && !(value instanceof Date)) {
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(value as Record<string, unknown>)) out[k] = copyPlain(v);
    return out as T;
  }
  return value;
}

/**
 * The text of a formula on sheet `formulaSheet` after `change`, with references to the moved rows
 * or columns of `change.sheet` adjusted as a spreadsheet application does (xlvane `adjustFormula`,
 * docs/decisions.md D-B01): the removal of `count` positions, then the insertion of `inserts`.
 * `null` when a reference cannot be adjusted safely.
 */
export function adjustFormulaText(text: string, formulaSheet: string, change: SpliceChange, sheets: readonly string[]): string | null {
  const base = { formulaSheet, sheet: change.sheet, axis: change.axis, at: change.at, sheets };
  let out: string | null = text;
  if (change.count > 0) out = adjustFormula(out, { ...base, count: -change.count });
  if (out !== null && change.inserts > 0) out = adjustFormula(out, { ...base, count: change.inserts });
  return out;
}

/**
 * The formula text of a cell copied `dRow` rows down: relative references move with it (xlvane
 * `shiftFormula`). `null` when the text cannot be moved safely.
 */
export function copyFormula(text: string, dRow: number): string | null {
  return shiftFormula(text, dRow, 0);
}

/** XLV_E_NOT_SUPPORTED for a formula that cannot follow a splice or copy. */
export function formulaNotAdjustable(method: string, sheet: string, cell: string, formula: string): never {
  throw new XlvError('XLV_E_NOT_SUPPORTED', `${method}: the formula in ${sheet}!${cell} cannot be adjusted safely, so nothing was changed.`, {
    location: { sheet, cell },
    hint: 'Its references use syntax xlvane does not rewrite (for example a 3-D reference across the changed sheet, the @ or spill # operator). Replace the formula with a plain value or simpler references, then repeat the call.',
    details: { method, formula: formula.slice(0, 200) },
  });
}
