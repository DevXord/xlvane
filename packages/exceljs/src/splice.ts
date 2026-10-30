// Helpers for row and column splicing (worksheet.spliceRows / spliceColumns, row.splice). xlvane
// behaves like a spreadsheet application here, not like exceljs 4.4.0, which moves only values,
// styles and heights (docs/facts/exceljs.md X4-15; docs/decisions.md D-B03).
import { adjustFormula, adjustRange, shiftFormula, XlvError } from 'xlvane';

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

/** The last row or column of the grid on the axis of `c`. */
export function axisMax(c: SpliceChange): number {
  return c.axis === 'row' ? 1_048_576 : 16_384;
}

function letters(n: number): string {
  let out = '';
  for (let v = n; v > 0; v = Math.floor((v - 1) / 26)) out = String.fromCharCode(65 + ((v - 1) % 26)) + out;
  return out;
}

function number(text: string): number {
  let n = 0;
  for (const ch of text) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}

/**
 * The removal and the insertion of a splice as two changes for xlvane's reference engine. Counts
 * are limited to what is left of the grid, so a removal past the last row or column is valid.
 */
function engineChanges(c: SpliceChange): { at: number; count: number }[] {
  const out: { at: number; count: number }[] = [];
  const max = axisMax(c);
  if (c.count > 0) out.push({ at: c.at, count: -Math.min(c.count, max - c.at + 1) });
  if (c.inserts > 0) out.push({ at: c.at, count: Math.min(c.inserts, max) });
  return out;
}

/**
 * The new span of `[lo, hi]` after the splice, with the rules of xlvane's `adjustRange`
 * (docs/facts/ref-adjust.md): the removed part is cut out (the span shrinks, and disappears when
 * nothing is left), then inserted positions at or above the first position move the span, inside
 * it they extend it; an end pushed beyond the grid stops at the last row or column, a start pushed
 * beyond it removes the span (R-09).
 */
export function moveSpan(lo: number, hi: number, c: SpliceChange): [number, number] | null {
  let text = c.axis === 'row' ? `A${String(lo)}:A${String(hi)}` : `${letters(lo)}1:${letters(hi)}1`;
  for (const step of engineChanges(c)) {
    const next = adjustRange(text, { sheet: c.sheet, axis: c.axis, at: step.at, count: step.count });
    if (next === null || next === '') return null;
    text = next;
  }
  const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(text);
  if (!m) return null;
  return c.axis === 'row' ? [Number(m[2]), Number(m[4])] : [number(m[1] ?? ''), number(m[3] ?? '')];
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
  if (!mayRefer(text, formulaSheet, change.sheet)) return text;
  let out: string | null = text;
  for (const step of engineChanges(change)) {
    if (out === null) break;
    out = adjustFormula(out, { formulaSheet, sheet: change.sheet, axis: change.axis, at: step.at, count: step.count, sheets });
  }
  return out;
}

/**
 * Whether a formula on `formulaSheet` can refer to `sheet` at all: it is on that sheet, it names
 * the sheet (also quoted, with '' for an apostrophe), or it may hold a 3-D prefix (a ":" before a
 * "!"). Formulas that cannot are left unchanged without being parsed.
 */
function mayRefer(text: string, formulaSheet: string, sheet: string): boolean {
  const name = sheet.toLowerCase();
  if (formulaSheet.toLowerCase() === name) return true;
  if (!text.includes('!')) return false;
  const lower = text.toLowerCase();
  return lower.includes(name) || lower.includes(name.replace(/'/g, "''")) || /:[^!]*!/.test(text);
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
