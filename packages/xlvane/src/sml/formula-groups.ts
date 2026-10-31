// Array and shared formulas of one worksheet, checked while its cells are written in row, then
// column order (both writers). A master is the top-left cell of its range, so it is always written
// before the cells of its range. Rules: ECMA-376 Part 1 §18.3.1.40 and §18.17.6.3 (docs/facts/
// ecma-376.md E-26 … E-28): array formulas are written only in the top-left cell and the other
// cells of the range hold values only; shared groups are numbered by `si`, dependents lie inside
// the master's `ref`; master ranges on a sheet do not overlap. docs/facts/formulas.md F2-13 … F2-16.
import { XlvError } from '../errors';
import type { StoredValue } from '../write';
import { parseRange, type CellRange } from './cell-ref';

interface Group {
  readonly t: 'array' | 'shared';
  readonly ref: string;
  readonly range: CellRange;
  readonly si: number;
}

const inside = (r: CellRange, row: number, col: number): boolean => row >= r.start.row && row <= r.end.row && col >= r.start.col && col <= r.end.col;
const overlaps = (a: CellRange, b: CellRange): boolean => a.start.row <= b.end.row && b.start.row <= a.end.row && a.start.col <= b.end.col && b.start.col <= a.end.col;

function invalid(sheet: string, cell: string, message: string, hint: string): XlvError {
  return new XlvError('XLV_E_INVALID_VALUE', message, { location: { sheet, cell }, hint, details: { cell } });
}

/** Tracks the formula groups of one worksheet while its cells are written in order. */
export class FormulaGroups {
  private readonly active: Group[] = [];
  private readonly masters = new Map<string, Group>();
  private next = 0;

  constructor(private readonly sheet: string) {}

  /**
   * The attributes of the cell's `f` element (`t`, `ref`, `si`), or `''` for a plain formula and
   * for other values. Throws `XLV_E_INVALID_VALUE` when a rule above is broken.
   */
  attributes(row: number, col: number, cell: string, value: StoredValue | null): string {
    // Groups that end above this row can no longer contain a cell (row, then column order).
    for (let i = this.active.length - 1; i >= 0; i--) if (this.active[i]!.range.end.row < row) this.active.splice(i, 1);
    const array = this.active.find((g) => g.t === 'array' && inside(g.range, row, col));
    if (value === null || value.kind !== 'formula') return '';
    if (array !== undefined) {
      throw invalid(this.sheet, cell, `${cell} lies in the range ${array.ref} of an array formula, so it cannot hold its own formula.`, `Give ${cell} its cached result as a plain value; the array formula in the top-left cell of ${array.ref} covers it.`);
    }
    if (value.master !== undefined) {
      const master = this.masters.get(value.master);
      if (master === undefined) {
        throw invalid(this.sheet, cell, `${cell} shares the formula of ${value.master}, which holds no shared formula master written before it.`, `Set ${value.master} to { formula, shared: "${value.master}:…" }, a range whose top-left cell is ${value.master} and that contains ${cell}.`);
      }
      if (!inside(master.range, row, col)) {
        throw invalid(this.sheet, cell, `${cell} lies outside the range ${master.ref} of the shared formula in ${value.master}.`, `Extend the shared range of ${value.master} to cover ${cell}, or give ${cell} its own formula.`);
      }
      return ` t="shared" si="${String(master.si)}"`;
    }
    const g = value.group;
    if (g === undefined) return '';
    const range = parseRange(g.ref)!;
    const clash = this.active.find((a) => overlaps(a.range, range));
    if (clash !== undefined) {
      throw invalid(this.sheet, cell, `The ${g.t} formula range ${g.ref} of ${cell} overlaps the ${clash.t} formula range ${clash.ref}.`, 'Array and shared formula ranges of a sheet must not overlap. Shrink one of the ranges.');
    }
    const group: Group = { t: g.t, ref: g.ref, range, si: g.t === 'shared' ? this.next++ : -1 };
    this.active.push(group);
    if (g.t === 'shared') {
      this.masters.set(cell, group);
      return ` t="shared" ref="${g.ref}" si="${String(group.si)}"`;
    }
    return ` t="array" ref="${g.ref}"`;
  }
}
