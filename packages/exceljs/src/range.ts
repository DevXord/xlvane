// The range object of `worksheet.dimensions` and `worksheetReader.dimensions`, as exceljs 4.4.0
// `Range` (lib/doc/range.js): the model keeps 0 for an empty range, the getters give at least 1,
// and `toString()` is the range text (docs/facts/exceljs.md, "Range").
import { columnLetter } from './model';
import type { Location, Range } from './compat-types';

/** Implements exceljs `Range` for sheet dimensions; created by the worksheet, not by users. */
export class Dimensions implements Range {
  readonly model: Location & { sheetName?: string } = { top: 0, left: 0, bottom: 0, right: 0 };

  get top(): number {
    return this.model.top || 1;
  }

  get left(): number {
    return this.model.left || 1;
  }

  get bottom(): number {
    return this.model.bottom || 1;
  }

  get right(): number {
    return this.model.right || 1;
  }

  /** Grows the range to include the given rectangle (exceljs `expand`, range.js L163). */
  expand(top: number, left: number, bottom: number, right: number): void {
    const m = this.model;
    if (!m.top || top < this.top) m.top = top;
    if (!m.left || left < this.left) m.left = left;
    if (!m.bottom || bottom > this.bottom) m.bottom = bottom;
    if (!m.right || right > this.right) m.right = right;
  }

  /** Grows the range to include the cells of a row (exceljs `expandRow`, range.js L170). */
  expandRow(row: { readonly number: number; readonly dimensions: { min: number; max: number } | null }): void {
    const d = row.dimensions;
    if (d) this.expand(row.number, d.min, row.number, d.max);
  }

  get tl(): string {
    return `${columnLetter(this.left)}${String(this.top)}`;
  }

  get $t$l(): string {
    return `$${columnLetter(this.left)}$${String(this.top)}`;
  }

  get br(): string {
    return `${columnLetter(this.right)}${String(this.bottom)}`;
  }

  get $b$r(): string {
    return `$${columnLetter(this.right)}$${String(this.bottom)}`;
  }

  get range(): string {
    return `${this.tl}:${this.br}`;
  }

  get $range(): string {
    return `${this.$t$l}:${this.$b$r}`;
  }

  get shortRange(): string {
    return this.count > 1 ? this.range : this.tl;
  }

  get $shortRange(): string {
    return this.count > 1 ? this.$range : this.$t$l;
  }

  get count(): number {
    return (1 + this.bottom - this.top) * (1 + this.right - this.left);
  }

  toString(): string {
    return this.range;
  }
}
