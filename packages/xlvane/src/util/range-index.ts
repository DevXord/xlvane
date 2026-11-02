// Finds whether a range overlaps any range added before, without comparing it with every one
// of them (D-H01): adding n merged ranges or tables one at a time must not take n² steps.
import type { CellRange } from '../sml/cell-ref';

/** Block heights of the row levels: 4, 16, …, 4^10 = 1 048 576 rows (the whole sheet). */
const ROW_SIZES = [4, 16, 64, 256, 1024, 4096, 16_384, 65_536, 262_144, 1_048_576] as const;
/** Block widths of the column levels: 4, 16, …, 4^7 = 16 384 columns (the whole sheet). */
const COL_SIZES = [4, 16, 64, 256, 1024, 4096, 16_384] as const;

function levelOf(sizes: readonly number[], length: number): number {
  for (let i = 0; i < sizes.length; i++) if (length <= sizes[i]!) return i;
  return sizes.length - 1;
}

function blockOf(index: number, size: number): number {
  return Math.floor((index - 1) / size);
}

function overlap(a: CellRange, b: CellRange): boolean {
  return a.start.row <= b.end.row && b.start.row <= a.end.row && a.start.col <= b.end.col && b.start.col <= a.end.col;
}

/**
 * A grid of blocks at several sizes. A range is stored in the level whose block height is the
 * smallest at least as tall as the range and whose block width is the smallest at least as wide,
 * in every block of that level it touches (at most 2 × 2), so a huge range costs no more than a
 * small one. A query visits, in each level that holds ranges, the blocks it touches that hold
 * ranges (the smaller of the blocks it spans and the blocks the level holds). Ranges of a level
 * are at least about a quarter of its block size in each dimension, so when the stored ranges do
 * not overlap each other (merged ranges, tables), each block holds a bounded number of them.
 */
export class RangeIndex {
  /** Level (row level × 16 + column level) → column block → row block → ranges. */
  private readonly levels = new Map<number, Map<number, Map<number, CellRange[]>>>();

  /** Adds a range. */
  add(range: CellRange): void {
    const a = levelOf(ROW_SIZES, range.end.row - range.start.row + 1);
    const b = levelOf(COL_SIZES, range.end.col - range.start.col + 1);
    const rows = ROW_SIZES[a]!;
    const cols = COL_SIZES[b]!;
    let level = this.levels.get(a * 16 + b);
    if (level === undefined) {
      level = new Map();
      this.levels.set(a * 16 + b, level);
    }
    for (let cb = blockOf(range.start.col, cols); cb <= blockOf(range.end.col, cols); cb++) {
      let column = level.get(cb);
      if (column === undefined) {
        column = new Map();
        level.set(cb, column);
      }
      for (let rb = blockOf(range.start.row, rows); rb <= blockOf(range.end.row, rows); rb++) {
        const list = column.get(rb);
        if (list === undefined) column.set(rb, [range]);
        else list.push(range);
      }
    }
  }

  /** Whether `range` shares a cell with a range added before. */
  overlaps(range: CellRange): boolean {
    const hit = (list: readonly CellRange[] | undefined): boolean => list !== undefined && list.some((r) => overlap(r, range));
    for (const [key, level] of this.levels) {
      const rows = ROW_SIZES[key >> 4]!;
      const cols = COL_SIZES[key & 15]!;
      const rb1 = blockOf(range.start.row, rows);
      const rb2 = blockOf(range.end.row, rows);
      const cb1 = blockOf(range.start.col, cols);
      const cb2 = blockOf(range.end.col, cols);
      const inColumn = (column: Map<number, CellRange[]>): boolean => {
        if (rb2 - rb1 + 1 <= column.size) {
          for (let rb = rb1; rb <= rb2; rb++) if (hit(column.get(rb))) return true;
          return false;
        }
        for (const [rb, list] of column) if (rb >= rb1 && rb <= rb2 && hit(list)) return true;
        return false;
      };
      if (cb2 - cb1 + 1 <= level.size) {
        for (let cb = cb1; cb <= cb2; cb++) {
          const column = level.get(cb);
          if (column !== undefined && inColumn(column)) return true;
        }
      } else {
        for (const [cb, column] of level) if (cb >= cb1 && cb <= cb2 && inColumn(column)) return true;
      }
    }
    return false;
  }
}
