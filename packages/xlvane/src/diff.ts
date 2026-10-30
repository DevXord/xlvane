// Comparing two workbooks cell by cell (docs/decisions.md D-308).
import { XlvError } from './errors';
import type { ReadLimits } from './limits';
import type { Cell, CellType, SheetState, Workbook } from './model';
import { stableKey, type CellStyle } from './style';
import { booleanOption, openInput, readOptions } from './util/options';

/**
 * Options for {@link diff}.
 *
 * @example
 * ```ts
 * import type { DiffOptions } from 'xlvane';
 *
 * const options: DiffOptions = { styles: true, maxCells: 1000 };
 * ```
 */
export interface DiffOptions {
  /** Also compare cell styles. Default `false`: values and formulas only. */
  readonly styles?: boolean;
  /** Most changed cells listed per sheet; the summary still counts all. Default 10 000. */
  readonly maxCells?: number;
  /** Read limits for byte input; see `ReadLimits`. */
  readonly limits?: Partial<ReadLimits>;
}

/**
 * One side of a changed cell, as JSON-safe data. Dates are ISO 8601 strings.
 *
 * @example
 * ```ts
 * import type { DiffCell } from 'xlvane';
 *
 * const cell: DiffCell = { type: 'number', value: 12, formula: 'B2*3' };
 * ```
 */
export interface DiffCell {
  readonly type: CellType;
  readonly value: string | number | boolean | null;
  /** Formula text without `=`, or `null`. */
  readonly formula: string | null;
  /** Present only with `styles: true`. */
  readonly style?: CellStyle;
}

/**
 * A cell that differs between the two workbooks.
 *
 * @example
 * ```ts
 * import { diff, type CellChange } from 'xlvane';
 *
 * declare const before: Uint8Array;
 * declare const after: Uint8Array;
 * const changes: CellChange[] = diff(before, after).sheets.flatMap((sheet) => sheet.cells);
 * ```
 */
export interface CellChange {
  /** A1 reference. */
  readonly ref: string;
  readonly kind: 'added' | 'removed' | 'changed';
  /** The cell in the first workbook, or `null` if it has none there. */
  readonly before: DiffCell | null;
  /** The cell in the second workbook, or `null` if it has none there. */
  readonly after: DiffCell | null;
}

/**
 * Differences in one sheet that exists in both workbooks.
 *
 * @example
 * ```ts
 * import { diff, type SheetDiff } from 'xlvane';
 *
 * declare const before: Uint8Array;
 * declare const after: Uint8Array;
 * const first: SheetDiff | undefined = diff(before, after).sheets[0];
 * console.log(first?.name, first?.cells.length);
 * ```
 */
export interface SheetDiff {
  readonly name: string;
  /** Present when the visibility changed. */
  readonly state?: { readonly before: SheetState; readonly after: SheetState };
  /** Changed cells in row order, at most `maxCells`. */
  readonly cells: CellChange[];
  /** `true` when more cells changed than are listed. */
  readonly truncated: boolean;
}

/**
 * The result of {@link diff}. Plain JSON data.
 *
 * @example
 * ```ts
 * import { diff, type WorkbookDiff } from 'xlvane';
 *
 * declare const before: Uint8Array;
 * declare const after: Uint8Array;
 * const result: WorkbookDiff = diff(before, after);
 * console.log(result.equal, result.summary);
 * ```
 */
export interface WorkbookDiff {
  /** `true` when nothing differs. */
  readonly equal: boolean;
  /** Sheets only in the second workbook, by name. */
  readonly sheetsAdded: string[];
  /** Sheets only in the first workbook, by name. */
  readonly sheetsRemoved: string[];
  /** Sheets in both workbooks that differ, in the order of the second workbook. */
  readonly sheets: SheetDiff[];
  readonly summary: {
    readonly sheetsAdded: number;
    readonly sheetsRemoved: number;
    readonly sheetsChanged: number;
    readonly cellsAdded: number;
    readonly cellsRemoved: number;
    readonly cellsChanged: number;
  };
}

const DEFAULT_MAX_CELLS = 10_000;

function toDiffCell(cell: Cell, styles: boolean): DiffCell {
  const value = cell.value instanceof Date ? cell.value.toISOString() : cell.value;
  const out: { type: CellType; value: string | number | boolean | null; formula: string | null; style?: CellStyle } = {
    type: cell.type,
    value,
    formula: cell.formula?.text ?? null,
  };
  if (styles) out.style = cell.style;
  return out;
}

function sameCell(a: Cell, b: Cell, styles: boolean): boolean {
  if (a.type !== b.type) return false;
  const av = a.value instanceof Date ? a.value.getTime() : a.value;
  const bv = b.value instanceof Date ? b.value.getTime() : b.value;
  if (!Object.is(av, bv)) return false;
  if ((a.formula?.text ?? null) !== (b.formula?.text ?? null)) return false;
  return !styles || stableKey(a.style) === stableKey(b.style);
}

/**
 * Compares two workbooks: added and removed sheets (by name), visibility changes, and cells
 * whose value, type or formula differ (and style, with `styles: true`). The result is plain
 * JSON, suitable for review by people or checks by programs. Formulas are compared as text and
 * values as stored; xlvane does not recalculate.
 *
 * @example
 * ```ts
 * import { diff } from 'xlvane';
 *
 * declare const before: Uint8Array;
 * declare const after: Uint8Array;
 * const result = diff(before, after);
 * for (const sheet of result.sheets) {
 *   for (const change of sheet.cells) console.log(sheet.name, change.ref, change.kind, change.before?.value, '->', change.after?.value);
 * }
 * ```
 */
export function diff(a: Uint8Array | ArrayBuffer | Workbook, b: Uint8Array | ArrayBuffer | Workbook, options?: DiffOptions): WorkbookDiff {
  const o = readOptions(options, ['styles', 'maxCells', 'limits'], 'diff options');
  const styles = booleanOption('styles', o.get('styles'), false);
  const maxCells = o.get('maxCells') ?? DEFAULT_MAX_CELLS;
  if (typeof maxCells !== 'number' || !(maxCells >= 0) || (!Number.isInteger(maxCells) && maxCells !== Infinity)) {
    throw new XlvError('XLV_E_INVALID_OPTION', 'The option "maxCells" must be a whole number of 0 or more.', {
      hint: 'For example { maxCells: 1000 }, or Infinity to list every change.',
      details: { option: 'maxCells' },
    });
  }
  const first = openInput(a, o.get('limits'));
  const second = openInput(b, o.get('limits'));
  const firstNames = new Map(first.sheets.map((s) => [s.name, s]));
  const secondNames = new Set(second.sheets.map((s) => s.name));
  const summary = { sheetsAdded: 0, sheetsRemoved: 0, sheetsChanged: 0, cellsAdded: 0, cellsRemoved: 0, cellsChanged: 0 };
  const sheetsAdded = second.sheets.filter((s) => !firstNames.has(s.name)).map((s) => s.name);
  const sheetsRemoved = first.sheets.filter((s) => !secondNames.has(s.name)).map((s) => s.name);
  summary.sheetsAdded = sheetsAdded.length;
  summary.sheetsRemoved = sheetsRemoved.length;

  const sheets: SheetDiff[] = [];
  for (const after of second.sheets) {
    const before = firstNames.get(after.name);
    if (!before) continue;
    // cells() yields row order, so each list is sorted; keeping the first maxCells of each and
    // merging them gives the first maxCells overall without holding every change in memory.
    type Entry = { row: number; col: number; change: CellChange };
    const fromFirst: Entry[] = [];
    const fromSecond: Entry[] = [];
    let count = 0;
    const record = (list: Entry[], row: number, col: number, change: CellChange): void => {
      count++;
      if (change.kind === 'added') summary.cellsAdded++;
      else if (change.kind === 'removed') summary.cellsRemoved++;
      else summary.cellsChanged++;
      if (list.length < maxCells) list.push({ row, col, change });
    };
    for (const cell of before.cells()) {
      const other = after.cell(cell.ref);
      if (other === undefined) record(fromFirst, cell.row, cell.col, { ref: cell.ref, kind: 'removed', before: toDiffCell(cell, styles), after: null });
      else if (!sameCell(cell, other, styles)) {
        record(fromFirst, cell.row, cell.col, { ref: cell.ref, kind: 'changed', before: toDiffCell(cell, styles), after: toDiffCell(other, styles) });
      }
    }
    for (const cell of after.cells()) {
      if (before.cell(cell.ref) === undefined) record(fromSecond, cell.row, cell.col, { ref: cell.ref, kind: 'added', before: null, after: toDiffCell(cell, styles) });
    }
    const stateChanged = before.state !== after.state;
    if (count === 0 && !stateChanged) continue;
    summary.sheetsChanged++;
    const listed = [...fromFirst, ...fromSecond]
      .sort((x, y) => x.row - y.row || x.col - y.col)
      .slice(0, maxCells)
      .map((e) => e.change);
    sheets.push({
      name: after.name,
      ...(stateChanged ? { state: { before: before.state, after: after.state } } : {}),
      cells: listed,
      truncated: listed.length < count,
    });
  }
  return {
    equal: sheetsAdded.length === 0 && sheetsRemoved.length === 0 && sheets.length === 0,
    sheetsAdded,
    sheetsRemoved,
    sheets,
    summary,
  };
}
