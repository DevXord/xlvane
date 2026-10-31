// Property test for surgical editing: after random edits, every edited cell holds its new value
// and every other cell is unchanged.
import fc from 'fast-check';
import { describe, expect, it } from 'vitest';
import { editCells, type CellEdit } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { formatCellRef } from '../../packages/xlvane/src/sml/cell-ref';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { NUM_RUNS } from './runs';

const RUNS = { numRuns: Math.max(1, Math.floor(NUM_RUNS / 20)) };

function original(): Uint8Array {
  const wb = createWorkbook();
  const s = wb.addSheet('S');
  for (let r = 2; r <= 12; r += 2) for (let c = 1; c <= 5; c += 2) s.set(formatCellRef(r, c), r * 100 + c, c === 3 ? { font: { bold: true } } : undefined);
  wb.addSheet('T').set('A1', 'other');
  return writeWorkbook(wb);
}

const value = fc.oneof(fc.integer(), fc.string({ maxLength: 8 }), fc.boolean(), fc.constant(null));
const edit = fc.record({ row: fc.integer({ min: 1, max: 14 }), col: fc.integer({ min: 1, max: 6 }), value });

describe('editCells properties', () => {
  const bytes = original();
  const before = readWorkbook(bytes).getSheet('S');

  it('changes exactly the edited cells', () => {
    fc.assert(
      fc.property(fc.array(edit, { maxLength: 15 }), (list) => {
        const edits: CellEdit[] = list.map((e) => ({ sheet: 'S', cell: formatCellRef(e.row, e.col), value: e.value }));
        const expected = new Map<string, unknown>();
        for (const cell of before.cells()) expected.set(cell.ref, cell.value);
        for (const e of edits) {
          if (e.value === null) expected.delete(e.cell);
          else expected.set(e.cell, e.value);
        }
        const after = readWorkbook(editCells(bytes, edits).bytes);
        const got = new Map([...after.getSheet('S').cells()].map((cell) => [cell.ref, cell.value]));
        expect(got).toEqual(expected);
        expect(after.getSheet('T').cell('A1')?.value).toBe('other');
        expect(after.warnings).toEqual([]);
      }),
      RUNS,
    );
  });
});
