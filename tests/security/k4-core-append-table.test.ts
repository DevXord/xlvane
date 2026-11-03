// appendRows after the K3 review (B8): array rows were placed from the sheet's first used column,
// while the table that ends on the last used row grew to include them. With a title in A1 and a
// table at B3, the new row landed one column left of the table that now included it. Array rows
// that extend a table now start at the table's first column; when tables with different first
// columns end on that row, the call is refused with XLV_E_INVALID_ARGUMENT.
import { describe, expect, it } from 'vitest';
import { appendRows } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';

function titled(): Uint8Array {
  const wb = createWorkbook();
  const t = wb.addSheet('T').set('A1', 'Report title');
  t.addTable({ name: 'Tb', ref: 'B3', columns: [{ name: 'k' }, { name: 'v' }], rows: [['a', 1]] });
  return writeWorkbook(wb);
}

describe('B8: array rows appended to a table start at its first column', () => {
  it('the reproduction: a title in A1 and a table at B3', () => {
    const { bytes } = appendRows(titled(), 'T', [['c', 3]]);
    const sheet = readWorkbook(bytes).getSheet('T');
    expect(sheet.tables[0]?.ref).toBe('B3:C5');
    expect([sheet.cell('A5')?.value, sheet.cell('B5')?.value, sheet.cell('C5')?.value]).toEqual([undefined, 'c', 3]);
  });

  it('object rows are placed under their header columns as before', () => {
    const { bytes } = appendRows(titled(), 'T', [{ k: 'd', v: 4 }], { header: 3 });
    const sheet = readWorkbook(bytes).getSheet('T');
    expect(sheet.tables[0]?.ref).toBe('B3:C5');
    expect([sheet.cell('B5')?.value, sheet.cell('C5')?.value]).toEqual(['d', 4]);
  });

  it('rows below a table that ends above the last used row start at the first used column', () => {
    const wb = createWorkbook();
    const t = wb.addSheet('T').set('A1', 'Report title');
    t.addTable({ name: 'Tb', ref: 'B3', columns: [{ name: 'k' }, { name: 'v' }], rows: [['a', 1]] });
    t.set('A8', 'note');
    const { bytes } = appendRows(writeWorkbook(wb), 'T', [['x', 'y']]);
    const sheet = readWorkbook(bytes).getSheet('T');
    expect(sheet.tables[0]?.ref).toBe('B3:C4');
    expect([sheet.cell('A9')?.value, sheet.cell('B9')?.value]).toEqual(['x', 'y']);
  });

  it('tables side by side that start in different columns are ambiguous', () => {
    const wb = createWorkbook();
    const t = wb.addSheet('T');
    t.addTable({ name: 'Left', ref: 'A1', columns: [{ name: 'a' }, { name: 'b' }], rows: [[1, 2]] });
    t.addTable({ name: 'Right', ref: 'D1', columns: [{ name: 'c' }, { name: 'd' }], rows: [[3, 4]] });
    const error = expectXlvError(() => appendRows(writeWorkbook(wb), 'T', [[5, 6]]), 'XLV_E_INVALID_ARGUMENT');
    expect(error.message).toContain('Left');
    expect(error.message).toContain('Right');
    expect(error.hint).toMatch(/object/i);
    // Object rows say where each value goes.
    const { bytes } = appendRows(writeWorkbook(wb), 'T', [{ a: 5, d: 6 }], { header: 1 });
    const sheet = readWorkbook(bytes).getSheet('T');
    expect([sheet.cell('A3')?.value, sheet.cell('E3')?.value]).toEqual([5, 6]);
  });

  it('a table that starts in the first used column grows as before', () => {
    const wb = createWorkbook();
    wb.addSheet('T').addTable({ name: 'Tb', ref: 'A1', columns: [{ name: 'k' }, { name: 'v' }], rows: [['a', 1]] });
    const { bytes } = appendRows(writeWorkbook(wb), 'T', [['b', 2], ['c', 3]]);
    const sheet = readWorkbook(bytes).getSheet('T');
    expect(sheet.tables[0]?.ref).toBe('A1:B4');
    expect([...sheet.rows()].map((r) => r.map((c) => c.value))).toEqual([['k', 'v'], ['a', 1], ['b', 2], ['c', 3]]);
  });
});
