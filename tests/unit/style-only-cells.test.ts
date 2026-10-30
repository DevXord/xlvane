// Cells that hold only a style (no value, no formula) are read and kept apart from the cells with
// data (docs/decisions.md D-E05).
import { describe, expect, it } from 'vitest';
import { diff } from '../../packages/xlvane/src/diff';
import { editCells } from '../../packages/xlvane/src/edit';
import { inspect } from '../../packages/xlvane/src/inspect';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { openWorkbookStream } from '../../packages/xlvane/src/stream-read';
import { readTable } from '../../packages/xlvane/src/table';
import { toCSV } from '../../packages/xlvane/src/csv';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { XlvError } from '../../packages/xlvane/src/errors';

function book(): Uint8Array {
  const wb = createWorkbook();
  const s = wb.addSheet('S');
  s.set('A1', 'Name').set('B1', 'Score').set('A2', 'Ann').set('B2', 7);
  s.style('C1:C2', { fill: { color: 'FFFF00' } });
  s.style('A5', { font: { bold: true } });
  return writeWorkbook(wb);
}

describe('style-only cells (D-E05)', () => {
  it('are read with their style through styledCells(), not through cells()', () => {
    const sheet = readWorkbook(book()).getSheet('S');
    expect([...sheet.cells()].map((c) => c.ref)).toEqual(['A1', 'B1', 'A2', 'B2']);
    expect(sheet.cellCount).toBe(4);
    const styled = [...sheet.styledCells()];
    expect(styled.map((c) => c.ref)).toEqual(['C1', 'C2', 'A5']);
    expect(styled[2]).toMatchObject({ type: 'empty', value: null, formula: null, style: { font: { bold: true } } });
    expect(styled[0]?.style.fill).toBeDefined();
    expect(sheet.cell('A5')).toBeUndefined();
  });

  it('do not change tables, CSV, inspect or diff', () => {
    const bytes = book();
    const wb = readWorkbook(bytes);
    expect(readTable(wb, { sheet: 'S' }).rows).toEqual([{ Name: 'Ann', Score: 7 }]);
    expect(toCSV(wb, { sheet: 'S' })).toBe('Name,Score\r\nAnn,7\r\n');
    const s = inspect(wb).sheets[0]!;
    expect(s.cellCount).toBe(4);
    expect(s.rowCount).toBe(2);
    expect(diff(wb, readWorkbook(bytes)).equal).toBe(true);
  });

  it('are given to the streaming reader per row', async () => {
    const stream = await openWorkbookStream(book());
    const rows = [];
    for await (const r of stream.rows('S')) rows.push({ n: r.number, cells: r.cells.map((c) => c.ref), styled: r.styledCells.map((c) => c.ref) });
    expect(rows).toEqual([
      { n: 1, cells: ['A1', 'B1'], styled: ['C1'] },
      { n: 2, cells: ['A2', 'B2'], styled: ['C2'] },
      { n: 5, cells: [], styled: ['A5'] },
    ]);
  });

  it('keep their style when editCells gives them a value', () => {
    const edited = editCells(book(), [{ sheet: 'S', cell: 'A5', value: 'x' }]);
    const cell = readWorkbook(edited.bytes).getSheet('S').cell('A5');
    expect(cell?.value).toBe('x');
    expect(cell?.style.font?.bold).toBe(true);
  });

  it('count against maxCells', () => {
    let code = '';
    try {
      readWorkbook(book(), { limits: { maxCells: 5 } });
    } catch (error) {
      code = error instanceof XlvError ? error.code : 'CRASH';
    }
    expect(code).toBe('XLV_E_LIMIT_EXCEEDED');
    expect(() => readWorkbook(book(), { limits: { maxCells: 7 } })).not.toThrow();
  });
});
