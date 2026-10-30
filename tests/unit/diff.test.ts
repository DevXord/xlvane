import { describe, expect, it } from 'vitest';
import { diff } from '../../packages/xlvane/src/diff';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';

function before(): Uint8Array {
  const wb = createWorkbook();
  const s = wb.addSheet('Data');
  s.set('A1', 'name').set('B1', 'total').set('A2', 'Alice').set('B2', 10).set('A3', 'Bob').set('B3', { formula: 'B2*2', result: 20 });
  s.set('C1', new Date(Date.UTC(2024, 0, 1)), { font: { bold: true } });
  wb.addSheet('Old').set('A1', 1);
  wb.addSheet('Same').set('A1', 'x');
  return writeWorkbook(wb);
}

function after(): Uint8Array {
  const wb = createWorkbook();
  const s = wb.addSheet('Data');
  s.set('A1', 'name').set('B1', 'total').set('A2', 'Alice').set('B2', 12).set('B3', { formula: 'B2*3', result: 36 });
  s.set('C1', new Date(Date.UTC(2024, 0, 1)), { font: { italic: true } });
  s.set('D4', true);
  wb.addSheet('Same').set('A1', 'x');
  wb.addSheet('New', { state: 'hidden' }).set('A1', 'n');
  return writeWorkbook(wb);
}

describe('diff', () => {
  it('reports added and removed sheets and changed cells', () => {
    const d = diff(before(), after());
    expect(d.equal).toBe(false);
    expect(d.sheetsAdded).toEqual(['New']);
    expect(d.sheetsRemoved).toEqual(['Old']);
    expect(d.sheets.map((s) => s.name)).toEqual(['Data']);
    const data = d.sheets[0]!;
    expect(data.cells).toEqual([
      { ref: 'B2', kind: 'changed', before: { type: 'number', value: 10, formula: null }, after: { type: 'number', value: 12, formula: null } },
      { ref: 'A3', kind: 'removed', before: { type: 'string', value: 'Bob', formula: null }, after: null },
      { ref: 'B3', kind: 'changed', before: { type: 'number', value: 20, formula: 'B2*2' }, after: { type: 'number', value: 36, formula: 'B2*3' } },
      { ref: 'D4', kind: 'added', before: null, after: { type: 'boolean', value: true, formula: null } },
    ]);
    expect(d.summary).toEqual({ sheetsAdded: 1, sheetsRemoved: 1, sheetsChanged: 1, cellsAdded: 1, cellsRemoved: 1, cellsChanged: 2 });
    expect(JSON.parse(JSON.stringify(d))).toEqual(d);
  });

  it('compares styles only when asked', () => {
    expect(diff(before(), after()).sheets[0]!.cells.some((c) => c.ref === 'C1')).toBe(false);
    const c1 = diff(before(), after(), { styles: true }).sheets[0]!.cells.find((c) => c.ref === 'C1');
    expect(c1?.kind).toBe('changed');
    expect(c1?.before?.style?.font).toMatchObject({ bold: true });
    expect(c1?.after?.style?.font).toMatchObject({ italic: true });
  });

  it('reports sheet visibility changes and equal workbooks', () => {
    const wb = readWorkbook(before());
    expect(diff(wb, before()).equal).toBe(true);
    const a = createWorkbook();
    a.addSheet('S');
    a.addSheet('T');
    const b = createWorkbook();
    b.addSheet('S');
    b.addSheet('T', { state: 'hidden' });
    const d = diff(writeWorkbook(a), writeWorkbook(b));
    expect(d.sheets).toEqual([{ name: 'T', state: { before: 'visible', after: 'hidden' }, cells: [], truncated: false }]);
  });

  it('limits the number of listed cells', () => {
    const a = createWorkbook();
    const s = a.addSheet('S');
    for (let r = 1; r <= 50; r++) s.set(`A${String(r)}`, r);
    const b = createWorkbook();
    b.addSheet('S');
    const d = diff(writeWorkbook(a), writeWorkbook(b), { maxCells: 10 });
    expect(d.sheets[0]?.cells).toHaveLength(10);
    expect(d.sheets[0]?.truncated).toBe(true);
    expect(d.summary.cellsRemoved).toBe(50);
    expectXlvError(() => diff(writeWorkbook(a), writeWorkbook(b), { maxCells: -1 }), 'XLV_E_INVALID_OPTION');
  });
});
