// Writing array and shared formulas: ECMA-376 Part 1 §18.3.1.40, §18.17.6.3, §18.18.6;
// [MS-OI29500] §2.1.620, §2.1.872; docs/facts/ecma-376.md E-26 … E-28, docs/facts/formulas.md
// F2-13 … F2-16; docs/decisions.md D-E02.
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { editCells } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError, expectXlvErrorAsync } from '../helpers/expect-error';

const sheetXml = (bytes: Uint8Array): string => strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']!);
const collect = async (stream: ReadableStream<Uint8Array>): Promise<Uint8Array> => new Uint8Array(await new Response(stream).arrayBuffer());

function arrayAndShared() {
  const wb = createWorkbook();
  const s = wb.addSheet('S');
  s.set('A1', 1).set('A2', 2).set('A3', 3);
  s.set('B1', { formula: 'A1:A3*2', result: 2, array: 'B1:B3' }).set('B2', 4).set('B3', 6);
  s.set('C1', { formula: 'A1+1', result: 2, shared: 'C1:C3' });
  s.set('C2', { sharedFormula: 'C1', result: 3 }).set('C3', { sharedFormula: 'C1', result: 4 });
  return wb;
}

describe('array formulas', () => {
  it('writes t="array" and ref on the top-left cell and only cached values in the other cells (P1 §18.17.6.3)', () => {
    const xml = sheetXml(writeWorkbook(arrayAndShared()));
    expect(xml).toContain('<c r="B1"><f t="array" ref="B1:B3">A1:A3*2</f><v>2</v></c>');
    expect(xml).toContain('<c r="B2"><v>4</v></c>');
  });

  it('reads back as an array formula with its range', () => {
    const s = readWorkbook(writeWorkbook(arrayAndShared())).getSheet('S');
    expect(s.cell('B1')?.formula).toEqual({ kind: 'array', text: 'A1:A3*2', ref: 'B1:B3', sharedIndex: null });
    expect(s.cell('B2')?.formula).toBeNull();
    expect(s.cell('B2')?.value).toBe(4);
  });

  it('accepts a one-cell array formula', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', { formula: 'SUM(B1:B3*C1:C3)', array: 'A1' });
    expect(sheetXml(writeWorkbook(wb))).toContain('<f t="array" ref="A1">SUM(B1:B3*C1:C3)</f>');
  });

  it('refuses a range that does not start at the cell, overlapping arrays and formulas inside an array', () => {
    expectXlvError(() => createWorkbook().addSheet('S').set('B2', { formula: 'A1', array: 'A1:B3' }), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => createWorkbook().addSheet('S').set('B2', { formula: 'A1', array: 'nonsense' }), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => createWorkbook().addSheet('S').set('A1', { formula: 'A1', array: 'A1:A2', shared: 'A1:A2' }), 'XLV_E_INVALID_VALUE');
    const overlap = createWorkbook();
    overlap.addSheet('S').set('A1', { formula: 'X1:X3', array: 'A1:A3' }).set('A2', { formula: 'Y1:Y2', array: 'A2:A3' });
    expectXlvError(() => writeWorkbook(overlap), 'XLV_E_INVALID_VALUE');
    const inside = createWorkbook();
    inside.addSheet('S').set('A1', { formula: 'X1:X3', array: 'A1:A3' }).set('A3', { formula: '1+1' });
    expectXlvError(() => writeWorkbook(inside), 'XLV_E_INVALID_VALUE');
  });
});

describe('shared formulas', () => {
  it('writes the master with t="shared", ref and si, and dependents with t="shared" and si only (E-28)', () => {
    const xml = sheetXml(writeWorkbook(arrayAndShared()));
    expect(xml).toContain('<c r="C1"><f t="shared" ref="C1:C3" si="0">A1+1</f><v>2</v></c>');
    expect(xml).toContain('<c r="C2"><f t="shared" si="0"/><v>3</v></c>');
    expect(xml).toContain('<c r="C3"><f t="shared" si="0"/><v>4</v></c>');
  });

  it('numbers groups per sheet in the order their masters are written', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('S');
    s.set('B2', { formula: 'A2*2', shared: 'B2:B3' }).set('B3', { sharedFormula: 'B2' });
    s.set('A1', { formula: 'Z1', shared: 'A1:C1' }).set('C1', { sharedFormula: 'A1' });
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml).toContain('<c r="A1"><f t="shared" ref="A1:C1" si="0">Z1</f></c>');
    expect(xml).toContain('<c r="C1"><f t="shared" si="0"/></c>');
    expect(xml).toContain('<c r="B2"><f t="shared" ref="B2:B3" si="1">A2*2</f></c>');
    expect(xml).toContain('<c r="B3"><f t="shared" si="1"/></c>');
  });

  it('reads back: the master with its range and group, dependents with the translated formula', () => {
    const s = readWorkbook(writeWorkbook(arrayAndShared())).getSheet('S');
    expect(s.cell('C1')?.formula).toEqual({ kind: 'shared', text: 'A1+1', ref: 'C1:C3', sharedIndex: 0 });
    expect(s.cell('C3')?.formula).toEqual({ kind: 'shared', text: 'A3+1', ref: null, sharedIndex: 0 });
    expect(s.cell('C3')?.value).toBe(4);
  });

  it('refuses dependents outside the range, a missing master and overlapping masters', () => {
    const outside = createWorkbook();
    outside.addSheet('S').set('A1', { formula: 'X1', shared: 'A1:A2' }).set('A3', { sharedFormula: 'A1' });
    expectXlvError(() => writeWorkbook(outside), 'XLV_E_INVALID_VALUE', { cell: 'A3' });
    const missing = createWorkbook();
    missing.addSheet('S').set('A2', { sharedFormula: 'A1' });
    expectXlvError(() => writeWorkbook(missing), 'XLV_E_INVALID_VALUE', { cell: 'A2' });
    const overlap = createWorkbook();
    overlap.addSheet('S').set('A1', { formula: 'X1', shared: 'A1:B2' }).set('B2', { formula: 'Y1', shared: 'B2:C3' });
    expectXlvError(() => writeWorkbook(overlap), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => createWorkbook().addSheet('S').set('A1', { sharedFormula: 'A1:A2' }), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => createWorkbook().addSheet('S').set('A1', { sharedFormula: 'A1', formula: 'B1' } as never), 'XLV_E_INVALID_VALUE');
  });
});

describe('streaming writer', () => {
  it('writes the same XML as the batch writer', async () => {
    const bytes = await collect(
      writeWorkbookStream({
        sheets: [
          {
            name: 'S',
            rows: [
              [1, { formula: 'A1:A3*2', result: 2, array: 'B1:B3' }, { formula: 'A1+1', result: 2, shared: 'C1:C3' }],
              [2, 4, { sharedFormula: 'C1', result: 3 }],
              [3, 6, { sharedFormula: 'C1', result: 4 }],
            ],
          },
        ],
      }),
    );
    const streamed = sheetXml(bytes);
    const batch = sheetXml(writeWorkbook(arrayAndShared()));
    const data = (xml: string) => xml.slice(xml.indexOf('<sheetData>'), xml.indexOf('</sheetData>'));
    expect(data(streamed)).toBe(data(batch));
  });

  it('refuses a dependent whose master has not been written', async () => {
    await expectXlvErrorAsync(() => collect(writeWorkbookStream({ sheets: [{ name: 'S', rows: [[{ sharedFormula: 'A2' }], [{ formula: '1', shared: 'A2:A3' }]] }] })), 'XLV_E_INVALID_VALUE');
  });
});

describe('editCells', () => {
  it('refuses array and shared formulas in edits, and edits inside an array or of a shared master', () => {
    const bytes = writeWorkbook(arrayAndShared());
    expectXlvError(() => editCells(bytes, [{ sheet: 'S', cell: 'D1', value: { formula: 'A1', array: 'D1:D2' } }]), 'XLV_E_NOT_SUPPORTED');
    expectXlvError(() => editCells(bytes, [{ sheet: 'S', cell: 'B2', value: 5 }]), 'XLV_E_NOT_SUPPORTED', { cell: 'B2' });
    expectXlvError(() => editCells(bytes, [{ sheet: 'S', cell: 'B1', value: 5 }]), 'XLV_E_NOT_SUPPORTED');
    expectXlvError(() => editCells(bytes, [{ sheet: 'S', cell: 'C1', value: 5 }]), 'XLV_E_NOT_SUPPORTED');
    // A dependent may be replaced: its own value overrides the group (E-28).
    const edited = readWorkbook(editCells(bytes, [{ sheet: 'S', cell: 'C3', value: 9 }]).bytes).getSheet('S');
    expect(edited.cell('C3')?.value).toBe(9);
    expect(edited.cell('C2')?.formula?.text).toBe('A2+1');
  });
});
