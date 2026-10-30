// Files written by LibreOffice's own XLSX export (tests/corpus/libreoffice, made by
// scripts/make-lo-corpus.mjs from the content in REQUESTED.md; docs/decisions.md D-902). Besides
// the inspect() snapshots in corpus.test.ts, the values are checked against what was entered.
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';

const read = (name: string) => readWorkbook(readFileSync(new URL(`./libreoffice/${name}.xlsx`, import.meta.url)));
const d = (y: number, m: number, day: number, h = 0, min = 0) => new Date(Date.UTC(y, m - 1, day, h, min));

describe('corpus: LibreOffice values', () => {
  it('basic-types', () => {
    const wb = read('basic-types');
    expect(wb.warnings).toEqual([]);
    const s = wb.getSheet('Data');
    const v = (ref: string) => s.cell(ref)?.value;
    expect([v('A2'), v('B2'), v('C2'), v('D2'), v('E2'), v('F2')]).toEqual(['Alice', 1234.5, d(2024, 1, 31), d(2024, 1, 31, 13, 45), true, 2469]);
    expect([v('A3'), v('B3'), v('C3'), v('D3'), v('E3'), v('F3')]).toEqual(['  spaces  ', -0.001, d(1900, 3, 1), d(2024, 2, 29), false, '#DIV/0!']);
    expect([v('A4'), v('B4'), v('C4'), v('F4')]).toEqual(['first\nsecond', 1e20, d(9999, 12, 31), 'Alice!']);
    expect([v('A5'), v('B5')]).toEqual(['zażółć 😀', 0.1]);
    expect(s.cell('F3')?.type).toBe('error');
    expect(s.cell('F2')?.formula?.text).toBe('B2*2');
    expect(s.cell('B5')?.numFmt).toBe('0.00%');
    // LibreOffice writes TRUE and FALSE as the formulas TRUE() and FALSE() with a cached boolean.
    expect(s.cell('E2')?.formula?.text).toBe('TRUE()');
    expect(wb.getSheet('Second').cell('A1')).toMatchObject({ value: 1234.5, formula: { text: 'Data!B2' } });
  });

  it('formulas, merged cells and a hidden sheet', () => {
    const wb = read('formulas-merged-hidden');
    expect(wb.sheets.map((s) => [s.name, s.state])).toEqual([
      ['Formulas', 'visible'],
      ['Report', 'visible'],
      ['Hidden', 'hidden'],
    ]);
    const f = wb.getSheet('Formulas');
    expect(f.cell('B10')).toMatchObject({ value: 20, formula: { text: 'A10*2' } });
    expect(f.cell('C1')?.value).toBe(110);
    expect(wb.getSheet('Report').merges).toEqual(['A1:D1']);
  });

  it('1904 date system', () => {
    const wb = read('date1904');
    expect(wb.dateSystem).toBe('1904');
    expect(wb.getSheet('Dates').cell('A2')?.value).toEqual(d(2024, 1, 31));
  });
});
