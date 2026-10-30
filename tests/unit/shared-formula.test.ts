import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { compileFormula, shiftFormula } from '../../packages/xlvane/src/sml/formula-shift';
import { buildXlsx, row } from '../helpers/xlsx-builder';

describe('shiftFormula', () => {
  it('shifts relative cell references and keeps absolute parts', () => {
    expect(shiftFormula('B1*2', 1, 0)).toBe('B2*2');
    expect(shiftFormula('A1+$A1+A$1+$A$1', 2, 3)).toBe('D3+$A3+D$1+$A$1');
    expect(shiftFormula('SUM(A1:B2)', 1, 1)).toBe('SUM(B2:C3)');
    expect(shiftFormula('a1+b2', 1, 0)).toBe('A2+B3');
  });

  it('shifts whole-column and whole-row ranges', () => {
    expect(shiftFormula('SUM(A:C)', 5, 1)).toBe('SUM(B:D)');
    expect(shiftFormula('SUM($A:C)', 0, 1)).toBe('SUM($A:D)');
    expect(shiftFormula('SUM(1:3)', 2, 9)).toBe('SUM(3:5)');
    expect(shiftFormula('SUM($1:3)', 2, 0)).toBe('SUM($1:5)');
  });

  it('shifts sheet-qualified references but not sheet names', () => {
    expect(shiftFormula("'My Sheet'!A1+Sheet2!B2:C3", 1, 0)).toBe("'My Sheet'!A2+Sheet2!B3:C4");
    expect(shiftFormula("'It''s'!A1", 0, 1)).toBe("'It''s'!B1");
    expect(shiftFormula('Sheet1:Sheet3!A1', 1, 0)).toBe('Sheet1:Sheet3!A2');
    expect(shiftFormula('[1]Data!A1', 1, 0)).toBe('[1]Data!A2');
  });

  it('never shifts strings, names, functions, numbers, errors or structured references', () => {
    expect(shiftFormula('"A1"&A1', 1, 0)).toBe('"A1"&A2');
    expect(shiftFormula('"say ""A1"""&B1', 1, 0)).toBe('"say ""A1"""&B2');
    expect(shiftFormula('TAX_RATE*A1', 1, 0)).toBe('TAX_RATE*A2');
    expect(shiftFormula('LOG10(A1)', 1, 0)).toBe('LOG10(A2)');
    expect(shiftFormula('_xlfn.XLOOKUP(A1,B:B,C:C)', 1, 0)).toBe('_xlfn.XLOOKUP(A2,B:B,C:C)');
    expect(shiftFormula('1E10+E10+1.5E-3+.5', 1, 0)).toBe('1E10+E11+1.5E-3+.5');
    expect(shiftFormula('IF(A1=0,#N/A,#DIV/0!)', 1, 0)).toBe('IF(A2=0,#N/A,#DIV/0!)');
    expect(shiftFormula('SUM(Table1[[#Headers],[A1]])+A1', 1, 0)).toBe('SUM(Table1[[#Headers],[A1]])+A2');
    expect(shiftFormula('TRUE+FALSE+A1%', 1, 0)).toBe('TRUE+FALSE+A2%');
    expect(shiftFormula('{1,2;3,4}*A1', 1, 0)).toBe('{1,2;3,4}*A2');
    expect(shiftFormula('A1#+@B1', 1, 0)).toBe('A2#+@B2');
    expect(shiftFormula('A1 B1', 1, 0)).toBe('A2 B2');
  });

  it('treats LOG10 without a parenthesis as a cell reference, and any name before "(" as a function', () => {
    expect(shiftFormula('LOG10+1', 1, 0)).toBe('LOG11+1');
    expect(shiftFormula('A1(2)', 1, 0)).toBe('A1(2)');
  });

  it('returns null when a reference would leave the grid', () => {
    expect(shiftFormula('A1', -1, 0)).toBeNull();
    expect(shiftFormula('A1', 0, -1)).toBeNull();
    expect(shiftFormula('XFD1', 0, 1)).toBeNull();
    expect(shiftFormula('A1048576', 1, 0)).toBeNull();
    expect(shiftFormula('$A$1', -1, -1)).toBe('$A$1');
  });

  it('returns null for text it cannot classify safely', () => {
    for (const bad of [
      '"unterminated',
      "'Sheet!A1",
      'Table1[Col',
      'A1]',
      'R1C1+1',
      'RC[-1]',
      'Jan:Feb!A1',
      'A1!B2',
      'A1 (2)',
      'B2[x]',
      '1A',
      '1E+',
      '$A$B',
      '#BOGUS!',
      '1:A',
      '$1',
      '~',
    ]) {
      expect(shiftFormula(bad, 1, 0), bad).toBeNull();
    }
  });

  it('accepts an empty shift unchanged for upper-case formulas', () => {
    const f = "SUM('Q1 data'!$B$2:B10)/COUNT(A:A)*IF(C1>0,\"x\",#REF!)";
    expect(shiftFormula(f, 0, 0)).toBe(f);
  });

  it('compiles once and reports the shift bounds', () => {
    const compiled = compileFormula('$A2+$B$1+C:C');
    expect(compiled).not.toBeNull();
    expect(compiled?.canShift(-1, 0)).toBe(true);
    expect(compiled?.canShift(-2, 0)).toBe(false);
    expect(compiled?.canShift(0, -2)).toBe(true);
    expect(compiled?.canShift(0, -3)).toBe(false);
    expect(compiled?.render(1, 1)).toBe('$A3+$B$1+D:D');
  });
});

describe('reading shared formulas', () => {
  const read = (data: string) => readWorkbook(buildXlsx({ sheets: [{ name: 'F', data }] }));

  it('expands dependent cells from the master formula', () => {
    const wb = read(
      row(1, '<c r="A1"><f t="shared" ref="A1:B3" si="0">C1*2+$D$1</f><v>2</v></c>', '<c r="B1"><f t="shared" si="0"/><v>3</v></c>') +
        row(2, '<c r="A2"><f t="shared" si="0"/><v>4</v></c>') +
        row(3, '<c r="B3"><f t="shared" si="0">IGNORED</f><v>5</v></c>'),
    );
    const s = wb.getSheet('F');
    expect(s.cell('A1')?.formula).toEqual({ kind: 'shared', text: 'C1*2+$D$1', ref: 'A1:B3', sharedIndex: 0 });
    expect(s.cell('B1')?.formula).toEqual({ kind: 'shared', text: 'D1*2+$D$1', ref: null, sharedIndex: 0 });
    expect(s.cell('A2')?.formula?.text).toBe('C2*2+$D$1');
    expect(s.cell('B3')?.formula?.text).toBe('D3*2+$D$1');
    expect(s.cell('A2')?.value).toBe(4);
    expect(wb.warnings.map((w) => w.code)).not.toContain('XLV_W_SHARED_FORMULA_NOT_EXPANDED');
  });

  it('expands a dependent that comes before its master', () => {
    const wb = read(row(1, '<c r="A1"><f t="shared" si="3"/><v>1</v></c>') + row(2, '<c r="A2"><f t="shared" ref="A1:A2" si="3">B2</f><v>1</v></c>'));
    expect(wb.getSheet('F').cell('A1')?.formula?.text).toBe('B1');
  });

  it('keeps text null and warns when the master is missing, the cell is outside its range or the shift fails', () => {
    const wb = read(
      row(1, '<c r="A1"><f t="shared" ref="A1:A2" si="0">A1048576</f><v>1</v></c>', '<c r="C1"><f t="shared" si="9"/><v>1</v></c>') +
        row(2, '<c r="A2"><f t="shared" si="0"/><v>1</v></c>') +
        row(5, '<c r="A5"><f t="shared" si="0"/><v>1</v></c>'),
    );
    const s = wb.getSheet('F');
    expect(s.cell('C1')?.formula?.text).toBeNull();
    expect(s.cell('A2')?.formula?.text).toBeNull();
    expect(s.cell('A5')?.formula?.text).toBeNull();
    const warning = wb.warnings.find((w) => w.code === 'XLV_W_SHARED_FORMULA_NOT_EXPANDED');
    expect(warning?.count).toBe(3);
    expect(warning?.cells).toEqual(['C1', 'A2', 'A5']);
  });

  it('does not let a large master formula multiply memory: dependents share the compiled master', () => {
    const master = `${'A1+'.repeat(2000)}A1`;
    let rows = row(1, `<c r="A1"><f t="shared" ref="A1:A2000" si="0">${master}</f><v>1</v></c>`);
    for (let r = 2; r <= 2000; r++) rows += row(r, `<c r="A${String(r)}"><f t="shared" si="0"/><v>1</v></c>`);
    const s = read(rows).getSheet('F');
    expect(s.cell('A2000')?.formula?.text).toBe(`${'A2000+'.repeat(2000)}A2000`);
  });
});
