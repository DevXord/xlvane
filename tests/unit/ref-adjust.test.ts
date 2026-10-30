// adjustFormula / adjustRange (docs/facts/ref-adjust.md, docs/decisions.md D-B01, D-B02).
import { describe, expect, it } from 'vitest';
import { adjustFormula, adjustRange, type ReferenceChange } from '../../packages/xlvane/src/index';
import { XlvError } from '../../packages/xlvane/src/errors';
import { OBSERVED } from '../libreoffice/ref-adjust-facts.cases';

const ins = (at: number, count = 2, extra: Partial<ReferenceChange> = {}): ReferenceChange => ({ sheet: 'Main', axis: 'row', at, count, ...extra });
const del = (at: number, count = 2, extra: Partial<ReferenceChange> = {}): ReferenceChange => ({ sheet: 'Main', axis: 'row', at, count: -count, ...extra });

describe('adjustFormula agrees with LibreOffice (R-01 … R-13)', () => {
  it('every recorded observation, except the documented differences', () => {
    const differences: string[] = [];
    for (const c of OBSERVED) {
      const got = adjustFormula(c.formula, { sheet: c.sheet, formulaSheet: c.formulaSheet, axis: c.axis, at: 5, count: c.op === 'insert' ? 2 : -2, sheets: ['Main', 'Other'] });
      if (got !== c.observed) differences.push(`${c.axis} ${c.op} ${c.sheet}/${c.formulaSheet}: ${c.formula} → ${String(got)} (LibreOffice: ${c.observed})`);
    }
    // R-13: LibreOffice leaves 3-D references unchanged even when they cover the changed sheet
    // (and saves Main:Other!B9 as Main:Other!B9:B9); xlvane refuses to adjust them. LibreOffice
    // also drops quotes that a sheet name does not need; xlvane keeps the text as written.
    expect(differences.filter((d) => !d.includes('Main:Other!') && !d.includes("'Main'!"))).toEqual([]);
    expect(differences.filter((d) => d.includes('Main:Other!')).every((d) => d.includes('→ null'))).toBe(true);
    expect(differences.filter((d) => d.includes("'Main'!"))).toEqual([
      "row insert Main/Other: 'Main'!B5 → 'Main'!B7 (LibreOffice: Main!B7)",
      "row insert Other/Other: 'Main'!B5 → 'Main'!B5 (LibreOffice: Main!B5)",
    ]);
  });
});

describe('adjustFormula', () => {
  it('moves references at or after the change, absolute ones too (R-01, R-02)', () => {
    expect(adjustFormula('B4+B5+$B$5+B$9+$B9', ins(5))).toBe('B4+B7+$B$7+B$11+$B11');
    expect(adjustFormula('C4+E4+$E$4', { sheet: 'Main', axis: 'col', at: 5, count: 1 })).toBe('C4+F4+$F$4');
  });

  it('grows ranges that the insertion falls into, moves ranges that start at it (R-03, R-04)', () => {
    expect(adjustFormula('SUM(B2:B10)', ins(5))).toBe('SUM(B2:B12)');
    expect(adjustFormula('SUM(B5:B8)', ins(5))).toBe('SUM(B7:B10)');
    expect(adjustFormula('SUM(B2:B4)', ins(5))).toBe('SUM(B2:B4)');
    expect(adjustFormula('SUM(B1:B5)', ins(5))).toBe('SUM(B1:B7)');
  });

  it('turns deleted references into #REF! without their sheet and shrinks ranges (R-05 … R-08)', () => {
    expect(adjustFormula('B5+B9', del(5))).toBe('#REF!+B7');
    expect(adjustFormula('SUM(Main!B5:B6)', del(5, 2, { formulaSheet: 'Other' }))).toBe('SUM(#REF!)');
    expect(adjustFormula("'Main'!B5", del(5, 2, { formulaSheet: 'Other' }))).toBe('#REF!');
    expect(adjustFormula('SUM(B4:B6)', del(5))).toBe('SUM(B4:B4)');
    expect(adjustFormula('SUM(B6:B10)', del(5))).toBe('SUM(B5:B8)');
    expect(adjustFormula('SUM(4:6)+SUM(5:6)', del(5))).toBe('SUM(4:4)+SUM(#REF!)');
  });

  it('handles the end of the grid (R-09, R-10)', () => {
    expect(adjustFormula('B1048576', ins(5))).toBe('#REF!');
    expect(adjustFormula('SUM(B1048574:B1048575)', ins(5))).toBe('SUM(B1048576:B1048576)');
    expect(adjustFormula('SUM(B1048575:B1048576)', ins(5))).toBe('SUM(#REF!)');
    expect(adjustFormula('SUM(B10:B1048576)', del(5))).toBe('SUM(B8:B1048576)');
    expect(adjustFormula('SUM(1048576:1048576)', del(5))).toBe('SUM(1048574:1048574)');
    expect(adjustFormula('XFD1', { sheet: 'Main', axis: 'col', at: 1, count: 1 })).toBe('#REF!');
  });

  it('only touches references to the changed sheet (R-12)', () => {
    expect(adjustFormula('B9+Other!B9+main!B9', ins(5))).toBe('B11+Other!B9+main!B11');
    expect(adjustFormula('B9+Main!B9', ins(5, 2, { formulaSheet: 'Other' }))).toBe('B9+Main!B11');
    expect(adjustFormula("'My Sheet'!A5+'It''s'!A5", { sheet: "it's", formulaSheet: 'X', axis: 'row', at: 1, count: 1 })).toBe("'My Sheet'!A5+'It''s'!A6");
    expect(adjustFormula('[1]Main!B9+SUM([1]Main!B2:B10)', ins(5))).toBe('[1]Main!B9+SUM([1]Main!B2:B10)');
  });

  it('adjusts 3-D references only when they do not cover the changed sheet (R-13)', () => {
    expect(adjustFormula('SUM(January:March!B9)', ins(5))).toBeNull();
    expect(adjustFormula('SUM(January:March!B9)', ins(5, 2, { sheet: 'Feb', sheets: ['January', 'Feb', 'March'] }))).toBeNull();
    expect(adjustFormula("SUM(January:March!B9)+SUM('January:March'!B9)", ins(5, 2, { sheet: 'Apr', formulaSheet: 'Apr', sheets: ['January', 'Feb', 'March', 'Apr'] }))).toBe("SUM(January:March!B9)+SUM('January:March'!B9)");
  });

  it('leaves names, strings, structured references, functions and numbers alone', () => {
    expect(adjustFormula('SUM(Total)+LOG10(B9)+"B9"+Table1[[#This Row],[B9]]+1E5', ins(5))).toBe('SUM(Total)+LOG10(B11)+"B9"+Table1[[#This Row],[B9]]+1E5');
  });

  it('returns null for formulas it cannot classify', () => {
    expect(adjustFormula('A1#', ins(1))).toBeNull();
    expect(adjustFormula('@A1', ins(1))).toBeNull();
    expect(adjustFormula('"open', ins(1))).toBeNull();
    expect(adjustFormula('Main!A1:Main!B2', ins(1))).toBeNull();
    expect(adjustFormula('SUM(B9:B2)', ins(5))).toBeNull();
  });

  it('rejects an invalid change with XLV_E_INVALID_OPTION', () => {
    const bad: unknown[] = [
      null,
      { sheet: '', axis: 'row', at: 1, count: 1 },
      { sheet: 'S', axis: 'rows', at: 1, count: 1 },
      { sheet: 'S', axis: 'row', at: 0, count: 1 },
      { sheet: 'S', axis: 'row', at: 1.5, count: 1 },
      { sheet: 'S', axis: 'row', at: 1, count: 0 },
      { sheet: 'S', axis: 'row', at: 1048576, count: -2 },
      { sheet: 'S', axis: 'col', at: 16385, count: 1 },
      { sheet: 'S', axis: 'row', at: 1, count: 1, sheets: 'S' },
    ];
    for (const change of bad) {
      expect(() => adjustFormula('A1', change as ReferenceChange)).toThrow(XlvError);
      try {
        adjustFormula('A1', change as ReferenceChange);
      } catch (error) {
        expect((error as XlvError).code).toBe('XLV_E_INVALID_OPTION');
      }
    }
  });
});

describe('adjustRange', () => {
  it('adjusts each range of a list and leaves out deleted ones', () => {
    expect(adjustRange('A1:B5 D7', ins(3, 1))).toBe('A1:B6 D8');
    expect(adjustRange('A5:A6', del(5))).toBe('');
    expect(adjustRange('A4:C8 E5 F9', del(5))).toBe('A4:C6 F7');
    expect(adjustRange('A1:A1048576', del(5))).toBe('A1:A1048576');
    expect(adjustRange('B2:D2', { sheet: 'Main', axis: 'col', at: 3, count: -1 })).toBe('B2:C2');
  });

  it('returns null for text that is not a list of A1 references', () => {
    expect(adjustRange('A1,B2', ins(1))).toBeNull();
    expect(adjustRange('Sheet1!A1', ins(1))).toBeNull();
    expect(adjustRange('1:3', ins(1))).toBeNull();
  });
});
