// Ported from exceljs v4.4.0 spec/integration/worksheet.spec.js (MIT, Copyright (c) 2014-2019
// Guyon Roche), changes: Mocha/Chai → Vitest; the 'Values' and 'Merge Cells' tests that repeat
// spec/unit/doc/worksheet.values.spec.js and worksheet.merge.spec.js are not repeated here (see
// worksheet-values.test.ts and worksheet-merge.test.ts); worksheet-name errors are checked with
// `toThrow()` without the exceljs message text; the `findCell`/`findRow` assertions (internal
// lookups outside the supported subset) are removed; insertRow tests expect
// XLV_E_NOT_SUPPORTED; `addRow(values, 'i')` (inherited style) expects XLV_E_NOT_SUPPORTED
// and the `'n'` half is kept; tests that read binary fixtures (.numbers, chartsheet, hidden sheets)
// are skipped because no fixtures were copied.
import { describe, expect, it } from 'vitest';
import { ExcelJS, expectNotSupported, styles } from './helpers';

describe('Worksheet', () => {
  describe('Values', () => {
    it('adds rows with style option', async () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      const dateValue1 = new Date(1970, 1, 1);
      const dateValue2 = new Date(1965, 1, 7);
      ws.addRow([1, 'John Doe', dateValue1]);
      ws.getRow(1).font = styles.fonts.comicSansUdB16;
      ws.addRow([2, 'Jane Doe', dateValue2], 'n');
      expect(ws.getCell('A1').font).toEqual(styles.fonts.comicSansUdB16);
      expect(ws.getCell('A2').font).not.toEqual(styles.fonts.comicSansUdB16);
      // Change for xlvane: the inherited style option 'i' is not supported.
      await expectNotSupported(() => ws.addRow([3, 'Jane Doe', dateValue2], 'i'));
    });

    it('inserts rows by object — not supported: insertRow', async () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      ws.columns = [{ header: 'Id', key: 'id', width: 10 }];
      await expectNotSupported(() => ws.insertRow(2, { id: 1 }));
    });
    it.skip('inserts rows by contiguous array — not supported: insertRow', () => {});
    it.skip('inserts rows by sparse array — not supported: insertRow', () => {});
    it.skip('inserts rows with style option — not supported: insertRow', () => {});
    it.skip('should style of the inserted row with inherited style be mutable — not supported: insertRow', () => {});

    it('returns undefined when row range is less than 1', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      expect(ws.getRows(1, 0)).toBeUndefined();
    });

    describe('when worksheet name is less than or equal 31', () => {
      it('save the original name', () => {
        const wb = new ExcelJS.Workbook();
        let ws = wb.addWorksheet();
        ws.name = 'ThisIsAWorksheetName';
        expect(ws.name).toBe('ThisIsAWorksheetName');
        ws = wb.addWorksheet();
        ws.name = 'ThisIsAWorksheetNameWith31Chars';
        expect(ws.name).toBe('ThisIsAWorksheetNameWith31Chars');
      });
    });

    describe('name is be not empty string', () => {
      it('when empty should thrown an error', () => {
        const wb = new ExcelJS.Workbook();
        expect(() => {
          wb.addWorksheet().name = '';
        }).toThrow();
      });
      it("when isn't string should thrown an error", () => {
        const wb = new ExcelJS.Workbook();
        expect(() => {
          wb.addWorksheet().name = 0;
        }).toThrow();
      });
    });

    describe('when worksheet name is `History`', () => {
      it('thrown an error', () => {
        const wb = new ExcelJS.Workbook();
        expect(() => {
          wb.addWorksheet().name = 'History';
        }).toThrow();
      });
    });

    describe('when worksheet name is longer than 31', () => {
      it('keep first 31 characters', () => {
        const ws = new ExcelJS.Workbook().addWorksheet();
        ws.name = 'ThisIsAWorksheetNameThatIsLongerThan31';
        expect(ws.name).toBe('ThisIsAWorksheetNameThatIsLonge');
      });
    });

    describe('when the worksheet name contains illegal characters', () => {
      it('throws an error (illegal characters)', () => {
        const wb = new ExcelJS.Workbook();
        for (const invalidCharacter of ['*', '?', ':', '/', '\\', '[', ']']) {
          expect(() => {
            wb.addWorksheet().name = invalidCharacter;
          }, invalidCharacter).toThrow();
        }
      });
      it('throws an error (leading or trailing quote)', () => {
        const wb = new ExcelJS.Workbook();
        for (const invalidName of ["'sheetName", "sheetName'"]) {
          expect(() => {
            wb.addWorksheet().name = invalidName;
          }, invalidName).toThrow();
        }
      });
    });

    describe('when worksheet name already exists', () => {
      it('throws an error (case-insensitive)', () => {
        const wb = new ExcelJS.Workbook();
        wb.addWorksheet().name = 'thisisaworksheetnameinuppercase';
        expect(() => {
          wb.addWorksheet().name = 'THISISAWORKSHEETNAMEINUPPERCASE';
        }).toThrow();
      });
      it('throws an error (after truncation to 31 characters)', () => {
        const wb = new ExcelJS.Workbook();
        const validName = 'ThisIsAWorksheetNameThatIsLonge';
        wb.addWorksheet().name = validName;
        expect(() => {
          wb.addWorksheet().name = validName;
        }).toThrow();
        expect(() => {
          wb.addWorksheet().name = 'ThisIsAWorksheetNameThatIsLongerThan31';
        }).toThrow();
      });
    });
  });

  it.skip('returns sheet values — not supported: getSheetValues', () => {});

  it('sets row styles', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('basket');
    ws.getCell('A1').value = 5;
    ws.getCell('A1').numFmt = styles.numFmts.numFmt1;
    ws.getCell('A1').font = styles.fonts.arialBlackUI14;
    ws.getCell('C1').value = 'Hello, World!';
    ws.getCell('C1').alignment = styles.namedAlignments.bottomRight;
    ws.getCell('C1').border = styles.borders.doubleRed;
    ws.getCell('C1').fill = styles.fills.redDarkVertical;

    ws.getRow(1).numFmt = styles.numFmts.numFmt2;
    ws.getRow(1).font = styles.fonts.comicSansUdB16;
    ws.getRow(1).alignment = styles.namedAlignments.middleCentre;
    ws.getRow(1).border = styles.borders.thin;
    ws.getRow(1).fill = styles.fills.redGreenDarkTrellis;

    // C1, then B1: when we 'get' the previously null cell, it should inherit the row styles
    for (const ref of ['A1', 'C1', 'B1']) {
      const cell = ws.getCell(ref);
      expect(cell.numFmt, ref).toBe(styles.numFmts.numFmt2);
      expect(cell.font, ref).toEqual(styles.fonts.comicSansUdB16);
      expect(cell.alignment, ref).toEqual(styles.namedAlignments.middleCentre);
      expect(cell.border, ref).toEqual(styles.borders.thin);
      expect(cell.fill, ref).toEqual(styles.fills.redGreenDarkTrellis);
    }
  });

  it('sets col styles', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('basket');
    ws.getCell('A1').value = 5;
    ws.getCell('A1').numFmt = styles.numFmts.numFmt1;
    ws.getCell('A1').font = styles.fonts.arialBlackUI14;
    ws.getCell('A3').value = 'Hello, World!';
    ws.getCell('A3').alignment = styles.namedAlignments.bottomRight;
    ws.getCell('A3').border = styles.borders.doubleRed;
    ws.getCell('A3').fill = styles.fills.redDarkVertical;

    ws.getColumn('A').numFmt = styles.numFmts.numFmt2;
    ws.getColumn('A').font = styles.fonts.comicSansUdB16;
    ws.getColumn('A').alignment = styles.namedAlignments.middleCentre;
    ws.getColumn('A').border = styles.borders.thin;
    ws.getColumn('A').fill = styles.fills.redGreenDarkTrellis;

    // A3, then A2: when we 'get' the previously null cell, it should inherit the column styles
    for (const ref of ['A1', 'A3', 'A2']) {
      const cell = ws.getCell(ref);
      expect(cell.numFmt, ref).toBe(styles.numFmts.numFmt2);
      expect(cell.font, ref).toEqual(styles.fonts.comicSansUdB16);
      expect(cell.alignment, ref).toEqual(styles.namedAlignments.middleCentre);
      expect(cell.border, ref).toEqual(styles.borders.thin);
      expect(cell.fill, ref).toEqual(styles.fills.redGreenDarkTrellis);
    }
  });

  it('puts the lotion in the basket', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('basket');
    ws.getCell('A1').value = 'lotion';
  });

  it.skip('Should not break when importing a .numbers file — fixture not copied', () => {});
  it.skip('Should not break when importing an Excel file that contains a chartsheet — fixture not copied', () => {});
  it.skip('Hidden: should set hidden attribute correctly — fixtures not copied', () => {});
});
