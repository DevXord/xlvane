// Ported from exceljs v4.4.0 spec/integration/worksheet.spec.js (MIT, Copyright (c) 2014-2019
// Guyon Roche), changes: Mocha/Chai → Vitest; the 'Values' and 'Merge Cells' tests that repeat
// spec/unit/doc/worksheet.values.spec.js and worksheet.merge.spec.js are not repeated here (see
// worksheet-values.test.ts and worksheet-merge.test.ts); worksheet-name errors are checked with
// `toThrow()` without the exceljs message text; the `findCell`/`findRow` assertions (internal
// lookups outside the supported subset) are removed; the insertRow and style-option tests are
// ported in full (phase B); tests that read binary fixtures (.numbers, chartsheet, hidden sheets)
// are skipped because no fixtures were copied.
import { describe, expect, it } from 'vitest';
import { ExcelJS, expectNotSupported, styles } from './helpers';

describe('Worksheet', () => {
  describe('Values', () => {
    it('adds rows with style option', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      const dateValue1 = new Date(1970, 1, 1);
      const dateValue2 = new Date(1965, 1, 7);
      ws.addRow([1, 'John Doe', dateValue1]);
      ws.getRow(1).font = styles.fonts.comicSansUdB16;
      ws.addRow([2, 'Jane Doe', dateValue2], 'i');
      ws.addRow([3, 'Jane Doe', dateValue2], 'n');
      ws.addRow([4, 'Jane Doe', dateValue2], 'i');
      expect(ws.getCell('A1').font).toEqual(styles.fonts.comicSansUdB16);
      expect(ws.getCell('A2').font).toEqual(styles.fonts.comicSansUdB16);
      expect(ws.getCell('A3').font).not.toEqual(styles.fonts.comicSansUdB16);
      expect(ws.getCell('A4').font).not.toEqual(styles.fonts.comicSansUdB16);
    });

    it('inserts rows by object', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      ws.columns = [
        { header: 'Id', key: 'id', width: 10 },
        { header: 'Name', key: 'name', width: 32 },
        { header: 'D.O.B.', key: 'dob', width: 10 },
      ];
      const dateValue1 = new Date(1970, 1, 1);
      const dateValue2 = new Date(1965, 1, 7);
      const dateValue3 = new Date(1965, 1, 10);
      ws.addRow({ id: 1, name: 'John Doe', dob: dateValue1 });
      ws.addRow({ id: 2, name: 'Jane Doe', dob: dateValue2 });
      ws.insertRow(3, { id: 3, name: 'Other Doe', dob: dateValue3 });

      expect(ws.getCell('A2').value).toBe(1);
      expect(ws.getCell('B2').value).toBe('John Doe');
      expect(ws.getCell('C2').value).toBe(dateValue1);
      expect(ws.getCell('A3').value).toBe(3);
      expect(ws.getCell('B3').value).toBe('Other Doe');
      expect(ws.getCell('C3').value).toBe(dateValue3);
      expect(ws.getCell('A4').value).toBe(2);
      expect(ws.getCell('B4').value).toBe('Jane Doe');
      expect(ws.getCell('C4').value).toBe(dateValue2);

      const values: any[] = [, [, 'Id', 'Name', 'D.O.B.'], [, 1, 'John Doe', dateValue1], [, 3, 'Other Doe', dateValue3], [, 2, 'Jane Doe', dateValue2]];
      ws.eachRow((row: any, rowNumber: number) => {
        expect(row.values).toEqual(values[rowNumber]);
        row.eachCell((cell: any, colNumber: number) => expect(cell.value).toBe(values[rowNumber][colNumber]));
      });
      const fetchedRows = ws.getRows(1, 2);
      for (let i = 0; i < 2; i++) expect(fetchedRows[i].values).toEqual(values[i + 1]);
    });

    it('inserts rows by contiguous array', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      const dateValue1 = new Date(1970, 1, 1);
      const dateValue2 = new Date(1965, 1, 7);
      const dateValue3 = new Date(1965, 1, 10);
      ws.addRow([1, 'John Doe', dateValue1]);
      ws.addRow([2, 'Jane Doe', dateValue2]);
      ws.insertRow(2, [3, 'Other Doe', dateValue3]);

      expect(ws.getCell('A1').value).toBe(1);
      expect(ws.getCell('B1').value).toBe('John Doe');
      expect(ws.getCell('C1').value).toBe(dateValue1);
      expect(ws.getCell('A2').value).toBe(3);
      expect(ws.getCell('B2').value).toBe('Other Doe');
      expect(ws.getCell('C2').value).toBe(dateValue3);
      expect(ws.getCell('A3').value).toBe(2);
      expect(ws.getCell('B3').value).toBe('Jane Doe');
      expect(ws.getCell('C3').value).toBe(dateValue2);

      const values: any[] = [[, 1, 'John Doe', dateValue1], [, 3, 'Other Doe', dateValue3], [, 2, 'Jane Doe', dateValue2]];
      for (let i = 0; i < 3; i++) expect(ws.getRow(i + 1).values).toEqual(values[i]);
      const fetchedRows = ws.getRows(1, 3);
      for (let i = 0; i < 3; i++) expect(fetchedRows[i].values).toEqual(values[i]);
    });

    it('inserts rows by sparse array', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      const dateValue1 = new Date(1970, 1, 1);
      const dateValue2 = new Date(1965, 1, 7);
      const dateValue3 = new Date(1965, 1, 10);
      const rows: any[] = [, [, 1, 'John Doe', , dateValue1], [, 2, 'Jane Doe', , dateValue2]];
      const row3: unknown[] = [];
      row3[1] = 3;
      row3[3] = 'Other Doe';
      row3[5] = dateValue3;
      rows.push(row3);
      rows.forEach((row) => {
        if (row) ws.insertRow(1, row);
      });

      expect(ws.getCell('A1').value).toBe(3);
      expect(ws.getCell('C1').value).toBe('Other Doe');
      expect(ws.getCell('E1').value).toBe(dateValue3);
      expect(ws.getCell('A2').value).toBe(2);
      expect(ws.getCell('B2').value).toBe('Jane Doe');
      expect(ws.getCell('D2').value).toBe(dateValue2);
      expect(ws.getCell('A3').value).toBe(1);
      expect(ws.getCell('B3').value).toBe('John Doe');
      expect(ws.getCell('D3').value).toBe(dateValue1);

      ws.eachRow((row: any, rowNumber: number) => {
        expect(row.values).toEqual(rows[rows.length - rowNumber]);
        row.eachCell((cell: any, colNumber: number) => expect(cell.value).toBe(rows[rows.length - rowNumber][colNumber]));
      });
      const fetchedRows = ws.getRows(1, 3);
      for (let i = 0; i < 3; i++) expect(fetchedRows[i].values).toEqual(rows[rows.length - i - 1]);
    });

    it('inserts rows with style option', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      const dateValue1 = new Date(1970, 1, 1);
      const dateValue2 = new Date(1965, 1, 7);
      const arr = [
        [5, 'Jane Doe', dateValue2],
        [5, 'Jane Doe', dateValue2],
        [5, 'Jane Doe', dateValue2],
      ];
      ws.addRow([5, 'John Doe', dateValue1]);
      ws.getRow(1).font = styles.fonts.comicSansUdB16;
      ws.insertRow(1, [5, 'Jane Doe', dateValue2], 'o');
      ws.insertRow(1, [4, 'Jane Doe', dateValue2], 'i');
      ws.insertRow(1, [3, 'Jane Doe', dateValue2], 'n');
      ws.insertRow(1, [2, 'Jane Doe', dateValue2], 'o');
      ws.addRow([6, 'Jane Doe', dateValue2]);
      ws.getRow(6).font = styles.fonts.comicSansUdB16;
      ws.insertRows(6, arr, 'o');
      ws.insertRows(10, arr, 'i');
      ws.insertRows(13, arr);

      for (const i of [1, 2, 3, 7, 8, 13, 14, 15]) expect(ws.getCell(`A${String(i)}`).font, `A${String(i)}`).not.toEqual(styles.fonts.comicSansUdB16);
      for (const i of [4, 5, 6, 9, 10, 11, 12]) expect(ws.getCell(`A${String(i)}`).font, `A${String(i)}`).toEqual(styles.fonts.comicSansUdB16);
    });

    it('should style of the inserted row with inherited style be mutable', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      const dateValue1 = new Date(1970, 1, 1);
      const dateValue2 = new Date(1965, 1, 7);
      ws.addRow([1, 'John Doe', dateValue1]);
      ws.getRow(1).font = styles.fonts.comicSansUdB16;
      ws.insertRow(2, [3, 'Jane Doe', dateValue2], 'i');
      ws.insertRow(2, [2, 'Jane Doe', dateValue2], 'o');
      ws.getRow(2).font = styles.fonts.broadwayRedOutline20;
      ws.getRow(3).font = styles.fonts.broadwayRedOutline20;
      ws.getCell('A2').font = styles.fonts.arialBlackUI14;
      ws.getCell('A3').font = styles.fonts.arialBlackUI14;
      expect(ws.getRow(2).font).not.toEqual(styles.fonts.comicSansUdB16);
      expect(ws.getRow(3).font).not.toEqual(styles.fonts.comicSansUdB16);
      expect(ws.getCell('A2').font).not.toEqual(styles.fonts.comicSansUdB16);
      expect(ws.getCell('A3').font).not.toEqual(styles.fonts.comicSansUdB16);
      // xlvane addition: the copied styles are copies, so row 1 keeps its font.
      expect(ws.getCell('A1').font).toEqual(styles.fonts.comicSansUdB16);
    });

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
