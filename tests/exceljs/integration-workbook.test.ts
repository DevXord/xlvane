// Ported from exceljs v4.4.0 spec/integration/workbook/workbook.spec.js (MIT, Copyright (c)
// 2014-2019 Guyon Roche), changes: Mocha/Chai → Vitest with async/await; files go to a temporary
// directory (`roundTrip`); the merge-styles test uses a pattern fill and a font without
// `family`/`outline` (gradient fills and those font attributes are not written by xlvane);
// the frozen-panes test writes and checks one view (only the first sheet view is read);
// 'wrong data type' checks that `load({})` rejects, not the JSZip message; tests that need
// unsupported features (createTestBook fixtures, compression options, print titles,
// document properties, data validation, CSV, defined names, duplicateRow, spliceRows, split
// panes, workbook views, fillFormula) are skipped or expect XLV_E_NOT_SUPPORTED.
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { ExcelJS, expectNotSupported, roundTrip, styles, withTempDir } from './helpers';

describe('Workbook', () => {
  describe('Serialise', () => {
    it.skip('xlsx file — not supported: createTestBook fixture (needs rich text, notes, data validation)', () => {});

    describe('Xlsx Zip Compression', () => {
      for (const level of ['best', 'default', 'fast', 'no']) {
        it.skip(`xlsx file with ${level} compression — not supported: zip compression options`, () => {});
      }
    });

    it('sheets with correct names', async () => {
      const wb = new ExcelJS.Workbook();
      const ws1 = wb.addWorksheet('Hello, World!');
      expect(ws1.name).toBe('Hello, World!');
      ws1.getCell('A1').value = 'Hello, World!';
      const ws2 = wb.addWorksheet();
      expect(ws2.name).toMatch(/sheet\d+/);
      ws2.getCell('A1').value = ws2.name;
      wb.addWorksheet('This & That');

      const wb2 = await roundTrip(wb);
      expect(wb2.getWorksheet('Hello, World!')).toBeTruthy();
      expect(wb2.getWorksheet('This & That')).toBeTruthy();
    });

    it.skip('creator, lastModifiedBy, etc — not supported: document properties are not written in phase 5', () => {});
    it.skip('printTitlesRow — not supported: print settings', () => {});
    it.skip('printTitlesColumn — not supported: print settings', () => {});
    it.skip('printTitlesRowAndColumn — not supported: print settings', () => {});
    it.skip('shared formula — not supported: fillFormula', () => {});

    it('auto filter', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = 1;
      ws.getCell('B1').value = 1;
      ws.getCell('A2').value = 2;
      ws.getCell('B2').value = 2;
      ws.getCell('A3').value = 3;
      ws.getCell('B3').value = 3;
      ws.autoFilter = 'A1:B1';

      const wb2 = await roundTrip(wb);
      expect(wb2.getWorksheet('Hello').autoFilter).toBe('A1:B1');
    });

    it.skip('company, manager, etc — not supported: document properties are not written in phase 5', () => {});
    it.skip('title, subject, etc — not supported: document properties are not written in phase 5', () => {});
    it.skip('language, revision and contentStatus — not supported: document properties are not written in phase 5', () => {});

    it('empty strings', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = 'Foo';
      ws.getCell('A2').value = '';
      ws.getCell('A3').value = 'Baz';

      const ws2 = (await roundTrip(wb)).getWorksheet('Hello');
      expect(ws2.getCell('A1').value).toBe('Foo');
      expect(ws2.getCell('A2').value).toBe('');
      expect(ws2.getCell('A3').value).toBe('Baz');
    });

    it.skip('dataValidations — not supported: data validation', () => {});

    it('empty string', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet();
      ws.columns = [
        { key: 'id', width: 10 },
        { key: 'name', width: 32 },
      ];
      ws.addRow({ id: 1, name: '' });
      await withTempDir((dir) => wb.xlsx.writeFile(join(dir, 'wb.test.xlsx')));
    });

    it.skip('a lot of sheets to xlsx file — not supported: createTestBook fixture', () => {});

    it('csv file — not supported: csv', async () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('blort').getCell('A1').value = 1;
      await expectNotSupported(() => wb.csv.writeBuffer());
    });

    it.skip('CSV file and its configuration — not supported: csv', () => {});
    it.skip('defined names — not supported: defined names', () => {});

    describe('Duplicate Rows', () => {
      it('Duplicate rows with styles properly — not supported: duplicateRow', async () => {
        const ws = new ExcelJS.Workbook().addWorksheet('duplicateTest');
        ws.getCell('A1').value = 'OneInfo';
        await expectNotSupported(() => ws.duplicateRow(1, 2, true));
      });
      it.skip('Duplicate rows replacing properly — not supported: duplicateRow', () => {});
      it.skip('Duplicate rows shifting properly — not supported: duplicateRow', () => {});
      it.skip('Duplicate rows with height properly — not supported: duplicateRow', () => {});
    });

    describe('Merge Cells', () => {
      it('serialises and deserialises properly', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('blort');
        ws.getCell('B2').value = 'B2';
        ws.mergeCells('B2:C3');

        const ws2 = (await roundTrip(wb)).getWorksheet('blort');
        expect(ws2.getCell('B2').value).toBe('B2');
        expect(ws2.getCell('B3').value).toBe('B2');
        expect(ws2.getCell('C2').value).toBe('B2');
        expect(ws2.getCell('C3').value).toBe('B2');
        expect(ws2.getCell('B2').type).toBe(ExcelJS.ValueType.String);
        expect(ws2.getCell('B3').type).toBe(ExcelJS.ValueType.Merge);
        expect(ws2.getCell('C2').type).toBe(ExcelJS.ValueType.Merge);
        expect(ws2.getCell('C3').type).toBe(ExcelJS.ValueType.Merge);
      });

      it('styles', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('blort');
        const B2 = ws.getCell('B2');
        B2.value = 5;
        B2.style.font = styles.roundTripFonts.broadwayRed20;
        B2.style.border = styles.borders.doubleRed;
        B2.style.fill = styles.fills.redDarkVertical;
        B2.style.alignment = styles.namedAlignments.middleCentre;
        B2.style.numFmt = styles.numFmts.numFmt1;
        // expecting styles to be copied (see worksheet spec)
        ws.mergeCells('B2:C3');

        const ws2 = (await roundTrip(wb)).getWorksheet('blort');
        for (const ref of ['B2', 'B3', 'C2', 'C3']) {
          const cell = ws2.getCell(ref);
          expect(cell.font, ref).toEqual(styles.roundTripFonts.broadwayRed20);
          expect(cell.border, ref).toEqual(styles.borders.doubleRed);
          expect(cell.fill, ref).toEqual(styles.fills.redDarkVertical);
          expect(cell.alignment, ref).toEqual(styles.namedAlignments.middleCentre);
          expect(cell.numFmt, ref).toBe(styles.numFmts.numFmt1);
        }
      });
    });
  });

  it.skip('spliced meat and ham — not supported: spliceRows', () => {});

  it('throws an error when xlsx file not found', async () => {
    const wb = new ExcelJS.Workbook();
    await expect(wb.xlsx.readFile('./wb.doesnotexist.xlsx')).rejects.toBeDefined();
  });

  it('throws an error when csv file not found — not supported: csv', async () => {
    const wb = new ExcelJS.Workbook();
    await expectNotSupported(() => wb.csv.readFile('./wb.doesnotexist.csv'));
  });

  it('throw an error for wrong data type', async () => {
    const wb = new ExcelJS.Workbook();
    await expect(wb.xlsx.load({})).rejects.toBeDefined();
  });

  describe('Sheet Views', () => {
    it('frozen panes', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('frozen');
      ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 3, topLeftCell: 'C4', activeCell: 'D5' }];
      ws.getCell('A1').value = 'Let it Snow!';

      const ws2 = (await roundTrip(wb)).getWorksheet('frozen');
      expect(ws2).toBeTruthy();
      expect(ws2.getCell('A1').value).toBe('Let it Snow!');
      expect(ws2.views[0]).toMatchObject({ state: 'frozen', xSplit: 2, ySplit: 3 });
    });

    it.skip('serialises split panes — not supported: split panes', () => {});
    it.skip('multiple book views — not supported: workbook views', () => {});
  });
});
