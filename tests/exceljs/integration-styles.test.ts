// Ported from exceljs v4.4.0 spec/integration/workbook/styles.spec.js (MIT, Copyright (c)
// 2014-2019 Guyon Roche), changes: Mocha/Chai → Vitest with async/await; the row style read-back (`ws2.getRow(2).font`)
// is removed because xlvane keeps no row default styles after reading (docs/decisions.md
// D-206); the rich-text fixture (spec/integration/data/rich-text-sample.js, decoded to
// tests/corpus/excel/exceljs-rich-text-sample.xlsx, and rich-text-sample-a1.json) is read
// from files and compared without the run fonts' `family`/`scheme`; 'sets row styles' and 'sets
// col styles' repeat spec/integration/worksheet.spec.js and are ported once, in
// integration-worksheet.test.ts.
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { ExcelJS, roundTrip, styles } from './helpers';
import richTextSampleA1 from './rich-text-sample-a1.json';

describe('Workbook', () => {
  describe('Styles', () => {
    it('row styles and columns properly', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('blort');
      ws.columns = [
        { header: 'A1', width: 10 },
        {
          header: 'B1',
          width: 20,
          style: { font: styles.fonts.comicSansUdB16, alignment: styles.alignments[1].alignment },
        },
        { header: 'C1', width: 30 },
      ];
      ws.getRow(2).font = styles.fonts.broadwayRedOutline20;
      ws.getCell('A2').value = 'A2';
      ws.getCell('B2').value = 'B2';
      ws.getCell('C2').value = 'C2';
      ws.getCell('A3').value = 'A3';
      ws.getCell('B3').value = 'B3';
      ws.getCell('C3').value = 'C3';

      const ws2 = (await roundTrip(wb)).getWorksheet('blort');
      for (const address of ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3']) {
        expect(ws2.getCell(address).value).toBe(address);
      }
      expect(ws2.getCell('B1').font).toEqual(styles.fonts.comicSansUdB16);
      expect(ws2.getCell('B1').alignment).toEqual(styles.alignments[1].alignment);
      expect(ws2.getCell('A2').font).toEqual(styles.fonts.broadwayRedOutline20);
      expect(ws2.getCell('B2').font).toEqual(styles.fonts.broadwayRedOutline20);
      expect(ws2.getCell('C2').font).toEqual(styles.fonts.broadwayRedOutline20);
      expect(ws2.getCell('B3').font).toEqual(styles.fonts.comicSansUdB16);
      expect(ws2.getCell('B3').alignment).toEqual(styles.alignments[1].alignment);

      expect(ws2.getColumn(2).font).toEqual(styles.fonts.comicSansUdB16);
      expect(ws2.getColumn(2).alignment).toEqual(styles.alignments[1].alignment);
    });

    it('in-cell formats properly in xlsx file', async () => {
      const bytes = readFileSync(new URL('../corpus/excel/exceljs-rich-text-sample.xlsx', import.meta.url));
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(bytes);
      const ws = wb.worksheets[0];
      // Run fonts keep `family`, `scheme` and the other font properties (D-E04), as in exceljs.
      expect(ws.getCell('A1').value).toEqual(richTextSampleA1);
      expect(ws.getCell('A1').text).toBe(ws.getCell('A2').value);
    });
    it('null cells retain style', async () => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('blort');
      // one value here
      ws.getCell('B2').value = 'hello';
      // style here
      ws.getCell('B4').fill = styles.fills.redDarkVertical;
      ws.getCell('B4').font = styles.fonts.broadwayRedOutline20;
      const ws2 = (await roundTrip(wb)).getWorksheet('blort');
      expect(ws2.getCell('B4').fill).toEqual(styles.fills.redDarkVertical);
      expect(ws2.getCell('B4').font).toEqual(styles.fonts.broadwayRedOutline20);
    });
    it.skip('sets row styles - duplicate of spec/integration/worksheet.spec.js, see integration-worksheet.test.ts', () => {});
    it.skip('sets col styles - duplicate of spec/integration/worksheet.spec.js, see integration-worksheet.test.ts', () => {});
  });
});
