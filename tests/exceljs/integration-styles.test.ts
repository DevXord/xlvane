// Ported from exceljs v4.4.0 spec/integration/workbook/styles.spec.js (MIT, Copyright (c)
// 2014-2019 Guyon Roche), changes: Mocha/Chai → Vitest with async/await; fonts without
// `family`/`outline` (not written by xlvane); the column and row style read-back
// (`ws2.getColumn(2).font`, `ws2.getRow(2).font`) is removed because xlvane keeps no row or
// column default styles after reading (docs/decisions.md D-206); 'sets row styles' and 'sets
// col styles' repeat spec/integration/worksheet.spec.js and are ported once, in
// integration-worksheet.test.ts.
import { describe, expect, it } from 'vitest';
import { ExcelJS, roundTrip, styles } from './helpers';

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
          style: { font: styles.roundTripFonts.comicSansUdB16, alignment: styles.alignments[1].alignment },
        },
        { header: 'C1', width: 30 },
      ];
      ws.getRow(2).font = styles.roundTripFonts.broadwayRed20;
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
      expect(ws2.getCell('B1').font).toEqual(styles.roundTripFonts.comicSansUdB16);
      expect(ws2.getCell('B1').alignment).toEqual(styles.alignments[1].alignment);
      expect(ws2.getCell('A2').font).toEqual(styles.roundTripFonts.broadwayRed20);
      expect(ws2.getCell('B2').font).toEqual(styles.roundTripFonts.broadwayRed20);
      expect(ws2.getCell('C2').font).toEqual(styles.roundTripFonts.broadwayRed20);
      expect(ws2.getCell('B3').font).toEqual(styles.roundTripFonts.comicSansUdB16);
      expect(ws2.getCell('B3').alignment).toEqual(styles.alignments[1].alignment);
    });

    it.skip('in-cell formats properly in xlsx file — not supported: rich text values (and the base64 fixture is not copied)', () => {});
    it.skip('null cells retain style — not supported: the reader drops cells without a value', () => {});
    it.skip('sets row styles — duplicate of spec/integration/worksheet.spec.js, see integration-worksheet.test.ts', () => {});
    it.skip('sets col styles — duplicate of spec/integration/worksheet.spec.js, see integration-worksheet.test.ts', () => {});
  });
});
