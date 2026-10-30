// xlvane-specific tests of cell notes in @xlvane/exceljs (D-A14; docs/facts/exceljs.md X4-14).
// exceljs writes the author "Author" for every note and gives a one-run note back as a string;
// the VML box settings (margins, protection, editAs) are not written by xlvane.
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { ExcelJS, roundTrip } from './helpers';

describe('cell notes', () => {
  it('writes string and rich notes with the author exceljs writes, and reads them back', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Notes');
    ws.getCell('A1').value = 1;
    ws.getCell('A1').note = 'plain';
    ws.getCell('B1').value = 2;
    ws.getCell('B1').note = { texts: [{ font: { bold: true }, text: 'Bold' }, { text: ' rest' }], editAs: 'oneCells' };
    const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
    const parts = unzipSync(bytes);
    expect(strFromU8(parts['xl/comments1.xml']!)).toContain('<author>Author</author>');
    expect(readWorkbook(bytes).getSheet('Notes').notes.map((n) => [n.ref, n.text, n.author])).toEqual([
      ['A1', 'plain', 'Author'],
      ['B1', 'Bold rest', 'Author'],
    ]);
    const ws2 = (await roundTrip(wb)).getWorksheet('Notes');
    expect(ws2.getCell('A1').note).toBe('plain');
    expect(ws2.getCell('B1').note).toEqual({ texts: [{ font: { bold: true }, text: 'Bold' }, { text: ' rest' }] });
  });

  it('keeps a note on a cell without a value (exceljs loses it on reading)', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S').getCell('C3').note = 'only a note';
    const ws2 = (await roundTrip(wb)).getWorksheet('S');
    expect(ws2.getCell('C3').value).toBeNull();
    expect(ws2.getCell('C3').note).toBe('only a note');
  });

  it('a note with no texts is written as an empty note; undefined removes a note', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').note = { texts: [] };
    ws.getCell('B1').note = 'gone';
    ws.getCell('B1').note = undefined;
    const ws2 = (await roundTrip(wb)).getWorksheet('S');
    expect(ws2.getCell('A1').note).toBe('');
    expect(ws2.getCell('B1').note).toBeUndefined();
  });

  it('a note belongs to its own cell inside a merged range', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'merged';
    ws.mergeCells('A1:B2');
    ws.getCell('B2').note = 'inside';
    const ws2 = (await roundTrip(wb)).getWorksheet('S');
    expect(ws2.getCell('B2').note).toBe('inside');
    expect(ws2.getCell('A1').note).toBeUndefined();
  });
});
