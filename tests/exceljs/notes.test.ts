// xlvane-specific tests of cell notes in @xlvane/compat (D-A14; docs/facts/exceljs.md X4-14).
// exceljs writes the author "Author" for every note and gives a one-run note back as a string;
// the VML box settings (margins, protection, editAs) are written and read back since D-E07.
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
    expect(ws2.getCell('B1').note).toEqual({
      texts: [{ font: { bold: true }, text: 'Bold' }, { text: ' rest' }],
      margins: { insetmode: 'auto', inset: [0.13, 0.13, 0.25, 0.25] },
      editAs: 'oneCells',
      protection: { locked: 'True', lockText: 'True' },
    });
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

  it('writes the box settings to the VML part as exceljs does and reads them back (D-E07, X4-31)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Notes');
    ws.getCell('A1').value = 1;
    ws.getCell('A1').note = 'plain';
    ws.getCell('B1').value = 2;
    const box = { margins: { insetmode: 'custom' as const, inset: [0.2, 0.15, 0.2, 0.1] }, protection: { locked: 'False' as const, lockText: 'False' as const }, editAs: 'twoCells' as const };
    ws.getCell('B1').note = { texts: [{ text: 'a' }, { text: 'b', font: { bold: true } }], ...box };
    const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
    const vml = strFromU8(unzipSync(bytes)['xl/drawings/vmlDrawing1.vml']!);
    // exceljs writes the default box for a plain note: automatic margins of 1.3/1.3/2.5/2.5 mm.
    expect(vml).toContain('inset="1.3mm,1.3mm,2.5mm,2.5mm"');
    expect(vml).toContain('inset="2mm,1.5mm,2mm,1mm"');
    expect(vml).toContain('o:insetmode="custom"');
    expect(vml).toContain('<x:Locked>False</x:Locked>');
    const note = (await roundTrip(wb)).getWorksheet('Notes').getCell('B1').note;
    expect({ margins: note.margins, protection: note.protection, editAs: note.editAs }).toEqual(box);
    expect(readWorkbook(bytes).getSheet('Notes').notes[1]!.box).toEqual({ autoMargins: false, margins: { left: 2, top: 1.5, right: 2, bottom: 1 }, locked: false, lockText: false, anchor: 'twoCell' });
  });
});
