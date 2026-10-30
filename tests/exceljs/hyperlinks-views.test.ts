// xlvane-specific tests of phase A2 in @xlvane/exceljs: hyperlink values, worksheet tab colour
// and views, company and manager (D-A13; docs/facts/exceljs.md X4-10 … X4-13). Links are
// written as the specification describes (D-A08), which differs from exceljs for links inside
// the workbook; the differences are listed in docs/exceljs-compat.md.
import { join } from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readWorkbookFile } from '../../packages/xlvane/src/node';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { ExcelJS, expectNotSupported, roundTrip, withTempDir } from './helpers';

async function bytesOf(wb: any): Promise<Uint8Array> {
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

describe('hyperlink values', () => {
  it('writes an external link as a relationship and a "#" link as a location', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Links');
    ws.getCell('A1').value = { text: 'Docs', hyperlink: 'https://example.com/?a=1&b=2', tooltip: 'Open the docs' };
    ws.getCell('A2').value = { text: 'Totals', hyperlink: "#'Other Sheet'!B2" };
    wb.addWorksheet('Other Sheet').getCell('B2').value = 1;
    const bytes = await bytesOf(wb);
    const parts = unzipSync(bytes);
    const sheet = strFromU8(parts['xl/worksheets/sheet1.xml']!);
    const rels = strFromU8(parts['xl/worksheets/_rels/sheet1.xml.rels']!);
    expect(sheet).toMatch(/<hyperlink ref="A1" r:id="rId1" tooltip="Open the docs"\/>/);
    expect(sheet).toContain(`<hyperlink ref="A2" location="'Other Sheet'!B2"/>`);
    expect(rels).toContain('Target="https://example.com/?a=1&amp;b=2" TargetMode="External"');
    expect(rels.match(/<Relationship /g)).toHaveLength(1);
    const s = readWorkbook(bytes).getSheet('Links');
    expect(s.cell('A1')?.value).toBe('Docs');
    expect(s.cell('A1')?.hyperlink).toMatchObject({ target: 'https://example.com/?a=1&b=2', location: null, tooltip: 'Open the docs' });
  });

  it('reads links back as { text, hyperlink, tooltip }', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Links');
    ws.getCell('A1').value = { text: 'Docs', hyperlink: 'https://example.com', tooltip: 'Open the docs' };
    ws.getCell('A2').value = { text: 'Totals', hyperlink: '#Links!B2' };
    ws.getCell('A3').value = { text: 'Mail', hyperlink: 'mailto:someone@example.com' };
    const ws2 = (await roundTrip(wb)).getWorksheet('Links');
    expect(ws2.getCell('A1').value).toEqual({ text: 'Docs', hyperlink: 'https://example.com', tooltip: 'Open the docs' });
    expect(ws2.getCell('A2').value).toEqual({ text: 'Totals', hyperlink: '#Links!B2' });
    expect(ws2.getCell('A3').type).toBe(ExcelJS.ValueType.Hyperlink);
    expect(ws2.getCell('A3').hyperlink).toBe('mailto:someone@example.com');
    expect(ws2.getCell('A3').text).toBe('Mail');
    expect(ws2.getCell('A3').isHyperlink).toBe(true);
  });

  it('keeps rich link text and the style of the cell', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Links');
    ws.getCell('A1').value = { text: { richText: [{ font: { bold: true }, text: 'Bold' }, { text: ' link' }] }, hyperlink: 'https://example.com' };
    ws.getCell('A1').font = { underline: true, color: { argb: 'FF0000FF' } };
    expect(ws.getCell('A1').text).toBe('Bold link');
    const ws2 = (await roundTrip(wb)).getWorksheet('Links');
    expect(ws2.getCell('A1').value).toEqual({ text: { richText: [{ font: { bold: true }, text: 'Bold' }, { text: ' link' }] }, hyperlink: 'https://example.com' });
    expect(ws2.getCell('A1').font).toMatchObject({ underline: 'single', color: { argb: 'FF0000FF' } });
  });

  it('is a hyperlink only when both text and hyperlink are set, as exceljs (cell.js L1078)', async () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    ws.getCell('A1').value = { text: 'x', hyperlink: 'https://example.com' };
    expect(ws.getCell('A1').type).toBe(ExcelJS.ValueType.Hyperlink);
    expect(ws.getCell('A1').hyperlink).toBe('https://example.com');
    ws.getCell('A2').value = 'plain';
    expect(ws.getCell('A2').hyperlink).toBeUndefined();
    // Without text it is no link: exceljs stores it as a JSON value of type String.
    ws.getCell('A3').value = { text: '', hyperlink: 'https://example.com' };
    expect(ws.getCell('A3').type).toBe(ExcelJS.ValueType.String);
    expect(ws.getCell('A3').hyperlink).toBeUndefined();
  });

  it('reports invalid links with XLV_E_INVALID_HYPERLINK when writing', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S').getCell('A1').value = { text: 'x', hyperlink: 'https://example.com', tooltip: 't'.repeat(256) };
    await expect(wb.xlsx.writeBuffer()).rejects.toMatchObject({ code: 'XLV_E_INVALID_HYPERLINK' });
  });

  it('streamed worksheets write links with their row', async () => {
    await withTempDir(async (dir) => {
      const filename = join(dir, 'links.xlsx');
      const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ filename });
      const sheet = writer.addWorksheet('S');
      sheet.addRow([{ text: 'Docs', hyperlink: 'https://example.com', tooltip: 'tip' }, { text: 'Back', hyperlink: '#S!A1' }]).commit();
      await writer.commit();
      const book = await readWorkbookFile(filename);
      expect(book.getSheet('S').hyperlinks).toEqual([
        { ref: 'A1', target: 'https://example.com', location: null, tooltip: 'tip', display: null },
        { ref: 'B1', target: null, location: 'S!A1', tooltip: null, display: null },
      ]);
    });
  });
});

describe('worksheet tab colour and views', () => {
  it('writes and reads properties.tabColor', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('Red', { properties: { tabColor: { argb: 'FFFF0000' } } }).getCell('A1').value = 1;
    wb.addWorksheet('Theme', { properties: { tabColor: { theme: 4, tint: 0.4 } } }).getCell('A1').value = 1;
    wb.addWorksheet('Plain').getCell('A1').value = 1;
    const wb2 = await roundTrip(wb);
    expect(wb2.getWorksheet('Red').properties.tabColor).toEqual({ argb: 'FFFF0000' });
    expect(wb2.getWorksheet('Theme').properties.tabColor).toEqual({ theme: 4, tint: 0.4 });
    expect(wb2.getWorksheet('Plain').properties.tabColor).toBeUndefined();
  });

  it('writes zoom, grid lines, headings, direction and layout of the first view', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('V', { views: [{ zoomScale: 150, showGridLines: false, showRowColHeaders: false, rightToLeft: true, style: 'pageBreakPreview' }] });
    ws.getCell('A1').value = 1;
    const bytes = await bytesOf(wb);
    expect(readWorkbook(bytes).getSheet('V').view).toEqual({ zoom: 150, showGridLines: false, showRowColHeaders: false, rightToLeft: true, layout: 'pageBreakPreview', split: null, topLeftCell: null, activeCell: null, zoomScaleNormal: 0, showRuler: true });
    const ws2 = (await roundTrip(wb)).getWorksheet('V');
    expect(ws2.views).toEqual([
      { workbookViewId: 0, state: 'normal', rightToLeft: true, showRuler: true, showRowColHeaders: false, showGridLines: false, zoomScale: 150, zoomScaleNormal: 100, style: 'pageBreakPreview' },
    ]);
  });

  it('reads a normal view for every worksheet, as exceljs', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S').getCell('A1').value = 1;
    const ws2 = (await roundTrip(wb)).getWorksheet('S');
    expect(ws2.views).toEqual([{ workbookViewId: 0, state: 'normal', rightToLeft: false, showRuler: true, showRowColHeaders: true, showGridLines: true, zoomScale: 100, zoomScaleNormal: 100 }]);
  });

  it('refuses a zoom outside 10–400 with XLV_E_INVALID_OPTION', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S', { views: [{ zoomScale: 500 }] }).getCell('A1').value = 1;
    await expect(wb.xlsx.writeBuffer()).rejects.toMatchObject({ code: 'XLV_E_INVALID_OPTION' });
  });

  it('the streaming writer writes tab colour and views set before the first commit', async () => {
    await withTempDir(async (dir) => {
      const filename = join(dir, 's.xlsx');
      const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ filename });
      const sheet = writer.addWorksheet('S', { properties: { tabColor: { argb: 'FF00B050' } }, views: [{ state: 'split', xSplit: 2400, zoomScale: 80 }] });
      sheet.addRow([1]).commit();
      sheet.commit();
      await writer.commit();
      const s = (await readWorkbookFile(filename)).getSheet('S');
      expect(s.tabColor).toBe('FF00B050');
      expect(s.view).toMatchObject({ zoom: 80, split: { x: 120, y: 0 } });
    });
  });

  it('the streaming writer refuses a view changed after rows were committed', async () => {
    const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: { write: (_c: unknown, cb: () => void) => cb(), end: (cb: () => void) => cb() } });
    const sheet = writer.addWorksheet('S');
    sheet.addRow([1]).commit();
    sheet.views = [{ zoomScale: 120 }];
    await expectNotSupported(() => sheet.commit());
    await expect(writer.commit()).rejects.toMatchObject({ code: 'XLV_E_NOT_SUPPORTED' });
  });
});

describe('company and manager', () => {
  it('are written to docProps/app.xml as exceljs writes them (Company always, Manager when set)', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S').getCell('A1').value = 1;
    wb.manager = undefined;
    const app = strFromU8(unzipSync(await bytesOf(wb))['docProps/app.xml']!);
    expect(app).toContain('<Company></Company>');
    expect(app).not.toContain('<Manager>');
    expect(app).not.toContain('<Application>');
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(await bytesOf(wb));
    expect(wb2.company).toBe('');
    expect(wb2.manager).toBeUndefined();
  });
});
