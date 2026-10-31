// xlvane tests for tables and images in @xlvane/compat (phase D, docs/decisions.md D-D04).
// The image tests are ported from exceljs v4.4.0 spec/integration/workbook/images.spec.js and
// spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js (MIT, Copyright (c)
// 2014-2019 Guyon Roche); changes: the image is passed as `buffer` (xlvane does not read files by
// name, D-D04) and expectations that differ on purpose are marked "xlvane:".
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { ExcelJS, withTempDir } from './helpers';

const IMAGE = new Uint8Array(readFileSync(new URL('./data/image.png', import.meta.url)));

/** Whether `bytes` equals the test image (exceljs's tests compare with Buffer.compare). */
function same(bytes: Uint8Array): boolean {
  return bytes.length === IMAGE.length && bytes.every((b, i) => b === IMAGE[i]);
}

async function reload(wb: any): Promise<any> {
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(await wb.xlsx.writeBuffer());
  return wb2;
}

describe('Workbook Images (ported)', () => {
  it('stores background image', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    const imageId = wb.addImage({ buffer: IMAGE, extension: 'jpeg' });
    ws.getCell('A1').value = 'Hello, World!';
    ws.addBackgroundImage(imageId);
    const wb2 = await reload(wb);
    const ws2 = wb2.getWorksheet('blort');
    expect(ws2).toBeDefined();
    const image = wb2.getImage(ws2.getBackgroundImageId());
    expect(same(image.buffer)).toBe(true);
  });

  it('stores embedded image and hyperlink', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    const imageId = wb.addImage({ buffer: IMAGE, extension: 'jpeg' });
    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('A2').value = { hyperlink: 'http://www.somewhere.com', text: 'www.somewhere.com' };
    ws.addImage(imageId, 'C3:E6');
    const wb2 = await reload(wb);
    const ws2 = wb2.getWorksheet('blort');
    expect(ws2.getCell('A1').value).toBe('Hello, World!');
    expect(ws2.getCell('A2').value).toEqual({ hyperlink: 'http://www.somewhere.com', text: 'www.somewhere.com' });
    const images = ws2.getImages();
    expect(images.length).toBe(1);
    const imageDesc = images[0];
    expect(imageDesc.range.tl.col).toBe(2);
    expect(imageDesc.range.tl.row).toBe(2);
    expect(imageDesc.range.br.col).toBe(5);
    expect(imageDesc.range.br.row).toBe(6);
    const image = wb2.getImage(imageDesc.imageId);
    expect(same(image.buffer)).toBe(true);
  });

  it('stores embedded image with oneCell', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    const imageId = wb.addImage({ buffer: IMAGE, extension: 'jpeg' });
    ws.addImage(imageId, { tl: { col: 0.1125, row: 0.4 }, br: { col: 2.101046875, row: 3.4 }, editAs: 'oneCell' });
    const ws2 = (await reload(wb)).getWorksheet('blort');
    const images = ws2.getImages();
    expect(images.length).toBe(1);
    expect(images[0].range.editAs).toBe('oneCell');
  });

  it('stores embedded image with one-cell-anchor', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    const imageId = wb.addImage({ buffer: IMAGE, extension: 'jpeg' });
    ws.addImage(imageId, { tl: { col: 0.1125, row: 0.4 }, ext: { width: 100, height: 100 }, editAs: 'oneCell' });
    const ws2 = (await reload(wb)).getWorksheet('blort');
    const images = ws2.getImages();
    expect(images.length).toBe(1);
    expect(images[0].range.editAs).toBe('oneCell');
    expect(images[0].range.ext.width).toBe(100);
    expect(images[0].range.ext.height).toBe(100);
  });

  it('stores embedded image with hyperlinks', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    const imageId = wb.addImage({ buffer: IMAGE, extension: 'jpeg' });
    ws.addImage(imageId, {
      tl: { col: 0.1125, row: 0.4 },
      ext: { width: 100, height: 100 },
      editAs: 'absolute',
      hyperlinks: { hyperlink: 'http://www.somewhere.com', tooltip: 'www.somewhere.com' },
    });
    const ws2 = (await reload(wb)).getWorksheet('blort');
    const images = ws2.getImages();
    expect(images.length).toBe(1);
    // xlvane: a picture with a size is a one-cell anchor, which has no editAs in the schema, and
    // picture tooltips are not written (D-D03); exceljs reads back editAs "absolute" and the tooltip.
    expect(images[0].range.editAs).toBe('oneCell');
    expect(images[0].range.ext.width).toBe(100);
    expect(images[0].range.ext.height).toBe(100);
    expect(images[0].range.hyperlinks).toEqual({ hyperlink: 'http://www.somewhere.com' });
  });

  it('image extensions should not be case sensitive', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    const imageId1 = wb.addImage({ buffer: IMAGE, extension: 'PNG' });
    const imageId2 = wb.addImage({ buffer: IMAGE, extension: 'JPEG' });
    ws.addImage(imageId1, { tl: { col: 0.1125, row: 0.4 }, ext: { width: 100, height: 100 } });
    ws.addImage(imageId2, { tl: { col: 0.1125, row: 0.4 }, br: { col: 2.101046875, row: 3.4 }, editAs: 'oneCell' });
    const wb2 = await reload(wb);
    const images = wb2.getWorksheet('blort').getImages();
    expect(images.length).toBe(2);
    expect(images[0].range.ext.width).toBe(100);
    expect(images[1].range.editAs).toBe('oneCell');
    expect(same(wb2.getImage(images[0].imageId).buffer)).toBe(true);
    expect(same(wb2.getImage(images[1].imageId).buffer)).toBe(true);
  });

  for (const advance of [false, true]) {
    it(`WorkbookWriter with background image${advance ? ' where worksheet is commited in advance' : ''}`, async () => {
      await withTempDir(async (dir) => {
        const filename = join(dir, 'wb.test.xlsx');
        const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename });
        const ws = wb.addWorksheet('Hello');
        const imageId = wb.addImage({ buffer: IMAGE, extension: 'jpeg' });
        ws.getCell('A1').value = 'Hello, World!';
        ws.addBackgroundImage(imageId);
        if (advance) await ws.commit();
        await wb.commit();
        const wb2 = new ExcelJS.Workbook();
        await wb2.xlsx.readFile(filename);
        const ws2 = wb2.getWorksheet('Hello');
        const image = wb2.getImage(ws2.getBackgroundImageId());
        expect(same(image.buffer)).toBe(true);
      });
    });
  }
});

describe('Images (xlvane)', () => {
  it('writes pictures of a streamed worksheet and refuses images added after the first row', async () => {
    await withTempDir(async (dir) => {
      const filename = join(dir, 's.xlsx');
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename });
      const id = wb.addImage({ buffer: IMAGE });
      const ws = wb.addWorksheet('S');
      ws.addImage(id, 'B2:C4');
      ws.addRow([1]).commit();
      expect(() => wb.addImage({ buffer: IMAGE })).toThrow(expect.objectContaining({ code: 'XLV_E_NOT_SUPPORTED' }));
      ws.commit();
      await wb.commit();
      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.readFile(filename);
      expect(wb2.getWorksheet('S').getImages()[0].range.br.nativeCol).toBe(3);
    });
  });

  it('stores an image used twice once, with one relationship per drawing', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    const id = wb.addImage({ base64: `data:image/png;base64,${btoa(String.fromCharCode(...IMAGE))}`, extension: 'png' });
    ws.addImage(id, 'A1:B2');
    ws.addImage(wb.addImage({ buffer: IMAGE, extension: 'png' }), 'D1:E2');
    ws.addImage(id, 'G1:H2');
    const files = unzipSync(new Uint8Array(await wb.xlsx.writeBuffer()));
    expect(Object.keys(files).filter((n) => n.startsWith('xl/media/'))).toEqual(['xl/media/image1.png']);
    const rels = strFromU8(files['xl/drawings/_rels/drawing1.xml.rels']!);
    expect(rels.match(/relationships\/image"/g)?.length).toBe(1);
    expect(strFromU8(files['xl/drawings/drawing1.xml']!)).not.toContain('creationId');
  });

  it('refuses unknown image data and file names', async () => {
    const wb = new ExcelJS.Workbook();
    expect(() => wb.addImage({ buffer: new Uint8Array([1, 2, 3]), extension: 'png' })).toThrow(expect.objectContaining({ code: 'XLV_E_INVALID_IMAGE' }));
    expect(() => wb.addImage({ filename: 'x.png', extension: 'png' })).toThrow(expect.objectContaining({ code: 'XLV_E_NOT_SUPPORTED' }));
  });

  it('pictures move with spliced rows; a removed corner row is refused', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.addImage(wb.addImage({ buffer: IMAGE }), 'C3:E6');
    ws.insertRow(1, ['new']);
    expect(ws.getImages()[0].range.tl.nativeRow).toBe(3);
    expect(() => ws.spliceRows(4, 1)).toThrow(expect.objectContaining({ code: 'XLV_E_NOT_SUPPORTED' }));
  });
});

describe('Tables', () => {
  const sales = () => ({
    name: 'Sales',
    ref: 'B2',
    totalsRow: true,
    style: { theme: 'TableStyleLight9', showRowStripes: true },
    columns: [{ name: 'Region', totalsRowLabel: 'Total' }, { name: 'Revenue', totalsRowFunction: 'sum' }, { name: 'Units', totalsRowFunction: 'stdDev' }],
    rows: [['North', 1200, 3], ['South', 980, 5]],
  });

  it('writes the cells, the table part and reads the table back with ref and rows', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    const t = ws.addTable(sales());
    expect(t.model.tableRef).toBe('B2:D5');
    expect(t.model.autoFilterRef).toBe('B2:D4');
    expect(ws.getCell('B2').value).toBe('Region');
    expect(ws.getCell('C3').value).toBe(1200);
    expect(ws.getCell('B5').value).toBe('Total');
    // xlvane: STDEV is SUBTOTAL 107 ([MS-OI29500] §2.1.1094); exceljs writes 106 (PRODUCT).
    expect(ws.getCell('D5').value).toEqual({ formula: 'SUBTOTAL(107,Sales[Units])' });
    const bytes = await wb.xlsx.writeBuffer();
    const files = unzipSync(new Uint8Array(bytes));
    const xml = strFromU8(files['xl/tables/table1.xml']!);
    expect(xml).toContain('ref="B2:D5"');
    expect(xml).not.toContain('dxfId');
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(bytes);
    const ws2 = wb2.getWorksheet('S');
    const t2 = ws2.getTable('Sales');
    expect(ws2.getTables().length).toBe(1);
    expect(t2.model.tableRef).toBe('B2:D5');
    expect(t2.ref).toBe('B2');
    expect(t2.headerRow).toBe(true);
    expect(t2.totalsRow).toBe(true);
    expect(t2.theme).toBe('TableStyleLight9');
    expect(t2.model.rows).toEqual([['North', 1200, 3], ['South', 980, 5]]);
    // A loaded table can be changed and written again.
    t2.addRow(['West', 700, 1]);
    t2.commit();
    expect(ws2.getCell('C5').value).toBe(700);
    const files3 = unzipSync(new Uint8Array(await wb2.xlsx.writeBuffer()));
    expect(strFromU8(files3['xl/tables/table1.xml']!)).toContain('ref="B2:D6"');
  });

  it('commits several changes, columns and the display name', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    const t = ws.addTable({ name: 'T', ref: 'A1', columns: [{ name: 'A' }, { name: 'B' }], rows: [[1, 2], [3, 4], [5, 6]] });
    t.removeRows(0, 2);
    t.commit();
    expect(ws.getCell('A2').value).toBe(5);
    expect(ws.getCell('A3').value).toBeNull();
    t.addColumn({ name: 'C' }, [9], 1);
    t.commit();
    expect(ws.getCell('B1').value).toBe('C');
    expect(ws.getCell('C2').value).toBe(6);
    t.removeColumns(0, 1);
    t.commit();
    expect(ws.getCell('A1').value).toBe('C');
    expect(ws.getCell('C1').value).toBeNull();
    t.displayName = 'Shown';
    expect(t.displayName).toBe('Shown');
    t.getColumn(0).name = 'Renamed';
    t.commit();
    expect(ws.getCell('A1').value).toBe('Renamed');
  });

  it('applies column styles to the cells and drops the label of a first column with a function', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.addTable({ name: 'T', ref: 'A1', totalsRow: true, columns: [{ name: 'N', totalsRowFunction: 'count', style: { font: { bold: true } } }, { name: 'V' }], rows: [[1, 2]] });
    expect(ws.getCell('A2').font).toEqual({ bold: true });
    expect(ws.getCell('A3').value).toEqual({ formula: 'SUBTOTAL(103,T[N])' });
    const xml = strFromU8(unzipSync(new Uint8Array(await wb.xlsx.writeBuffer()))['xl/tables/table1.xml']!);
    expect(xml).toContain('totalsRowFunction="count"');
    expect(xml).not.toContain('totalsRowLabel="Total"');
  });

  it("enforces Excel's names and refuses splices through a table", () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    expect(() => ws.addTable({ name: 'My Table', ref: 'A1', columns: [{ name: 'A' }], rows: [[1]] })).toThrow(expect.objectContaining({ code: 'XLV_E_INVALID_TABLE' }));
    expect(() => ws.addTable({ name: 'T', ref: 'A1', columns: [{ name: 'A' }, { name: 'a' }], rows: [[1, 2]] })).toThrow(expect.objectContaining({ code: 'XLV_E_INVALID_TABLE' }));
    ws.addTable({ name: 'T', ref: 'A1', columns: [{ name: 'A' }], rows: [[1]] });
    expect(() => wb.addWorksheet('S2').addTable({ name: 't', ref: 'A1', columns: [{ name: 'A' }], rows: [[1]] })).toThrow(expect.objectContaining({ code: 'XLV_E_INVALID_TABLE' }));
    expect(() => ws.insertRow(1, ['x'])).toThrow(expect.objectContaining({ code: 'XLV_E_NOT_SUPPORTED' }));
    ws.insertRow(5, ['below']);
    ws.removeTable('T');
    expect(ws.getTables()).toEqual([]);
    expect(ws.getCell('A1').value).toBe('A');
  });

  it('a table without rows gets one empty data row in the file', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S').addTable({ name: 'T', ref: 'A1', columns: [{ name: 'A' }], rows: [] });
    const xml = strFromU8(unzipSync(new Uint8Array(await wb.xlsx.writeBuffer()))['xl/tables/table1.xml']!);
    expect(xml).toContain('ref="A1:A2"');
  });
});
