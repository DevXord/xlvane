// worksheet.id is written as sheet@sheetId and read back from it, as exceljs 4.4.0 does
// (lib/xlsx/xform/book/sheet-xform.js L7, L18; workbook-xform.js reconcile L161-186;
// lib/doc/workbook.js model setter L201-212; D-F03).
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { buildXlsx } from '../helpers/xlsx-builder';
import { ExcelJS, withTempDir } from './helpers';

async function reload(wb: any): Promise<any> {
  const back = new ExcelJS.Workbook();
  await back.xlsx.load(await wb.xlsx.writeBuffer());
  return back;
}

const sheetsXml = (bytes: Uint8Array): string => strFromU8(unzipSync(bytes)['xl/workbook.xml']!).replace(/^[\s\S]*<sheets>|<\/sheets>[\s\S]*$/g, '');

describe('worksheet ids', () => {
  it('survive removing a sheet, writing and loading, so getWorksheet(id) finds the same sheet', async () => {
    const wb = new ExcelJS.Workbook();
    const a = wb.addWorksheet('A');
    wb.addWorksheet('B');
    wb.addWorksheet('C');
    wb.removeWorksheet(a.id);
    const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
    expect(sheetsXml(bytes)).toBe('<sheet name="B" sheetId="2" r:id="rId1"/><sheet name="C" sheetId="3" r:id="rId2"/>');
    const back = await reload(wb);
    expect(back.worksheets.map((w: any) => [w.name, w.id])).toEqual([
      ['B', 2],
      ['C', 3],
    ]);
    expect(back.getWorksheet(3).name).toBe('C');
    expect(back.getWorksheet(1)).toBeUndefined();
    // The next sheet takes the free id 1, as in exceljs (nextId).
    expect(back.addWorksheet('D').id).toBe(1);
  });

  it('are read from a file that does not number its sheets 1, 2, 3', async () => {
    const bytes = buildXlsx({ sheets: [{ name: 'First', sheetId: '7' }, { name: 'Second', sheetId: '3' }] });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(bytes);
    expect(wb.worksheets.map((w: any) => [w.name, w.id])).toEqual([
      ['First', 7],
      ['Second', 3],
    ]);
    expect(wb.getWorksheet(7).name).toBe('First');
  });

  it('a missing or repeated id in a file takes the lowest free one', async () => {
    const bytes = buildXlsx({ sheets: [{ name: 'A', sheetId: '2' }, { name: 'B', sheetId: '2' }, { name: 'C', sheetId: null }] });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(bytes);
    expect(wb.worksheets.map((w: any) => [w.name, w.id])).toEqual([
      ['A', 2],
      ['B', 1],
      ['C', 3],
    ]);
  });

  it('above 65534 (from an untrusted file) take the lowest free id, without losing a sheet or slowing down', async () => {
    const bytes = buildXlsx({ sheets: [{ name: 'A', sheetId: '1' }, { name: 'B', sheetId: '4294967295' }, { name: 'C', sheetId: '20000000' }] });
    const wb = new ExcelJS.Workbook();
    const t0 = performance.now();
    await wb.xlsx.load(bytes);
    expect(wb.worksheets.map((w: any) => [w.name, w.id])).toEqual([
      ['A', 1],
      ['B', 2],
      ['C', 3],
    ]);
    const back = await reload(wb);
    expect(back.worksheets.map((w: any) => w.name)).toEqual(['A', 'B', 'C']);
    expect(performance.now() - t0).toBeLessThan(5000);
    const copy = new ExcelJS.Workbook();
    copy.model = { ...wb.model, worksheets: wb.model.worksheets.map((m: any, i: number) => ({ ...m, id: i === 0 ? 70000 : m.id })) };
    expect(copy.worksheets.map((w: any) => [w.name, w.id])).toEqual([
      ['A', 1],
      ['B', 2],
      ['C', 3],
    ]);
  });

  it('are kept by the model setter (lib/doc/workbook.js L201-212)', () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('A');
    wb.addWorksheet('B');
    wb.removeWorksheet(1);
    const copy = new ExcelJS.Workbook();
    copy.model = wb.model;
    expect(copy.worksheets.map((w: any) => [w.name, w.id])).toEqual([['B', 2]]);
  });

  it('are written by the streaming WorkbookWriter', async () => {
    await withTempDir(async (dir) => {
      const filename = join(dir, 'stream.xlsx');
      const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ filename });
      writer.addWorksheet('A').addRow([1]).commit();
      writer.addWorksheet('B').addRow([2]).commit();
      await writer.commit();
      expect(sheetsXml(new Uint8Array(await readFile(filename)))).toBe('<sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2"/>');
    });
  });
});
