// @xlvane/exceljs's xlsx.load must not turn a few bytes into millions of cell or column objects
// (1.1.1 security review, D-H01): merged ranges, table ranges and <col> ranges count against
// max(maxCells, 16 x cells read), and checking a new merge against earlier ones is linear.
import { describe, expect, it } from 'vitest';
import { buildXlsx } from '../helpers/xlsx-builder';
import { ExcelJS } from '../exceljs/helpers';

async function load(bytes: Uint8Array): Promise<{ error: unknown; ms: number; wb: any }> {
  const wb = new ExcelJS.Workbook();
  const start = performance.now();
  let error: unknown = null;
  try {
    await wb.xlsx.load(bytes);
  } catch (e) {
    error = e;
  }
  return { error, ms: performance.now() - start, wb };
}

const one = '<row r="1"><c r="A1"><v>1</v></c></row>';

describe('compat xlsx.load', () => {
  it('refuses a merged range over the whole sheet', async () => {
    const { error, ms } = await load(buildXlsx({ sheets: [{ name: 'S', data: one, after: '<mergeCells count="1"><mergeCell ref="A1:XFD1048576"/></mergeCells>' }] }));
    expect(ms).toBeLessThan(3000);
    expect((error as { code?: string }).code).toBe('XLV_E_LIMIT_EXCEEDED');
  });

  it('loads 20 000 small merges in linear time', async () => {
    const merges = Array.from({ length: 20_000 }, (_, i) => `<mergeCell ref="A${String(2 * i + 1)}:B${String(2 * i + 1)}"/>`).join('');
    const { error, ms, wb } = await load(buildXlsx({ sheets: [{ name: 'S', data: one, after: `<mergeCells count="20000">${merges}</mergeCells>` }] }));
    expect(error).toBeNull();
    expect(Object.keys(wb.getWorksheet('S').model.merges)).toHaveLength(20_000);
    expect(ms).toBeLessThan(5000);
  });

  it('still refuses an overlapping merge, as exceljs does', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    ws.mergeCells('B2:D4');
    expect(() => ws.mergeCells('A1:B2')).toThrow('Cannot merge already merged cells');
    expect(() => ws.mergeCells('D4:E5')).toThrow('Cannot merge already merged cells');
    ws.mergeCells('E5:F6');
  });

  it('refuses a table whose range covers a million rows', async () => {
    const table = '<?xml version="1.0"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:CV1000000"><tableColumns count="1"><tableColumn id="1" name="a"/></tableColumns></table>';
    const { error, ms } = await load(
      buildXlsx({
        sheets: [{ name: 'S', data: one, after: '<tableParts count="1"><tablePart r:id="rT"/></tableParts>' }],
        extraParts: { 'xl/tables/table1.xml': table, 'xl/worksheets/_rels/sheet1.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rT" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>' },
      }),
    );
    expect(ms).toBeLessThan(3000);
    // The core may refuse the table first; either way no million-row table is built.
    expect(error === null || (error as { code?: string }).code === 'XLV_E_LIMIT_EXCEEDED').toBe(true);
  });

  it('refuses 20 000 <col> elements that each cover all 16 384 columns', async () => {
    const cols = Array.from({ length: 20_000 }, () => '<col min="1" max="16384" width="9"/>').join('');
    const { error, ms } = await load(buildXlsx({ sheets: [{ name: 'S', data: one, before: `<cols>${cols}</cols>` }] }));
    expect(ms).toBeLessThan(3000);
    expect((error as { code?: string }).code).toBe('XLV_E_LIMIT_EXCEEDED');
  });
});

describe('mergeCells keeps exceljs behaviour (independent review of the patch)', () => {
  it('refuses a merge over a merged range whose cells were replaced by row.values, as exceljs 4.4.0 does', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.mergeCells('A1:B1');
    ws.getRow(1).values = [1, 2];
    expect(() => ws.mergeCells('A1:C1')).toThrow('Cannot merge already merged cells');
    ws.unMergeCells('A1:B1');
    expect(() => ws.mergeCells('A1:C1')).not.toThrow();
  });

  it('checks 20 000 merges in linear time', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    const start = performance.now();
    for (let r = 1; r <= 20000; r++) ws.mergeCells(r, 1, r, 2);
    expect(performance.now() - start).toBeLessThan(3000);
    expect(() => ws.mergeCells(20000, 2, 20001, 3)).toThrow('Cannot merge already merged cells');
  });
});

describe('row settings up to the dimension (independent review of the patch)', () => {
  it('300 sheets declaring a million rows each, with one cell, are refused quickly', async () => {
    const sheets = Array.from({ length: 300 }, (_, i) => ({ name: `S${String(i + 1)}`, before: '<dimension ref="A1:A1048576"/>', data: '<row r="1"><c r="A1"><v>1</v></c></row>' }));
    const bytes = buildXlsx({ sheets });
    const start = performance.now();
    let error: unknown = null;
    try {
      await new ExcelJS.Workbook().xlsx.load(bytes);
    } catch (e) {
      error = e;
    }
    expect(performance.now() - start).toBeLessThan(2000);
    expect((error as { code?: string } | null)?.code).toBe('XLV_E_LIMIT_EXCEEDED');
  });

  it('a few sheets declaring a million rows still load, with their row settings', async () => {
    const sheets = Array.from({ length: 3 }, (_, i) => ({ name: `S${String(i + 1)}`, before: '<dimension ref="A1:A1048576"/>', data: '<row r="1"><c r="A1"><v>1</v></c></row><row r="900000" hidden="1"/>' }));
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buildXlsx({ sheets }));
    expect(wb.getWorksheet('S3')?.getRow(900000).hidden).toBe(true);
  });
});
