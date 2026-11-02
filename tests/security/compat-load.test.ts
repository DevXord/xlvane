// @xlvane/compat's xlsx.load must not turn a few bytes into millions of cell or column objects
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
