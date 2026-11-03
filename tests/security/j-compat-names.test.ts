// @xlvane/compat keeps a defined name as one object per cell, as exceljs does (cell-matrix.js).
// A name read from a file counts its cells against the same budget as merged ranges and tables
// (D-H01), so a 1.6 kB file naming the whole sheet cannot exhaust memory; a range with rows past
// the sheet is dropped on load, and refused by definedNames.add (security review J, C1).
import { describe, expect, it } from 'vitest';
import { buildXlsx } from '../helpers/xlsx-builder';
import { ExcelJS } from '../exceljs/helpers';

const one = '<row r="1"><c r="A1"><v>1</v></c></row>';

function namedFile(refersTo: string, name = 'Big'): Uint8Array {
  return buildXlsx({ sheets: [{ name: 'S', data: one }], workbookAfterSheets: `<definedNames><definedName name="${name}">${refersTo}</definedName></definedNames>` });
}

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

describe('defined names read by xlsx.load', () => {
  it('refuses a name over the whole sheet quickly', async () => {
    const bytes = namedFile('S!$A$1:$XFD$1048576');
    expect(bytes.length).toBeLessThan(2000);
    const { error, ms } = await load(bytes);
    expect(ms).toBeLessThan(3000);
    expect((error as { code?: string }).code).toBe('XLV_E_LIMIT_EXCEEDED');
    expect((error as { details?: { limit?: string } }).details?.limit).toBe('maxCells');
  });

  it('counts the cells of all names together', async () => {
    const names = Array.from({ length: 4 }, (_, i) => `<definedName name="N${String(i)}">S!$A$1:$XFD$100</definedName>`).join('');
    const { error, ms } = await load(buildXlsx({ sheets: [{ name: 'S', data: one }], workbookAfterSheets: `<definedNames>${names}</definedNames>` }));
    expect(ms).toBeLessThan(5000);
    expect((error as { code?: string }).code).toBe('XLV_E_LIMIT_EXCEEDED');
  });

  it('drops a range whose rows lie past the sheet, as it drops other invalid ranges', async () => {
    const { error, ms, wb } = await load(namedFile('S!$A$1:$A$99999999999,S!$B$2'));
    expect(error).toBeNull();
    expect(ms).toBeLessThan(3000);
    expect(wb.definedNames.model).toEqual([{ name: 'Big', ranges: ['S!$B$2'] }]);
  });

  it('round-trips a normal named range', async () => {
    const { error, wb } = await load(namedFile('S!$A$1:$C$3', 'Block'));
    expect(error).toBeNull();
    expect(wb.definedNames.model).toEqual([{ name: 'Block', ranges: ['S!$A$1:$C$3'] }]);
    const again = new ExcelJS.Workbook();
    await again.xlsx.load(await wb.xlsx.writeBuffer());
    expect(again.definedNames.model).toEqual([{ name: 'Block', ranges: ['S!$A$1:$C$3'] }]);
  });
});

describe('definedNames.add and addEx', () => {
  it('refuse a range with rows past the sheet instead of looping over them', () => {
    const wb = new ExcelJS.Workbook();
    const start = performance.now();
    let error: unknown = null;
    try {
      wb.definedNames.add('S!$A$1:$A$99999999999', 'huge');
    } catch (e) {
      error = e;
    }
    expect(performance.now() - start).toBeLessThan(1000);
    expect((error as { code?: string }).code).toBe('XLV_E_INVALID_CELL_REF');
    expect(() => wb.definedNames.addEx({ sheetName: 'S', top: 1, left: 1, bottom: 2_000_000, right: 1 }, 'huge')).toThrow(expect.objectContaining({ code: 'XLV_E_INVALID_CELL_REF' }));
    expect(() => wb.definedNames.getMatrix('m').addCell('S!$A$1:$A$99999999999')).toThrow(expect.objectContaining({ code: 'XLV_E_INVALID_CELL_REF' }));
  });

  it('still add ranges inside the sheet, also the last row', () => {
    const wb = new ExcelJS.Workbook();
    wb.definedNames.add('S!$A$1048575:$B$1048576', 'bottom');
    expect(wb.definedNames.getRanges('bottom')).toEqual({ name: 'bottom', ranges: ['S!$A$1048575:$B$1048576'] });
  });
});
