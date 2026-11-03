// Many worksheets: the duplicate check of a sheet name, the next id and the next orderNo of
// @xlvane/compat take constant time, so 16 000 sheets (a 72 kB file) load in linear time
// (security review J, C3). Sheet name and merge errors of a loaded file are XlvErrors with a
// stable code; the public addWorksheet and mergeCells keep exceljs's plain Errors.
import { describe, expect, it } from 'vitest';
import { XlvError } from '../../packages/xlvane/src/errors';
import { buildXlsx } from '../helpers/xlsx-builder';
import { ExcelJS } from '../exceljs/helpers';

const one = '<row r="1"><c r="A1"><v>1</v></c></row>';
const N = 16_000;

/** A workbook listing `n` sheets that all lead to the same worksheet part. */
function manySheets(n: number): Uint8Array {
  const sheets = Array.from({ length: n }, (_, i) => (i === 0 ? { name: 'S1', data: one } : { name: `S${String(i + 1)}`, missingPart: true }));
  const rels = Array.from({ length: n }, (_, i) => `<Relationship Id="rId${String(i + 1)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>`).join('');
  return buildXlsx({
    sheets,
    extraParts: { 'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>` },
  });
}

async function loadError(bytes: Uint8Array): Promise<any> {
  try {
    await new ExcelJS.Workbook().xlsx.load(bytes);
  } catch (e) {
    return e;
  }
  return null;
}

async function streamError(bytes: Uint8Array): Promise<any> {
  try {
    for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(bytes, {})) for await (const row of ws) void row;
  } catch (e) {
    return e;
  }
  return null;
}

describe('many worksheets', () => {
  it(`xlsx.load reads ${String(N)} sheets in a few seconds`, async () => {
    const bytes = manySheets(N);
    const wb = new ExcelJS.Workbook();
    const start = performance.now();
    await wb.xlsx.load(bytes);
    expect(performance.now() - start).toBeLessThan(5000);
    expect(wb.worksheets).toHaveLength(N);
    expect(wb.getWorksheet('S1').getCell('A1').value).toBe(1);
    expect(wb.getWorksheet(`S${String(N)}`).orderNo).toBe(N - 1);
  });

  it(`stream.xlsx.WorkbookReader reads ${String(N)} sheets in a few seconds`, async () => {
    const bytes = manySheets(N);
    const start = performance.now();
    let sheets = 0;
    for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(bytes, {})) {
      void ws;
      sheets++;
    }
    expect(performance.now() - start).toBeLessThan(5000);
    expect(sheets).toBe(N);
  });

  it(`addWorksheet ${String(N)} times takes a few seconds`, () => {
    const wb = new ExcelJS.Workbook();
    const start = performance.now();
    for (let i = 1; i <= N; i++) wb.addWorksheet(`Sheet ${String(i)}`);
    expect(performance.now() - start).toBeLessThan(3000);
    expect(wb.worksheets).toHaveLength(N);
    expect(wb.getWorksheet(N).orderNo).toBe(N);
    expect(wb.nextId).toBe(N + 1);
  });

  it('keeps exceljs duplicate checks, ids and orderNo through renames, removals and orderNo changes', () => {
    const wb = new ExcelJS.Workbook();
    const a = wb.addWorksheet('Alpha');
    const b = wb.addWorksheet('Beta');
    expect(() => wb.addWorksheet('ALPHA')).toThrow(new Error('Worksheet name already exists: ALPHA'));
    expect(() => wb.addWorksheet('beta')).toThrow('Worksheet name already exists: beta');
    a.name = 'Gamma';
    expect(wb.addWorksheet('alpha').id).toBe(3);
    expect(() => wb.addWorksheet('GAMMA')).toThrow('Worksheet name already exists: GAMMA');
    expect(wb.getWorksheet('Gamma')).toBe(a);
    expect(wb.getWorksheet('gamma')).toBeUndefined();
    a.name = 'GAMMA';
    expect(a.name).toBe('GAMMA');
    expect(() => (b.name = 'gamma')).toThrow('Worksheet name already exists: gamma');
    // Truncated to 31 characters before the check, as exceljs.
    wb.addWorksheet('x'.repeat(40));
    expect(() => wb.addWorksheet('X'.repeat(35))).toThrow(`Worksheet name already exists: ${'X'.repeat(31)}`);
    wb.removeWorksheet(b.id);
    expect(wb.getWorksheet('Beta')).toBeUndefined();
    const b2 = wb.addWorksheet('Beta');
    expect(b2.id).toBe(2);
    expect(b2.orderNo).toBe(5);
    // A removed sheet's name is free, and renaming the removed sheet changes nothing.
    b.name = 'Delta';
    wb.addWorksheet('Delta');
    // orderNo: one more than the largest, also after it was changed by hand or its sheet removed.
    b2.orderNo = 50;
    expect(wb.addWorksheet('E').orderNo).toBe(51);
    b2.orderNo = 1;
    expect(wb.addWorksheet('F').orderNo).toBe(52);
    wb.removeWorksheet('F');
    wb.removeWorksheet('E');
    expect(wb.addWorksheet('G').orderNo).toBe(7);
  });

  it('keeps the duplicate check of the streaming writer', () => {
    const writer = new ExcelJS.stream.xlsx.WorkbookWriter({});
    writer.addWorksheet('Data');
    expect(() => writer.addWorksheet('DATA')).toThrow('Worksheet name already exists: DATA');
  });
});

describe('sheet name errors of a loaded file', () => {
  const cases: [string, string[], string, string][] = [
    ['an empty name', ['A', ''], 'XLV_E_INVALID_SHEET_NAME', "The name can't be empty."],
    ['the protected name History', ['History'], 'XLV_E_INVALID_SHEET_NAME', 'The name "History" is protected. Please use a different name.'],
    ['a forbidden character', ['a*b'], 'XLV_E_INVALID_SHEET_NAME', 'Worksheet name a*b cannot include any of the following characters: * ? : \\ / [ ]'],
    ['a leading apostrophe', ["'quoted"], 'XLV_E_INVALID_SHEET_NAME', "The first or last character of worksheet name cannot be a single quotation mark: 'quoted"],
    ['a duplicate ignoring case', ['Data', 'DATA'], 'XLV_E_DUPLICATE_SHEET_NAME', 'Worksheet name already exists: DATA'],
    ['a duplicate after truncation to 31 characters', ['n'.repeat(31), `${'n'.repeat(31)}more`], 'XLV_E_DUPLICATE_SHEET_NAME', `Worksheet name already exists: ${'n'.repeat(31)}`],
  ];
  for (const [what, names, code, message] of cases) {
    it(`gives ${code} for ${what}`, async () => {
      const bytes = buildXlsx({ sheets: names.map((name) => ({ name, data: one })) });
      const bad = names[names.length - 1]!;
      for (const error of [await loadError(bytes), await streamError(bytes)]) {
        expect(error).toBeInstanceOf(XlvError);
        expect(error.code).toBe(code);
        expect(error.message).toContain(message);
        expect(error.location?.sheet).toBe(bad);
        expect(typeof error.hint).toBe('string');
      }
    });
  }

  it('gives XLV_E_OVERLAPPING_MERGE for overlapping merged ranges', async () => {
    const error = await loadError(buildXlsx({ sheets: [{ name: 'S', data: one, after: '<mergeCells count="2"><mergeCell ref="A1:B2"/><mergeCell ref="B2:C3"/></mergeCells>' }] }));
    expect(error.code).toBe('XLV_E_OVERLAPPING_MERGE');
    expect(error.message).toContain('Cannot merge already merged cells');
    expect(error.location).toMatchObject({ sheet: 'S', cell: 'B2:C3' });
    expect(error.details?.other).toBe('A1:B2');
    expect(typeof error.hint).toBe('string');
  });

  it('keeps plain exceljs Errors for addWorksheet and mergeCells', () => {
    const wb = new ExcelJS.Workbook();
    for (const name of ['History', 'a*b', "'q"]) {
      let error: any = null;
      try {
        wb.addWorksheet(name);
      } catch (e) {
        error = e;
      }
      expect(error).toBeInstanceOf(Error);
      expect(error).not.toBeInstanceOf(XlvError);
    }
    const ws = wb.addWorksheet('S');
    ws.mergeCells('A1:B2');
    let error: any = null;
    try {
      ws.mergeCells('B2:C3');
    } catch (e) {
      error = e;
    }
    expect(error.message).toBe('Cannot merge already merged cells');
    expect(error.code).toBeUndefined();
  });
});
