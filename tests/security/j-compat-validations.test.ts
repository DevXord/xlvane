// Data validations read by @xlvane/compat's xlsx.load are given to every cell they cover, as
// exceljs does, up to 100 000 cells. The budget is shared by the whole workbook, so 400 sheets
// each covering 98 304 cells cannot create 39 million entries (security review J, C2).
import { describe, expect, it } from 'vitest';
import { buildXlsx } from '../helpers/xlsx-builder';
import { ExcelJS } from '../exceljs/helpers';

const one = '<row r="1"><c r="A1"><v>1</v></c></row>';
const validation = '<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" sqref="A1:XFD6"><formula1>0</formula1></dataValidation></dataValidations>';

describe('data validations read by xlsx.load', () => {
  it('expand at most 100 000 cells over all sheets of a workbook', async () => {
    const sheets = Array.from({ length: 400 }, (_, i) => ({ name: `S${String(i + 1)}`, data: one, after: validation }));
    const bytes = buildXlsx({ sheets });
    const wb = new ExcelJS.Workbook();
    const start = performance.now();
    await wb.xlsx.load(bytes);
    expect(performance.now() - start).toBeLessThan(10_000);
    let keys = 0;
    for (const ws of wb.worksheets) keys += Object.keys(ws.dataValidations.model).length;
    expect(keys).toBeLessThanOrEqual(100_000 + 400);
    // The first sheet's validation is given to each cell; later ones keep their range, and every
    // cell still finds its validation.
    expect(Object.keys(wb.getWorksheet('S1').dataValidations.model)).toHaveLength(6 * 16_384);
    expect(Object.keys(wb.getWorksheet('S400').dataValidations.model)).toEqual(['A1:XFD6']);
    for (const name of ['S1', 'S2', 'S400']) {
      expect(wb.getWorksheet(name).getCell('XFD6').dataValidation).toEqual({ type: 'whole', operator: 'greaterThan', formulae: [0] });
      expect(wb.getWorksheet(name).getCell('A7').dataValidation).toBeUndefined();
    }
  });

  it('still expand the validations of a few sheets', async () => {
    const small = '<dataValidations count="1"><dataValidation type="list" sqref="B2:B4"><formula1>"a,b"</formula1></dataValidation></dataValidations>';
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buildXlsx({ sheets: [1, 2, 3].map((i) => ({ name: `S${String(i)}`, data: one, after: small })) }));
    for (const ws of wb.worksheets) expect(Object.keys(ws.dataValidations.model)).toEqual(['B2', 'B3', 'B4']);
  });
});
