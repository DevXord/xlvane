// Ported from exceljs v4.4.0 spec/unit/doc/workbook.spec.js (MIT, Copyright (c) 2014-2019 Guyon Roche),
// changes: Mocha/Chai → Vitest; `verquire('exceljs')` → @xlvane/exceljs; data validation
// removed from the helper workbook (not supported); the
// model-serialisation test stays skipped as upstream; `duplicateRow` tests expect
// XLV_E_NOT_SUPPORTED; hyperlink values throw XLV_E_NOT_SUPPORTED when assigned, so they are
// removed from the helper workbook and checked separately; the rich-text fonts are written with
// spread objects (same values).
import { describe, expect, it } from 'vitest';
import { ExcelJS, expectNotSupported } from './helpers';

function createSimpleWorkbook() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('blort');
  ws.getCell('A1').value = 7;
  ws.getCell('A1').name = 'Seven';
  ws.getCell('B1').value = 'Hello, World!';
  ws.getCell('B1').name = 'Hello';
  ws.getCell('C1').value = 3.14;
  ws.getCell('D1').value = new Date();
  ws.getCell('A2').value = { formula: 'A1', result: 7 };
  ws.getCell('A2').name = 'TheFormula';
  ws.getCell('B2').value = { formula: 'CONCATENATE("Hello", ", ", "World!")', result: 'Hello, World!' };
  ws.getCell('B2').name = 'TheFormula';
  ws.getCell('C2').value = { formula: 'D1', result: new Date() };
  ws.getCell('C3').value = { formula: 'D1' };
  return wb;
}

describe('Workbook', () => {
  it('stores shared string values properly', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('A2').value = 'Hello';
    ws.getCell('B2').value = 'World';
    ws.getCell('C2').value = { formula: 'CONCATENATE(A2, ", ", B2, "!")', result: 'Hello, World!' };
    ws.getCell('A3').value = `${['Hello', 'World'].join(', ')}!`;
    expect(ws.getCell('A1').value).toBe(ws.getCell('A3').value);
    expect(ws.getCell('A1').value).toBe(ws.getCell('C2').value.result);
  });

  it('assigns cell types properly', () => {
    const ws = createSimpleWorkbook().getWorksheet('blort');
    expect(ws.getCell('A1').type).toBe(ExcelJS.ValueType.Number);
    expect(ws.getCell('B1').type).toBe(ExcelJS.ValueType.String);
    expect(ws.getCell('C1').type).toBe(ExcelJS.ValueType.Number);
    expect(ws.getCell('D1').type).toBe(ExcelJS.ValueType.Date);
    expect(ws.getCell('A2').type).toBe(ExcelJS.ValueType.Formula);
    expect(ws.getCell('B2').type).toBe(ExcelJS.ValueType.Formula);
    expect(ws.getCell('C2').type).toBe(ExcelJS.ValueType.Formula);
  });

  it('assigns rich text', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    const base = { size: 12, name: 'Calibri', scheme: 'minor' };
    ws.getCell('A1').value = {
      richText: [
        { font: { ...base, color: { theme: 0 }, family: 2 }, text: 'This is ' },
        { font: { ...base, italic: true, color: { theme: 0 } }, text: 'a' },
        { font: { ...base, color: { theme: 1 }, family: 2 }, text: ' ' },
        { font: { ...base, color: { argb: 'FFFF6600' } }, text: 'colorful' },
        { font: { ...base, color: { theme: 1 }, family: 2 }, text: ' text ' },
        { font: { ...base, color: { argb: 'FFCCFFCC' } }, text: 'with' },
        { font: { ...base, color: { theme: 1 }, family: 2 }, text: ' in-cell ' },
        { font: { ...base, bold: true, color: { theme: 1 }, family: 2 }, text: 'format' },
      ],
    };
    expect(ws.getCell('A1').text).toBe('This is a colorful text with in-cell format');
    expect(ws.getCell('A1').type).toBe(ExcelJS.ValueType.RichText);
  });

  it('assigns hyperlinks — not supported: hyperlink values', async () => {
    const ws = new ExcelJS.Workbook().addWorksheet('blort');
    await expectNotSupported(() => {
      ws.getCell('E1').value = { text: 'www.google.com', hyperlink: 'http://www.google.com' };
    });
  });

  it.skip('serialises to model (skipped upstream as well)', () => {});

  it('returns undefined for non-existant sheet', () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('first');
    expect(wb.getWorksheet('w00t')).toBeUndefined();
  });

  it('returns undefined for sheet 0', () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('first');
    expect(wb.getWorksheet(0)).toBeUndefined();
  });

  it('returns undefined for sheet 0 after accessing wb.worksheets or wb.eachSheet ', () => {
    const wb = new ExcelJS.Workbook();
    const sheet = wb.addWorksheet('first');
    wb.eachSheet(() => {});
    expect(wb.worksheets.length).toBe(1);
    expect(wb.getWorksheet(0)).toBeUndefined();
    expect(wb.getWorksheet(1) === sheet).toBe(true);
  });

  describe('duplicateRows', () => {
    it('inserts duplicates — not supported: duplicateRow', async () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      ws.getCell('A1').value = '1.1';
      await expectNotSupported(() => ws.duplicateRow(1, 2, true));
    });

    it('overwrites with duplicates — not supported: duplicateRow', async () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      ws.getCell('A1').value = '1.1';
      await expectNotSupported(() => ws.duplicateRow(1, 1, false));
    });
  });
});
