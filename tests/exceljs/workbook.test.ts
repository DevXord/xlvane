// Ported from exceljs v4.4.0 spec/unit/doc/workbook.spec.js (MIT, Copyright (c) 2014-2019 Guyon Roche),
// changes: Mocha/Chai → Vitest; `verquire('exceljs')` → @xlvane/exceljs; data validation
// removed from the helper workbook (not supported); the
// model-serialisation test stays skipped as upstream; the `duplicateRow` tests are ported in
// full (phase B); the rich-text fonts are written with spread objects (same values).
import { describe, expect, it } from 'vitest';
import { ExcelJS, expectNotSupported, styles } from './helpers';

function createSimpleWorkbook() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('blort');
  ws.getCell('A1').value = 7;
  ws.getCell('A1').name = 'Seven';
  ws.getCell('B1').value = 'Hello, World!';
  ws.getCell('B1').name = 'Hello';
  ws.getCell('C1').value = 3.14;
  ws.getCell('D1').value = new Date();
  ws.getCell('E1').value = { text: 'www.google.com', hyperlink: 'http://www.google.com' };
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
    expect(ws.getCell('E1').type).toBe(ExcelJS.ValueType.Hyperlink);
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
    const twoRows = () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      ws.getCell('A1').value = '1.1';
      ws.getCell('A1').font = styles.fonts.arialBlackUI14;
      ws.getCell('B1').value = '1.2';
      ws.getCell('B1').font = styles.fonts.comicSansUdB16;
      ws.getCell('C1').value = '1.3';
      ws.getCell('C1').fill = styles.fills.redDarkVertical;
      ws.getRow(1).numFmt = styles.numFmts.numFmt1;
      ws.getCell('A2').value = '2.1';
      ws.getCell('A2').alignment = styles.namedAlignments.topLeft;
      ws.getCell('B2').value = '2.2';
      ws.getCell('B2').alignment = styles.namedAlignments.middleCentre;
      ws.getCell('C2').value = '2.3';
      ws.getCell('C2').alignment = styles.namedAlignments.bottomRight;
      ws.getRow(2).numFmt = styles.numFmts.numFmt2;
      return ws;
    };

    it('inserts duplicates', () => {
      const ws = twoRows();
      ws.duplicateRow(1, 2, true);
      expect(ws.getRow(1).values).toEqual([, '1.1', '1.2', '1.3']);
      expect(ws.getRow(2).values).toEqual([, '1.1', '1.2', '1.3']);
      expect(ws.getRow(3).values).toEqual([, '1.1', '1.2', '1.3']);
      expect(ws.getRow(4).values).toEqual([, '2.1', '2.2', '2.3']);
      for (let i = 1; i <= 3; i++) {
        expect(ws.getCell(`A${String(i)}`).font).toEqual(styles.fonts.arialBlackUI14);
        expect(ws.getCell(`B${String(i)}`).font).toEqual(styles.fonts.comicSansUdB16);
        expect(ws.getCell(`C${String(i)}`).fill).toEqual(styles.fills.redDarkVertical);
      }
      expect(ws.getCell('A4').alignment).toEqual(styles.namedAlignments.topLeft);
      expect(ws.getCell('B4').alignment).toEqual(styles.namedAlignments.middleCentre);
      expect(ws.getCell('C4').alignment).toEqual(styles.namedAlignments.bottomRight);
      expect(ws.getRow(1).numFmt).toBe(styles.numFmts.numFmt1);
      expect(ws.getRow(2).numFmt).toBe(styles.numFmts.numFmt1);
      expect(ws.getRow(3).numFmt).toBe(styles.numFmts.numFmt1);
      expect(ws.getRow(4).numFmt).toBe(styles.numFmts.numFmt2);
    });

    it('overwrites with duplicates', () => {
      const ws = twoRows();
      ws.getCell('A3').value = '3.1';
      ws.getCell('A3').fill = styles.fills.redGreenDarkTrellis;
      ws.getCell('B3').value = '3.2';
      ws.getCell('B3').fill = styles.fills.blueWhiteHGrad;
      ws.getCell('C3').value = '3.3';
      ws.getCell('C3').fill = styles.fills.rgbPathGrad;
      ws.getRow(3).font = styles.fonts.broadwayRedOutline20;
      ws.duplicateRow(1, 1, false);
      expect(ws.getRow(1).values).toEqual([, '1.1', '1.2', '1.3']);
      expect(ws.getRow(2).values).toEqual([, '1.1', '1.2', '1.3']);
      expect(ws.getRow(3).values).toEqual([, '3.1', '3.2', '3.3']);
      for (let i = 1; i <= 2; i++) {
        expect(ws.getCell(`A${String(i)}`).font).toEqual(styles.fonts.arialBlackUI14);
        expect(ws.getCell(`A${String(i)}`).alignment).toBeUndefined();
        expect(ws.getCell(`B${String(i)}`).font).toEqual(styles.fonts.comicSansUdB16);
        expect(ws.getCell(`B${String(i)}`).alignment).toBeUndefined();
        expect(ws.getCell(`C${String(i)}`).fill).toEqual(styles.fills.redDarkVertical);
        expect(ws.getCell(`C${String(i)}`).alignment).toBeUndefined();
      }
      expect(ws.getRow(1).numFmt).toBe(styles.numFmts.numFmt1);
      expect(ws.getRow(2).numFmt).toBe(styles.numFmts.numFmt1);
      expect(ws.getRow(3).numFmt).toBeUndefined();
      expect(ws.getRow(3).font).toEqual(styles.fonts.broadwayRedOutline20);
    });

    // xlvane additions (phase B, D-B03).
    it('copies styles instead of sharing them, and copies merged ranges inside the row', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      ws.getCell('A1').value = 'a';
      ws.getCell('A1').font = { bold: true };
      ws.mergeCells('A1:B1');
      ws.duplicateRow(1, 1, true);
      ws.getCell('A2').font.italic = true;
      expect(ws.getCell('A1').font).toEqual({ bold: true });
      expect(ws.getCell('B2').master.address).toBe('A2');
      expect(ws.getCell('A2').value).toBe('a');
    });

    it('refuses arguments exceljs would corrupt data with', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      ws.getCell('A1').value = 1;
      for (const call of [() => ws.duplicateRow(9, 1, true), () => ws.duplicateRow(1, -1, true), () => ws.duplicateRow(1), () => ws.insertRow(0, [1]), () => ws.insertRow('2', [1]), () => ws.insertRows(2), () => ws.spliceRows(2, -1), () => ws.spliceColumns(0, 1)]) {
        expect(call).toThrow(expect.objectContaining({ code: 'XLV_E_INVALID_ARGUMENT' }));
      }
      expect(ws.getCell('A1').value).toBe(1);
      expect(ws.rowCount).toBe(1);
    });
  });
});
