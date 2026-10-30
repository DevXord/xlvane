// Ported from exceljs v4.4.0 spec/unit/doc/worksheet.merge.spec.js (MIT, Copyright (c) 2014-2019
// Guyon Roche), changes: Mocha/Chai → Vitest; `Dimensions` (internal Range class) replaced by a
// small loop over the range; the duplicated "does not allow overlapping merges" test is kept
// once; "preserves merges after row inserts" inspects internal value classes and uses
// `insertRow`, so it now expects XLV_E_NOT_SUPPORTED.
import { describe, expect, it } from 'vitest';
import { ExcelJS, expectNotSupported, styles } from './helpers';

function eachAddress(range: string, f: (row: number, col: number) => void): void {
  const [a, b] = range.split(':');
  const ws = new ExcelJS.Workbook().addWorksheet('x');
  const tl = ws.getCell(a);
  const br = ws.getCell(b ?? a);
  for (let r = tl.row; r <= br.row; r++) for (let c = tl.col; c <= br.col; c++) f(r, c);
}

describe('Worksheet', () => {
  describe('Merge Cells', () => {
    it('references the same top-left value', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      ws.getCell('A1').value = 'A1';
      ws.getCell('B1').value = 'B1';
      ws.getCell('A2').value = 'A2';
      ws.getCell('B2').value = 'B2';
      ws.mergeCells('A1:B2');
      expect(ws.getCell('A1').value).toBe('A1');
      expect(ws.getCell('B1').value).toBe('A1');
      expect(ws.getCell('A2').value).toBe('A1');
      expect(ws.getCell('B2').value).toBe('A1');
      expect(ws.getCell('A1').type).toBe(ExcelJS.ValueType.String);
      expect(ws.getCell('B1').type).toBe(ExcelJS.ValueType.Merge);
      expect(ws.getCell('A2').type).toBe(ExcelJS.ValueType.Merge);
      expect(ws.getCell('B2').type).toBe(ExcelJS.ValueType.Merge);
    });

    it('does not allow overlapping merges', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      ws.mergeCells('B2:C3');
      expect(() => ws.mergeCells('A1:B2')).toThrow(Error);
      expect(() => ws.mergeCells('C1:D2')).toThrow(Error);
      expect(() => ws.mergeCells('C3:D4')).toThrow(Error);
      expect(() => ws.mergeCells('A3:B4')).toThrow(Error);
      expect(() => ws.mergeCells('A1:D4')).toThrow(Error);
    });

    it('merges and unmerges', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      const expectMaster = (range: string, master: string | null): void => {
        eachAddress(range, (r, c) => {
          const cell = ws.getCell(r, c);
          const masterCell = master ? ws.getCell(master) : cell;
          expect(cell.master.address).toBe(masterCell.address);
        });
      };
      ws.mergeCells('A1:B2');
      expectMaster('A1:B2', 'A1');
      ws.unMergeCells('A1:B2');
      expectMaster('A1:B2', null);

      ws.mergeCells('A1:B2');
      expectMaster('A1:B2', 'A1');
      ws.unMergeCells('A1');
      expectMaster('A1:B2', null);

      ws.mergeCells('A1:B2');
      expectMaster('A1:B2', 'A1');
      ws.unMergeCells('B2');
      expectMaster('A1:B2', null);

      ws.mergeCells('A1:B2');
      ws.mergeCells('D1:E2');
      ws.mergeCells('A4:B5');
      ws.mergeCells('D4:E5');
      expectMaster('A1:B2', 'A1');
      expectMaster('D1:E2', 'D1');
      expectMaster('A4:B5', 'A4');
      expectMaster('D4:E5', 'D4');

      ws.unMergeCells('B2:D4');
      expectMaster('A1:B2', null);
      expectMaster('D1:E2', null);
      expectMaster('A4:B5', null);
      expectMaster('D4:E5', null);
    });

    it('merges styles', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      const B2 = ws.getCell('B2');
      B2.value = 5;
      B2.style.font = styles.fonts.broadwayRedOutline20;
      B2.style.border = styles.borders.doubleRed;
      B2.style.fill = styles.fills.blueWhiteHGrad;
      B2.style.alignment = styles.namedAlignments.middleCentre;
      B2.style.numFmt = styles.numFmts.numFmt1;

      ws.mergeCells('B2:C3');
      for (const address of ['B2', 'B3', 'C2', 'C3']) {
        expect(ws.getCell(address).font).toEqual(styles.fonts.broadwayRedOutline20);
        expect(ws.getCell(address).border).toEqual(styles.borders.doubleRed);
        expect(ws.getCell(address).fill).toEqual(styles.fills.blueWhiteHGrad);
        expect(ws.getCell(address).alignment).toEqual(styles.namedAlignments.middleCentre);
        expect(ws.getCell(address).numFmt).toEqual(styles.numFmts.numFmt1);
      }
    });

    it('preserves merges after row inserts — not supported: insertRow', async () => {
      const ws = new ExcelJS.Workbook().addWorksheet('testMergeAfterInsert');
      ws.addRow([1, 2]);
      ws.addRow([3, 4]);
      ws.mergeCells('A1:B2');
      await expectNotSupported(() => ws.insertRow(1, ['Inserted Row Text']));
    });
  });
});
