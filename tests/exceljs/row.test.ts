// Ported from exceljs v4.4.0 spec/unit/doc/row.spec.js (MIT, Copyright (c) 2014-2019 Guyon Roche),
// changes: Mocha/Chai → Vitest; the internal sheet mock (`createSheetMock`, `addColumn`) is
// replaced by a real worksheet and `ws.getColumn(n).key = …`; `row.hasValues`, `row.dimensions`
// and `row.model` (internal model) assertions are skipped. The `row.splice` tests are ported in
// full (phase B).
import { describe, expect, it } from 'vitest';
import { ExcelJS, expectNotSupported } from './helpers';

const sheet = () => new ExcelJS.Workbook().addWorksheet('blort');

describe('Row', () => {
  it('stores cells', () => {
    const ws = sheet();
    ws.getColumn(1).key = 'name';
    const row1 = ws.getRow(1);
    expect(row1.number).toBe(1);

    const a1 = row1.getCell(1);
    expect(a1.address).toBe('A1');
    expect(a1.type).toBe(ExcelJS.ValueType.Null);
    expect(row1.getCell('A')).toBe(a1);
    expect(row1.getCell('name')).toBe(a1);

    a1.value = 5;
    expect(a1.type).toBe(ExcelJS.ValueType.Number);

    const b1 = row1.getCell(2);
    expect(b1.address).toBe('B1');
    expect(b1.type).toBe(ExcelJS.ValueType.Null);
    b1.value = 'Hello, World!';
    const d1 = row1.getCell(4);
    d1.value = { hyperlink: 'http://www.hyperlink.com', text: 'www.hyperlink.com' };

    // Sparse array: index 0 and 3 are holes, as in exceljs.
    const values: unknown[] = [];
    values[1] = 5;
    values[2] = 'Hello, World!';
    values[4] = { hyperlink: 'http://www.hyperlink.com', text: 'www.hyperlink.com' };
    expect(row1.values).toEqual(values);

    let count = 0;
    row1.eachCell((cell: { type: number; value: unknown }, colNumber: number) => {
      expect(cell.type).not.toBe(ExcelJS.ValueType.Null);
      expect(cell.value).toEqual(values[colNumber]);
      count++;
    });
    // eachCell should just cover non-null cells
    expect(count).toBe(3);
  });

  it.skip('stores cells: hasValues and dimensions - not supported: internal row dimensions', () => {});

  it('stores values by whole row', () => {
    const ws = sheet();
    ws.getColumn(1).key = 'id';
    ws.getColumn(2).key = 'name';
    ws.getColumn(3).key = 'dob';
    const now = new Date();
    const row1 = ws.getRow(1);

    row1.values = [5, 'Hello, World!', null];
    expect(row1.getCell(1).value).toBe(5);
    expect(row1.getCell(2).value).toBe('Hello, World!');
    expect(row1.getCell(3).value).toBeNull();
    const expected1: unknown[] = [];
    expected1[1] = 5;
    expected1[2] = 'Hello, World!';
    expect(row1.values).toEqual(expected1);

    const values: unknown[] = [];
    values[1] = 7;
    values[3] = 'Not Null!';
    values[5] = now;
    row1.values = values;
    expect(row1.getCell(1).value).toBe(7);
    expect(row1.getCell(2).value).toBeNull();
    expect(row1.getCell(3).value).toBe('Not Null!');
    expect(row1.getCell(5).type).toBe(ExcelJS.ValueType.Date);
    expect(row1.values).toEqual(values);

    row1.values = { id: 9, name: 'Dobbie', dob: now };
    expect(row1.getCell(1).value).toBe(9);
    expect(row1.getCell(2).value).toBe('Dobbie');
    expect(row1.getCell(3).type).toBe(ExcelJS.ValueType.Date);
    expect(row1.getCell(5).value).toBeNull();
    const expected3: unknown[] = [];
    expected3[1] = 9;
    expected3[2] = 'Dobbie';
    expected3[3] = now;
    expect(row1.values).toEqual(expected3);
  });

  describe('Splice', () => {
    it('remove only', () => {
      const row = sheet().getRow(1);
      row.getCell(2).value = 2;
      row.getCell(3).value = 3;
      row.getCell(5).value = 5;
      row.getCell(7).value = 7;
      row.getCell(11).value = 'eleven';
      row.getCell(13).value = 13;

      row.splice(3, 3);
      expect(row.getCell(2).value).toBe(2);
      expect(row.getCell(2).type).toBe(ExcelJS.ValueType.Number);
      expect(row.getCell(2).address).toBe('B1');
      expect(row.getCell(4).value).toBe(7);
      expect(row.getCell(4).type).toBe(ExcelJS.ValueType.Number);
      expect(row.getCell(4).address).toBe('D1');
      expect(row.getCell(8).value).toBe('eleven');
      expect(row.getCell(8).type).toBe(ExcelJS.ValueType.String);
      expect(row.getCell(8).address).toBe('H1');
    });

    it('remove to end', () => {
      const row = sheet().getRow(1);
      for (let i = 1; i <= 5; i++) row.getCell(i).value = i;
      row.splice(4, 2);
      expect(row.getCell(1).value).toBe(1);
      expect(row.getCell(2).value).toBe(2);
      expect(row.getCell(3).value).toBe(3);
      expect(row.getCell(4).value).toBeNull();
      expect(row.getCell(5).value).toBeNull();
      expect(row.getCell(6).value).toBeNull();
    });

    it('remove almost to end', () => {
      const row = sheet().getRow(1);
      for (let i = 1; i <= 6; i++) row.getCell(i).value = i;
      row.splice(4, 2);
      expect(row.getCell(1).value).toBe(1);
      expect(row.getCell(2).value).toBe(2);
      expect(row.getCell(3).value).toBe(3);
      expect(row.getCell(4).value).toBe(6);
      expect(row.getCell(5).value).toBeNull();
      expect(row.getCell(6).value).toBeNull();
    });

    it('remove past end', () => {
      const row = sheet().getRow(1);
      for (let i = 1; i <= 6; i++) row.getCell(i).value = i;
      row.splice(4, 4);
      expect(row.getCell(1).value).toBe(1);
      expect(row.getCell(2).value).toBe(2);
      expect(row.getCell(3).value).toBe(3);
      for (let i = 4; i <= 8; i++) expect(row.getCell(i).value).toBeNull();
    });

    it('remove and insert fewer', () => {
      const row = sheet().getRow(1);
      row.values = [1, 2, 3, 4, 5, 6, 7, 8];
      row.splice(4, 3, 'four', 'five');
      expect(row.values).toEqual([, 1, 2, 3, 'four', 'five', 7, 8]);
    });

    it('remove and insert replacements', () => {
      const row = sheet().getRow(1);
      row.values = [1, 2, 3, 4, 5, 6, 7, 8];
      row.splice(4, 3, 'four', 'five', 'six');
      expect(row.values).toEqual([, 1, 2, 3, 'four', 'five', 'six', 7, 8]);
    });

    it('remove and insert more', () => {
      const row = sheet().getRow(1);
      row.values = [1, 2, 3, 4, 5, 6, 7, 8];
      row.splice(4, 3, 'four', 'five', 'six', 'six and a half');
      expect(row.values).toEqual([, 1, 2, 3, 'four', 'five', 'six', 'six and a half', 7, 8]);
    });

    // xlvane additions (phase B, D-B03).
    it('moves notes and styles with their cells', () => {
      const row = sheet().getRow(1);
      row.values = [1, 2, 3];
      row.getCell(2).note = 'two';
      row.getCell(3).font = { bold: true };
      row.splice(1, 1);
      expect(row.getCell(1).note).toBe('two');
      expect(row.getCell(2).font).toEqual({ bold: true });
      expect(row.getCell(3).note).toBeUndefined();
    });

    it('refuses a row that crosses a merged range, and invalid arguments', async () => {
      const ws = sheet();
      ws.mergeCells('C1:D2');
      await expectNotSupported(() => ws.getRow(1).splice(2, 1));
      expect(() => ws.getRow(3).splice(0, 1)).toThrow(expect.objectContaining({ code: 'XLV_E_INVALID_ARGUMENT' }));
      expect(() => ws.getRow(3).splice(1, -1)).toThrow(expect.objectContaining({ code: 'XLV_E_INVALID_ARGUMENT' }));
    });
  });

  it('iterates over cells', () => {
    const row1 = sheet().getRow(1);
    row1.getCell(1).value = 1;
    row1.getCell(2).value = 2;
    row1.getCell(4).value = 4;
    row1.getCell(6).value = 6;
    row1.eachCell((cell: { value: unknown }, colNumber: number) => {
      expect(colNumber).not.toBe(3);
      expect(colNumber).not.toBe(5);
      expect(cell.value).toBe(colNumber);
    });
    let count = 1;
    row1.eachCell({ includeEmpty: true }, (_cell: unknown, colNumber: number) => {
      expect(colNumber).toBe(count++);
    });
    expect(count).toBe(7);
  });

  it.skip('builds a model - not supported: internal row model', () => {});
  it.skip('builds from model - not supported: internal row model', () => {});

  it('counts cells', () => {
    const row1 = sheet().getRow(1);
    row1.getCell(1).value = 'one';
    row1.getCell(2).value = 'two';
    row1.getCell(4).value = 'four';
    row1.getCell(5).value = 'five';
    expect(row1.cellCount).toBe(5);
    expect(row1.actualCellCount).toBe(4);
  });
});
