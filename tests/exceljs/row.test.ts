// Ported from exceljs v4.4.0 spec/unit/doc/row.spec.js (MIT, Copyright (c) 2014-2019 Guyon Roche),
// changes: Mocha/Chai → Vitest; the internal sheet mock (`createSheetMock`, `addColumn`) is
// replaced by a real worksheet and `ws.getColumn(n).key = …`; `row.hasValues`, `row.dimensions`
// and `row.model` (internal model) assertions are skipped; `row.splice` tests expect
// XLV_E_NOT_SUPPORTED (row splicing is not supported).
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

  it.skip('stores cells: hasValues and dimensions — not supported: internal row dimensions', () => {});

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
    it('remove only — not supported: row.splice', async () => {
      const row = sheet().getRow(1);
      row.getCell(2).value = 2;
      await expectNotSupported(() => row.splice(3, 3));
    });
    for (const name of ['remove to end', 'remove almost to end', 'remove past end', 'remove and insert fewer', 'remove and insert replacements', 'remove and insert more']) {
      it.skip(`${name} — not supported: row.splice`, () => {});
    }
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

  it.skip('builds a model — not supported: internal row model', () => {});
  it.skip('builds from model — not supported: internal row model', () => {});

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
