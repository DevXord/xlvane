// Ported from exceljs v4.4.0 spec/unit/doc/worksheet.values.spec.js (MIT, Copyright (c)
// 2014-2019 Guyon Roche), changes: Mocha/Chai → Vitest; sparse array literals written out
// explicitly; column `defn` and `headers` assertions removed (multi-row headers are not
// supported); the Splice tests (`spliceRows`, `spliceColumns`) expect XLV_E_NOT_SUPPORTED;
// `getSheetValues` and `actualColumnCount` are skipped (outside the supported subset).
import { describe, expect, it } from 'vitest';
import { ExcelJS, expectNotSupported } from './helpers';

function sparse(entries: Record<number, unknown>): unknown[] {
  const out: unknown[] = [];
  for (const [k, v] of Object.entries(entries)) out[Number(k)] = v;
  return out;
}

describe('Worksheet', () => {
  describe('Values', () => {
    it('stores values properly', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      const now = new Date();
      ws.getCell('A1').value = 7;
      ws.getCell('B1').value = 'Hello, World!';
      ws.getCell('C1').value = 3.14;
      ws.getCell('D1').value = 5;
      ws.getCell('D1').value = now;
      ws.getCell('E1').value = `${['Hello', 'World'].join(', ')}!`;
      ws.getCell('F1').value = { text: 'www.google.com', hyperlink: 'http://www.google.com' };
      ws.getCell('A2').value = { formula: 'A1', result: 7 };
      ws.getCell('B2').value = { formula: 'CONCATENATE("Hello", ", ", "World!")', result: 'Hello, World!' };
      ws.getCell('C2').value = { formula: 'D1', result: now };

      expect(ws.getCell('A1').value).toBe(7);
      expect(ws.getCell('B1').value).toBe('Hello, World!');
      expect(ws.getCell('C1').value).toBe(3.14);
      expect(ws.getCell('D1').value).toBe(now);
      expect(ws.getCell('E1').value).toBe('Hello, World!');
      expect(ws.getCell('F1').value.text).toBe('www.google.com');
      expect(ws.getCell('F1').value.hyperlink).toBe('http://www.google.com');
      expect(ws.getCell('A2').value.formula).toBe('A1');
      expect(ws.getCell('A2').value.result).toBe(7);
      expect(ws.getCell('B2').value.formula).toBe('CONCATENATE("Hello", ", ", "World!")');
      expect(ws.getCell('B2').value.result).toBe('Hello, World!');
      expect(ws.getCell('C2').value.formula).toBe('D1');
      expect(ws.getCell('C2').value.result).toBe(now);
    });

    it('stores shared string values properly', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      ws.getCell('A1').value = 'Hello, World!';
      ws.getCell('A2').value = 'Hello';
      ws.getCell('B2').value = 'World';
      ws.getCell('C2').value = { formula: 'CONCATENATE(A2, ", ", B2, "!")', result: 'Hello, World!' };
      ws.getCell('A3').value = `${['Hello', 'World'].join(', ')}!`;
      expect(ws.getCell('A1').value).toBe(ws.getCell('A3').value);
      expect(ws.getCell('A1').value).toBe(ws.getCell('C2').value.result);
    });

    it('assigns cell types properly', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      ws.getCell('A1').value = 7;
      ws.getCell('B1').value = 'Hello, World!';
      ws.getCell('C1').value = 3.14;
      ws.getCell('D1').value = new Date();
      ws.getCell('E1').value = { text: 'www.google.com', hyperlink: 'http://www.google.com' };
      ws.getCell('A2').value = { formula: 'A1', result: 7 };
      ws.getCell('B2').value = { formula: 'CONCATENATE("Hello", ", ", "World!")', result: 'Hello, World!' };
      ws.getCell('C2').value = { formula: 'D1', result: new Date() };
      expect(ws.getCell('A1').type).toBe(ExcelJS.ValueType.Number);
      expect(ws.getCell('B1').type).toBe(ExcelJS.ValueType.String);
      expect(ws.getCell('C1').type).toBe(ExcelJS.ValueType.Number);
      expect(ws.getCell('D1').type).toBe(ExcelJS.ValueType.Date);
      expect(ws.getCell('E1').type).toBe(ExcelJS.ValueType.Hyperlink);
      expect(ws.getCell('A2').type).toBe(ExcelJS.ValueType.Formula);
      expect(ws.getCell('B2').type).toBe(ExcelJS.ValueType.Formula);
      expect(ws.getCell('C2').type).toBe(ExcelJS.ValueType.Formula);
    });

    it('adds columns', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      ws.columns = [
        { key: 'id', width: 10 },
        { key: 'name', width: 32 },
        { key: 'dob', width: 10 },
      ];
      expect(ws.getColumn('id').number).toBe(1);
      expect(ws.getColumn('id').width).toBe(10);
      expect(ws.getColumn('A')).toBe(ws.getColumn('id'));
      expect(ws.getColumn(1)).toBe(ws.getColumn('id'));
      expect(ws.getColumn('name').number).toBe(2);
      expect(ws.getColumn('name').width).toBe(32);
      expect(ws.getColumn('B')).toBe(ws.getColumn('name'));
      expect(ws.getColumn(2)).toBe(ws.getColumn('name'));
      expect(ws.getColumn('dob').number).toBe(3);
      expect(ws.getColumn('dob').width).toBe(10);
      expect(ws.getColumn('C')).toBe(ws.getColumn('dob'));
      expect(ws.getColumn(3)).toBe(ws.getColumn('dob'));
    });

    it('adds column headers', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      ws.columns = [
        { header: 'Id', width: 10 },
        { header: 'Name', width: 32 },
        { header: 'D.O.B.', width: 10 },
      ];
      expect(ws.getCell('A1').value).toBe('Id');
      expect(ws.getCell('B1').value).toBe('Name');
      expect(ws.getCell('C1').value).toBe('D.O.B.');
    });

    for (const by of ['number', 'letter'] as const) {
      it(`adds column headers by ${by}`, () => {
        const ws = new ExcelJS.Workbook().addWorksheet('blort');
        const first = by === 'number' ? 1 : 'A';
        const second = by === 'number' ? 2 : 'B';
        ws.getColumn(first).key = 'id';
        ws.getColumn(first).header = 'Id';
        ws.getColumn(first).width = 10;
        ws.getColumn(second).key = 'name';
        ws.getColumn(second).header = 'Name';
        ws.getColumn(second).width = 32;
        expect(ws.getCell('A1').value).toBe('Id');
        expect(ws.getCell('B1').value).toBe('Name');
        expect(ws.getColumn('A').key).toBe('id');
        expect(ws.getColumn(1).key).toBe('id');
        expect(ws.getColumn(first).header).toBe('Id');
        expect(ws.getColumn(first).width).toBe(10);
        expect(ws.getColumn(second).key).toBe('name');
        expect(ws.getColumn(second).header).toBe('Name');
        expect(ws.getColumn(second).width).toBe(32);
      });
    }

    it('adds rows by object', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      ws.columns = [
        { header: 'Id', key: 'id', width: 10 },
        { header: 'Name', key: 'name', width: 32 },
        { header: 'D.O.B.', key: 'dob', width: 10 },
      ];
      const dateValue1 = new Date(1970, 1, 1);
      const dateValue2 = new Date(1965, 1, 7);
      ws.addRow({ id: 1, name: 'John Doe', dob: dateValue1 });
      ws.addRow({ id: 2, name: 'Jane Doe', dob: dateValue2 });
      expect(ws.getCell('A2').value).toBe(1);
      expect(ws.getCell('B2').value).toBe('John Doe');
      expect(ws.getCell('C2').value).toBe(dateValue1);
      expect(ws.getCell('A3').value).toBe(2);
      expect(ws.getCell('B3').value).toBe('Jane Doe');
      expect(ws.getCell('C3').value).toBe(dateValue2);
      expect(ws.getRow(2).values).toEqual(sparse({ 1: 1, 2: 'John Doe', 3: dateValue1 }));
      expect(ws.getRow(3).values).toEqual(sparse({ 1: 2, 2: 'Jane Doe', 3: dateValue2 }));
      const values = sparse({
        1: sparse({ 1: 'Id', 2: 'Name', 3: 'D.O.B.' }),
        2: sparse({ 1: 1, 2: 'John Doe', 3: dateValue1 }),
        3: sparse({ 1: 2, 2: 'Jane Doe', 3: dateValue2 }),
      }) as unknown[][];
      ws.eachRow((row: { values: unknown; eachCell: (f: (cell: { value: unknown }, c: number) => void) => void }, rowNumber: number) => {
        expect(row.values).toEqual(values[rowNumber]);
        row.eachCell((cell, colNumber) => expect(cell.value).toBe(values[rowNumber]?.[colNumber]));
      });
    });

    it('adds rows by contiguous array', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      const dateValue1 = new Date(1970, 1, 1);
      const dateValue2 = new Date(1965, 1, 7);
      ws.addRow([1, 'John Doe', dateValue1]);
      ws.addRow([2, 'Jane Doe', dateValue2]);
      expect(ws.getCell('A1').value).toBe(1);
      expect(ws.getCell('B1').value).toBe('John Doe');
      expect(ws.getCell('C1').value).toBe(dateValue1);
      expect(ws.getCell('A2').value).toBe(2);
      expect(ws.getCell('B2').value).toBe('Jane Doe');
      expect(ws.getCell('C2').value).toBe(dateValue2);
      expect(ws.getRow(1).values).toEqual(sparse({ 1: 1, 2: 'John Doe', 3: dateValue1 }));
      expect(ws.getRow(2).values).toEqual(sparse({ 1: 2, 2: 'Jane Doe', 3: dateValue2 }));
    });

    it('adds rows by sparse array', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      const dateValue1 = new Date(1970, 1, 1);
      const dateValue2 = new Date(1965, 1, 7);
      const rows = [
        undefined,
        sparse({ 1: 1, 2: 'John Doe', 4: dateValue1 }),
        sparse({ 1: 2, 2: 'Jane Doe', 4: dateValue2 }),
        sparse({ 1: 3, 3: 'Sam', 5: dateValue1 }),
      ];
      for (const row of rows) if (row) ws.addRow(row);
      expect(ws.getCell('A1').value).toBe(1);
      expect(ws.getCell('B1').value).toBe('John Doe');
      expect(ws.getCell('D1').value).toBe(dateValue1);
      expect(ws.getCell('A2').value).toBe(2);
      expect(ws.getCell('B2').value).toBe('Jane Doe');
      expect(ws.getCell('D2').value).toBe(dateValue2);
      expect(ws.getCell('A3').value).toBe(3);
      expect(ws.getCell('C3').value).toBe('Sam');
      expect(ws.getCell('E3').value).toBe(dateValue1);
      expect(ws.getRow(1).values).toEqual(rows[1]);
      expect(ws.getRow(2).values).toEqual(rows[2]);
      expect(ws.getRow(3).values).toEqual(rows[3]);
      ws.eachRow((row: { values: unknown; eachCell: (f: (cell: { value: unknown }, c: number) => void) => void }, rowNumber: number) => {
        expect(row.values).toEqual(rows[rowNumber]);
        row.eachCell((cell, colNumber) => expect(cell.value).toBe(rows[rowNumber]?.[colNumber]));
      });
    });

    describe('Splice', () => {
      it('Rows — not supported: spliceRows', async () => {
        const ws = new ExcelJS.Workbook().addWorksheet('splice');
        ws.addRow([1, 2]);
        await expectNotSupported(() => ws.spliceRows(1, 1));
      });
      it('Columns — not supported: spliceColumns', async () => {
        const ws = new ExcelJS.Workbook().addWorksheet('splice');
        ws.addRow([1, 2]);
        await expectNotSupported(() => ws.spliceColumns(1, 1));
      });
      for (const name of [
        'Rows: Remove and insert fewer',
        'Rows: Remove and insert same',
        'Rows: Remove and insert more',
        'Rows: Remove/Insert/Replace style',
        'Rows: Remove/Insert/Replace defined names',
        'Columns: Remove and insert fewer/same/more',
        'Columns: handles column keys',
        'Columns: Splices to end, past end, almost to end',
        'Columns: Remove/Insert/Replace style',
        'Columns: Remove/Insert/Replace defined names',
      ]) {
        it.skip(`${name} — not supported: spliceRows/spliceColumns`, () => {});
      }
    });

    it('iterates over rows', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      ws.getCell('A1').value = 1;
      ws.getCell('B2').value = 2;
      ws.getCell('D4').value = 4;
      ws.getCell('F6').value = 6;
      ws.eachRow((_row: unknown, rowNumber: number) => {
        expect(rowNumber).not.toBe(3);
        expect(rowNumber).not.toBe(5);
      });
      let count = 1;
      ws.eachRow({ includeEmpty: true }, (_row: unknown, rowNumber: number) => {
        expect(rowNumber).toBe(count++);
      });
    });

    it('iterates over collumn cells', () => {
      const ws = new ExcelJS.Workbook().addWorksheet('blort');
      ws.getCell('A1').value = 1;
      ws.getCell('A2').value = 2;
      ws.getCell('A4').value = 4;
      ws.getCell('A6').value = 6;
      const colA = ws.getColumn('A');
      colA.eachCell((cell: { value: unknown }, rowNumber: number) => {
        expect(rowNumber).not.toBe(3);
        expect(rowNumber).not.toBe(5);
        expect(cell.value).toBe(rowNumber);
      });
      let count = 1;
      colA.eachCell({ includeEmpty: true }, (_cell: unknown, rowNumber: number) => {
        expect(rowNumber).toBe(count++);
      });
      expect(count).toBe(7);
    });

    it.skip('returns sheet values — not supported: getSheetValues', () => {});

    it('calculates rowCount and actualRowCount', () => {
      const ws = new ExcelJS.Workbook().addWorksheet();
      ws.getCell('A1').value = 'A1';
      ws.getCell('C1').value = 'C1';
      ws.getCell('A3').value = 'A3';
      ws.getCell('D3').value = 'D3';
      ws.getCell('A4').value = null;
      ws.getCell('B5').value = 'B5';
      expect(ws.rowCount).toBe(5);
      expect(ws.actualRowCount).toBe(3);
    });

    it('calculates columnCount', () => {
      const ws = new ExcelJS.Workbook().addWorksheet();
      ws.getCell('A1').value = 'A1';
      ws.getCell('C1').value = 'C1';
      ws.getCell('A3').value = 'A3';
      ws.getCell('D3').value = 'D3';
      ws.getCell('E4').value = null;
      ws.getCell('F5').value = 'F5';
      expect(ws.columnCount).toBe(6);
    });

    it.skip('calculates actualColumnCount — not supported: actualColumnCount', () => {});
  });
});
