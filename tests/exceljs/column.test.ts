// Ported from exceljs v4.4.0 spec/unit/doc/column.spec.js (MIT, Copyright (c) 2014-2019 Guyon Roche),
// changes: Mocha/Chai → Vitest; the internal sheet mock (`addColumn(n, defn)`) is replaced by a
// real worksheet (`ws.columns = [...]`, `ws.getColumn(n)`); the multi-row header assertions are
// a separate test; the internal `Column.toModel` tests are skipped.
import { describe, expect, it } from 'vitest';
import { ExcelJS } from './helpers';

const sheet = () => new ExcelJS.Workbook().addWorksheet('blort');

describe('Column', () => {
  it('creates by defn', () => {
    const ws = sheet();
    ws.columns = [{ header: 'Col 1', key: 'id1', width: 10 }];
    expect(ws.getColumn(1).header).toBe('Col 1');
    expect(ws.getCell(1, 1).value).toBe('Col 1');
    expect(ws.getColumn('id1')).toBe(ws.getColumn(1));
    ws.getRow(2).values = { id1: 'Hello, World!' };
    expect(ws.getCell(2, 1).value).toBe('Hello, World!');
  });

  it('maintains properties', () => {
    const ws = sheet();
    const column = ws.getColumn(1);
    column.key = 'id1';
    expect(ws.getColumn('id1')).toBe(column);
    expect(column.number).toBe(1);
    expect(column.letter).toBe('A');
    column.header = 'Col 1';
    expect(ws.getColumn(1).header).toBe('Col 1');
    expect(ws.getCell(1, 1).value).toBe('Col 1');
    ws.getRow(3).values = { id1: 'Hello, World!' };
    expect(ws.getCell(3, 1).value).toBe('Hello, World!');
  });

  it('maintains properties: multi-row headers', () => {
    const ws = sheet();
    const column = ws.getColumn(1);
    column.key = 'id1';
    column.header = 'Col 1';
    expect(ws.getColumn(1).header).toBe('Col 1');
    expect(ws.getColumn(1).headers).toEqual(['Col 1']);
    expect(ws.getCell(1, 1).value).toBe('Col 1');
    column.header = ['Col A1', 'Col A2'];
    expect(ws.getColumn(1).header).toEqual(['Col A1', 'Col A2']);
    expect(ws.getColumn(1).headers).toEqual(['Col A1', 'Col A2']);
    expect(ws.getCell(1, 1).value).toBe('Col A1');
    expect(ws.getCell(2, 1).value).toBe('Col A2');
    ws.getRow(3).values = { id1: 'Hello, World!' };
    expect(ws.getCell(3, 1).value).toBe('Hello, World!');
  });
  it.skip('creates model - not supported: internal Column.toModel', () => {});

  it('gets column values', () => {
    const ws = sheet();
    ws.getCell(1, 1).value = 'a';
    ws.getCell(2, 1).value = 'b';
    ws.getCell(4, 1).value = 'd';
    const expected: unknown[] = [];
    expected[1] = 'a';
    expected[2] = 'b';
    expected[4] = 'd';
    expect(ws.getColumn(1).values).toEqual(expected);
  });

  it('sets column values', () => {
    const ws = sheet();
    ws.getColumn(1).values = [2, 3, 5, 7, 11];
    expect(ws.getCell(1, 1).value).toBe(2);
    expect(ws.getCell(2, 1).value).toBe(3);
    expect(ws.getCell(3, 1).value).toBe(5);
    expect(ws.getCell(4, 1).value).toBe(7);
    expect(ws.getCell(5, 1).value).toBe(11);
    expect(ws.getCell(6, 1).value).toBe(null);
  });

  it('sets sparse column values', () => {
    const ws = sheet();
    const values: unknown[] = [];
    values[2] = 2;
    values[3] = 3;
    values[5] = 5;
    values[11] = 11;
    ws.getColumn(1).values = values;
    for (let r = 1; r <= 12; r++) expect(ws.getCell(r, 1).value).toBe(values[r] ?? null);
  });

  it('sets sparse column values (array literal with holes)', () => {
    const ws = sheet();
    const values: unknown[] = [];
    values[2] = 2;
    values[3] = 3;
    values[5] = 5;
    values[7] = 7;
    values[11] = 11;
    ws.getColumn(1).values = values;
    for (let r = 1; r <= 12; r++) expect(ws.getCell(r, 1).value).toBe(values[r] ?? null);
  });

  it.skip('sets default column width - not supported: internal Column.toModel', () => {});
});
