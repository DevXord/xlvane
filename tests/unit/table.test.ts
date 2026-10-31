import { type } from 'arktype';
import * as v from 'valibot';
import { describe, expect, it } from 'vitest';
import { z } from 'zod';
import { readWorkbook } from '../../packages/xlvane/src/read';
import type { StandardSchemaV1 } from '../../packages/xlvane/src/standard-schema';
import { fromJSON, readTable, readTableAsync, toJSON, writeTable } from '../../packages/xlvane/src/table';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError, expectXlvErrorAsync } from '../helpers/expect-error';

function sample(): Uint8Array {
  const wb = createWorkbook();
  const s = wb.addSheet('Sales');
  s.set('A1', 'Report title');
  s.set('A3', 'Name').set('B3', 'Amount').set('C3', 'Date').set('D3', ' Note ').set('E3', 'Paid');
  s.set('A4', 'Alice').set('B4', 10.5).set('C4', new Date(Date.UTC(2024, 0, 31))).set('D4', 'x').set('E4', true);
  s.set('A5', 'Bob').set('B5', 'n/a').set('E5', false);
  s.set('A7', 'Carol').set('B7', 7).set('E7', { error: '#N/A' });
  wb.addSheet('Empty');
  return writeWorkbook(wb);
}

describe('readTable', () => {
  it('detects the header row and returns rows as objects', () => {
    const t = readTable(sample());
    expect(t.sheet).toBe('Sales');
    expect(t.headerRow).toBe(3);
    expect(t.columns.map((c) => c.key)).toEqual(['Name', 'Amount', 'Date', 'Note', 'Paid']);
    expect(t.rows).toEqual([
      { Name: 'Alice', Amount: 10.5, Date: new Date(Date.UTC(2024, 0, 31)), Note: 'x', Paid: true },
      { Name: 'Bob', Amount: 'n/a', Date: null, Note: null, Paid: false },
      { Name: 'Carol', Amount: 7, Date: null, Note: null, Paid: '#N/A' },
    ]);
    expect(t.rowNumbers).toEqual([4, 5, 7]);
    expect(t.range).toBe('A3:E7');
    expect(t.issues).toEqual([]);
  });

  it('makes keys unique and names empty headers after their column', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 'id').set('B1', 'ID').set('D1', 'id').set('E1', '__proto__').set('A2', 1).set('C2', 3).set('E2', 'p');
    const t = readTable(writeWorkbook(wb), { header: 1 });
    expect(t.columns.map((c) => c.key)).toEqual(['id', 'ID', 'C', 'id_2', '__proto__']);
    const row = t.rows[0]!;
    expect(Object.keys(row)).toEqual(['id', 'ID', 'C', 'id_2', '__proto__']);
    expect(Object.getPrototypeOf(row)).toBe(Object.prototype);
    expect(row['__proto__']).toBe('p');
  });

  it('keeps empty rows when asked', () => {
    expect(readTable(sample(), { skipEmptyRows: false }).rowNumbers).toEqual([4, 5, 6, 7]);
  });

  it('can return dates as ISO strings', () => {
    expect(readTable(sample(), { dates: 'iso' }).rows[0]?.['Date']).toBe('2024-01-31');
  });

  it('uses a given header row, or column letters without a header', () => {
    expect(readTable(sample(), { header: 3, range: 'A3:B5' }).rows).toEqual([
      { Name: 'Alice', Amount: 10.5 },
      { Name: 'Bob', Amount: 'n/a' },
    ]);
    const noHeader = readTable(sample(), { header: false, range: 'A4:B5' });
    expect(noHeader.headerRow).toBeNull();
    expect(noHeader.rows).toEqual([
      { A: 'Alice', B: 10.5 },
      { A: 'Bob', B: 'n/a' },
    ]);
  });

  it('reads a workbook object as well as bytes, and selects sheets by name', () => {
    const wb = readWorkbook(sample());
    expect(readTable(wb, { sheet: 'Sales' }).rows).toHaveLength(3);
    expect(readTable(wb, { sheet: 'Empty' }).rows).toEqual([]);
    expectXlvError(() => readTable(wb, { sheet: 'sales' }), 'XLV_E_SHEET_NOT_FOUND');
  });

  it('reports a missing header row and bad options', () => {
    expectXlvError(() => readTable(sample(), { header: 2 }), 'XLV_E_HEADER_NOT_FOUND');
    expectXlvError(() => readTable(sample(), { header: 'yes' as never }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => readTable(sample(), { rnage: 'A1' } as never), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => readTable(sample(), { range: 'A1:' }), 'XLV_E_INVALID_CELL_REF');
    expectXlvError(() => readTable(sample(), { schema: {} as never }), 'XLV_E_SCHEMA_INVALID');
  });
});

describe('readTable with schemas', () => {
  const rowsOf = (bytes: Uint8Array) => bytes;

  it('validates with Zod and maps issues to cells', () => {
    const schema = z.object({ Name: z.string(), Amount: z.number() });
    const t = readTable(rowsOf(sample()), { schema });
    expect(t.rows).toEqual([
      { Name: 'Alice', Amount: 10.5 },
      { Name: 'Carol', Amount: 7 },
    ]);
    expect(t.rowNumbers).toEqual([4, 7]);
    expect(t.issues).toHaveLength(1);
    expect(t.issues[0]).toMatchObject({ row: 5, cell: 'B5', column: 'Amount', path: ['Amount'] });
    expect(t.issues[0]?.message.length).toBeGreaterThan(0);
  });

  it('validates with Valibot', () => {
    const schema = v.object({ Name: v.string(), Amount: v.number() });
    const t = readTable(sample(), { schema });
    expect(t.rows.map((r) => r.Name)).toEqual(['Alice', 'Carol']);
    expect(t.issues[0]).toMatchObject({ cell: 'B5', column: 'Amount' });
  });

  it('validates with ArkType', () => {
    const schema = type({ Name: 'string', Amount: 'number' });
    const t = readTable(sample(), { schema });
    expect(t.rows.map((r) => r.Name)).toEqual(['Alice', 'Carol']);
    expect(t.issues[0]).toMatchObject({ cell: 'B5', column: 'Amount' });
  });

  it('reports row-level issues without a cell', () => {
    const schema = z.object({ Name: z.string() }).refine((r) => r.Name !== 'Bob', { message: 'No Bob' });
    const t = readTable(sample(), { schema });
    expect(t.issues).toEqual([{ row: 5, cell: null, column: null, message: 'No Bob', path: [] }]);
  });

  it('needs readTableAsync for asynchronous schemas', async () => {
    const asyncSchema: StandardSchemaV1<unknown, { Name: string }> = {
      '~standard': {
        version: 1,
        vendor: 'test',
        validate: async (value) => ({ value: value as { Name: string } }),
      },
    };
    expectXlvError(() => readTable(sample(), { schema: asyncSchema }), 'XLV_E_ASYNC_SCHEMA');
    const t = await readTableAsync(sample(), { schema: asyncSchema });
    expect(t.rows).toHaveLength(3);
    await expectXlvErrorAsync(() => readTableAsync(sample(), { sheet: 'Nope' }), 'XLV_E_SHEET_NOT_FOUND');
  });
});

describe('writeTable', () => {
  const rows = [
    { name: 'Alice', amount: 10.5, when: new Date(Date.UTC(2024, 0, 31)), ok: true },
    { name: 'Bob', amount: null, when: undefined, ok: false },
  ];

  it('writes a report that reads back as the same table', () => {
    const bytes = writeTable(rows);
    const t = readTable(bytes);
    expect(t.rows).toEqual([
      { name: 'Alice', amount: 10.5, when: new Date(Date.UTC(2024, 0, 31)), ok: true },
      { name: 'Bob', amount: null, when: null, ok: false },
    ]);
    const s = readWorkbook(bytes).getSheet('Sheet1');
    expect(s.freeze).toEqual({ rows: 1, cols: 0 });
    expect(s.autoFilter).toBe('A1:D3');
    expect(s.cell('A1')?.style.font?.bold).toBe(true);
    expect(s.cell('A1')?.style.fill).toBeDefined();
    expect(s.columns.length).toBe(4);
  });

  it('writes a plain table without styles', () => {
    const s = readWorkbook(writeTable(rows, { preset: 'plain', sheet: 'Data' })).getSheet('Data');
    expect(s.cell('A1')?.style).toEqual({});
    expect(s.freeze).toBeNull();
    expect(s.autoFilter).toBeNull();
  });

  it('uses column definitions for order, headers and formats', () => {
    const bytes = writeTable(rows, { columns: [{ key: 'amount', header: 'Amount (EUR)', numFmt: '#,##0.00', width: 15 }, 'name'] });
    const s = readWorkbook(bytes).getSheet('Sheet1');
    expect(s.cell('A1')?.value).toBe('Amount (EUR)');
    expect(s.cell('B1')?.value).toBe('name');
    expect(s.cell('A2')?.numFmt).toBe('#,##0.00');
    expect(s.columns[0]).toMatchObject({ min: 1, width: 15 });
    expect(s.cell('C1')).toBeUndefined();
  });

  it('writes arrays of arrays', () => {
    const t = readTable(
      writeTable(
        [
          [1, 'a'],
          [2, 'b'],
        ],
        { columns: ['n', 's'] },
      ),
    );
    expect(t.rows).toEqual([
      { n: 1, s: 'a' },
      { n: 2, s: 'b' },
    ]);
  });

  it('rejects values that are not cell values', () => {
    expectXlvError(() => writeTable([{ a: { nested: true } }]), 'XLV_E_INVALID_ROWS', { row: 2, key: 'a' });
    expectXlvError(() => writeTable([5 as never]), 'XLV_E_INVALID_ROWS');
    expectXlvError(() => writeTable('rows' as never), 'XLV_E_INVALID_ROWS');
    expectXlvError(() => writeTable(rows, { preset: 'fancy' as never }), 'XLV_E_INVALID_OPTION');
  });
});

describe('toJSON / fromJSON', () => {
  it('returns JSON-safe rows', () => {
    const json = toJSON(sample());
    expect(json.sheet).toBe('Sales');
    expect(json.columns).toEqual(['Name', 'Amount', 'Date', 'Note', 'Paid']);
    expect(json.rows[0]).toEqual({ Name: 'Alice', Amount: 10.5, Date: '2024-01-31', Note: 'x', Paid: true });
    expect(JSON.parse(JSON.stringify(json))).toEqual(json);
  });

  it('writes JSON text or parsed rows', () => {
    const bytes = fromJSON('[{"a":1,"b":"x"},{"a":2,"c":true}]');
    expect(readTable(bytes).rows).toEqual([
      { a: 1, b: 'x', c: null },
      { a: 2, b: null, c: true },
    ]);
    expect(readTable(fromJSON([{ a: 1 }])).rows).toEqual([{ a: 1 }]);
  });

  it('rejects invalid JSON with its position', () => {
    const error = expectXlvError(() => fromJSON('[{"a":1},'), 'XLV_E_JSON_INVALID');
    expect(error.hint.length).toBeGreaterThan(0);
    expectXlvError(() => fromJSON('{"a":1}'), 'XLV_E_INVALID_ROWS');
  });
});
