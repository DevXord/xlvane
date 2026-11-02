// MCP server tests on an in-memory client (acceptance criterion of phase 7).
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { SERVER_VERSION } from '../../packages/mcp/src/server';
import { readTable } from '../../packages/xlvane/src/table';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { connect, withData } from './helpers';

const READ_TOOLS = ['inspect_workbook', 'read_table', 'diff', 'validate'];
const WRITE_TOOLS = ['write_table', 'edit_cells', 'convert'];

describe('tools/list', () => {
  it('offers only the read tools without --allow-write, all seven with it, in a fixed order', async () => {
    await withData(async (dir) => {
      const ro = await connect({ roots: [dir] });
      const tools = (await ro.client.listTools()).tools;
      expect(tools.map((t) => t.name)).toEqual(READ_TOOLS);
      for (const t of tools) {
        expect(t.annotations?.readOnlyHint, t.name).toBe(true);
        expect(t.description!.length, t.name).toBeGreaterThan(40);
        expect(t.inputSchema.type).toBe('object');
        expect(t.inputSchema.additionalProperties, t.name).toBe(false);
      }
      expect(ro.client.getInstructions()).toContain('Writing is disabled');
      await ro.close();
      const rw = await connect({ roots: [dir], allowWrite: true });
      const all = (await rw.client.listTools()).tools;
      expect(all.map((t) => t.name)).toEqual([...READ_TOOLS, ...WRITE_TOOLS]);
      for (const t of all.slice(4)) expect(t.annotations, t.name).toMatchObject({ readOnlyHint: false, destructiveHint: false });
      expect(rw.client.getInstructions()).toContain('Writing is enabled');
      await rw.close();
    });
  });

  it('reports the package version', async () => {
    const pkg = JSON.parse(await readFile('packages/mcp/package.json', 'utf8')) as { version: string };
    expect(SERVER_VERSION).toBe(pkg.version);
  });
});

describe('read tools', () => {
  it('inspect_workbook, read_table, diff and validate', async () => {
    await withData(async (dir) => {
      const c = await connect({ roots: [dir], maxRows: 2 });
      const inspect = await c.call('inspect_workbook', { path: 'report.xlsx' });
      expect(inspect.isError).toBe(false);
      expect(inspect.result).toMatchObject({ schemaVersion: 1, sheetCount: 2 });
      expect(JSON.parse(inspect.text)).toEqual(inspect.result);
      expect((await c.call('inspect_workbook', { path: 'report.xlsx', sheet: 'Costs' })).result['sheets']).toHaveLength(1);

      const page1 = await c.call('read_table', { path: 'report.xlsx', sheet: 'Sales' });
      expect(page1.result).toMatchObject({ totalRows: 3, truncated: true, offset: 0, headerRow: 1 });
      expect(page1.result['rows']).toEqual([
        { Region: 'North', Revenue: 1200, Date: '2024-01-31T00:00:00.000Z' },
        { Region: 'South', Revenue: 980, Date: '2024-02-29T00:00:00.000Z' },
      ]);
      const page2 = await c.call('read_table', { path: 'report.xlsx', sheet: 'Sales', offset: 2 });
      expect(page2.result).toMatchObject({ truncated: false, rowNumbers: [4] });

      const d = await c.call('diff', { before: 'old.xlsx', after: 'new.xlsx' });
      expect(d.result).toMatchObject({ equal: false, summary: { cellsChanged: 2 } });

      const v = await c.call('validate', { path: 'report.xlsx', sheet: 'Sales', columns: [{ name: 'Revenue', type: 'number' }, { name: 'Date', type: 'date' }, { name: 'Owner' }] });
      expect(v.result['ok']).toBe(false);
      expect((v.result['problems'] as { cell: string | null; kind: string }[]).map((p) => [p.kind, p.cell])).toEqual([
        ['empty', 'C4'],
        ['missing-column', null],
      ]);
      await c.close();
    });
  });

  it('returns errors with code and hint so the call can be corrected', async () => {
    await withData(async (dir) => {
      const c = await connect({ roots: [dir] });
      const r = await c.call('read_table', { path: 'report.xlsx', sheet: 'Salse' });
      expect(r.isError).toBe(true);
      expect(r.result['error']).toMatchObject({ code: 'XLV_E_SHEET_NOT_FOUND', hint: expect.stringContaining('Did you mean "Sales"?') as unknown });
      const missing = await c.call('inspect_workbook', { path: 'nope.xlsx' });
      expect(missing.result['error']).toMatchObject({ code: 'XLV_E_FILE_READ' });
      const notXlsx = await c.call('inspect_workbook', { path: 'data.json' });
      expect(notXlsx.result['error']).toMatchObject({ code: 'XLV_E_NOT_ZIP', details: { path: 'data.json' } });
      await c.close();
    });
  });

  it('rejects input that does not match the schema', async () => {
    await withData(async (dir) => {
      const c = await connect({ roots: [dir] });
      for (const args of [{}, { path: 5 }, { path: 'report.xlsx', bogus: true }, { path: 'report.xlsx', limit: 0 }]) {
        const r = await c.client.callTool({ name: 'read_table', arguments: args });
        expect(r.isError, JSON.stringify(args)).toBe(true);
      }
      await c.close();
    });
  });
});

describe('write tools', () => {
  it('write_table, edit_cells and convert write inside the root', async () => {
    await withData(async (dir) => {
      const c = await connect({ roots: [dir], allowWrite: true });
      const w = await c.call('write_table', { path: 'people.xlsx', rows: [{ name: 'Ann', joined: { date: '2024-03-01' }, score: 7.5 }], sheet: 'People' });
      expect(w).toMatchObject({ isError: false, result: { path: 'people.xlsx', replaced: false, rows: 1 } });
      expect(readTable(await readFile(join(dir, 'people.xlsx'))).rows).toEqual([{ name: 'Ann', joined: new Date(Date.UTC(2024, 2, 1)), score: 7.5 }]);

      const e = await c.call('edit_cells', {
        path: 'report.xlsx',
        edits: [{ sheet: 'Sales', cell: 'b2', value: 1300 }, { sheet: 'Sales', cell: 'D2', value: { formula: 'B2*2', result: 2600 } }],
        append: { sheet: 'Sales', rows: [{ Region: 'West', Revenue: 700 }], header: 1 },
        output: 'updated.xlsx',
      });
      expect(e).toMatchObject({ isError: false, result: { path: 'updated.xlsx', cellsSet: 2, rowsAppended: 1 } });
      const s = readWorkbook(await readFile(join(dir, 'updated.xlsx'))).getSheet('Sales');
      expect(s.cell('B2')?.value).toBe(1300);
      expect(s.cell('D2')?.formula?.text).toBe('B2*2');
      expect(s.cell('A5')?.value).toBe('West');

      const inPlace = await c.call('edit_cells', { path: 'report.xlsx', edits: [{ sheet: 'Sales', cell: 'B3', value: 1 }] });
      expect(inPlace.result['error']).toMatchObject({ code: 'XLV_E_FILE_EXISTS' });
      expect((await c.call('edit_cells', { path: 'report.xlsx', edits: [{ sheet: 'Sales', cell: 'B3', value: 1 }], overwrite: true })).result).toMatchObject({ replaced: true });

      expect((await c.call('convert', { input: 'report.xlsx', output: 'costs.csv', sheet: 'Costs' })).isError).toBe(false);
      expect(await readFile(join(dir, 'costs.csv'), 'utf8')).toContain(`"'=cmd|calc"`);
      expect((await c.call('convert', { input: 'sales.csv', output: 'sales.xlsx' })).isError).toBe(false);
      expect((await c.call('convert', { input: 'report.xlsx', output: 'x.pdf' })).result['error']).toMatchObject({ code: 'XLV_E_INVALID_INPUT' });
      await c.close();
    });
  });
});
