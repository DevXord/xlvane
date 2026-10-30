// Security of the MCP server (acceptance criterion: writing outside an allowed directory is
// rejected). Each case checks the error code and that nothing was written outside the root.
import { mkdir, mkdtemp, readdir, readFile, rm, stat, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { connect, type Connected } from './helpers';
import { writeFixtures } from '../cli/helpers';

let base = '';
let root = '';
let outside = '';
let c: Connected;

beforeEach(async () => {
  base = await mkdtemp(join(tmpdir(), 'xlvane-mcp-sec-'));
  root = join(base, 'root');
  outside = join(base, 'outside');
  await mkdir(root);
  await mkdir(outside);
  await writeFixtures(root);
  await writeFile(join(outside, 'secret.xlsx'), await readFile(join(root, 'report.xlsx')));
  await symlink(outside, join(root, 'link-dir'));
  await symlink(join(outside, 'secret.xlsx'), join(root, 'link-file.xlsx'));
  // A dangling link: writing through it would create a file outside the root.
  await symlink(join(outside, 'new.csv'), join(root, 'link-file.csv'));
  c = await connect({ roots: [root], allowWrite: true, maxFileBytes: 64 * 1024 });
});

afterEach(async () => {
  await c.close();
  await rm(base, { recursive: true, force: true });
});

const ROWS = [{ a: 1 }];

async function code(name: string, args: Record<string, unknown>): Promise<unknown> {
  const r = await c.call(name, args);
  expect(r.isError, `${name} ${JSON.stringify(args)}`).toBe(true);
  return (r.result['error'] as { code: string }).code;
}

describe('writing outside the allowed directories is rejected', () => {
  it.each([
    ['an absolute path', () => join(outside, 'x.xlsx')],
    ['a path with ..', () => '../outside/x.xlsx'],
    ['a symlinked directory', () => 'link-dir/x.xlsx'],
    ['a symlinked file', () => 'link-file.xlsx'],
  ])('%s', async (_title, path) => {
    for (const [tool, args] of [
      ['write_table', { path: path(), rows: ROWS, overwrite: true }],
      ['edit_cells', { path: 'report.xlsx', edits: [{ sheet: 'Sales', cell: 'A1', value: 'x' }], output: path(), overwrite: true }],
      ['convert', { input: 'report.xlsx', output: path().replace(/\.xlsx$/, '.csv'), overwrite: true }],
    ] as const) {
      expect(await code(tool, args)).toBe('XLV_E_PATH_NOT_ALLOWED');
    }
    expect((await readdir(outside)).sort()).toEqual(['secret.xlsx']);
    expect((await stat(join(outside, 'secret.xlsx'))).size).toBe((await stat(join(root, 'report.xlsx'))).size);
  });
});

describe('reading outside the allowed directories is rejected', () => {
  it.each([
    ['an absolute path', () => join(outside, 'secret.xlsx')],
    ['a path with ..', () => '../outside/secret.xlsx'],
    ['a symlinked directory', () => 'link-dir/secret.xlsx'],
    ['a symlinked file', () => 'link-file.xlsx'],
  ])('%s', async (_title, path) => {
    expect(await code('inspect_workbook', { path: path() })).toBe('XLV_E_PATH_NOT_ALLOWED');
    expect(await code('read_table', { path: path() })).toBe('XLV_E_PATH_NOT_ALLOWED');
    expect(await code('diff', { before: 'report.xlsx', after: path() })).toBe('XLV_E_PATH_NOT_ALLOWED');
    expect(await code('convert', { input: path(), output: 'leak.csv' })).toBe('XLV_E_PATH_NOT_ALLOWED');
  });

  it('the error names the allowed directories', async () => {
    const r = await c.call('read_table', { path: '../outside/secret.xlsx' });
    expect(r.result['error']).toMatchObject({ details: { roots: [expect.stringContaining('root') as unknown] } });
  });
});

describe('limits and overwrite protection', () => {
  it('never replaces a file without overwrite: true', async () => {
    const before = await readFile(join(root, 'report.xlsx'));
    expect(await code('write_table', { path: 'report.xlsx', rows: ROWS })).toBe('XLV_E_FILE_EXISTS');
    expect(await code('convert', { input: 'sales.csv', output: 'report.xlsx' })).toBe('XLV_E_FILE_EXISTS');
    expect(await readFile(join(root, 'report.xlsx'))).toEqual(before);
  });

  it('refuses files above the size limit, for reading and writing', async () => {
    await writeFile(join(root, 'big.xlsx'), new Uint8Array(65 * 1024));
    expect(await code('inspect_workbook', { path: 'big.xlsx' })).toBe('XLV_E_LIMIT_EXCEEDED');
    const rows = Array.from({ length: 5000 }, (_, i) => ({ id: i, text: `row ${String(i)} ${'x'.repeat(20)}` }));
    expect(await code('write_table', { path: 'huge.xlsx', rows })).toBe('XLV_E_LIMIT_EXCEEDED');
    await expect(stat(join(root, 'huge.xlsx'))).rejects.toThrow();
  });

  it('does not offer write tools without allowWrite', async () => {
    const ro = await connect({ roots: [root] });
    // The SDK answers a call to an unregistered tool with a protocol error or an error result.
    const r = await ro.client.callTool({ name: 'write_table', arguments: { path: 'x.xlsx', rows: ROWS } }).catch((error: unknown) => ({ isError: true, error }));
    expect(r.isError).toBe(true);
    await expect(stat(join(root, 'x.xlsx'))).rejects.toThrow();
    await ro.close();
  });

  it('refuses a root that does not exist', async () => {
    await expect(connect({ roots: [join(base, 'missing')] })).rejects.toMatchObject({ code: 'XLV_E_FILE_READ' });
    await expect(connect({ roots: [] })).rejects.toMatchObject({ code: 'XLV_E_INVALID_OPTION' });
  });
});
