// Findings of the independent review of the MCP server (docs/decisions.md D-I04). Each test failed
// before its fix.
import { chmod, mkdir, mkdtemp, readdir, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { readAllowedWith, resolveRoots, writeAllowed } from '../../packages/mcp/src/paths';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/index';
import { cli } from '../cli/helpers';
import { connect, type Connected } from './helpers';

let base = '';
let root = '';
let c: Connected | null = null;

beforeEach(async () => {
  base = await mkdtemp(join(tmpdir(), 'xlvane-mcp-review-'));
  root = join(base, 'root');
  await mkdir(root);
});

afterEach(async () => {
  await c?.close();
  c = null;
  await rm(base, { recursive: true, force: true });
});

function texts(rows: number): Uint8Array {
  const wb = createWorkbook();
  const s = wb.addSheet('S');
  s.set('A1', 'x');
  for (let r = 2; r <= rows + 1; r++) s.set(`A${String(r)}`, 'text');
  return writeWorkbook(wb);
}

describe('validate', () => {
  it('stops collecting problems at the limit: 20 000 rows of wrong type take little time', async () => {
    await writeFile(join(root, 'v.xlsx'), texts(20_000));
    c = await connect({ roots: [root] });
    const start = performance.now();
    const r = await c.call('validate', { path: 'v.xlsx', columns: [{ name: 'x', type: 'number' }] });
    expect(performance.now() - start).toBeLessThan(5000);
    expect((r.result['problems'] as unknown[]).length).toBe(1000);
    expect(r.result['problemsTotal']).toBe(20_000);
  });

  it('refuses repeated column names in the rules', async () => {
    await writeFile(join(root, 'v.xlsx'), texts(2));
    c = await connect({ roots: [root] });
    const r = await c.call('validate', { path: 'v.xlsx', columns: [{ name: 'x' }, { name: 'x', type: 'number' }] });
    expect(r.isError).toBe(true);
    expect((r.result['error'] as { code: string }).code).toBe('XLV_E_INVALID_INPUT');
  });
});

describe('replacing a file', () => {
  it('uses a temporary name nobody else has, and never deletes a file it did not create', async () => {
    const roots = await resolveRoots([root], base, 1 << 20);
    await writeFile(join(root, 'a.csv'), 'old');
    for (let i = 0; i < 20; i++) await writeFile(join(root, `.a.csv.xlvane-${String(Date.now() + i)}.tmp`), 'mine');
    const before = (await readdir(root)).filter((n) => n.endsWith('.tmp')).sort();
    await writeAllowed(roots, 'a.csv', 'new', true);
    expect((await readdir(root)).filter((n) => n.endsWith('.tmp')).sort()).toEqual(before);
  });

  it('two replacements of one file at the same time both succeed', async () => {
    const roots = await resolveRoots([root], base, 1 << 20);
    await writeFile(join(root, 'a.csv'), 'old');
    const results = await Promise.allSettled([writeAllowed(roots, 'a.csv', 'one', true), writeAllowed(roots, 'a.csv', 'two', true)]);
    expect(results.map((r) => r.status)).toEqual(['fulfilled', 'fulfilled']);
  });

  it('keeps the permissions of the replaced file', async () => {
    const roots = await resolveRoots([root], base, 1 << 20);
    await writeFile(join(root, 'a.csv'), 'old');
    await chmod(join(root, 'a.csv'), 0o600);
    await writeAllowed(roots, 'a.csv', 'new', true);
    expect((await stat(join(root, 'a.csv'))).mode & 0o777).toBe(0o600);
  });

  it('refuses to replace a read-only file', async () => {
    const roots = await resolveRoots([root], base, 1 << 20);
    await writeFile(join(root, 'a.csv'), 'old');
    await chmod(join(root, 'a.csv'), 0o444);
    await expect(writeAllowed(roots, 'a.csv', 'new', true)).rejects.toMatchObject({ code: 'XLV_E_FILE_WRITE' });
  });
});

describe('paths outside the roots give one answer, whatever exists there', () => {
  it('reading and writing', async () => {
    c = await connect({ roots: [root], allowWrite: true });
    const read = async (path: string) => (await c!.call('read_table', { path })).result['error'] as { code: string; message: string };
    const write = async (path: string) => (await c!.call('write_table', { path, rows: [{ a: 1 }] })).result['error'] as { code: string; message: string };
    const missing = join(base, 'nope', 'x.xlsx');
    const existing = join(base, 'here.xlsx');
    await writeFile(existing, texts(1));
    for (const result of [await read(missing), await read(existing), await read(join(existing, 'x')), await write(missing), await write(join(base, 'x.xlsx'))]) {
      expect(result.code).toBe('XLV_E_PATH_NOT_ALLOWED');
      expect(result.message).not.toContain('symbolic');
    }
  });

  it('a file swapped for a link to a large outside file reveals nothing about it', async () => {
    await writeFile(join(root, 'a.xlsx'), texts(1));
    await writeFile(join(base, 'big.bin'), new Uint8Array(123_457));
    const roots = await resolveRoots([root], base, 1000);
    const { unlink, symlink } = await import('node:fs/promises');
    const swap = async (): Promise<void> => {
      await unlink(join(root, 'a.xlsx'));
      await symlink(join(base, 'big.bin'), join(root, 'a.xlsx'));
    };
    const error = await readAllowedWith(roots, 'a.xlsx', { afterCheck: swap }).catch((e: unknown) => e);
    expect(error).toMatchObject({ code: 'XLV_E_PATH_NOT_ALLOWED' });
    expect(String((error as Error).message)).not.toContain('123457');
  });
});

describe('special files', () => {
  it('a FIFO inside a root is refused without blocking', async () => {
    const { spawnSync } = await import('node:child_process');
    if (spawnSync('mkfifo', [join(root, 'f.xlsx')], { encoding: 'utf8' }).status !== 0) return;
    const roots = await resolveRoots([root], base, 1 << 20);
    const start = performance.now();
    await expect(readAllowedWith(roots, 'f.xlsx', {})).rejects.toMatchObject({ code: 'XLV_E_FILE_READ' });
    expect(performance.now() - start).toBeLessThan(2000);
  });
});

describe('tool errors and inputs', () => {
  it('an error result is limited in size too', async () => {
    c = await connect({ roots: [root], maxResponseChars: 2000 });
    const r = await c.call('read_table', { path: `${'p'.repeat(100_000)}.xlsx` });
    expect(r.isError).toBe(true);
    expect(r.text.length).toBeLessThanOrEqual(2000);
  });

  it('invalid UTF-8 in a CSV file is an input error, not an internal one', async () => {
    await writeFile(join(root, 'bad.csv'), new Uint8Array([0x61, 0x0a, 0xff, 0xfe, 0x0a]));
    c = await connect({ roots: [root], allowWrite: true });
    const r = await c.call('convert', { input: 'bad.csv', output: 'bad.xlsx' });
    expect((r.result['error'] as { code: string }).code).toBe('XLV_E_INVALID_INPUT');
  });

  it('an impossible date is refused instead of moved to another day', async () => {
    c = await connect({ roots: [root], allowWrite: true });
    for (const date of ['2024-02-30', '2023-02-29', '2024-01-31T25:00']) {
      const r = await c.call('write_table', { path: `d-${date.replace(/:/g, '')}.xlsx`, rows: [{ a: { date } }] });
      expect((r.result['error'] as { code: string } | undefined)?.code, date).toBe('XLV_E_INVALID_VALUE');
    }
    const ok = await c.call('write_table', { path: 'ok.xlsx', rows: [{ a: { date: '2024-02-29T23:59:59' } }] });
    expect(ok.isError).toBe(false);
  });
});

describe('xlvane mcp', () => {
  const stub = (code: number, seen: string[][]) => async (specifier: string) => {
    if (specifier !== '@xlvane/mcp') throw new Error('unexpected');
    return { runCli: async (args: readonly string[]) => (seen.push([...args]), code) };
  };

  it('refuses --json, whose output would mix with the protocol on stdout', async () => {
    const seen: string[][] = [];
    const r = await cli(['mcp', '--json', '--root', '.'], { cwd: base, importModule: stub(0, seen) });
    expect(r.code).toBe(2);
    expect(seen).toEqual([]);
  });

  it('passes the server arguments without --quiet and --verbose, and keeps the server exit code', async () => {
    const seen: string[][] = [];
    expect((await cli(['mcp', '--quiet', '--root', '.', '--verbose'], { cwd: base, importModule: stub(0, seen) })).code).toBe(0);
    expect(seen).toEqual([['--root', '.']]);
    expect((await cli(['mcp', '--root', '.'], { cwd: base, importModule: stub(2, seen) })).code).toBe(2);
    expect((await cli(['mcp', '--root', '.'], { cwd: base, importModule: stub(3, seen) })).code).toBe(3);
  });

  it('help names the server options', async () => {
    const r = await cli(['mcp', '--help'], { cwd: base });
    for (const option of ['--root', '--allow-write', '--max-file-size', '--max-rows', '--max-response-chars']) expect(r.stdout).toContain(option);
  });

  it('a missing dependency of @xlvane/mcp is not reported as @xlvane/mcp missing', async () => {
    const importModule = async (): Promise<unknown> => {
      throw Object.assign(new Error("Cannot find package '@modelcontextprotocol/server' imported from /x/node_modules/@xlvane/mcp/dist/index.js"), { code: 'ERR_MODULE_NOT_FOUND' });
    };
    const r = await cli(['mcp', '--root', '.'], { cwd: base, importModule });
    expect(r.stderr).not.toContain('XLV_E_NOT_INSTALLED');
    expect(r.stderr).toContain('@modelcontextprotocol/server');
  });
});
