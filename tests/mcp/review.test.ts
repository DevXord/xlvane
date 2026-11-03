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
    for (const option of ['--root', '--allow-write', '--max-file-size', '--max-rows', '--max-response-chars', '--max-cells']) expect(r.stdout).toContain(option);
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

describe('1.2.3: a symbolic link inside a root reveals nothing about what lies outside (D-J01)', () => {
  it('reading and writing through a link to an outside directory give one answer', async () => {
    const outside = join(base, 'outside');
    await mkdir(outside);
    await writeFile(join(outside, 'secret.xlsx'), texts(1));
    await writeFile(join(outside, 'secret.txt'), 'x');
    await mkdir(join(outside, 'sub'));
    const { symlink } = await import('node:fs/promises');
    await symlink(outside, join(root, 'link-dir'));
    c = await connect({ roots: [root], allowWrite: true });
    const answers: string[] = [];
    for (const path of ['link-dir/secret.xlsx', 'link-dir/missing.xlsx', 'link-dir/secret.txt/x.xlsx', 'link-dir/nodir/x.xlsx']) {
      const r = await c.call('inspect_workbook', { path });
      const e = r.result['error'] as { code: string; message: string };
      answers.push(`${e.code} ${e.message.replace(path, 'P')}`);
    }
    for (const path of ['link-dir/new.xlsx', 'link-dir/nodir/new.xlsx', 'link-dir/secret.txt/new.xlsx', 'link-dir/sub/new.xlsx']) {
      const r = await c.call('write_table', { path, rows: [{ a: 1 }] });
      const e = r.result['error'] as { code: string; message: string };
      answers.push(`${e.code} ${e.message.replace(path, 'P')}`);
    }
    expect(new Set(answers.map((a) => a.split(' ')[0]))).toEqual(new Set(['XLV_E_PATH_NOT_ALLOWED']));
    expect(answers.some((a) => /ENOENT|ENOTDIR|does not exist/.test(a))).toBe(false);
  });

  it('a missing file inside a root is still reported as missing', async () => {
    c = await connect({ roots: [root], allowWrite: true });
    const r = await c.call('inspect_workbook', { path: 'missing.xlsx' });
    expect((r.result['error'] as { code: string }).code).toBe('XLV_E_FILE_READ');
    const w = await c.call('write_table', { path: 'nodir/new.xlsx', rows: [{ a: 1 }] });
    expect((w.result['error'] as { code: string }).code).toBe('XLV_E_FILE_WRITE');
  });
});

describe('1.2.3: stderr of xlvane-mcp, --max-cells and JSON output (D-J01)', () => {
  const run = async (args: string[]): Promise<{ code: number; err: string; out: string }> => {
    const { runWith } = await import('../../packages/mcp/src/server');
    let err = '';
    let out = '';
    const code = await runWith(args, { stdout: (t) => (out += t), stderr: (t) => (err += t), cwd: base, serve: async () => {} });
    return { code, err, out };
  };

  it('writes no escape sequence from arguments or directory names to stderr', async () => {
    const esc = '\u001b]0;pwned\u0007';
    for (const args of [[`--bogus${esc}`], ['--root', `nodir${esc}`], ['--root', root, '--max-rows', `x${esc}`]]) {
      const r = await run(args);
      expect(r.code).not.toBe(0);
      expect(r.err).not.toMatch(/[\u0000-\u0008\u000b-\u001f\u007f-\u009f]/);
      expect(r.err).toContain('\\u001b');
    }
    const odd = join(base, `dir${esc}`);
    await mkdir(odd);
    const r = await run(['--root', odd]);
    expect(r.code).toBe(0);
    expect(r.err).not.toMatch(/[\u0000-\u0008\u000b-\u001f\u007f-\u009f]/);
  });

  it('--max-cells limits the cells of every workbook read', async () => {
    expect((await run(['--help'])).out).toContain('--max-cells');
    await writeFile(join(root, 'v.xlsx'), texts(30));
    c = await connect({ roots: [root], maxCells: 10 });
    const r = await c.call('read_table', { path: 'v.xlsx' });
    expect((r.result['error'] as { code: string }).code).toBe('XLV_E_LIMIT_EXCEEDED');
    await c.close();
    c = await connect({ roots: [root] });
    expect((await c.call('read_table', { path: 'v.xlsx' })).isError).toBe(false);
  });

  it('convert to JSON of a header repeated in many rows ends as a limit error, not an internal one', async () => {
    // 4 000 000 header characters in 140 rows: 560 million characters of JSON, more than a
    // JavaScript string can hold (before: "Invalid string length" as XLV_E_INTERNAL).
    let seed = 7;
    let header = '';
    for (let i = 0; i < 4_000_000; i++) header += ((seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648) < 0.01 ? 'x' : 'h';
    let data = `<row r="1"><c r="A1" t="inlineStr"><is><t>${header}</t></is></c></row>`;
    for (let r = 2; r <= 141; r++) data += `<row r="${String(r)}"><c r="A${String(r)}"><v>${String(r)}</v></c></row>`;
    const { buildXlsx } = await import('../helpers/xlsx-builder');
    await writeFile(join(root, 'wide.xlsx'), buildXlsx({ sheets: [{ name: 'S', data }] }));
    c = await connect({ roots: [root], allowWrite: true });
    const start = performance.now();
    const r = await c.call('convert', { input: 'wide.xlsx', output: 'wide.json' });
    expect((r.result['error'] as { code: string }).code).toBe('XLV_E_LIMIT_EXCEEDED');
    expect(performance.now() - start).toBeLessThan(10_000);
  }, 60_000);

  it('the JSON written row by row is the text JSON.stringify gives', async () => {
    await writeFile(join(root, 'v.xlsx'), texts(3));
    c = await connect({ roots: [root], allowWrite: true });
    expect((await c.call('convert', { input: 'v.xlsx', output: 'v.json' })).isError).toBe(false);
    const { readFile } = await import('node:fs/promises');
    const text = await readFile(join(root, 'v.json'), 'utf8');
    expect(text).toBe(`${JSON.stringify(JSON.parse(text), null, 2)}\n`);
    expect(JSON.parse(text)).toEqual([{ x: 'text' }, { x: 'text' }, { x: 'text' }]);
  });
});

describe('1.2.3: dangling links and --max-cells for every tool (independent review of the patch, D-J01)', () => {
  it('a link inside a root to a missing outside file or directory gives the same answer as to an existing one', async () => {
    const outside = join(base, 'outside');
    await mkdir(outside);
    const { symlink } = await import('node:fs/promises');
    await symlink(join(outside, 'secret.xlsx'), join(root, 'l.xlsx'));
    await symlink(join(outside, 'dir'), join(root, 'd'));
    c = await connect({ roots: [root], allowWrite: true });
    const codes: string[] = [];
    codes.push(((await c.call('inspect_workbook', { path: 'l.xlsx' })).result['error'] as { code: string }).code);
    codes.push(((await c.call('write_table', { path: 'd/x.xlsx', rows: [{ a: 1 }] })).result['error'] as { code: string }).code);
    codes.push(((await c.call('write_table', { path: 'l.xlsx', rows: [{ a: 1 }] })).result['error'] as { code: string }).code);
    await writeFile(join(outside, 'secret.xlsx'), texts(1));
    await mkdir(join(outside, 'dir'));
    codes.push(((await c.call('inspect_workbook', { path: 'l.xlsx' })).result['error'] as { code: string }).code);
    codes.push(((await c.call('write_table', { path: 'd/x.xlsx', rows: [{ a: 1 }] })).result['error'] as { code: string }).code);
    expect(codes).toEqual(Array(5).fill('XLV_E_PATH_NOT_ALLOWED'));
  });

  it('--max-cells also limits edit_cells and convert from CSV', async () => {
    await writeFile(join(root, 'v.xlsx'), texts(30));
    await writeFile(join(root, 'v.csv'), `a,b\n${'1,2\n'.repeat(30)}`);
    c = await connect({ roots: [root], allowWrite: true, maxCells: 10 });
    const e = await c.call('edit_cells', { path: 'v.xlsx', output: 'w.xlsx', edits: [{ sheet: 'S', cell: 'A1', value: 'x' }] });
    expect((e.result['error'] as { code: string } | undefined)?.code).toBe('XLV_E_LIMIT_EXCEEDED');
    const v = await c.call('convert', { input: 'v.csv', output: 'v2.xlsx' });
    expect((v.result['error'] as { code: string } | undefined)?.code).toBe('XLV_E_LIMIT_EXCEEDED');
  });
});

describe('1.3.0: ".." in the text of a link is resolved as the system does (D-K03)', () => {
  it('a link whose text climbs out through another link gives one answer', async () => {
    const outside = join(base, 'out');
    await mkdir(join(outside, 'dir'), { recursive: true });
    await writeFile(join(outside, 'secret.xlsx'), texts(1));
    const { symlink } = await import('node:fs/promises');
    await symlink(join(outside, 'dir'), join(root, 's'));
    await symlink('s/..', join(root, 'L'));
    await mkdir(join(base, 'home', 'proj'), { recursive: true });
    await mkdir(join(base, 'home', '.ssh'));
    await symlink(join(base, 'home', 'proj'), join(root, 'proj'));
    await symlink('..', join(base, 'home', 'proj', 'parent'));
    c = await connect({ roots: [root], allowWrite: true });
    const codes: string[] = [];
    for (const path of ['L/secret.xlsx', 'L/missing.xlsx', 'proj/parent/.ssh/x.xlsx', 'proj/parent/.gnupg/x.xlsx']) {
      codes.push(((await c.call('inspect_workbook', { path })).result['error'] as { code: string }).code);
    }
    for (const path of ['proj/parent/.ssh/new.xlsx', 'proj/parent/.gnupg/new.xlsx']) {
      codes.push(((await c.call('write_table', { path, rows: [{ a: 1 }] })).result['error'] as { code: string }).code);
    }
    expect(codes).toEqual(Array(6).fill('XLV_E_PATH_NOT_ALLOWED'));
  });

  it('a link inside the root whose text uses ".." to stay inside still works', async () => {
    await mkdir(join(root, 'a', 'b'), { recursive: true });
    await writeFile(join(root, 'a', 'x.xlsx'), texts(1));
    const { symlink } = await import('node:fs/promises');
    await symlink('b/..', join(root, 'a', 'up'));
    c = await connect({ roots: [root] });
    expect((await c.call('inspect_workbook', { path: 'a/up/x.xlsx' })).isError).toBe(false);
    expect(((await c.call('inspect_workbook', { path: 'a/up/missing.xlsx' })).result['error'] as { code: string }).code).toBe('XLV_E_FILE_READ');
  });
});

describe('1.3.0: --max-cells limits the cells the write tools create (D-K03)', () => {
  it('write_table and convert from JSON refuse more cells than --max-cells', async () => {
    const rows = Array.from({ length: 30 }, (_, i) => ({ a: i, b: `x${String(i)}` }));
    await writeFile(join(root, 'r.json'), JSON.stringify(rows));
    c = await connect({ roots: [root], allowWrite: true, maxCells: 50 });
    const w = await c.call('write_table', { path: 'w.xlsx', rows });
    expect(w.result['error']).toMatchObject({ code: 'XLV_E_LIMIT_EXCEEDED', details: { limit: 'maxCells' } });
    const v = await c.call('convert', { input: 'r.json', output: 'r.xlsx' });
    expect(v.result['error']).toMatchObject({ code: 'XLV_E_LIMIT_EXCEEDED', details: { limit: 'maxCells' } });
    const small = await c.call('write_table', { path: 's.xlsx', rows: rows.slice(0, 10) });
    expect(small.isError).toBe(false);
    const bad = await c.call('convert', { input: 'missing-brace.json', output: 'x.xlsx' });
    expect(bad.isError).toBe(true);
  });
});

describe('1.3.0: review of D-K03 (MCP paths and names)', () => {
  it('overwrites a file whose name is 245 bytes of CJK characters', async () => {
    const roots = await resolveRoots([root], base, 1 << 20);
    const name = `${'表'.repeat(80)}.csv`;
    await writeFile(join(root, name), 'old');
    await expect(writeAllowed(roots, name, 'new', true)).resolves.toMatchObject({ replaced: true });
  });

  it('a backslash in the text of a link is part of a name on POSIX, not a separator', async () => {
    if (process.platform === 'win32') return;
    const { symlink } = await import('node:fs/promises');
    await symlink('..\\..', join(root, 'L'));
    c = await connect({ roots: [root] });
    const r = await c.call('inspect_workbook', { path: 'L/x.xlsx' });
    expect((r.result['error'] as { code: string }).code).toBe('XLV_E_FILE_READ');
  });
});
