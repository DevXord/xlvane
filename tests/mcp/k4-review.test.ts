// Findings of the full review of 1.3.0 for the MCP server (docs/decisions.md D-K04). Each test
// failed before its fix.
import { mkdir, mkdtemp, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, relative } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/index';
import { connect, type Connected } from './helpers';

let base = '';
let root = '';
let c: Connected | null = null;

beforeEach(async () => {
  base = await mkdtemp(join(tmpdir(), 'xlvane-mcp-k4-'));
  root = join(base, 'root');
  await mkdir(root);
});

afterEach(async () => {
  await c?.close();
  c = null;
  await rm(base, { recursive: true, force: true });
});

function book(): Uint8Array {
  const wb = createWorkbook();
  wb.addSheet('S').set('A1', 'x');
  return writeWorkbook(wb);
}

const code = (r: { result: Record<string, unknown> }): string | undefined => (r.result['error'] as { code?: string } | undefined)?.code;

describe('a link that leaves the roots and comes back reveals nothing about what lies outside', () => {
  it.skipIf(process.platform === 'win32')('reads: one answer whether the outside name is a directory, a file or missing', async () => {
    const outside = join(base, 'outside');
    await mkdir(join(outside, 'dir'), { recursive: true });
    await writeFile(join(outside, 'file'), 'x');
    await writeFile(join(root, 'ok.xlsx'), book());
    const back = relative(outside, root);
    for (const probe of ['dir', 'file', 'missing']) {
      await symlink(`${outside}/${probe}/../${back}/ok.xlsx`, join(root, `abs-${probe}`));
      await symlink(`../outside/${probe}/../${back}/ok.xlsx`, join(root, `rel-${probe}`));
    }
    // A directory outside that does not exist at all, on the way back in.
    await symlink(`../nowhere/../root/ok.xlsx`, join(root, 'rel-nowhere'));
    c = await connect({ roots: [root] });
    const codes: (string | undefined)[] = [];
    for (const p of ['abs-dir', 'abs-file', 'abs-missing', 'rel-dir', 'rel-file', 'rel-missing', 'rel-nowhere']) {
      codes.push(code(await c.call('inspect_workbook', { path: p })));
    }
    expect(codes).toEqual(Array(7).fill('XLV_E_PATH_NOT_ALLOWED'));
  });

  it.skipIf(process.platform === 'win32')('writes: one answer whether the outside name is a directory, a file or missing', async () => {
    const outside = join(base, 'outside');
    await mkdir(join(outside, 'dir'), { recursive: true });
    await writeFile(join(outside, 'file'), 'x');
    await mkdir(join(root, 'sub'));
    for (const probe of ['dir', 'file', 'missing']) await symlink(`../outside/${probe}/../../root/sub`, join(root, `d-${probe}`));
    c = await connect({ roots: [root], allowWrite: true });
    const codes: (string | undefined)[] = [];
    for (const probe of ['dir', 'file', 'missing']) codes.push(code(await c.call('write_table', { path: `d-${probe}/w.xlsx`, rows: [{ a: 1 }] })));
    expect(codes).toEqual(Array(3).fill('XLV_E_PATH_NOT_ALLOWED'));
  });

  it.skipIf(process.platform === 'win32')('links that stay inside, or pass only through parents of the root, still work', async () => {
    await mkdir(join(root, 'a', 'b'), { recursive: true });
    await writeFile(join(root, 'a', 'x.xlsx'), book());
    await symlink('b/..', join(root, 'a', 'up'));
    await symlink(`../${relative(base, root)}/a/x.xlsx`, join(root, 'via-parent.xlsx'));
    await symlink(join(root, 'a', 'x.xlsx'), join(root, 'absolute.xlsx'));
    c = await connect({ roots: [root] });
    for (const path of ['a/up/x.xlsx', 'via-parent.xlsx', 'absolute.xlsx']) expect((await c.call('inspect_workbook', { path })).isError).toBe(false);
    expect(code(await c.call('inspect_workbook', { path: 'a/up/missing.xlsx' }))).toBe('XLV_E_FILE_READ');
  });

  it.skipIf(process.platform === 'win32')('a root given through a link still works', async () => {
    await writeFile(join(root, 'ok.xlsx'), book());
    await symlink(root, join(base, 'root-link'));
    c = await connect({ roots: [join(base, 'root-link')] });
    expect((await c.call('inspect_workbook', { path: 'ok.xlsx' })).isError).toBe(false);
    expect((await c.call('inspect_workbook', { path: join(base, 'root-link', 'ok.xlsx') })).isError).toBe(false);
  });
});

describe('edit_cells keeps within --max-cells', () => {
  it('appending more rows than --max-cells allows is refused before the file is read', async () => {
    c = await connect({ roots: [root], allowWrite: true, maxCells: 50 });
    const rows = Array.from({ length: 60 }, (_, i) => [i]);
    const r = await c.call('edit_cells', { path: 'missing.xlsx', append: { sheet: 'S', rows }, output: 'x.xlsx' });
    expect(r.result['error']).toMatchObject({ code: 'XLV_E_LIMIT_EXCEEDED', details: { limit: 'maxCells' } });
  });

  it('more edits than --max-cells are refused', async () => {
    await writeFile(join(root, 'b.xlsx'), book());
    c = await connect({ roots: [root], allowWrite: true, maxCells: 50 });
    const edits = Array.from({ length: 80 }, (_, i) => ({ sheet: 'S', cell: `B${String(i + 1)}`, value: i }));
    const r = await c.call('edit_cells', { path: 'b.xlsx', edits, output: 'x.xlsx' });
    expect(r.result['error']).toMatchObject({ code: 'XLV_E_LIMIT_EXCEEDED', details: { limit: 'maxCells' } });
  });

  it('a file the read tools refuse is not written: existing cells count too', async () => {
    const wb = createWorkbook();
    const s = wb.addSheet('S');
    for (let r = 1; r <= 45; r++) s.set(`A${String(r)}`, r);
    await writeFile(join(root, 'c.xlsx'), writeWorkbook(wb));
    c = await connect({ roots: [root], allowWrite: true, maxCells: 50 });
    const r = await c.call('edit_cells', { path: 'c.xlsx', append: { sheet: 'S', rows: Array.from({ length: 10 }, (_, i) => [i]) }, output: 'x.xlsx' });
    expect(r.result['error']).toMatchObject({ code: 'XLV_E_LIMIT_EXCEEDED' });
    const ok = await c.call('edit_cells', { path: 'c.xlsx', append: { sheet: 'S', rows: [[1], [2]] }, output: 'y.xlsx' });
    expect(ok.isError).toBe(false);
    expect((await c.call('inspect_workbook', { path: 'y.xlsx' })).isError).toBe(false);
  });
});

describe('messages', () => {
  it('validate refuses a sheet that does not exist, also without columns', async () => {
    await writeFile(join(root, 'b.xlsx'), book());
    c = await connect({ roots: [root] });
    expect(code(await c.call('validate', { path: 'b.xlsx', sheet: 'Nope' }))).toBe('XLV_E_SHEET_NOT_FOUND');
    expect((await c.call('validate', { path: 'b.xlsx', sheet: 'S' })).isError).toBe(false);
  });

  it.skipIf(process.platform === 'win32')('writing to a link whose target is inside does not say the path is outside', async () => {
    await writeFile(join(root, 'b.xlsx'), book());
    await symlink(join(root, 'b.xlsx'), join(root, 'link.xlsx'));
    c = await connect({ roots: [root], allowWrite: true });
    const r = await c.call('write_table', { path: 'link.xlsx', rows: [{ a: 1 }], overwrite: true });
    expect(code(r)).toBe('XLV_E_PATH_NOT_ALLOWED');
    expect(r.text).not.toContain('outside');
    expect(r.text).toContain('symbolic link');
  });
});
