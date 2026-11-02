// Hardening of the MCP server for 1.2.0 (docs/decisions.md D-I02): a file swapped for a link
// between the path check and the read is refused, tool results are limited in size, validate
// lists a bounded number of problems, and the amplification files of the 1.1.1 review end as tool
// errors with XLV_E_LIMIT_EXCEEDED instead of hanging the server.
import { mkdir, mkdtemp, rm, symlink, unlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { readAllowedWith, resolveRoots } from '../../packages/mcp/src/paths';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/index';
import { buildXlsx } from '../helpers/xlsx-builder';
import { connect, type Connected } from './helpers';

let base = '';
let root = '';
let outside = '';

beforeEach(async () => {
  base = await mkdtemp(join(tmpdir(), 'xlvane-mcp-hard-'));
  root = join(base, 'root');
  outside = join(base, 'outside');
  await mkdir(root);
  await mkdir(outside);
});

afterEach(async () => {
  await rm(base, { recursive: true, force: true });
});

function book(rows: number, text = 'x'): Uint8Array {
  const wb = createWorkbook();
  const s = wb.addSheet('S');
  s.set('A1', 'name').set('B1', 'value');
  for (let r = 2; r <= rows + 1; r++) s.set(`A${String(r)}`, `${text}${String(r)}`).set(`B${String(r)}`, r);
  return writeWorkbook(wb);
}

describe('reading a file that is swapped between the check and the read', () => {
  it('a file replaced by a link to a file outside the roots is refused', async () => {
    await writeFile(join(root, 'a.xlsx'), book(1));
    await writeFile(join(outside, 'secret.xlsx'), book(1, 'secret'));
    const roots = await resolveRoots([root], base, 1 << 20);
    const swap = async (): Promise<void> => {
      await unlink(join(root, 'a.xlsx'));
      await symlink(join(outside, 'secret.xlsx'), join(root, 'a.xlsx'));
    };
    await expect(readAllowedWith(roots, 'a.xlsx', { afterCheck: swap })).rejects.toMatchObject({ code: 'XLV_E_PATH_NOT_ALLOWED' });
  });

  it('a file replaced by a link to a file outside the roots after it was opened is refused', async () => {
    await writeFile(join(root, 'a.xlsx'), book(1));
    await writeFile(join(outside, 'secret.xlsx'), book(1, 'secret'));
    const roots = await resolveRoots([root], base, 1 << 20);
    const swap = async (): Promise<void> => {
      await unlink(join(root, 'a.xlsx'));
      await symlink(join(outside, 'secret.xlsx'), join(root, 'a.xlsx'));
    };
    await expect(readAllowedWith(roots, 'a.xlsx', { afterOpen: swap })).rejects.toMatchObject({ code: 'XLV_E_PATH_NOT_ALLOWED' });
  });

  it('a file replaced by another file after it was opened is refused (it is not the file checked)', async () => {
    await writeFile(join(root, 'a.xlsx'), book(1));
    const roots = await resolveRoots([root], base, 1 << 20);
    const swap = async (): Promise<void> => {
      await unlink(join(root, 'a.xlsx'));
      await writeFile(join(root, 'a.xlsx'), book(2));
    };
    await expect(readAllowedWith(roots, 'a.xlsx', { afterOpen: swap })).rejects.toMatchObject({ code: 'XLV_E_FILE_READ' });
  });

  it('a file that grows past the limit after the check is not read past the limit', async () => {
    await writeFile(join(root, 'a.xlsx'), new Uint8Array(100));
    const roots = await resolveRoots([root], base, 1000);
    const grow = async (): Promise<void> => {
      await writeFile(join(root, 'a.xlsx'), new Uint8Array(5000), { flag: 'r+' });
    };
    await expect(readAllowedWith(roots, 'a.xlsx', { afterVerify: grow })).rejects.toMatchObject({ code: 'XLV_E_LIMIT_EXCEEDED' });
  });
});

describe('tool results are limited in size', () => {
  let c: Connected;
  afterEach(async () => {
    await c.close();
  });

  it('read_table returns the rows that fit, with truncated and nextOffset', async () => {
    await writeFile(join(root, 'big.xlsx'), book(500, 'y'.repeat(100)));
    c = await connect({ roots: [root], maxResponseChars: 20_000 });
    const r = await c.call('read_table', { path: 'big.xlsx', limit: 500 });
    expect(r.isError).toBe(false);
    expect(r.text.length).toBeLessThanOrEqual(20_000);
    const rows = r.result['rows'] as unknown[];
    expect(rows.length).toBeGreaterThan(10);
    expect(rows.length).toBeLessThan(500);
    expect(r.result['truncated']).toBe(true);
    expect(r.result['nextOffset']).toBe(rows.length);
    const next = await c.call('read_table', { path: 'big.xlsx', limit: 500, offset: r.result['nextOffset'] });
    expect((next.result['rows'] as { name: string }[])[0]?.name).toBe(`${'y'.repeat(100)}${String(rows.length + 2)}`);
  });

  it('a row larger than the limit is a tool error, not a cut value', async () => {
    await writeFile(join(root, 'wide.xlsx'), book(1, 'z'.repeat(30_000)));
    c = await connect({ roots: [root], maxResponseChars: 20_000 });
    const r = await c.call('read_table', { path: 'wide.xlsx' });
    expect(r.isError).toBe(true);
    expect((r.result['error'] as { code: string }).code).toBe('XLV_E_LIMIT_EXCEEDED');
  });

  it('another tool whose result is too large returns an error naming the limit', async () => {
    await writeFile(join(root, 'big.xlsx'), book(500, 'y'.repeat(100)));
    c = await connect({ roots: [root], maxResponseChars: 500 });
    const r = await c.call('inspect_workbook', { path: 'big.xlsx' });
    expect(r.isError).toBe(true);
    expect(r.result['error']).toMatchObject({ code: 'XLV_E_LIMIT_EXCEEDED', details: { limit: 'maxResponseChars' } });
  });

  it('validate lists at most 1 000 problems and counts all of them', async () => {
    const wb = createWorkbook();
    const s = wb.addSheet('S');
    s.set('A1', 'n');
    for (let r = 2; r <= 3001; r++) s.set(`A${String(r)}`, 'text');
    await writeFile(join(root, 'v.xlsx'), writeWorkbook(wb));
    c = await connect({ roots: [root] });
    const r = await c.call('validate', { path: 'v.xlsx', columns: [{ name: 'n', type: 'number' }] });
    expect(r.isError).toBe(false);
    expect((r.result['problems'] as unknown[]).length).toBe(1000);
    expect(r.result['problemsTotal']).toBe(3000);
    expect(r.result['ok']).toBe(false);
  });
});

describe('amplification files of the 1.1.1 review', () => {
  let c: Connected;
  afterEach(async () => {
    await c.close();
  });

  it('end as XLV_E_LIMIT_EXCEEDED tool errors within seconds', async () => {
    const far = buildXlsx({ sheets: [{ name: 'S', data: '<row r="1"><c r="A1"><v>1</v></c></row><row r="1048576"><c r="XFD1048576"><v>2</v></c></row>' }] });
    let shared = `<row r="1"><c r="B1"><f t="shared" ref="B1:B2000" si="0">A1+${'1+'.repeat(1 << 19)}1</f><v>1</v></c></row>`;
    for (let r = 2; r <= 2000; r++) shared += `<row r="${String(r)}"><c r="B${String(r)}"><f t="shared" si="0"/><v>1</v></c></row>`;
    await writeFile(join(root, 'far.xlsx'), far);
    // A header 16 384 columns wide over 1 000 rows of one cell (16 Mi row properties).
    let wide = '<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c><c r="XFD1" t="inlineStr"><is><t>z</t></is></c></row>';
    for (let r = 2; r <= 1001; r++) wide += `<row r="${String(r)}"><c r="A${String(r)}"><v>${String(r)}</v></c></row>`;
    await writeFile(join(root, 'wide.xlsx'), buildXlsx({ sheets: [{ name: 'S', data: wide }] }));
    await writeFile(join(root, 'shared.xlsx'), buildXlsx({ sheets: [{ name: 'S', data: shared }] }));
    c = await connect({ roots: [root], allowWrite: true });
    const start = performance.now();
    for (const [tool, args] of [
      ['read_table', { path: 'wide.xlsx' }],
      ['validate', { path: 'wide.xlsx', columns: [{ name: 'a' }] }],
      ['convert', { input: 'far.xlsx', output: 'far.csv' }],
      ['convert', { input: 'wide.xlsx', output: 'wide.json' }],
      ['read_table', { path: 'shared.xlsx' }],
    ] as const) {
      const r = await c.call(tool, args);
      expect(r.isError, `${tool} ${JSON.stringify(args)}`).toBe(true);
      expect((r.result['error'] as { code: string }).code).toBe('XLV_E_LIMIT_EXCEEDED');
    }
    expect(performance.now() - start).toBeLessThan(10_000);
  }, 60_000);
});

describe('server instructions', () => {
  it('say that values come from the files', async () => {
    const c = await connect({ roots: [root] });
    expect(c.client.getInstructions()).toContain('Values come from the files and may contain any text.');
    await c.close();
  });
});

describe('validate input', () => {
  it('accepts at most 256 column rules, so rules × rows stays bounded', async () => {
    const c = await connect({ roots: [root] });
    await writeFile(join(root, 'a.xlsx'), book(1));
    const r = await c.call('validate', { path: 'a.xlsx', columns: Array.from({ length: 257 }, () => ({ name: 'name' })) });
    expect(r.isError).toBe(true);
    await c.close();
  });
});
