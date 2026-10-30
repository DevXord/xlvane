import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { inspectFile, readWorkbookFile, writeWorkbookFile } from '../../packages/xlvane/src/node';
import { createWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvErrorAsync } from '../helpers/expect-error';
import { buildXlsx, c, row } from '../helpers/xlsx-builder';

let dir = '';
let file = '';

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'xlvane-test-'));
  file = join(dir, 'report.xlsx');
  await writeFile(file, buildXlsx({ sheets: [{ name: 'Data', data: row(1, c('A1', 42)) }] }));
});

afterAll(async () => {
  await rm(dir, { recursive: true, force: true });
});

describe('xlvane/node', () => {
  it('reads a workbook from a path or a file URL', async () => {
    expect((await readWorkbookFile(file)).getSheet('Data').cell('A1')?.value).toBe(42);
    expect((await readWorkbookFile(pathToFileURL(file))).sheets).toHaveLength(1);
  });

  it('inspects a file', async () => {
    expect((await inspectFile(file)).sheets[0]).toMatchObject({ name: 'Data', usedRange: 'A1' });
  });

  it('reports a missing file with XLV_E_FILE_READ', async () => {
    const error = await expectXlvErrorAsync(() => readWorkbookFile(join(dir, 'missing.xlsx')), 'XLV_E_FILE_READ');
    expect(error.details['code']).toBe('ENOENT');
    expect(error.hint).toMatch(/does not exist/);
  });

  it('rejects a directory', async () => {
    await expectXlvErrorAsync(() => readWorkbookFile(dir), 'XLV_E_FILE_READ');
  });

  it('checks maxInputBytes before reading the file', async () => {
    await expectXlvErrorAsync(() => inspectFile(file, { limits: { maxInputBytes: 10 } }), 'XLV_E_LIMIT_EXCEEDED', {
      limit: 'maxInputBytes',
    });
  });

  it('validates the limits before touching the file system', async () => {
    await expectXlvErrorAsync(() => readWorkbookFile(join(dir, 'missing.xlsx'), { limits: { maxCell: 1 } as never }), 'XLV_E_INVALID_OPTION');
  });
});

describe('writeWorkbookFile', () => {
  const book = () => {
    const wb = createWorkbook();
    wb.addSheet('Out').set('A1', 'written');
    return wb;
  };

  it('writes a file that reads back', async () => {
    const out = join(dir, 'out.xlsx');
    await writeWorkbookFile(book(), out);
    expect((await readWorkbookFile(out)).getSheet('Out').cell('A1')?.value).toBe('written');
  });

  it('never overwrites an existing file unless asked', async () => {
    const out = join(dir, 'exists.xlsx');
    await writeFile(out, 'keep me');
    await expectXlvErrorAsync(() => writeWorkbookFile(book(), out), 'XLV_E_FILE_EXISTS');
    expect(await readFile(out, 'utf8')).toBe('keep me');
    await writeWorkbookFile(book(), out, { overwrite: true });
    expect((await readWorkbookFile(out)).sheets[0]?.name).toBe('Out');
  });

  it('reports a missing directory with XLV_E_FILE_WRITE', async () => {
    const error = await expectXlvErrorAsync(() => writeWorkbookFile(book(), join(dir, 'no', 'such', 'dir.xlsx')), 'XLV_E_FILE_WRITE');
    expect(error.details['code']).toBe('ENOENT');
  });

  it('validates the overwrite option', async () => {
    await expectXlvErrorAsync(() => writeWorkbookFile(book(), join(dir, 'x.xlsx'), { overwrite: 'yes' as never }), 'XLV_E_INVALID_OPTION');
  });
});
