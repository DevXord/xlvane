// Observes how LibreOffice treats a row format (`row@s` with `customFormat="1"`) and a default
// font written by xlvane when it re-saves the file (W2-04a, W2-12a): it keeps the row formats
// and font 0, and gives a cell of the row that has no format of its own (`s` 0) the row's format,
// while a cell with its own format keeps it.
// Skipped when `soffice` is not installed, unless XLV_REQUIRE_LIBREOFFICE=1 (set in CI).
import { spawnSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';

const SOFFICE = process.env['XLV_SOFFICE'] ?? 'soffice';
const available = spawnSync(SOFFICE, ['--version'], { encoding: 'utf8', timeout: 60_000 }).status === 0;
const required = process.env['XLV_REQUIRE_LIBREOFFICE'] === '1';

describe.skipIf(!available && !required)('LibreOffice and row formats', () => {
  it('keeps row formats and the default font when it re-saves the file', { timeout: 200_000 }, async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-lo-rowstyle-'));
    try {
      const wb = createWorkbook({ defaultFont: { name: 'Arial', size: 10 } });
      wb.addSheet('S')
        .set('A1', 'plain')
        .set('B1', 'italic', { font: { italic: true } })
        .row(1, { style: { font: { bold: true } } })
        .row(3, { style: { fill: { color: 'FFFF00' } } });
      const file = join(dir, 'rows.xlsx');
      await writeFile(file, writeWorkbook(wb));
      const profile = pathToFileURL(join(dir, 'profile')).href;
      const result = spawnSync(SOFFICE, [`-env:UserInstallation=${profile}`, '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', join(dir, 'out'), file], {
        encoding: 'utf8',
        timeout: 180_000,
      });
      expect(result.status, result.stderr).toBe(0);
      const resaved = readWorkbook(await readFile(join(dir, 'out', 'rows.xlsx')));
      expect(resaved.defaultFont).toMatchObject({ name: 'Arial', size: 10 });
      const s = resaved.getSheet('S');
      expect(s.rowInfo(1)?.style?.font).toMatchObject({ name: 'Arial', size: 10, bold: true });
      // A row with only a format is kept (LibreOffice 24.2.7.2 also writes its other attributes).
      expect(s.rowInfo(3)?.style?.fill).toMatchObject({ pattern: 'solid', fgColor: 'FFFFFF00' });
      // LibreOffice 24.2.7.2 (observed 2026-11-03): the cell without a format of its own is
      // written back with the row's format; the italic cell keeps its own format.
      expect(s.cell('A1')?.style.font).toMatchObject({ bold: true });
      expect(s.cell('B1')?.style.font).toMatchObject({ italic: true });
      expect(s.cell('B1')?.style.font?.bold).toBeUndefined();
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
