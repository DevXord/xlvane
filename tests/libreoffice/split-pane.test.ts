// LibreOffice reads pane@topLeftCell of a split window as the position of the right and bottom
// panes (sc/source/filter/oox/viewsettings.cxx: maSecondPos, PositionRight, PositionBottom) and
// writes it back from them (sc/source/filter/excel/xeview.cxx XclExpPane::SaveXml); this test
// observes that a split view written by xlvane keeps its pane@topLeftCell when LibreOffice
// re-saves the file (W2-69, D-F02).
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

describe.skipIf(!available && !required)('LibreOffice and split windows', () => {
  it('keeps the top-left cell of the bottom-right pane when it re-saves the file', { timeout: 200_000 }, async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-lo-split-'));
    try {
      const wb = createWorkbook();
      wb.addSheet('Both').view({ split: { x: 100, y: 60, topLeftCell: 'F20' } }).set('A1', 1);
      wb.addSheet('Rows').view({ split: { y: 60, topLeftCell: 'A30' } }).set('A1', 1);
      wb.addSheet('Cols').view({ split: { x: 100, topLeftCell: 'H1' } }).set('A1', 1);
      const file = join(dir, 'split.xlsx');
      await writeFile(file, writeWorkbook(wb));
      const profile = pathToFileURL(join(dir, 'profile')).href;
      const result = spawnSync(SOFFICE, [`-env:UserInstallation=${profile}`, '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', join(dir, 'out'), file], {
        encoding: 'utf8',
        timeout: 180_000,
      });
      expect(result.status, result.stderr).toBe(0);
      const resaved = readWorkbook(await readFile(join(dir, 'out', 'split.xlsx')));
      const both = resaved.getSheet('Both').view.split;
      expect(both?.topLeftCell).toBe('F20');
      expect(both?.x).toBeGreaterThan(0);
      expect(both?.y).toBeGreaterThan(0);
      // LibreOffice 24.2.7.2 writes back the same cells (observed 2026-11-02; W2-69).
      expect(resaved.getSheet('Rows').view.split?.topLeftCell).toBe('A30');
      expect(resaved.getSheet('Cols').view.split?.topLeftCell).toBe('H1');
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
