// LibreOffice reads array and shared formulas written by xlvane and computes them (D-E02,
// F2-13 … F2-16). Skipped when `soffice` is not installed, unless XLV_REQUIRE_LIBREOFFICE=1.
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

describe.skipIf(!available && !required)('LibreOffice and array or shared formulas', () => {
  it('computes an array formula and a shared formula written without cached results', { timeout: 200_000 }, async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-lo-formulas-'));
    try {
      const wb = createWorkbook();
      const s = wb.addSheet('S');
      s.set('A1', 1).set('A2', 2).set('A3', 3);
      s.set('B1', { formula: 'A1:A3*2', array: 'B1:B3' });
      s.set('C1', { formula: 'A1+10', shared: 'C1:C3' }).set('C2', { sharedFormula: 'C1' }).set('C3', { sharedFormula: 'C1' });
      const file = join(dir, 'formulas.xlsx');
      await writeFile(file, writeWorkbook(wb));
      const profile = pathToFileURL(join(dir, 'profile')).href;
      const result = spawnSync(SOFFICE, [`-env:UserInstallation=${profile}`, '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', join(dir, 'out'), file], {
        encoding: 'utf8',
        timeout: 180_000,
      });
      expect(result.status, result.stderr).toBe(0);
      const sheet = readWorkbook(await readFile(join(dir, 'out', 'formulas.xlsx'))).getSheet('S');
      expect([1, 2, 3].map((r) => sheet.cell(`B${String(r)}`)?.value)).toEqual([2, 4, 6]);
      expect(sheet.cell('B1')?.formula).toMatchObject({ kind: 'array', ref: 'B1:B3' });
      expect([1, 2, 3].map((r) => sheet.cell(`C${String(r)}`)?.value)).toEqual([11, 12, 13]);
      // LibreOffice 24.2.7.2 writes the shared group back as one ordinary formula per cell
      // (F2-16), and the other cells of the array range as values only (F2-14).
      expect(sheet.cell('C3')?.formula).toMatchObject({ kind: 'normal', text: 'A3+10' });
      expect(sheet.cell('B2')?.formula).toBeNull();
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
