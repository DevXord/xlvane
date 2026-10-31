// LibreOffice re-saves the style extensions written by xlvane (D-E01, W2-60 … W2-63): a gradient
// fill, a diagonal border, a superscript rich text run, shrinkToFit and an unlocked cell. What
// LibreOffice keeps is recorded in docs/facts/sml-write.md W2-64: everything except the gradient,
// which it turns into a solid fill. Skipped when `soffice` is not
// installed, unless XLV_REQUIRE_LIBREOFFICE=1.
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

describe.skipIf(!available && !required)('LibreOffice and the style extensions', () => {
  it('keeps a diagonal border, a superscript run, shrinkToFit and an unlocked cell; turns a gradient into a solid fill', { timeout: 200_000 }, async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-lo-styles-'));
    try {
      const wb = createWorkbook();
      const s = wb.addSheet('S');
      s.set('A1', 'gradient', { fill: { type: 'gradient', gradient: 'linear', degree: 90, stops: [{ position: 0, color: 'FFFFFF' }, { position: 1, color: '4472C4' }] } });
      s.set('A2', 'diagonal', { border: { diagonal: { style: 'thin', color: 'FF0000' }, diagonalDown: true } });
      s.set('A3', { richText: [{ text: 'm' }, { text: '2', font: { vertAlign: 'superscript' } }] });
      s.set('A4', 'a long text that shrinks', { alignment: { shrinkToFit: true } });
      s.set('A5', 'input', { protection: { locked: false } });
      const file = join(dir, 'styles.xlsx');
      await writeFile(file, writeWorkbook(wb));
      const profile = pathToFileURL(join(dir, 'profile')).href;
      const result = spawnSync(SOFFICE, [`-env:UserInstallation=${profile}`, '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', join(dir, 'out'), file], {
        encoding: 'utf8',
        timeout: 180_000,
      });
      expect(result.status, result.stderr).toBe(0);
      const back = readWorkbook(await readFile(join(dir, 'out', 'styles.xlsx'))).getSheet('S');
      // LibreOffice 24.2.7.2 reads the gradient and writes it back as a solid fill of a blended
      // colour (observed: FFA1B8E1 for white to 4472C4 at 90 degrees; W2-64).
      expect(back.cell('A1')?.style.fill).toMatchObject({ pattern: 'solid' });
      expect(back.cell('A2')?.style.border).toMatchObject({ diagonalDown: true, diagonal: { style: 'thin' } });
      expect(back.cell('A3')?.richText?.find((r) => r.text === '2')?.font?.vertAlign).toBe('superscript');
      expect(back.cell('A4')?.style.alignment?.shrinkToFit).toBe(true);
      expect(back.cell('A5')?.style.protection?.locked).toBe(false);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
