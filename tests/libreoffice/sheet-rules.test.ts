// LibreOffice re-saves a file with a list validation and a cellIs conditional format written by
// xlvane and keeps both (D-C01, D-C02). Skipped when `soffice` is not installed, unless
// XLV_REQUIRE_LIBREOFFICE=1 (set in CI).
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

describe.skipIf(!available && !required)('LibreOffice, data validation and conditional formatting', () => {
  it('keeps a list validation and a cellIs rule with its fill when it re-saves the file', { timeout: 200_000 }, async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-lo-rules-'));
    try {
      const wb = createWorkbook();
      const sheet = wb.addSheet('Form');
      for (let r = 1; r <= 5; r++) sheet.set(`A${String(r)}`, r * 50);
      sheet.validate('B1:B5', { type: 'list', formula1: '"Yes,No"', showErrorMessage: true, error: 'Choose Yes or No.' });
      sheet.conditionalFormat('A1:A5', [{ type: 'cellIs', operator: 'greaterThan', formulas: [100], style: { fill: { pattern: 'solid', fgColor: 'FFC7CE' } } }]);
      const file = join(dir, 'rules.xlsx');
      await writeFile(file, writeWorkbook(wb));
      const profile = pathToFileURL(join(dir, 'profile')).href;
      const result = spawnSync(SOFFICE, [`-env:UserInstallation=${profile}`, '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', join(dir, 'out'), file], {
        encoding: 'utf8',
        timeout: 180_000,
      });
      expect(result.status, result.stderr).toBe(0);
      const back = readWorkbook(await readFile(join(dir, 'out', 'rules.xlsx'))).getSheet('Form');
      expect(back.dataValidations).toHaveLength(1);
      expect(back.dataValidations[0]).toMatchObject({ sqref: 'B1:B5', type: 'list', formula1: '"Yes,No"', showErrorMessage: true, error: 'Choose Yes or No.' });
      const rule = back.conditionalFormats.flatMap((c) => c.rules)[0];
      expect(back.conditionalFormats[0]?.sqref).toBe('A1:A5');
      expect(rule).toMatchObject({ type: 'cellIs', operator: 'greaterThan', formulas: ['100'] });
      // The fill colour survives; LibreOffice may write it as foreground, background or both.
      const fill = rule?.style?.fill as { fgColor?: unknown; bgColor?: unknown } | undefined;
      expect([fill?.fgColor, fill?.bgColor]).toContain('FFFFC7CE');
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
