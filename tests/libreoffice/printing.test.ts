// LibreOffice re-saves a file with page setup, headers and footers, page breaks, print area and
// titles and sheet protection written by xlvane, and keeps them (D-C03, D-C04). Skipped when
// `soffice` is not installed, unless XLV_REQUIRE_LIBREOFFICE=1 (set in CI).
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

describe.skipIf(!available && !required)('LibreOffice and printing settings', () => {
  it('keeps margins, orientation, paper size, headers, breaks, print area, titles and protection', { timeout: 300_000 }, async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-lo-print-'));
    try {
      const wb = createWorkbook();
      wb.addSheet("Bob's Report")
        .set('A1', 'Title')
        .set('A30', 'End')
        .pageSetup({ orientation: 'landscape', paperSize: 9, margins: { left: 0.5, right: 0.5, top: 1, bottom: 1, header: 0.4, footer: 0.4 } })
        .headerFooter({ oddHeader: '&CQuarterly', oddFooter: '&RPage &P of &N' })
        .pageBreaks({ rows: [20] })
        .printArea('A1:F30')
        .printTitles({ rows: '1:2' })
        .protect({ sort: false });
      await writeFile(join(dir, 'print.xlsx'), writeWorkbook(wb));
      const profile = pathToFileURL(join(dir, 'profile')).href;
      const result = spawnSync(SOFFICE, [`-env:UserInstallation=${profile}`, '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', join(dir, 'out'), join(dir, 'print.xlsx')], {
        encoding: 'utf8',
        timeout: 180_000,
      });
      expect(result.status, result.stderr).toBe(0);

      const s = readWorkbook(await readFile(join(dir, 'out', 'print.xlsx'))).getSheet("Bob's Report");
      expect(s.pageSetup.orientation).toBe('landscape');
      expect(s.pageSetup.paperSize).toBe(9);
      expect(s.pageSetup.margins?.left).toBeCloseTo(0.5, 2);
      expect(s.pageSetup.margins?.top).toBeCloseTo(1, 2);
      expect(s.headerFooter?.oddHeader).toContain('Quarterly');
      expect(s.headerFooter?.oddFooter).toContain('&P');
      expect(s.pageBreaks.rows).toContain(20);
      expect(s.printArea).toEqual(['A1:F30']);
      expect(s.printTitles?.rows).toBe('1:2');
      expect(s.protection?.sheet).toBe(true);
      expect(s.protection?.sort).toBe(false);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
