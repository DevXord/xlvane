// LibreOffice re-saves a file with a table written by xlvane and keeps the table, its columns and
// its totals row (D-D02, TD-10). Skipped when `soffice` is not installed, unless
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

describe.skipIf(!available && !required)('LibreOffice and tables', () => {
  it('keeps the table, its columns, its totals row and the computed totals', { timeout: 200_000 }, async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-lo-tables-'));
    try {
      const wb = createWorkbook();
      wb.addSheet('Data').addTable({
        name: 'Sales',
        ref: 'B2',
        totalsRow: true,
        style: { theme: 'TableStyleMedium9', showRowStripes: true },
        columns: [{ name: 'Region', totalsRowLabel: 'Total' }, { name: 'Revenue', totalsRowFunction: 'sum' }, { name: 'Units', totalsRowFunction: 'average' }],
        rows: [
          ['North', 1200, 3],
          ['South', 980, 5],
        ],
      });
      const file = join(dir, 'tables.xlsx');
      await writeFile(file, writeWorkbook(wb));
      const profile = pathToFileURL(join(dir, 'profile')).href;
      const result = spawnSync(SOFFICE, [`-env:UserInstallation=${profile}`, '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', join(dir, 'out'), file], {
        encoding: 'utf8',
        timeout: 180_000,
      });
      expect(result.status, result.stderr).toBe(0);
      const sheet = readWorkbook(await readFile(join(dir, 'out', 'tables.xlsx'))).getSheet('Data');
      const t = sheet.tables[0];
      expect(t?.ref).toBe('B2:D5');
      expect(t?.totalsRowCount).toBe(1);
      expect(t?.columns.map((c) => [c.name, c.totalsRowFunction])).toEqual([
        ['Region', 'none'],
        ['Revenue', 'sum'],
        ['Units', 'average'],
      ]);
      // Observed with LibreOffice 24.2.7.2 (TD-10): the table style and the totals label are not
      // written back, and the autofilter then includes the totals row.
      expect(t?.style).toBeNull();
      expect(t?.columns[0]?.totalsRowLabel).toBeNull();
      expect(t?.autoFilter).toBe('B2:D5');
      // LibreOffice calculates the totals from the SUBTOTAL formulas xlvane wrote.
      expect(sheet.cell('C5')?.value).toBe(2180);
      expect(sheet.cell('D5')?.value).toBe(4);
      expect(sheet.cell('B5')?.value).toBe('Total');
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
