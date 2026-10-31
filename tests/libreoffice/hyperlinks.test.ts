// LibreOffice re-saves a file with hyperlinks written by xlvane and keeps them (D-A08).
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

describe.skipIf(!available && !required)('LibreOffice and hyperlinks', () => {
  it('keeps an external and an internal link when it re-saves the file', { timeout: 200_000 }, async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-lo-links-'));
    try {
      const wb = createWorkbook();
      wb.addSheet('Links')
        .set('A1', 'Example')
        .link('A1', { url: 'https://example.com/page?a=1&b=2', tooltip: 'Open example.com' })
        .set('A2', 'Target')
        .link('A2', { location: "'Second sheet'!B2" });
      wb.addSheet('Second sheet').set('B2', 'here');
      const file = join(dir, 'links.xlsx');
      await writeFile(file, writeWorkbook(wb));
      const profile = pathToFileURL(join(dir, 'profile')).href;
      const result = spawnSync(SOFFICE, [`-env:UserInstallation=${profile}`, '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', join(dir, 'out'), file], {
        encoding: 'utf8',
        timeout: 180_000,
      });
      expect(result.status, result.stderr).toBe(0);
      const sheet = readWorkbook(await readFile(join(dir, 'out', 'links.xlsx'))).getSheet('Links');
      expect(sheet.cell('A1')?.value).toBe('Example');
      const external = sheet.hyperlinks.find((l) => l.ref === 'A1');
      const internal = sheet.hyperlinks.find((l) => l.ref === 'A2');
      expect(external?.target).toBe('https://example.com/page?a=1&b=2');
      // LibreOffice 24.2 also writes an in-document target in `location` only, without r:id.
      expect(internal?.target).toBeNull();
      expect(internal?.location).toBe("'Second sheet'!B2");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
