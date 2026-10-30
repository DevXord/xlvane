// LibreOffice re-saves a file with [MS-XLSX] extensions written by xlvane: an extended data bar,
// an extension-only icon set and a list validation from another sheet (D-C07, SR-45). Observed
// with LibreOffice 24.2.7.2: it keeps the extension link, the negative fill and axis colours, the
// solid fill, autoMin/autoMax and the 3Stars set; it drops the bar border and writes the lengths
// as 10 and 90. Skipped when `soffice` is not installed, unless XLV_REQUIRE_LIBREOFFICE=1.
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

describe.skipIf(!available && !required)('LibreOffice and the [MS-XLSX] extensions', () => {
  it('keeps an extended data bar, an extension icon set and a list from another sheet', { timeout: 200_000 }, async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-lo-x14-'));
    try {
      const wb = createWorkbook();
      const lists = wb.addSheet('Lists');
      ['Red', 'Green', 'Blue'].forEach((v, i) => lists.set(`A${String(i + 1)}`, v));
      const form = wb.addSheet('Form');
      for (let r = 1; r <= 5; r++) form.set(`A${String(r)}`, r * 10 - 25);
      form.validate('B1:B5', { type: 'list', formula1: 'Lists!$A$1:$A$3' });
      form.conditionalFormat('A1:A5', [
        {
          type: 'dataBar',
          dataBar: { cfvo: [{ type: 'autoMin' }, { type: 'autoMax' }], color: '638EC6', minLength: 0, maxLength: 100, gradient: false, borderColor: '0000FF', negativeFillColor: 'FF0000', axisColor: '000000', axisPosition: 'middle' },
        },
      ]);
      form.conditionalFormat('C1:C5', [{ type: 'iconSet', iconSet: { iconSet: '3Stars', cfvo: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }] } }]);
      const file = join(dir, 'x14.xlsx');
      await writeFile(file, writeWorkbook(wb));
      const profile = pathToFileURL(join(dir, 'profile')).href;
      const result = spawnSync(SOFFICE, [`-env:UserInstallation=${profile}`, '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', join(dir, 'out'), file], {
        encoding: 'utf8',
        timeout: 180_000,
      });
      expect(result.status, result.stderr).toBe(0);
      const back = readWorkbook(await readFile(join(dir, 'out', 'x14.xlsx'))).getSheet('Form');
      expect(back.dataValidations[0]).toMatchObject({ sqref: 'B1:B5', type: 'list', formula1: 'Lists!$A$1:$A$3' });
      const rules = back.conditionalFormats.flatMap((c) => c.rules);
      expect(rules.find((r) => r.type === 'dataBar')?.dataBar).toMatchObject({
        extended: true,
        color: 'FF638EC6',
        gradient: false,
        negativeFillColor: 'FFFF0000',
        axisPosition: 'middle',
        axisColor: 'FF000000',
        cfvo: [{ type: 'autoMin' }, { type: 'autoMax' }],
      });
      expect(rules.find((r) => r.type === 'iconSet')?.iconSet).toMatchObject({ iconSet: '3Stars', cfvo: [{ value: '0' }, { value: '33' }, { value: '67' }] });
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
