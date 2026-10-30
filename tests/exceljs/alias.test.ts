// Phase 5 acceptance criterion: code written for exceljs runs unchanged when `exceljs` is
// @xlvane/exceljs, from ESM and from CommonJS.
//
// How the alias is emulated: users switch with `"exceljs": "npm:@xlvane/exceljs@^1"` in
// package.json. npm documents that for an alias `<alias>@npm:<name>`, "the `<alias>` is the
// name of the package as it is reified in the `node_modules` folder" (npm/cli
// docs/lib/content/using-npm/package-spec.md, section "Aliases", branch `latest` at 158315cd,
// read 2026-10-30). So after installation the package simply sits in node_modules/exceljs.
// Tests must not use the network, so this test builds the packages (into its own directory)
// and places the built @xlvane/exceljs at node_modules/exceljs and xlvane at node_modules/xlvane, which is what module resolution sees
// after `npm install` with the alias. The scripts in sample-project are then run with `node`.
import { spawnSync } from 'node:child_process';
import { readFile, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { buildPackage } from '../helpers/build-package';
import { copyDir, expectedSummary, makeProject, ROOT, runScript, withoutKnownDifferences } from './project-helper';

const SOFFICE = process.env['XLV_SOFFICE'] ?? 'soffice';
const hasSoffice = spawnSync(SOFFICE, ['--version'], { encoding: 'utf8', timeout: 60_000 }).status === 0;

let dir = '';

beforeAll(async () => {
  dir = await makeProject('xlvane-alias-');
  const modules = join(dir, 'node_modules');
  // node_modules/exceljs = @xlvane/exceljs (package.json + dist), as the npm alias installs it.
  await buildPackage('@xlvane/exceljs', modules, 'exceljs');
  await buildPackage('xlvane', modules);
  await copyDir(join(ROOT, 'node_modules/fflate'), join(modules, 'fflate'));
}, 400_000);

afterAll(async () => {
  if (dir !== '') await rm(dir, { recursive: true, force: true });
});

describe('exceljs alias: code written for exceljs runs unchanged', () => {
  it('installs @xlvane/exceljs under the name exceljs', async () => {
    const pkg = JSON.parse(await readFile(join(dir, 'node_modules/exceljs/package.json'), 'utf8')) as { name: string; engines?: { node?: string } };
    expect(pkg.name).toBe('@xlvane/exceljs');
    expect(pkg.engines?.node).toBe('>=22.12');
  });

  it('runs the ESM sample (default import) with the same results as exceljs 4.4.0', async () => {
    const summary = runScript(dir, 'esm.mjs');
    expect(withoutKnownDifferences(summary)).toEqual(withoutKnownDifferences(await expectedSummary()));
  });

  it('runs the CommonJS sample (require) with the same results', async () => {
    const summary = runScript(dir, 'cjs.cjs');
    expect(summary['requireType']).toBe('object');
    delete summary['requireType'];
    expect(withoutKnownDifferences(summary)).toEqual(withoutKnownDifferences(await expectedSummary()));
  });

  it('also supports named imports, which exceljs 4.4.0 does not', () => {
    const summary = runScript(dir, 'esm-named.mjs');
    expect(summary['namedImportsAreSame']).toBe(true);
  });

  it('writes files that xlvane reads with the expected values, styles and layout', async () => {
    runScript(dir, 'esm.mjs');
    const wb = readWorkbook(await readFile(join(dir, 'out/report.xlsx')));
    const s = wb.getSheet('Sales');
    expect(s.cell('B3')?.value).toBe('Coffee');
    expect(s.cell('D2')?.value).toEqual(new Date(Date.UTC(2024, 0, 31)));
    expect(s.cell('C5')?.formula?.text).toBe('SUM(C2:C4)');
    expect(s.cell('C5')?.value).toBe(39.75);
    expect(s.cell('C2')?.numFmt).toBe('#,##0.00');
    expect(s.cell('A1')?.style.font?.bold).toBe(true);
    expect(s.cell('A1')?.style.fill).toEqual({ pattern: 'solid', fgColor: 'FFD9E1F2' });
    expect(s.cell('A1')?.style.border?.bottom?.style).toBe('thin');
    expect(s.cell('B2')?.style.alignment).toMatchObject({ horizontal: 'center', wrapText: true });
    expect(s.merges).toEqual(['A7:D7']);
    expect(s.freeze).toEqual({ rows: 1, cols: 0 });
    expect(s.autoFilter).toBe('A1:D4');
    expect(s.columns.find((c) => c.min <= 2 && c.max >= 2)?.width).toBe(24);
    const log = readWorkbook(await readFile(join(dir, 'out/streamed.xlsx'))).getSheet('Log');
    expect(log.cellCount).toBe(2002);
    expect(log.cell('B1001')?.value).toBe(1_000_000);
  });

  it.skipIf(!hasSoffice)('writes files that LibreOffice opens', async () => {
    runScript(dir, 'esm.mjs');
    const profile = pathToFileURL(join(dir, 'lo-profile')).href;
    const filter = 'csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,true,false,false,-1';
    const result = spawnSync(
      SOFFICE,
      [`-env:UserInstallation=${profile}`, '--headless', '--convert-to', filter, '--outdir', join(dir, 'csv'), join(dir, 'out/report.xlsx'), join(dir, 'out/streamed.xlsx')],
      { encoding: 'utf8', timeout: 180_000 },
    );
    expect(result.status, result.stderr).toBe(0);
    const sales = (await readFile(join(dir, 'csv/report-Sales.csv'), 'utf8')).split(/\r?\n/);
    expect(sales[0]).toBe('Id,Product,Amount,Date');
    // The date's display format is not compared: exceljs 4.4.0 writes dates with built-in format 14
    // (LibreOffice shows 1/31/2024), xlvane with yyyy-mm-dd (docs/exceljs-compat.md).
    expect(sales[1]?.split(',').slice(0, 3)).toEqual(['1', 'Tea', '12.50']);
    expect(sales[1]?.split(',')[3]).not.toBe('');
    expect(sales[4]).toBe('Total,,39.75,');
    const log = (await readFile(join(dir, 'csv/streamed-Log.csv'), 'utf8')).split(/\r?\n/);
    expect(log[1000]).toBe('1000,1000000');
  }, 200_000);
});
