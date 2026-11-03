// Command-line findings of the full review of 1.3.0 (docs/decisions.md D-K04). Each test failed
// before its fix.
import { spawnSync } from 'node:child_process';
import { mkdir, symlink, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/index';
import { cli, withFixtures } from './helpers';

/** A workbook whose names, title, chart title, formula and link target hold line feeds followed by "FAKE". */
function crafted(formula: string): Uint8Array {
  const wb = createWorkbook({ properties: { title: 'Sales report' } });
  const s = wb.addSheet('Costs');
  s.set('A1', 'k').set('B1', 'v').set('A2', 'a').set('B2', 1).set('C2', { formula }).set('A3', 'link').link('A3', { url: 'https://example.com/x' });
  s.addChart({ type: 'column', anchor: 'E2:J12', title: 'Sales', series: [{ values: 'Costs!B2:B2' }] });
  wb.defineName('Total', 'Costs!$B$2');
  const z = unzipSync(writeWorkbook(wb));
  const patch = (name: string, from: string | RegExp, to: string): void => {
    const text = strFromU8(z[name]!);
    const next = text.replace(from, to);
    if (next === text) throw new Error(`no ${String(from)} in ${name}`);
    z[name] = strToU8(next);
  };
  patch('xl/workbook.xml', 'name="Costs"', 'name="Costs&#10;FAKE sheet"');
  patch('xl/workbook.xml', /Costs!\$B\$2/, "'Costs&#10;FAKE sheet'!$B$2");
  patch('xl/workbook.xml', 'name="Total"', 'name="Total&#10;FAKE name"');
  patch('docProps/core.xml', 'Sales report', 'Sales report&#10;FAKE title');
  const chart = Object.keys(z).find((n) => /^xl\/charts\/chart\d+\.xml$/.test(n))!;
  patch(chart, '<a:t>Sales</a:t>', '<a:t>Sales&#10;FAKE chart</a:t>');
  patch(chart, /Costs!/g, "'Costs&#10;FAKE sheet'!");
  const rels = Object.keys(z).find((n) => /^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/.test(n))!;
  patch(rels, 'https://example.com/x', 'https://example.com/x&#10;FAKE link');
  return zipSync(z);
}

/** Lines of the output that start with injected text. */
const mkfifo = (path: string): boolean => process.platform !== 'win32' && spawnSync('mkfifo', [path]).status === 0;

const fake = (r: { stdout: string; stderr: string }): string[] => `${r.stdout}\n${r.stderr}`.split('\n').filter((line) => /^\s*FAKE/.test(line));

describe('line feeds from a file do not start new lines of human-readable output', () => {
  it('inspect, read, validate and diff', async () => {
    await withFixtures(async (dir) => {
      await writeFile(join(dir, 'a.xlsx'), crafted('B2*2'));
      await writeFile(join(dir, 'b.xlsx'), crafted('B2*3\nFAKE formula'));
      for (const argv of [
        ['inspect', 'a.xlsx'],
        ['inspect', 'b.xlsx', '--verbose'],
        ['read', 'a.xlsx'],
        ['validate', 'a.xlsx'],
        ['diff', 'a.xlsx', 'b.xlsx'],
      ]) {
        const r = await cli(argv, { cwd: dir });
        expect(r.stdout + r.stderr, argv.join(' ')).toContain('FAKE');
        expect(fake(r), argv.join(' ')).toEqual([]);
      }
    });
  });

  it('read of an empty sheet, validate, edit --append and migrate (review of D-K04)', async () => {
    await withFixtures(async (dir) => {
      const wb = createWorkbook();
      wb.addSheet('Costs').set('A1', 'k').set('A2', 'a');
      wb.addSheet('Empty');
      const z = unzipSync(writeWorkbook(wb));
      const text = strFromU8(z['xl/workbook.xml']!).replace('name="Costs"', 'name="Costs&#10;FAKE sheet"').replace('name="Empty"', 'name="Empty&#10;FAKE empty"');
      z['xl/workbook.xml'] = strToU8(text);
      await writeFile(join(dir, 'e.xlsx'), zipSync(z));
      await writeFile(join(dir, 'rows.json'), '[["b"]]');
      await mkdir(join(dir, 'app', 'src'), { recursive: true });
      await writeFile(join(dir, 'app', 'package.json'), '{"dependencies":{"exceljs":"^4.0.0\\nFAKE version"},"engines":{"node":">=22\\nFAKE engines"}}');
      await writeFile(join(dir, 'app', 'src', 'a\nFAKE file.js'), "require('exceljs/lib/x');\n");
      for (const argv of [
        ['read', 'e.xlsx', '--sheet', 'Empty\nFAKE empty'],
        ['validate', 'e.xlsx', '--columns', 'k:string'],
        ['edit', 'e.xlsx', '--append', 'rows.json', '-o', 'out.xlsx'],
        ['migrate', 'app'],
      ]) {
        const r = await cli(argv, { cwd: dir });
        expect(r.stdout + r.stderr, argv.join(' ')).toContain('FAKE');
        expect(fake(r), argv.join(' ')).toEqual([]);
      }
    });
  });

  it('--json output keeps the line feeds as data', async () => {
    await withFixtures(async (dir) => {
      await writeFile(join(dir, 'a.xlsx'), crafted('B2*2'));
      const r = await cli(['inspect', 'a.xlsx', '--json'], { cwd: dir });
      expect(r.stdout).toContain('Costs\\nFAKE sheet');
    });
  });
});

describe('messages', () => {
  it('a dangling link is not suggested as the file to read instead', async () => {
    await withFixtures(async (dir) => {
      await symlink(join(dir, 'nowhere.xlsx'), join(dir, 'dangling.xlsx'));
      const r = await cli(['read', 'dangling.xlsx'], { cwd: dir });
      expect(r.code).not.toBe(0);
      expect(r.stderr).not.toContain('Did you mean "dangling.xlsx"');
      expect(r.stderr).not.toContain('Try:   xlvane read dangling.xlsx');
    });
  });

  it('reading a FIFO does not say "not a directory"', async () => {
    await withFixtures(async (dir) => {
      if (!mkfifo(join(dir, 'f.xlsx'))) return;
      const r = await cli(['read', 'f.xlsx'], { cwd: dir });
      expect(r.code).not.toBe(0);
      expect(r.stderr).not.toContain('not a directory');
      expect(r.stderr).toContain('is not a file');
    });
  });

  it('validate refuses a sheet that does not exist, also without --columns', async () => {
    await withFixtures(async (dir) => {
      const r = await cli(['validate', 'report.xlsx', '--sheet', 'Nope', '--json'], { cwd: dir });
      expect(r.stdout).toContain('XLV_E_SHEET_NOT_FOUND');
      expect(r.code).not.toBe(0);
    });
  });
});
