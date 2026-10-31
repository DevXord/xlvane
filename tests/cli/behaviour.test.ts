// CLI behaviour: colour rules, overwrite protection, --dry-run, CSV injection, value typing of
// --set and migrate (docs/PLAN.md section 5; docs/decisions.md D-601 … D-607).
import { readFile, stat, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { parseSetValue } from '../../packages/xlvane/src/cli/commands/edit';
import { scanSource } from '../../packages/xlvane/src/cli/commands/migrate';
import { Output, useColor } from '../../packages/xlvane/src/cli/output';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { readTable } from '../../packages/xlvane/src/table';
import { cli, withFixtures } from './helpers';

const ESC = '\u001b[';
const strip = (s: string): string => s.replace(/\u001b\[[0-9;]*m/g, '');

describe('colour', () => {
  it('follows NO_COLOR, NODE_DISABLE_COLORS, FORCE_COLOR and the terminal as Node.js does', () => {
    expect(useColor({}, true)).toBe(true);
    expect(useColor({}, false)).toBe(false);
    expect(useColor({ NO_COLOR: '1' }, true)).toBe(false);
    expect(useColor({ NO_COLOR: '' }, true)).toBe(false);
    expect(useColor({ NODE_DISABLE_COLORS: '1' }, true)).toBe(false);
    expect(useColor({ FORCE_COLOR: '1' }, false)).toBe(true);
    expect(useColor({ FORCE_COLOR: '' }, false)).toBe(true);
    expect(useColor({ FORCE_COLOR: '3' }, false)).toBe(true);
    expect(useColor({ FORCE_COLOR: '0' }, true)).toBe(false);
    expect(useColor({ FORCE_COLOR: 'no' }, true)).toBe(false);
    // FORCE_COLOR with a supported value wins over NO_COLOR (Node.js doc/api/cli.md).
    expect(useColor({ NO_COLOR: '1', FORCE_COLOR: '1' }, false)).toBe(true);
  });

  it('is used on a terminal and never carries information alone', async () => {
    await withFixtures(async (dir) => {
      const coloured = await cli(['diff', 'old.xlsx', 'new.xlsx'], { cwd: dir, env: {}, tty: true });
      const plain = await cli(['diff', 'old.xlsx', 'new.xlsx'], { cwd: dir, env: {}, tty: false });
      expect(coloured.stdout).toContain(ESC);
      expect(plain.stdout).not.toContain(ESC);
      expect(strip(coloured.stdout)).toBe(plain.stdout);
      expect(plain.stdout).toMatch(/changed\s+B2\s+1200 → 1250/);
      const noColor = await cli(['inspect', 'report.xlsx'], { cwd: dir, env: { NO_COLOR: '1' }, tty: true });
      expect(noColor.stdout).not.toContain(ESC);
      const error = await cli(['inspect', 'nope.xlsx'], { cwd: dir, env: {}, tty: true });
      expect(error.stderr).toContain(ESC);
      expect(strip(error.stderr)).toMatch(/^Error: /);
    });
  });

  it('tables fit the terminal width', () => {
    const out = new Output(false, 30);
    const text = out.table(['A', 'Long header'], [['1', 'x'.repeat(100)]]);
    for (const line of text.split('\n')) expect(line.length).toBeLessThanOrEqual(30);
    expect(text).toContain('…');
  });
});

describe('writing', () => {
  it('never replaces a file without --force, and --force replaces it atomically', async () => {
    await withFixtures(async (dir) => {
      const before = await readFile(join(dir, 'sales.csv'), 'utf8');
      const refused = await cli(['convert', 'report.xlsx', 'sales.csv'], { cwd: dir });
      expect(refused.code).toBe(3);
      expect(await readFile(join(dir, 'sales.csv'), 'utf8')).toBe(before);
      const forced = await cli(['convert', 'report.xlsx', 'sales.csv', '--force'], { cwd: dir });
      expect(forced.code).toBe(0);
      expect(forced.stdout).toBe('Replaced sales.csv (' + String((await stat(join(dir, 'sales.csv'))).size) + ' bytes).\n');
      expect(await readFile(join(dir, 'sales.csv'), 'utf8')).toMatch(/^Region,Revenue,Date\r\nNorth,1200,2024-01-31\r\n/);
    });
  });

  it('--dry-run writes nothing and says what it would do', async () => {
    await withFixtures(async (dir) => {
      const r = await cli(['create', 'data.json', 'out.xlsx', '--dry-run'], { cwd: dir });
      expect(r.code).toBe(0);
      expect(r.stdout).toMatch(/^Dry run: would write out\.xlsx \(\d+\.\d kB\)\. Nothing was written\.\n$/);
      await expect(stat(join(dir, 'out.xlsx'))).rejects.toThrow();
      const inPlace = await cli(['edit', 'report.xlsx', '--set', 'Sales!B2=1', '--dry-run'], { cwd: dir });
      expect(inPlace.stdout).toContain('Dry run: would replace report.xlsx');
    });
  });

  it('--quiet prints nothing on success', async () => {
    await withFixtures(async (dir) => {
      expect(await cli(['convert', 'report.xlsx', 'out.json', '-q'], { cwd: dir })).toEqual({ code: 0, stdout: '', stderr: '' });
    });
  });

  it('CSV output escapes formula-like text (OWASP CSV injection)', async () => {
    await withFixtures(async (dir) => {
      await cli(['convert', 'report.xlsx', 'costs.csv', '--sheet', 'Costs'], { cwd: dir });
      expect(await readFile(join(dir, 'costs.csv'), 'utf8')).toContain(`"'=cmd|calc"`);
    });
  });

  it('create builds a report and convert a plain table', async () => {
    await withFixtures(async (dir) => {
      expect((await cli(['create', 'data.json', 'r.xlsx', '--sheet', 'People', '--columns', 'name,score'], { cwd: dir })).code).toBe(0);
      const report = readWorkbook(await readFile(join(dir, 'r.xlsx'))).getSheet('People');
      expect(report.freeze).toEqual({ rows: 1, cols: 0 });
      expect(report.cell('A1')?.style.font?.bold).toBe(true);
      expect(readTable(await readFile(join(dir, 'r.xlsx'))).rows).toEqual([
        { name: 'Ann', score: 7.5 },
        { name: 'Ben', score: 9 },
      ]);
      expect((await cli(['convert', 'sales.csv', 's.xlsx'], { cwd: dir })).code).toBe(0);
      expect(readTable(await readFile(join(dir, 's.xlsx'))).rows[0]).toEqual({ Region: 'North', Revenue: 1200, Date: new Date(Date.UTC(2024, 0, 31)) });
    });
  });
});

describe('edit', () => {
  it('types --set values like CSV fields', () => {
    expect(parseSetValue('12.5')).toBe(12.5);
    expect(parseSetValue('007')).toBe('007');
    expect(parseSetValue('true')).toBe(true);
    expect(parseSetValue('2024-02-29')).toEqual(new Date(Date.UTC(2024, 1, 29)));
    expect(parseSetValue('=SUM(A1:A2)')).toEqual({ formula: 'SUM(A1:A2)' });
    expect(parseSetValue("'=not a formula")).toBe('=not a formula');
    expect(parseSetValue("'42")).toBe('42');
    expect(parseSetValue('')).toBeNull();
    expect(parseSetValue('North-East')).toBe('North-East');
  });

  it('sets cells, appends rows and keeps the input unless --force', async () => {
    await withFixtures(async (dir) => {
      const original = await readFile(join(dir, 'report.xlsx'));
      const r = await cli(['edit', 'report.xlsx', '--set', "'Sales'!B2=1300", '--set', 'Sales!D1=Note', '--append', 'rows.json', '--sheet', 'Sales', '-o', 'out.xlsx'], { cwd: dir });
      expect(r.code).toBe(0);
      expect(r.stdout).toContain('Set 2 cells, appended 1 row to "Sales".');
      expect(await readFile(join(dir, 'report.xlsx'))).toEqual(original);
      const s = readWorkbook(await readFile(join(dir, 'out.xlsx'))).getSheet('Sales');
      expect(s.cell('B2')?.value).toBe(1300);
      expect(s.cell('D1')?.value).toBe('Note');
      expect(s.cell('A5')?.value).toBe('West');
      expect((await cli(['edit', 'report.xlsx', '--sheet', 'Sales', '--set', 'B3=1', '--force'], { cwd: dir })).code).toBe(0);
      expect(readWorkbook(await readFile(join(dir, 'report.xlsx'))).getSheet('Sales').cell('B3')?.value).toBe(1);
    });
  });
});

describe('edit: rows', () => {
  it('inserts and deletes rows and moves formulas and names with them', async () => {
    await withFixtures(async (dir) => {
      const r = await cli(['edit', 'report.xlsx', '--sheet', 'Sales', '--insert-rows', '3', '--count', '2', '--insert-style', 'above', '-o', 'more.xlsx'], { cwd: dir });
      expect(r.code).toBe(0);
      expect(r.stdout).toContain('Inserted 2 rows before row 3 of "Sales".');
      const more = readWorkbook(await readFile(join(dir, 'more.xlsx')));
      expect(more.getSheet('Sales').cell('A5')?.value).toBe('South');
      expect(more.getSheet('Sales').cell('B6')?.formula?.text).toBe('SUM(B2:B5)');
      expect(more.definedNames.find((n) => n.name === 'Total')?.refersTo).toBe('Sales!$B$6');
      const back = await cli(['edit', 'more.xlsx', '--sheet', 'Sales', '--delete-rows', '3:4', '--json', '-o', 'back.xlsx'], { cwd: dir });
      const doc = JSON.parse(back.stdout) as { result: { rowsDeleted: number; rowSheet: string } };
      expect([doc.result.rowsDeleted, doc.result.rowSheet]).toEqual([2, 'Sales']);
      expect(readWorkbook(await readFile(join(dir, 'back.xlsx'))).getSheet('Sales').cell('B4')?.formula?.text).toBe('SUM(B2:B3)');
    });
  });

  it('applies row changes before --set and needs --sheet', async () => {
    await withFixtures(async (dir) => {
      const r = await cli(['edit', 'report.xlsx', '--sheet', 'Sales', '--insert-rows', '2', '--set', 'A2=New', '-o', 'out.xlsx'], { cwd: dir });
      expect(r.code).toBe(0);
      const s = readWorkbook(await readFile(join(dir, 'out.xlsx'))).getSheet('Sales');
      expect([s.cell('A2')?.value, s.cell('A3')?.value]).toEqual(['New', 'North']);
      expect((await cli(['edit', 'report.xlsx', '--delete-rows', '2', '-o', 'x.xlsx'], { cwd: dir })).code).toBe(2);
      expect((await cli(['edit', 'report.xlsx', '--sheet', 'Sales', '--delete-rows', '3:2', '-o', 'x.xlsx'], { cwd: dir })).code).toBe(2);
      expect((await cli(['edit', 'report.xlsx', '--sheet', 'Sales', '--count', '2', '--set', 'A1=x', '-o', 'x.xlsx'], { cwd: dir })).code).toBe(2);
    });
  });
});

describe('validate', () => {
  it('names the cell of each problem and counts warnings only with --strict', async () => {
    await withFixtures(async (dir) => {
      const r = await cli(['validate', 'report.xlsx', '--sheet', 'Sales', '--columns', 'Region:string,Revenue:number,Date:date,Owner:string'], { cwd: dir });
      expect(r.code).toBe(1);
      expect(r.stdout).toContain('  problem  C4  is empty; "Date" needs a value.');
      expect(r.stdout).toContain('Column "Owner" is missing.');
      const ok = await cli(['validate', 'report.xlsx', '--sheet', 'Sales', '--columns', 'Region:string,Revenue:number,Date:date?'], { cwd: dir });
      expect(ok.code).toBe(0);
      expect(ok.stdout).toContain('OK: report.xlsx, 3 rows of sheet "Sales" checked against 3 column rules.');
    });
  });
});

describe('migrate', () => {
  it('reports without writing, and --force changes only the dependency value', async () => {
    await withFixtures(async (dir) => {
      const before = await readFile(join(dir, 'my-app/package.json'), 'utf8');
      const report = await cli(['migrate', 'my-app'], { cwd: dir });
      expect(report.code).toBe(0);
      expect(report.stdout).toContain('- "exceljs": "^4.4.0"');
      expect(report.stdout).toContain('+ "exceljs": "npm:@xlvane/exceljs@^1"');
      expect(report.stdout).toContain('src/report.js:4:3  addImage({ filename })');
      expect(await readFile(join(dir, 'my-app/package.json'), 'utf8')).toBe(before);
      const forced = await cli(['migrate', 'my-app', '--force'], { cwd: dir });
      expect(forced.code).toBe(0);
      expect(await readFile(join(dir, 'my-app/package.json'), 'utf8')).toBe(before.replace('"^4.4.0"', '"npm:@xlvane/exceljs@^1"'));
      expect((await cli(['migrate', 'my-app'], { cwd: dir })).stdout).toContain('already uses npm:@xlvane/exceljs@^1');
    });
  });

  it('scans only files that mention exceljs', () => {
    expect(scanSource('a.js', 'ws.insertRow(1, [])')).toEqual([]);
    expect(scanSource('a.js', "import ExcelJS from 'exceljs';\nwb.csv.writeFile('x')").map((f) => f.api)).toEqual([]);
    expect(scanSource('a.js', "const ExcelJS = require('exceljs/dist/es5');\nconst cell = require('exceljs/lib/doc/cell');").map((f) => f.api)).toEqual(['deep import (exceljs/lib/...)']);
  });

  it('skips node_modules', async () => {
    await withFixtures(async (dir) => {
      await writeFile(join(dir, 'my-app/src/ok.js'), "require('exceljs');\n");
      const r = await cli(['migrate', 'my-app', '--json'], { cwd: dir });
      const doc = JSON.parse(r.stdout) as { result: { filesScanned: number } };
      expect(doc.result.filesScanned).toBe(2);
    });
  });
});

describe('read and inspect', () => {
  it('read prints an aligned table with the row count', async () => {
    await withFixtures(async (dir) => {
      const r = await cli(['read', 'report.xlsx', '--sheet', 'Sales'], { cwd: dir });
      expect(r.stdout).toBe(
        [
          'Region  Revenue  Date',
          'North      1200  2024-01-31',
          'South       980  2024-02-29',
          'Total      2180',
          '',
        ].join('\n'),
      );
      // The summary goes to stderr, so that redirected output holds only the table.
      expect(r.stderr).toBe('Sheet "Sales", A1:C4, header in row 1: 3 rows.\n');
    });
  });

  it('inspect describes sheets and columns', async () => {
    await withFixtures(async (dir) => {
      const r = await cli(['inspect', 'report.xlsx'], { cwd: dir });
      await expect(r.stdout).toMatchFileSnapshot('__snapshots__/inspect-report.txt');
    });
  });
});
