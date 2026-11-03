// Command-line findings of the 1.2.2 penetration tests (docs/decisions.md D-J01): replacing a file
// with --force removed a temporary file it had not created (its name came from Date.now()), the
// replacement kept setuid/setgid bits, migrate read and replaced a package.json that is a symbolic
// link, and "Try:" commands put options after "--", where they are positional arguments.
import { chmod, copyFile, lstat, mkdir, readdir, readFile, stat, symlink, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import { cli, withFixtures } from './helpers';

describe('--force', () => {
  it('never removes a temporary-looking file it did not create', async () => {
    await withFixtures(async (dir) => {
      await writeFile(join(dir, 'out.xlsx'), 'old');
      const now = 1_700_000_000_000;
      const spy = vi.spyOn(Date, 'now').mockReturnValue(now);
      try {
        await writeFile(join(dir, `.out.xlsx.xlvane-${String(now)}.tmp`), 'not mine');
        const r = await cli(['convert', 'sales.csv', 'out.xlsx', '--force'], { cwd: dir });
        expect(r.code).toBe(0);
      } finally {
        spy.mockRestore();
      }
      expect(await readFile(join(dir, `.out.xlsx.xlvane-${String(now)}.tmp`), 'utf8')).toBe('not mine');
      expect((await readdir(dir)).filter((n) => n.endsWith('.tmp'))).toEqual([`.out.xlsx.xlvane-${String(now)}.tmp`]);
    });
  });

  it('keeps the permission bits of the replaced file but not setuid or setgid', async () => {
    await withFixtures(async (dir) => {
      await writeFile(join(dir, 'out.xlsx'), 'old');
      await chmod(join(dir, 'out.xlsx'), 0o6750);
      if (((await stat(join(dir, 'out.xlsx'))).mode & 0o6000) === 0) return; // the file system keeps no such bits
      expect((await cli(['convert', 'sales.csv', 'out.xlsx', '--force'], { cwd: dir })).code).toBe(0);
      expect((await stat(join(dir, 'out.xlsx'))).mode & 0o7777).toBe(0o750);
    });
  });
});

describe('migrate and a package.json that is a symbolic link', () => {
  it('refuses to read it and leaves the link and its target alone with --force', async () => {
    await withFixtures(async (dir) => {
      await mkdir(join(dir, 'outside'));
      await mkdir(join(dir, 'app'));
      const outside = '{\n  "name": "outside",\n  "dependencies": { "exceljs": "^4.4.0" }\n}\n';
      await writeFile(join(dir, 'outside', 'package.json'), outside);
      await symlink(join(dir, 'outside', 'package.json'), join(dir, 'app', 'package.json'));
      const r = await cli(['migrate', 'app', '--json'], { cwd: dir });
      expect(r.code).not.toBe(0);
      expect(r.stdout + r.stderr).toContain('XLV_E_FILE_READ');
      expect(r.stdout).not.toContain('outside');
      const f = await cli(['migrate', 'app', '--force'], { cwd: dir });
      expect(f.code).not.toBe(0);
      expect((await lstat(join(dir, 'app', 'package.json'))).isSymbolicLink()).toBe(true);
      expect(await readFile(join(dir, 'outside', 'package.json'), 'utf8')).toBe(outside);
    });
  });
});

describe('"Try:" commands with "--"', () => {
  it('put --force before "--"', async () => {
    await withFixtures(async (dir) => {
      await copyFile(join(dir, 'report.xlsx'), join(dir, '-r.xlsx'));
      await writeFile(join(dir, '-out.csv'), 'old');
      const r = await cli(['convert', '--', '-r.xlsx', '-out.csv'], { cwd: dir });
      expect(r.code).not.toBe(0);
      expect(r.stderr).toContain('xlvane convert --force -- -r.xlsx -out.csv');
    });
  });

  it('migrate names a directory starting with "-" after "--"', async () => {
    await withFixtures(async (dir) => {
      await mkdir(join(dir, '-app'));
      await writeFile(join(dir, '-app', 'package.json'), '{ "dependencies": { "exceljs": "^4.4.0" } }\n');
      const r = await cli(['migrate', '--', '-app'], { cwd: dir });
      expect(r.code).toBe(0);
      expect(r.stdout).toContain('xlvane migrate --force -- -app');
    });
  });

  it('the delimiter note puts --delimiter before "--"', async () => {
    await withFixtures(async (dir) => {
      await writeFile(join(dir, '-s.csv'), 'a;b\r\n1;2\r\n');
      const r = await cli(['convert', '--', '-s.csv', '-s.xlsx'], { cwd: dir });
      expect(r.code).toBe(0);
      expect(r.stderr).toMatch(/xlvane convert --delimiter ';' -- -s\.csv -s\.xlsx|xlvane convert --delimiter ";" -- -s\.csv -s\.xlsx/);
    });
  });
});

describe('convert to JSON of a header repeated in many rows', () => {
  it('ends as a limit error, not an internal one', async () => {
    const { buildXlsx } = await import('../helpers/xlsx-builder');
    let seed = 7;
    let header = '';
    for (let i = 0; i < 4_000_000; i++) header += ((seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648) < 0.01 ? 'x' : 'h';
    let data = `<row r="1"><c r="A1" t="inlineStr"><is><t>${header}</t></is></c></row>`;
    for (let r = 2; r <= 141; r++) data += `<row r="${String(r)}"><c r="A${String(r)}"><v>${String(r)}</v></c></row>`;
    await withFixtures(async (dir) => {
      await writeFile(join(dir, 'wide.xlsx'), buildXlsx({ sheets: [{ name: 'S', data }] }));
      const r = await cli(['convert', 'wide.xlsx', 'wide.json', '--json'], { cwd: dir });
      expect(r.stdout).toContain('XLV_E_LIMIT_EXCEEDED');
      expect(r.stdout).not.toContain('XLV_E_INTERNAL');
    });
  }, 60_000);
});
