// Command-line findings of the 1.2.2 penetration tests (docs/decisions.md D-J01): replacing a file
// with --force removed a temporary file it had not created (its name came from Date.now()), the
// replacement kept setuid/setgid bits, migrate read and replaced a package.json that is a symbolic
// link, and "Try:" commands put options after "--", where they are positional arguments.
import { existsSync } from 'node:fs';
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

describe('1.3.0: output paths that are not plain files (D-K03)', () => {
  it('a dangling symbolic link as output is reported as existing and replaced with --force', async () => {
    await withFixtures(async (dir) => {
      await symlink(join(dir, 'nowhere.csv'), join(dir, 'out.csv'));
      const r = await cli(['convert', 'report.xlsx', 'out.csv'], { cwd: dir });
      expect(r.stderr).toContain('already exists');
      expect(r.stderr).toContain('xlvane convert report.xlsx out.csv --force');
      expect(r.stderr).not.toContain('--force --force');
      const f = await cli(['convert', 'report.xlsx', 'out.csv', '--force'], { cwd: dir });
      expect(f.code).toBe(0);
      expect((await lstat(join(dir, 'out.csv'))).isFile()).toBe(true);
      expect(existsSync(join(dir, 'nowhere.csv'))).toBe(false);
    });
  });

  it('a FIFO as output is "not a regular file", not "a directory"', async () => {
    await withFixtures(async (dir) => {
      const { spawnSync } = await import('node:child_process');
      if (spawnSync('mkfifo', [join(dir, 'f.csv')]).status !== 0) return;
      const r = await cli(['convert', 'report.xlsx', 'f.csv', '--force'], { cwd: dir });
      expect(r.stderr).toContain('is not a regular file');
    });
  });

  it('--force replaces a file whose name is 250 characters long', async () => {
    await withFixtures(async (dir) => {
      const name = `${'n'.repeat(245)}.csv`;
      await writeFile(join(dir, name), 'old');
      const r = await cli(['convert', 'report.xlsx', name, '--force'], { cwd: dir });
      expect(r.code, r.stderr).toBe(0);
      expect(await readFile(join(dir, name), 'utf8')).not.toBe('old');
    });
  });
});
