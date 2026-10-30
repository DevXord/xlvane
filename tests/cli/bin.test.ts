// End to end: the built `xlvane` command (dist/cli.js), run by node as a user would run it.
// It is built into its own directory under node_modules/.cache (so that fflate resolves), not
// into packages/xlvane/dist, which other tests rebuild at the same time.
import { spawnSync } from 'node:child_process';
import { mkdir, mkdtemp, readFile, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { buildPackage } from '../helpers/build-package';
import { withFixtures } from './helpers';

const ROOT = fileURLToPath(new URL('../..', import.meta.url));
let dir = '';
let bin = '';

beforeAll(async () => {
  await mkdir(join(ROOT, 'node_modules/.cache'), { recursive: true });
  dir = await mkdtemp(join(ROOT, 'node_modules/.cache/xlvane-cli-'));
  bin = join(await buildPackage('xlvane', dir), 'dist/cli.js');
}, 180_000);

afterAll(async () => {
  if (dir) await rm(dir, { recursive: true, force: true });
});

function xlvane(args: readonly string[], cwd: string, env: Record<string, string | undefined> = {}): { status: number | null; stdout: string; stderr: string } {
  return spawnSync('node', [bin, ...args], { cwd, encoding: 'utf8', timeout: 60_000, env: { PATH: process.env['PATH'], ...env } });
}

describe('dist/cli.js', () => {
  it('starts with a shebang and prints the package version', async () => {
    expect((await readFile(bin, 'utf8')).startsWith('#!/usr/bin/env node\n')).toBe(true);
    const pkg = JSON.parse(await readFile(join(ROOT, 'packages/xlvane/package.json'), 'utf8')) as { version: string; bin: Record<string, string> };
    expect(pkg.bin).toEqual({ xlvane: './dist/cli.js' });
    expect(xlvane(['--version'], dir)).toMatchObject({ status: 0, stdout: `${pkg.version}\n` });
  });

  it('runs commands, without colour when piped, with documented exit codes', async () => {
    await withFixtures(async (cwd) => {
      const welcome = xlvane([], cwd);
      expect(welcome.status).toBe(0);
      expect(welcome.stdout).toContain('Common tasks:');
      expect(welcome.stdout).not.toContain('\u001b[');
      const inspect = xlvane(['inspect', 'report.xlsx', '--json'], cwd);
      expect(inspect.status).toBe(0);
      expect((JSON.parse(inspect.stdout) as { result: { sheetCount: number } }).result.sheetCount).toBe(2);
      expect(xlvane(['diff', 'old.xlsx', 'new.xlsx'], cwd).status).toBe(1);
      expect(xlvane(['raed'], cwd).status).toBe(2);
      const missing = xlvane(['read', 'nope.xlsx'], cwd);
      expect(missing.status).toBe(3);
      expect(missing.stderr).toMatch(/^Error: Cannot read "nope\.xlsx"/);
      expect(xlvane(['read', 'report.xlsx'], cwd, { FORCE_COLOR: '1' }).stdout).toContain('\u001b[');
    });
  });
});
