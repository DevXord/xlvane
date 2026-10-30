// Builds a workspace package into a separate directory, laid out as npm installs it
// (<modules>/<name>/package.json + dist). Tests that need built packages use this instead of
// packages/*/dist, so that no test rewrites the dist of another test running at the same time.
import { spawnSync } from 'node:child_process';
import { copyFile, mkdir } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

export const REPO: string = fileURLToPath(new URL('../..', import.meta.url));

const DIRS = { xlvane: 'packages/xlvane', '@xlvane/exceljs': 'packages/exceljs', '@xlvane/mcp': 'packages/mcp' } as const;

/** Builds `name` into `<modules>/<install name>`; returns that directory. */
export async function buildPackage(name: keyof typeof DIRS, modules: string, installAs: string = name): Promise<string> {
  const target = join(modules, installAs);
  await mkdir(target, { recursive: true });
  const cwd = join(REPO, DIRS[name]);
  const build = spawnSync('npx', ['tsdown', '--out-dir', join(target, 'dist'), '--no-dts'], { cwd, encoding: 'utf8', timeout: 180_000 });
  if (build.status !== 0) throw new Error(`build of ${name} failed:\n${build.stdout}\n${build.stderr}`);
  await copyFile(join(cwd, 'package.json'), join(target, 'package.json'));
  return target;
}
