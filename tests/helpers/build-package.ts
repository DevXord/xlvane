// Builds a workspace package into a separate directory, laid out as npm installs it
// (<modules>/<name>/package.json + dist). Tests that need built packages use this instead of
// packages/*/dist, so that no test rewrites the dist of another test running at the same time.
import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { copyFile, mkdir } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

export const REPO: string = fileURLToPath(new URL('../..', import.meta.url));

/** The JS file of a package's `bin` entry in the root node_modules (portable: no .cmd shims). */
export function binOf(pkg: string, bin: string): string {
  const dir = join(REPO, 'node_modules', pkg);
  const manifest = JSON.parse(readFileSync(join(dir, 'package.json'), 'utf8')) as { bin?: string | Record<string, string> };
  const rel = typeof manifest.bin === 'string' ? manifest.bin : manifest.bin?.[bin];
  if (rel === undefined) throw new Error(`${pkg} has no bin "${bin}"`);
  return join(dir, rel);
}

const DIRS = { xlvane: 'packages/xlvane', '@xlvane/exceljs': 'packages/exceljs' } as const;

/**
 * Builds `name` into `<modules>/<install name>`; returns that directory. Declaration files are
 * built only with `{ dts: true }` (slower; needed by type tests of the published package).
 */
export async function buildPackage(name: keyof typeof DIRS, modules: string, installAs: string = name, options: { readonly dts?: boolean } = {}): Promise<string> {
  const target = join(modules, installAs);
  await mkdir(target, { recursive: true });
  const cwd = join(REPO, DIRS[name]);
  // tsdown's JS entry run by this Node.js binary: no shell and no `npx`, which is `npx.cmd` on
  // Windows and cannot be spawned without a shell.
  const build = spawnSync(process.execPath, [binOf('tsdown', 'tsdown'), '--out-dir', join(target, 'dist'), ...(options.dts === true ? [] : ['--no-dts'])], { cwd, encoding: 'utf8', timeout: 180_000 });
  if (build.status !== 0) throw new Error(`build of ${name} failed (${build.error?.message ?? `exit ${String(build.status)}`}):\n${build.stdout}\n${build.stderr}`);
  await copyFile(join(cwd, 'package.json'), join(target, 'package.json'));
  return target;
}
