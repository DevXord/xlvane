// Builds a temporary project in which `exceljs` resolves to a given package directory, and runs
// the sample scripts of tests/exceljs/sample-project in it. Used by alias.test.ts (exceljs =
// @xlvane/exceljs, as installed by the alias "exceljs": "npm:@xlvane/exceljs@^1") and by
// cross-check.test.ts (exceljs = the real exceljs 4.4.0 from bench/node_modules).
import { spawnSync } from 'node:child_process';
import { mkdir, mkdtemp, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

export const ROOT = resolve('.');
export const SAMPLE = join(ROOT, 'tests/exceljs/sample-project');

/** Copies a directory tree (files and folders only). */
export async function copyDir(from: string, to: string, skip: (name: string) => boolean = () => false): Promise<void> {
  await mkdir(to, { recursive: true });
  for (const name of await readdir(from)) {
    if (skip(name)) continue;
    const source = join(from, name);
    const target = join(to, name);
    if ((await stat(source)).isDirectory()) await copyDir(source, target, skip);
    else await writeFile(target, await readFile(source));
  }
}

/** A fresh project directory with the sample scripts and an `out` folder. */
export async function makeProject(prefix: string): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), prefix));
  await copyDir(SAMPLE, dir, (name) => name === 'node_modules' || name.startsWith('expected-'));
  await mkdir(join(dir, 'out'), { recursive: true });
  await mkdir(join(dir, 'node_modules'), { recursive: true });
  return dir;
}

/** Runs one sample script with node in `dir`; returns its parsed JSON summary. */
export function runScript(dir: string, script: string, env: Record<string, string> = {}): Record<string, unknown> {
  const envPrefix = Object.entries(env).map(([k, v]) => `${k}=${JSON.stringify(v)}`);
  const result = spawnSync(envPrefix.length > 0 ? 'env' : 'node', [...envPrefix, ...(envPrefix.length > 0 ? ['node'] : []), script, join(dir, 'out')], {
    cwd: dir,
    encoding: 'utf8',
    timeout: 120_000,
  });
  if (result.status !== 0) throw new Error(`${script} failed (${String(result.status)}):\n${result.stderr}\n${result.stdout}`);
  const lines = result.stdout.trim().split('\n');
  return JSON.parse(lines[lines.length - 1] ?? '{}') as Record<string, unknown>;
}

/** The summary printed by exceljs 4.4.0 itself, kept in the repository (see cross-check.test.ts). */
export async function expectedSummary(): Promise<Record<string, unknown>> {
  return JSON.parse(await readFile(join(SAMPLE, 'expected-exceljs-4.4.0.json'), 'utf8')) as Record<string, unknown>;
}

/**
 * Differences between @xlvane/exceljs and exceljs 4.4.0 that are documented in
 * docs/exceljs-compat.md. Each entry removes one key of the summary before comparing; the
 * list is empty when the two agree on everything the sample checks.
 */
export const KNOWN_DIFFERENCES: readonly { readonly key: string; readonly reason: string }[] = [];

export function withoutKnownDifferences(summary: Record<string, unknown>): Record<string, unknown> {
  const out = { ...summary };
  for (const d of KNOWN_DIFFERENCES) delete out[d.key];
  return out;
}
