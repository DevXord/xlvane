// Runs the sample scripts against the real exceljs 4.4.0 (from bench/, installed by
// `npm ci --prefix bench`) and checks that they still produce the stored summary
// sample-project/expected-exceljs-4.4.0.json, which alias.test.ts compares @xlvane/exceljs
// with. Skipped when bench/node_modules is not installed (CI and `npm test` do not install it).
//
// Differences between the two libraries are not hidden by loosening assertions: they are
// listed in KNOWN_DIFFERENCES (project-helper.ts) and in docs/exceljs-compat.md.
import { existsSync } from 'node:fs';
import { rm } from 'node:fs/promises';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { copyDir, expectedSummary, KNOWN_DIFFERENCES, makeProject, ROOT, runScript } from './project-helper';

const BENCH_MODULES = join(ROOT, 'bench/node_modules');
const installed = existsSync(join(BENCH_MODULES, 'exceljs/package.json'));
if (!installed) console.warn('exceljs cross-check skipped: run `npm ci --prefix bench` to install exceljs 4.4.0.');

let dir = '';

beforeAll(async () => {
  if (!installed) return;
  dir = await makeProject('xlvane-exceljs-real-');
  // The whole bench/node_modules tree, so that exceljs finds its own dependencies.
  await copyDir(BENCH_MODULES, join(dir, 'node_modules'));
}, 300_000);

afterAll(async () => {
  if (dir !== '') await rm(dir, { recursive: true, force: true });
});

describe.skipIf(!installed)('cross-check with exceljs 4.4.0', () => {
  it('produces the stored summary from ESM and CommonJS', async () => {
    const expected = await expectedSummary();
    expect(runScript(dir, 'esm.mjs')).toEqual(expected);
    const cjs = runScript(dir, 'cjs.cjs');
    expect(cjs['requireType']).toBe('object');
    delete cjs['requireType'];
    expect(cjs).toEqual(expected);
  });

  it('resolves the documented deep imports', () => {
    const summary = runScript(dir, 'deep-imports.cjs');
    for (const value of Object.values(summary)) expect(value).toEqual({ workbook: 'function', value: 1, valueTypeNumber: 2 });
  });

  it('rejects named ESM imports, which @xlvane/exceljs adds', () => {
    expect(() => runScript(dir, 'esm-named.mjs')).toThrow(/Named export '\w+' not found/);
  });

  it('lists each known difference with a reason', () => {
    for (const d of KNOWN_DIFFERENCES) expect(d.reason.length).toBeGreaterThan(10);
  });
});
