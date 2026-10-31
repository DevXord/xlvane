// TypeScript users of exceljs write `ExcelJS.Worksheet`, `ExcelJS.Cell`, `ExcelJS.Fill` and so on
// as types. This test builds both packages with their declaration files, installs
// @xlvane/compat as node_modules/exceljs (as the npm alias does, see alias.test.ts) and
// type-checks code from the exceljs 4.4.0 README (types-sample/readme-body.txt) with the
// import forms and module settings TypeScript projects use:
// - `import * as ExcelJS` and `import ExcelJS from` with `moduleResolution: bundler`;
// - an ES module and a CommonJS module (`import ExcelJS = require('exceljs')`) with `module: nodenext`;
// - `module: commonjs` with its default resolution (TypeScript 7 removed `node10`, TS5108), which
//   reads the top-level `types` field of package.json.
// The sample also compiles against exceljs 4.4.0's own index.d.ts (checked when it was written,
// docs/decisions.md D-R01). No @types/node is used: the declarations must not need it.
import { spawnSync } from 'node:child_process';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { buildPackage, REPO } from '../helpers/build-package';

let dir = '';
let body = '';

beforeAll(async () => {
  const cache = join(REPO, 'node_modules', '.cache');
  await mkdir(cache, { recursive: true });
  dir = await mkdtemp(join(cache, 'xlvane-types-'));
  const modules = join(dir, 'node_modules');
  await buildPackage('@xlvane/compat', modules, 'exceljs', { dts: true });
  await buildPackage('xlvane', modules, 'xlvane', { dts: true });
  body = await readFile(join(REPO, 'tests/exceljs/types-sample/readme-body.txt'), 'utf8');
}, 400_000);

afterAll(async () => {
  if (dir !== '') await rm(dir, { recursive: true, force: true });
});

interface Variant {
  readonly name: string;
  readonly packageType?: 'module' | 'commonjs';
  readonly compilerOptions: Record<string, unknown>;
  /** File name → import line put before the README body. */
  readonly files: Readonly<Record<string, string>>;
}

const COMMON = {
  strict: true,
  noEmit: true,
  target: 'es2022',
  lib: ['es2022', 'dom', 'dom.asynciterable'],
  types: [],
  skipLibCheck: false,
};

const VARIANTS: readonly Variant[] = [
  {
    name: 'bundler resolution, namespace and default imports',
    compilerOptions: { module: 'esnext', moduleResolution: 'bundler', esModuleInterop: true },
    files: { 'namespace.ts': "import * as ExcelJS from 'exceljs';", 'default.ts': "import ExcelJS from 'exceljs';" },
  },
  {
    name: 'nodenext, ES module',
    packageType: 'module',
    compilerOptions: { module: 'nodenext', moduleResolution: 'nodenext' },
    files: { 'namespace.mts': "import * as ExcelJS from 'exceljs';", 'default.mts': "import ExcelJS from 'exceljs';" },
  },
  {
    name: 'nodenext, CommonJS module (import = require)',
    packageType: 'commonjs',
    compilerOptions: { module: 'nodenext', moduleResolution: 'nodenext' },
    files: { 'require.cts': "import ExcelJS = require('exceljs');" },
  },
  {
    name: 'module commonjs with its default resolution',
    packageType: 'commonjs',
    compilerOptions: { module: 'commonjs', esModuleInterop: true },
    files: { 'namespace.ts': "import * as ExcelJS from 'exceljs';", 'default.ts': "import ExcelJS from 'exceljs';" },
  },
];

describe('exceljs README code type-checks against @xlvane/compat', () => {
  for (const [i, variant] of VARIANTS.entries()) {
    it(variant.name, { timeout: 120_000 }, async () => {
      const project = join(dir, `variant-${String(i)}`);
      await mkdir(project, { recursive: true });
      if (variant.packageType !== undefined) await writeFile(join(project, 'package.json'), JSON.stringify({ type: variant.packageType }));
      for (const [file, header] of Object.entries(variant.files)) await writeFile(join(project, file), `${header}\n${body}`);
      await writeFile(
        join(project, 'tsconfig.json'),
        JSON.stringify({ compilerOptions: { ...COMMON, ...variant.compilerOptions }, files: Object.keys(variant.files) }, null, 2),
      );
      const tsc = join(REPO, 'node_modules/typescript/bin/tsc');
      const result = spawnSync(process.execPath, [tsc, '-p', join(project, 'tsconfig.json'), '--pretty', 'false'], { encoding: 'utf8', timeout: 110_000 });
      if (result.error) throw result.error;
      const problems = `${result.stdout}${result.stderr}`.split('\n').filter((l) => l.trim() !== '');
      expect(problems, problems.join('\n')).toEqual([]);
      expect(result.status).toBe(0);
    });
  }
});
