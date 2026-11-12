// Repository policy from CLAUDE.md, checked on package manifests and sources.
import { readdir, readFile } from 'node:fs/promises';
import { dirname, join, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = fileURLToPath(new URL('../../', import.meta.url));

interface Manifest {
  name: string;
  type?: string;
  license?: string;
  files?: string[];
  exports?: unknown;
  dependencies?: Record<string, string>;
  peerDependencies?: Record<string, string>;
  optionalDependencies?: Record<string, string>;
  bundleDependencies?: unknown;
  bundledDependencies?: unknown;
  devDependencies?: Record<string, string>;
}

/** Runtime dependencies each package may have (CLAUDE.md, "Hard rules"). */
const ALLOWED_RUNTIME_DEPENDENCIES: Record<string, readonly string[]> = {
  xlvane: ['fflate'],
  '@xlvane/exceljs': ['xlvane'],
  '@xlvane/mcp': ['xlvane', '@modelcontextprotocol/server'],
};

/** Development dependencies approved by the maintainer (docs/decisions.md, D-005, D-301, D-706, D-906, D-I01). */
const APPROVED_DEV_DEPENDENCIES = [
  '@arethetypeswrong/cli',
  '@modelcontextprotocol/client',
  'arktype',
  'deno',
  'fast-check',
  'playwright',
  'publint',
  'skills',
  'tsdown',
  'typescript',
  'valibot',
  'vite',
  'vitest',
  'workerd',
  'zod',
];

const PACKAGE_DIRS = ['packages/xlvane', 'packages/exceljs', 'packages/mcp'];

async function readManifest(dir: string): Promise<Manifest> {
  return JSON.parse(await readFile(join(root, dir, 'package.json'), 'utf8')) as Manifest;
}

/** Collects every export condition name used anywhere in an `exports` map. */
function conditionNames(exportsField: unknown, names: string[] = []): string[] {
  if (exportsField !== null && typeof exportsField === 'object') {
    for (const [key, value] of Object.entries(exportsField)) {
      if (!key.startsWith('.')) names.push(key);
      conditionNames(value, names);
    }
  }
  return names;
}

async function listTsFiles(dir: string): Promise<string[]> {
  const out: string[] = [];
  for (const entry of await readdir(dir)) {
    const path = join(dir, entry);
    if (entry.endsWith('.ts')) out.push(path);
    else if (!entry.includes('.')) out.push(...(await listTsFiles(path)));
  }
  return out;
}

describe('package manifests', () => {
  it.each(PACKAGE_DIRS)('%s has only the allowed runtime dependencies', async (dir) => {
    const manifest = await readManifest(dir);
    const allowed = ALLOWED_RUNTIME_DEPENDENCIES[manifest.name];
    expect(allowed, `no dependency policy for ${manifest.name}`).toBeDefined();
    for (const dep of Object.keys(manifest.dependencies ?? {})) {
      expect(allowed).toContain(dep);
    }
    expect(manifest.peerDependencies ?? {}).toEqual({});
    expect(manifest.optionalDependencies ?? {}).toEqual({});
    expect(manifest.bundleDependencies ?? manifest.bundledDependencies).toBeUndefined();
  });

  it('xlvane depends on fflate and nothing else', async () => {
    const manifest = await readManifest('packages/xlvane');
    expect(Object.keys(manifest.dependencies ?? {})).toEqual(['fflate']);
  });

  /** Files in the tarball: dist, and for xlvane the docs, the Agent Skill and llms-full.txt (D-707). */
  const FILES: Record<string, readonly string[]> = {
    'packages/xlvane': ['dist', 'docs', 'skills', 'llms-full.txt'],
    'packages/exceljs': ['dist'],
    'packages/mcp': ['dist'],
  };

  it.each(PACKAGE_DIRS)('%s is ESM-only, MIT-licensed and ships only its listed files', async (dir) => {
    const manifest = await readManifest(dir);
    expect(manifest.type).toBe('module');
    expect(manifest.license).toBe('MIT');
    expect(manifest.files).toEqual(FILES[dir]);
    expect(conditionNames(manifest.exports)).not.toContain('require');
    expect(conditionNames(manifest.exports)).not.toContain('node-addons');
    const license = await readFile(join(root, dir, 'LICENSE'), 'utf8');
    expect(license.startsWith('MIT License\n')).toBe(true);
  });

  it('@xlvane/exceljs carries the exceljs license notice and says it is not affiliated (D-R05)', async () => {
    const license = await readFile(join(root, 'packages/exceljs', 'LICENSE'), 'utf8');
    expect(license).toContain('Copyright (c) 2014-2019 Guyon Roche');
    const readme = await readFile(join(root, 'packages/exceljs', 'README.md'), 'utf8');
    expect(readme).toContain('not affiliated with or endorsed by the exceljs project');
    // Parts of the implementation are adapted from exceljs 4.4.0 (IP review of 2026-11-11).
    expect(readme).not.toContain('contains no exceljs code');
  });

  it('@xlvane/exceljs carries the dayjs license notice for the date token expressions copied from dayjs 1.11', async () => {
    const license = await readFile(join(root, 'packages/exceljs', 'LICENSE'), 'utf8');
    expect(license).toContain('Copyright (c) 2018-present, iamkun');
    expect(license).toContain('https://github.com/iamkun/dayjs');
    const csv = await readFile(join(root, 'packages/exceljs/src/csv.ts'), 'utf8');
    for (const name of ['FORMAT_TOKENS', 'PARSE_TOKENS']) {
      const at = csv.indexOf(`const ${name} =`);
      expect(at, name).toBeGreaterThan(0);
      expect(csv.slice(Math.max(0, at - 400), at), name).toMatch(/copied from dayjs[\s\S]*MIT[\s\S]*LICENSE/);
    }
  });

  it('xlvane carries the Unicode License v3 notice for the East Asian Width table generated from the Unicode Character Database', async () => {
    // Text of the LICENSE file of github.com/unicode-org/unicodetools (D-O03).
    const license = await readFile(join(root, 'packages/xlvane', 'LICENSE'), 'utf8');
    expect(license).toContain('UNICODE LICENSE V3');
    expect(license).toContain('Copyright © 2001-2024 Unicode, Inc.');
    expect(license).toContain('provided that either (a)\nthis copyright and permission notice appear with all copies of the Data\nFiles or Software, or (b) this copyright and permission notice appear in\nassociated Documentation.');
    expect(license).toContain('SPDX-License-Identifier: Unicode-3.0');
    expect(license).toContain('EastAsianWidth-17.0.0.txt');
    for (const file of ['packages/xlvane/src/cli/east-asian-width.ts', 'scripts/gen-east-asian-width.mjs']) {
      expect(await readFile(join(root, file), 'utf8'), file).toContain('packages/xlvane/LICENSE');
    }
  });

  it.each(PACKAGE_DIRS)('%s names its bin targets as npm normalises them, without "./"', async (dir) => {
    // @npmcli/package-json 7.0.5 lib/normalize.js rewrites "./dist/cli.js" to "dist/cli.js" and warns (F-NPM-09).
    const manifest = JSON.parse(await readFile(join(root, dir, 'package.json'), 'utf8')) as { bin?: Record<string, string> };
    for (const [name, target] of Object.entries(manifest.bin ?? {})) {
      expect(target, name).toBe(target.replace(/^(\.\/)+/, ''));
    }
  });

  it('root devDependencies are all approved', async () => {
    const manifest = await readManifest('.');
    for (const dep of Object.keys(manifest.devDependencies ?? {})) {
      expect(APPROVED_DEV_DEPENDENCIES).toContain(dep);
    }
  });

  it.each(PACKAGE_DIRS)('%s has no devDependencies of its own', async (dir) => {
    const manifest = await readManifest(dir);
    expect(manifest.devDependencies ?? {}).toEqual({});
  });
});

describe('core sources', () => {
  const NODE_BUILTINS = [
    'buffer', 'child_process', 'crypto', 'fs', 'fs/promises', 'module', 'os', 'path',
    'process', 'stream', 'url', 'util', 'worker_threads', 'zlib',
  ];

  // Node.js-only modules: the xlvane/node entry, its file helpers and the CLI (D-601). The core
  // entry (src/index.ts) must not reach them, so it keeps running in browsers, Deno and Workers.
  const isNodeOnly = (srcDir: string, file: string): boolean =>
    file === join(srcDir, 'node.ts') || file === join(srcDir, 'util', 'node-files.ts') || file.startsWith(join(srcDir, 'cli') + sep);

  it('do not import Node.js built-ins outside the Node.js-only modules', async () => {
    const srcDir = join(root, 'packages/xlvane/src');
    const offenders: string[] = [];
    for (const file of await listTsFiles(srcDir)) {
      if (isNodeOnly(srcDir, file)) continue;
      const text = await readFile(file, 'utf8');
      for (const match of text.matchAll(/(?:from|import)\s*\(?\s*['"]([^'"]+)['"]/g)) {
        const specifier = match[1] ?? '';
        if (specifier.startsWith('node:') || NODE_BUILTINS.includes(specifier)) {
          offenders.push(`${file}: ${specifier}`);
        }
      }
      if (/\bprocess\.|\bBuffer\b|\brequire\(/.test(text)) offenders.push(`${file}: Node global`);
    }
    expect(offenders).toEqual([]);
  });

  it('src/index.ts does not reach the Node.js-only modules', async () => {
    const srcDir = join(root, 'packages/xlvane/src');
    const seen = new Set<string>();
    const reached: string[] = [];
    const visit = async (file: string): Promise<void> => {
      if (seen.has(file)) return;
      seen.add(file);
      if (isNodeOnly(srcDir, file)) reached.push(file);
      const text = await readFile(file, 'utf8');
      for (const match of text.matchAll(/(?:from|import)\s*\(?\s*['"](\.[^'"]+)['"]/g)) {
        const target = join(dirname(file), match[1]!);
        await visit(target.endsWith('.ts') ? target : `${target}.ts`);
      }
    };
    await visit(join(srcDir, 'index.ts'));
    expect(seen.size).toBeGreaterThan(10);
    expect(reached).toEqual([]);
  });
});
