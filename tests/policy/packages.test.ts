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
};

/** Development dependencies approved by the maintainer (docs/decisions.md, D-005, D-301, D-706, D-906). */
const APPROVED_DEV_DEPENDENCIES = [
  '@arethetypeswrong/cli',
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

const PACKAGE_DIRS = ['packages/xlvane', 'packages/exceljs'];

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
