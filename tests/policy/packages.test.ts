// Repository policy from CLAUDE.md, checked on package manifests and sources.
import { readdir, readFile } from 'node:fs/promises';
import { join } from 'node:path';
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
  '@xlvane/mcp': ['xlvane', '@modelcontextprotocol/sdk'],
};

/** Development dependencies approved by the maintainer (docs/decisions.md, D-005, D-301). */
const APPROVED_DEV_DEPENDENCIES = [
  '@arethetypeswrong/cli',
  'arktype',
  'fast-check',
  'publint',
  'tsdown',
  'typescript',
  'valibot',
  'vite',
  'vitest',
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

  it.each(PACKAGE_DIRS)('%s is ESM-only, MIT-licensed and ships only dist', async (dir) => {
    const manifest = await readManifest(dir);
    expect(manifest.type).toBe('module');
    expect(manifest.license).toBe('MIT');
    expect(manifest.files).toEqual(['dist']);
    expect(conditionNames(manifest.exports)).not.toContain('require');
    expect(conditionNames(manifest.exports)).not.toContain('node-addons');
    const license = await readFile(join(root, dir, 'LICENSE'), 'utf8');
    expect(license.startsWith('MIT License\n')).toBe(true);
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

  it('do not import Node.js built-ins outside src/node.ts', async () => {
    const srcDir = join(root, 'packages/xlvane/src');
    const offenders: string[] = [];
    for (const file of await listTsFiles(srcDir)) {
      if (file === join(srcDir, 'node.ts')) continue;
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
});
