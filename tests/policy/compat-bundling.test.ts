// @xlvane/compat reaches xlvane/node and node:fs/promises only through its own package imports
// (#xlvane-node, #node-fs), which resolve to dist/no-node.js without the `node` condition, so a
// bundler for browsers or Workers (esbuild, vite, wrangler) never follows a node: import (D-R07).
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';

const DIR = 'packages/compat';
const manifest = JSON.parse(readFileSync(join(DIR, 'package.json'), 'utf8')) as { imports?: Record<string, Record<string, string>> };

describe('@xlvane/compat bundling (D-R07)', () => {
  it('maps its Node-only imports with a node condition and a throwing default', () => {
    expect(manifest.imports).toEqual({
      '#xlvane-node': { node: 'xlvane/node', default: './dist/no-node.js' },
      '#node-fs': { node: './dist/node-fs.js', default: './dist/no-node.js' },
    });
  });

  it('imports Node-only modules in its sources only through those package imports', () => {
    // src/node-fs.ts is the one module that imports node:fs/promises; it is reached only through #node-fs.
    const sources = readdirSync(join(DIR, 'src')).filter((f) => f.endsWith('.ts') && f !== 'node-fs.ts');
    const dynamic: string[] = [];
    for (const file of sources) {
      const text = readFileSync(join(DIR, 'src', file), 'utf8');
      expect(text, file).not.toMatch(/from ['"]node:/);
      expect(text, file).not.toMatch(/^import (?!type )[^;]*from ['"]xlvane\/node['"]/m);
      for (const m of text.matchAll(/\bimport\(([^)]*)\)/g)) dynamic.push(`${file}: ${m[1]!.trim()}`);
    }
    expect(dynamic.sort()).toEqual(["csv.ts: '#node-fs'", "node-api.ts: '#xlvane-node'"]);
  });

  it('src/node-fs.ts only re-exports node:fs/promises', () => {
    expect(readFileSync(join(DIR, 'src', 'node-fs.ts'), 'utf8')).toMatch(/^export \{ readFile, writeFile \} from 'node:fs\/promises';$/m);
  });

  it('the fallback module fails to load', async () => {
    const source = readFileSync(join(DIR, 'src', 'no-node.ts'), 'utf8');
    expect(source).toMatch(/^throw new Error\(/m);
  });
});
