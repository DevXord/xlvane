// The ```ts blocks of the user guides in docs/ must compile, as the TSDoc and README examples do
// (tests/policy/tsdoc-examples.test.ts, D-314). Each block is written to its own module in a
// temporary project where `xlvane`, `xlvane/node`, `@xlvane/exceljs` and `exceljs` (the npm alias
// of @xlvane/exceljs) resolve to the sources, and the project is type-checked with one run of the
// repository's `tsc`. docs/decisions.md D-K02.
import { spawnSync } from 'node:child_process';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { codeBlocks } from './markdown';

const ROOT = fileURLToPath(new URL('../..', import.meta.url));
const SRC = join(ROOT, 'packages/xlvane/src');
const COMPAT_SRC = join(ROOT, 'packages/exceljs/src');
/** Guides whose ```ts blocks must compile, with the least number of blocks each must have. */
const GUIDES: readonly (readonly [string, number])[] = [
  ['docs/migrate-from-exceljs.md', 2],
  ['docs/xlsx-cves.md', 1],
];

interface Example {
  readonly file: string;
  /** 1-based line of the first line of code. */
  readonly line: number;
  readonly code: string;
}

describe('TypeScript examples of the guides in docs/', () => {
  it('compile', { timeout: 120_000 }, async () => {
    const examples: Example[] = [];
    for (const [file, least] of GUIDES) {
      const blocks = codeBlocks(await readFile(join(ROOT, file), 'utf8'), /^(?:ts|typescript)$/);
      expect(blocks.length, file).toBeGreaterThanOrEqual(least);
      for (const b of blocks) examples.push({ file, ...b });
    }
    // Inside the repository, so that node_modules resolution works as in the TSDoc test.
    const cache = join(ROOT, 'node_modules', '.cache');
    await mkdir(cache, { recursive: true });
    const dir = await mkdtemp(join(cache, 'xlvane-guides-'));
    try {
      await mkdir(join(dir, 'examples'));
      const names = new Map<string, Example>();
      for (const [i, example] of examples.entries()) {
        const name = `example-${String(i).padStart(3, '0')}.ts`;
        names.set(name, example);
        // `export {}` makes every block a module, so top-level await and repeated names are fine.
        await writeFile(join(dir, 'examples', name), `${example.code}\nexport {};\n`);
      }
      const tsconfig = {
        compilerOptions: {
          target: 'es2022',
          module: 'esnext',
          moduleResolution: 'bundler',
          lib: ['ES2023', 'DOM'],
          types: [],
          strict: true,
          noUncheckedIndexedAccess: true,
          exactOptionalPropertyTypes: true,
          isolatedModules: true,
          skipLibCheck: true,
          noEmit: true,
          paths: {
            xlvane: [join(SRC, 'index.ts')],
            'xlvane/node': [join(SRC, 'node.ts')],
            '@xlvane/exceljs': [join(COMPAT_SRC, 'index.ts')],
            exceljs: [join(COMPAT_SRC, 'index.ts')],
            // @xlvane/exceljs's package imports (D-R07), mapped to the sources as in vitest.config.ts,
            // so the check does not need a build.
            '#xlvane-node': [join(SRC, 'node.ts')],
            '#node-fs': [join(COMPAT_SRC, 'node-fs.ts')],
          },
        },
        files: [join(ROOT, 'types/node-modules.d.ts')],
        include: [join(dir, 'examples', '*.ts')],
      };
      await writeFile(join(dir, 'tsconfig.json'), JSON.stringify(tsconfig, null, 2));
      const tsc = join(ROOT, 'node_modules/typescript/bin/tsc');
      const result = spawnSync('node', [tsc, '-p', join(dir, 'tsconfig.json'), '--pretty', 'false'], { encoding: 'utf8', timeout: 110_000 });
      if (result.error) throw result.error;
      const problems = `${result.stdout}${result.stderr}`
        .split('\n')
        .filter((l) => l.trim() !== '')
        .map((l) => {
          const m = /example-(\d{3})\.ts\((\d+),(\d+)\): (.*)$/.exec(l);
          if (!m) return l;
          const example = names.get(`example-${m[1]!}.ts`)!;
          return `${example.file}:${String(example.line + Number(m[2]) - 1)}: ${m[4]!}`;
        });
      expect(problems, `Guide examples that do not compile:\n${problems.join('\n')}`).toEqual([]);
      expect(result.status).toBe(0);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
