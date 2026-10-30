// CLAUDE.md: "Every public symbol has TSDoc with a description and a working example; a test
// checks that the examples compile." This test does both:
// 1. Every ```ts block inside a /** ... */ comment in packages/xlvane/src is written to its own
//    module in a temporary project, where `xlvane` and `xlvane/node` resolve to the sources, and
//    the project is type-checked with one run of the repository's `tsc` (TypeScript 7 CLI).
// 2. Every symbol exported from the public entry points (src/index.ts, src/node.ts) has a TSDoc
//    comment with an @example at its declaration.
import { spawnSync } from 'node:child_process';
import { mkdtemp, mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const ROOT = fileURLToPath(new URL('../..', import.meta.url));
const SRC = join(ROOT, 'packages/xlvane/src');
const ENTRY_POINTS = ['index.ts', 'node.ts'];

interface Example {
  /** Source file relative to the repository root. */
  readonly file: string;
  /** 1-based line in the source file of the first line of code in the block. */
  readonly line: number;
  readonly code: string;
}

async function sourceFiles(dir: string): Promise<string[]> {
  const out: string[] = [];
  for (const name of (await readdir(dir)).sort()) {
    const path = join(dir, name);
    if (name.endsWith('.ts')) out.push(path);
    else if (!name.includes('.')) out.push(...(await sourceFiles(path)));
  }
  return out;
}

/** Extracts the ```ts blocks of all /** comments of a file. */
function extractExamples(file: string, text: string): Example[] {
  const examples: Example[] = [];
  const lines = text.split('\n');
  let inComment = false;
  let block: string[] | null = null;
  let blockLine = 0;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]!;
    if (!inComment && /^\s*\/\*\*/.test(line)) inComment = true;
    if (inComment) {
      const content = line.replace(/^\s*\/?\*+\/?\s?/, '').replace(/\s*\*\/\s*$/, '');
      if (block === null && /^```(ts|typescript)\s*$/.test(content.trim())) {
        block = [];
        blockLine = i + 2;
      } else if (block !== null && content.trim() === '```') {
        examples.push({ file, line: blockLine, code: block.join('\n') });
        block = null;
      } else if (block !== null) {
        block.push(line.replace(/^\s*\* ?/, ''));
      }
      if (line.includes('*/')) {
        inComment = false;
        block = null;
      }
    }
  }
  return examples;
}

async function allExamples(): Promise<Example[]> {
  const out: Example[] = [];
  for (const path of await sourceFiles(SRC)) {
    out.push(...extractExamples(path.slice(ROOT.length).replace(/\\/g, '/'), await readFile(path, 'utf8')));
  }
  return out;
}

describe('TSDoc examples', () => {
  it('compile', { timeout: 120_000 }, async () => {
    const examples = await allExamples();
    expect(examples.length).toBeGreaterThan(10);
    // Inside the repository, so that examples can import development dependencies such as zod
    // (they show how readTable is used with schema libraries) through node_modules resolution.
    const cache = join(ROOT, 'node_modules', '.cache');
    await mkdir(cache, { recursive: true });
    const dir = await mkdtemp(join(cache, 'xlvane-tsdoc-'));
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
          },
        },
        files: [join(ROOT, 'types/node-modules.d.ts')],
        include: [join(dir, 'examples', '*.ts')],
      };
      await writeFile(join(dir, 'tsconfig.json'), JSON.stringify(tsconfig, null, 2));
      const tsc = join(ROOT, 'node_modules/typescript/bin/tsc');
      const result = spawnSync('node', [tsc, '-p', join(dir, 'tsconfig.json'), '--pretty', 'false'], {
        encoding: 'utf8',
        timeout: 110_000,
      });
      if (result.error) throw result.error;
      const output = `${result.stdout}${result.stderr}`;
      // Diagnostics look like "path/example-007.ts(3,5): error TS2304: Cannot find name 'x'."
      const problems = output
        .split('\n')
        .filter((l) => l.trim() !== '')
        .map((l) => {
          const m = /example-(\d{3})\.ts\((\d+),(\d+)\): (.*)$/.exec(l);
          if (!m) return l;
          const example = names.get(`example-${m[1]!}.ts`)!;
          return `${example.file}:${String(example.line + Number(m[2]) - 1)}: ${m[4]!}`;
        });
      expect(problems, `TSDoc examples that do not compile:\n${problems.join('\n')}`).toEqual([]);
      expect(result.status).toBe(0);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});

/** Names exported from an entry point, with the source module of each (resolved relative to src). */
function exportedNames(text: string, ownModule: string): Map<string, string> {
  const names = new Map<string, string>();
  for (const m of text.matchAll(/export\s+(?:type\s+)?\{([^}]*)\}\s*from\s*'(\.[^']*)'/g)) {
    for (const raw of m[1]!.split(',')) {
      const name = raw.trim().replace(/^type\s+/, '').split(/\s+as\s+/).pop()!.trim();
      if (name !== '') names.set(name, `${m[2]!.replace(/^\.\//, '')}.ts`);
    }
  }
  for (const m of text.matchAll(/^export\s+(?:declare\s+)?(?:async\s+)?(?:function|class|interface|type|const|let|enum)\s+([A-Za-z_$][\w$]*)/gm)) {
    names.set(m[1]!, ownModule);
  }
  return names;
}

/** Whether the declaration of `name` in `text` is preceded by a /** comment with @example. */
function hasExample(text: string, name: string): boolean {
  const decl = new RegExp(
    `^export\\s+(?:declare\\s+)?(?:abstract\\s+)?(?:async\\s+)?(?:function|class|interface|type|const|let|enum)\\s+${name.replace(/\$/g, '\\$')}\\b`,
    'm',
  );
  const m = decl.exec(text);
  if (!m) return false;
  const before = text.slice(0, m.index).trimEnd();
  if (!before.endsWith('*/')) return false;
  const start = before.lastIndexOf('/**');
  return start !== -1 && before.slice(start).includes('@example');
}

describe('public symbols', () => {
  it('have TSDoc with an @example', async () => {
    const missing: string[] = [];
    for (const entry of ENTRY_POINTS) {
      const names = exportedNames(await readFile(join(SRC, entry), 'utf8'), entry);
      for (const [name, module] of names) {
        const text = await readFile(join(SRC, module), 'utf8');
        if (!hasExample(text, name)) missing.push(`${name} (packages/xlvane/src/${module})`);
      }
    }
    expect(missing, `Public symbols without an @example:\n${missing.join('\n')}`).toEqual([]);
  });
});
