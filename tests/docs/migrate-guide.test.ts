// docs/migrate-from-exceljs.md shows `xlvane migrate` commands and a report. This test runs every
// `npx xlvane ...` command of the guide on the example project the guide describes, compares the
// report with the guide's ```text block, and checks the facts the guide states about the command
// against its source. docs/decisions.md D-K02.
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { ALIAS } from '../../packages/xlvane/src/cli/commands/migrate';
import { cli, splitCommand } from '../cli/helpers';
import { codeBlocks } from './markdown';

const ROOT = fileURLToPath(new URL('../..', import.meta.url));
const GUIDE = 'docs/migrate-from-exceljs.md';

/** The example project of the guide: "exceljs": "^4.4.0", src/report.js and src/legacy.cjs. */
const PACKAGE_JSON = `${JSON.stringify({ name: 'my-app', version: '1.0.0', dependencies: { exceljs: '^4.4.0' } }, null, 2)}\n`;

async function withProject<T>(fn: (dir: string) => Promise<T>): Promise<T> {
  const dir = await mkdtemp(join(tmpdir(), 'xlvane-guide-'));
  try {
    await mkdir(join(dir, 'src'));
    await writeFile(join(dir, 'package.json'), PACKAGE_JSON);
    await writeFile(join(dir, 'src/report.js'), "import ExcelJS from 'exceljs';\n\nconst workbook = new ExcelJS.Workbook();\n");
    await writeFile(join(dir, 'src/legacy.cjs'), "const BaseXform = require('exceljs/lib/xlsx/xform/base-xform');\n");
    return await fn(dir);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

const guide = async () => readFile(join(ROOT, GUIDE), 'utf8');

/** `npx xlvane ...` lines of the guide's ```sh blocks, without `npx` and trailing comments. */
async function xlvaneCommands(): Promise<string[][]> {
  const out: string[][] = [];
  for (const block of codeBlocks(await guide(), /^sh$/)) {
    for (const line of block.code.split('\n')) {
      const command = line.replace(/\s+#.*$/, '').trim();
      if (command.startsWith('npx xlvane ')) out.push(splitCommand(command).slice(2));
    }
  }
  return out;
}

describe('docs/migrate-from-exceljs.md', () => {
  it('shows the report that xlvane migrate prints for the example project', async () => {
    const text = codeBlocks(await guide(), /^text$/);
    expect(text.length).toBeGreaterThanOrEqual(1);
    await withProject(async (dir) => {
      const result = await cli(['migrate'], { cwd: dir });
      expect(result.code).toBe(0);
      expect(result.stdout.trimEnd()).toBe(text[0]!.code);
      expect(await readFile(join(dir, 'package.json'), 'utf8')).toBe(PACKAGE_JSON);
    });
  });

  it('runs every xlvane command of the guide, and only --force without --dry-run writes', async () => {
    const commands = await xlvaneCommands();
    expect(commands.map((c) => c.join(' '))).toEqual(['migrate', 'migrate --json', 'migrate --force --dry-run', 'migrate --force']);
    for (const argv of commands) {
      await withProject(async (dir) => {
        const result = await cli(argv, { cwd: dir });
        expect(result.code, argv.join(' ')).toBe(0);
        if (argv.includes('--json')) expect(JSON.parse(result.stdout)).toMatchObject({ ok: true, command: 'migrate' });
        const written = argv.includes('--force') && !argv.includes('--dry-run');
        const after = await readFile(join(dir, 'package.json'), 'utf8');
        // --force replaces only the value of the "exceljs" entry.
        expect(after, argv.join(' ')).toBe(written ? PACKAGE_JSON.replace('"^4.4.0"', `"${ALIAS}"`) : PACKAGE_JSON);
      });
    }
  });

  it('states the alias and the limits of the command as the source has them', async () => {
    const text = await guide();
    expect(ALIAS).toBe('npm:@xlvane/compat@^1');
    expect(codeBlocks(text, /^json$/)[0]!.code).toContain(`"exceljs": "${ALIAS}"`);
    const source = await readFile(join(ROOT, 'packages/xlvane/src/cli/commands/migrate.ts'), 'utf8');
    expect(source).toContain("const SECTIONS = ['dependencies', 'devDependencies', 'optionalDependencies', 'peerDependencies'] as const;");
    expect(text).toContain('`dependencies`, `devDependencies`, `optionalDependencies` and `peerDependencies`');
    expect(source).toContain('const SOURCE = /\\.(?:js|mjs|cjs|jsx|ts|mts|cts|tsx)$/;');
    expect(text).toContain('(`.js`, `.mjs`, `.cjs`, `.jsx`, `.ts`, `.mts`, `.cts`, `.tsx`, up to 2 MiB each and 10 000 files)');
    expect(source).toContain('const MAX_FILES = 10_000;');
    expect(source).toContain('const MAX_FILE_BYTES = 2 * 1024 * 1024;');
    const compat = JSON.parse(await readFile(join(ROOT, 'packages/compat/package.json'), 'utf8')) as { engines: { node: string } };
    expect(compat.engines.node).toBe('>=22.12');
    expect(text).toContain('`engines.node` `>=22.12`');
  });
});
