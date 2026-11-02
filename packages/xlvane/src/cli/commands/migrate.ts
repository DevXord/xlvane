import { lstat, readdir, readFile } from 'node:fs/promises';
import { join, relative } from 'node:path';
import { XlvError } from '../../errors';
import { flag } from '../args';
import { readText, resolvePath, writeOutput, wroteLine, type WriteOutcome } from '../files';
import { commandLine, plural, type Output } from '../output';
import type { CommandContext, CommandSpec } from '../types';

/** The dependency value that installs @xlvane/compat under the name exceljs (npm alias). */
export const ALIAS = 'npm:@xlvane/compat@^1';
const SECTIONS = ['dependencies', 'devDependencies', 'optionalDependencies', 'peerDependencies'] as const;
const SOURCE = /\.(?:js|mjs|cjs|jsx|ts|mts|cts|tsx)$/;
const MAX_FILES = 10_000;
const MAX_FILE_BYTES = 2 * 1024 * 1024;

/**
 * exceljs API that @xlvane/compat does not support or supports only partly
 * (docs/exceljs-compat.md). Matching is textual, so a hit is a place to look at, not proof.
 */
const PATTERNS: readonly { readonly api: string; readonly status: 'not supported' | 'partly'; readonly re: RegExp }[] = [
  { api: 'ModelContainer', status: 'not supported', re: /\bModelContainer\b/ },
  { api: 'deep import (exceljs/lib/...)', status: 'not supported', re: /['"]exceljs\/lib\/[^'"]*['"]/ },
];

/** A place in the source code that uses exceljs API outside the supported subset. */
export interface MigrateFinding {
  readonly file: string;
  readonly line: number;
  readonly column: number;
  readonly api: string;
  readonly status: 'not supported' | 'partly';
}

async function sourceFiles(root: string): Promise<{ files: string[]; limited: boolean }> {
  const files: string[] = [];
  let limited = false;
  const walk = async (dir: string): Promise<void> => {
    let names: string[];
    try {
      names = (await readdir(dir)).sort();
    } catch {
      return;
    }
    for (const name of names) {
      if (files.length >= MAX_FILES) {
        limited = true;
        return;
      }
      if (name === 'node_modules' || name.startsWith('.')) continue;
      const path = join(dir, name);
      // lstat: symbolic links are not followed, so a link loop cannot make the walk exponential
      // and a link cannot lead it outside the project (D-H01).
      let info;
      try {
        info = await lstat(path);
      } catch {
        continue;
      }
      if (info.isSymbolicLink()) continue;
      if (info.isDirectory()) await walk(path);
      else if (info.isFile() && SOURCE.test(name) && info.size <= MAX_FILE_BYTES) files.push(path);
    }
  };
  await walk(root);
  return { files, limited };
}

/** Finds unsupported exceljs API in files that mention exceljs. */
export function scanSource(file: string, text: string): MigrateFinding[] {
  if (!text.includes('exceljs')) return [];
  const findings: MigrateFinding[] = [];
  text.split(/\r?\n/).forEach((line, i) => {
    for (const p of PATTERNS) {
      const m = p.re.exec(line);
      if (m) findings.push({ file, line: i + 1, column: m.index + 1, api: p.api, status: p.status });
    }
  });
  return findings;
}

/**
 * Offsets `[start, end)` of the string value (with its quotes) at `path` in valid JSON text, the
 * one JSON.parse reads when keys repeat, or `null`. Keys are compared after decoding their
 * escapes. Values off the path are skipped without recursion.
 */
export function stringValueSpan(text: string, path: readonly string[]): [number, number] | null {
  let i = 0;
  const space = (): void => {
    while (i < text.length && ' \t\r\n'.includes(text[i]!)) i++;
  };
  /** Skips a string starting at i (on its opening quote); returns its raw text. */
  const string = (): string => {
    const start = i++;
    while (text[i] !== '"') i += text[i] === '\\' ? 2 : 1;
    i++;
    return text.slice(start, i);
  };
  /** Skips any value starting at i, iteratively. */
  const skip = (): void => {
    let depth = 0;
    do {
      const c = text[i]!;
      if (c === '"') string();
      else {
        if (c === '{' || c === '[') depth++;
        else if (c === '}' || c === ']') depth--;
        i++;
        if (depth === 0 && c !== '{' && c !== '[') {
          // A literal or number: skip to the next delimiter.
          while (i < text.length && !',]} \t\r\n'.includes(text[i]!)) i++;
        }
      }
    } while (depth > 0 && i < text.length);
  };
  const value = (level: number): [number, number] | null => {
    space();
    if (level === path.length) {
      if (text[i] !== '"') {
        skip();
        return null;
      }
      const start = i;
      string();
      return [start, i];
    }
    if (text[i] !== '{') {
      skip();
      return null;
    }
    i++;
    let found: [number, number] | null = null;
    space();
    if (text[i] === '}') {
      i++;
      return null;
    }
    for (;;) {
      space();
      const key = JSON.parse(string()) as string;
      space();
      i++; // ':'
      if (key === path[level]) found = value(level + 1);
      else {
        space();
        skip();
      }
      space();
      if (text[i++] === '}') return found;
    }
  };
  return value(0);
}

function jsonError(path: string, cause?: unknown): XlvError {
  return new XlvError('XLV_E_JSON_INVALID', `"${path}" is not a valid package.json.`, { hint: 'Fix the JSON syntax of package.json and run the command again.', details: { path }, cause });
}

async function run(ctx: CommandContext) {
  const dir = ctx.args.positionals[0] ?? '.';
  const pkgPath = join(dir, 'package.json');
  const text = await readText(ctx, pkgPath);
  let pkg: unknown;
  try {
    pkg = JSON.parse(text);
  } catch (error) {
    throw jsonError(pkgPath, error);
  }
  if (typeof pkg !== 'object' || pkg === null || Array.isArray(pkg)) throw jsonError(pkgPath);
  const p = pkg as Record<string, unknown>;
  const deps: { section: string; from: string }[] = [];
  for (const section of SECTIONS) {
    const s = p[section];
    if (typeof s === 'object' && s !== null && typeof (s as Record<string, unknown>)['exceljs'] === 'string') deps.push({ section, from: (s as Record<string, string>)['exceljs']! });
  }
  const engines = typeof p['engines'] === 'object' && p['engines'] !== null ? (p['engines'] as Record<string, unknown>)['node'] : undefined;
  const changes = deps.filter((d) => d.from !== ALIAS).map((d) => ({ ...d, to: ALIAS }));
  const alreadyMigrated = deps.length > 0 && changes.length === 0;

  const root = resolvePath(ctx, dir);
  const { files, limited } = await sourceFiles(root);
  const findings: MigrateFinding[] = [];
  for (const file of files) {
    let source: string;
    try {
      source = new TextDecoder().decode(await readFile(file));
    } catch {
      continue;
    }
    findings.push(...scanSource(relative(root, file).split('\\').join('/'), source));
  }

  let write: WriteOutcome | null = null;
  if (changes.length > 0 && flag(ctx.args, 'force')) {
    // The values are replaced in place, so the file keeps its formatting and key order; each is
    // the one JSON.parse read, located by its section (not the first "exceljs" in the file, which
    // could be in "overrides" or a string, D-H01).
    const spans = changes.map((c) => stringValueSpan(text, [c.section, 'exceljs'])).filter((s): s is [number, number] => s !== null);
    let updated = '';
    let at = 0;
    for (const [start, end] of spans.sort((a, b) => a[0] - b[0])) {
      updated += `${text.slice(at, start)}"${ALIAS}"`;
      at = end;
    }
    updated += text.slice(at);
    write = await writeOutput(ctx, pkgPath, updated, { force: true, dryRun: flag(ctx.args, 'dry-run') });
  }
  const result = {
    directory: dir,
    packageJson: pkgPath,
    changes,
    alreadyMigrated,
    enginesNode: typeof engines === 'string' ? engines : null,
    filesScanned: files.length,
    scanLimited: limited,
    findings,
    written: write?.written ?? false,
  };
  return {
    result,
    exitCode: 0 as const,
    render: (out: Output) => {
      const lines: string[] = [];
      if (deps.length === 0) lines.push(`${pkgPath} has no "exceljs" dependency.`);
      else if (alreadyMigrated) lines.push(`${pkgPath} already uses ${ALIAS} for "exceljs".`);
      else {
        lines.push(`${out.bold('Change in')} ${pkgPath}:`);
        for (const c of changes) lines.push(`  ${c.section}: ${out.red(`- "exceljs": "${c.from}"`)}`, `  ${' '.repeat(c.section.length)}  ${out.green(`+ "exceljs": "${c.to}"`)}`);
      }
      lines.push(
        `Node.js: require('exceljs') of @xlvane/compat needs Node.js 22.12 or later (package.json engines.node: ${typeof engines === 'string' ? `"${engines}"` : 'not set'}).`,
      );
      lines.push('', `Scanned ${plural(files.length, 'source file')}${limited ? ` (stopped at ${String(MAX_FILES)})` : ''} for exceljs API that @xlvane/compat does not fully support:`);
      if (findings.length === 0) lines.push('  none found.');
      for (const f of findings.slice(0, 200)) lines.push(`  ${f.status === 'partly' ? out.yellow('partly       ') : out.red('not supported')} ${f.file}:${String(f.line)}:${String(f.column)}  ${f.api}`);
      if (findings.length > 200) lines.push(out.dim(`  … ${String(findings.length - 200)} more; use --json for all`));
      lines.push('  Details: docs/exceljs-compat.md. Unsupported calls throw XLV_E_NOT_SUPPORTED at run time.');
      lines.push('');
      if (write) lines.push(wroteLine(write));
      else if (changes.length > 0) lines.push(`Nothing was written. Run "${commandLine(['migrate', ...(dir === '.' ? [] : [dir]), '--force'])}" to change package.json.`);
      return lines.join('\n');
    },
  };
}

export const migrateCommand: CommandSpec = {
  name: 'migrate',
  summary: 'Switch a project from exceljs to @xlvane/compat',
  description:
    'Prepares a project for @xlvane/compat, which offers the commonly used exceljs API. It shows the change to package.json ("exceljs": "npm:@xlvane/compat@^1", an npm alias, so the code keeps importing "exceljs") and lists places in the source code that use exceljs API outside the supported subset. Only files that mention exceljs are scanned (node_modules and dot folders are skipped), and matching is textual. package.json is changed only with --force; nothing else is changed.',
  usage: '[project-dir] [options]',
  positionals: { min: 0, max: 1, names: ['project-dir'] },
  writes: true,
  options: {},
  examples: [
    { command: 'xlvane migrate', description: 'Report for the project in the current directory' },
    { command: 'xlvane migrate my-app --force', description: 'Change my-app/package.json' },
    { command: 'xlvane migrate my-app --json', description: 'The report as JSON' },
  ],
  related: ['help'],
  run,
};
