import { readdir, readFile, stat } from 'node:fs/promises';
import { join, relative } from 'node:path';
import { XlvError } from '../../errors';
import { flag } from '../args';
import { readText, resolvePath, writeOutput, wroteLine, type WriteOutcome } from '../files';
import { commandLine, plural, type Output } from '../output';
import type { CommandContext, CommandSpec } from '../types';

/** The dependency value that installs @xlvane/exceljs under the name exceljs (npm alias). */
export const ALIAS = 'npm:@xlvane/exceljs@^1';
const SECTIONS = ['dependencies', 'devDependencies', 'optionalDependencies', 'peerDependencies'] as const;
const SOURCE = /\.(?:js|mjs|cjs|jsx|ts|mts|cts|tsx)$/;
const MAX_FILES = 10_000;
const MAX_FILE_BYTES = 2 * 1024 * 1024;

/**
 * exceljs API that @xlvane/exceljs does not support or supports only partly
 * (docs/exceljs-compat.md). Matching is textual, so a hit is a place to look at, not proof.
 */
const PATTERNS: readonly { readonly api: string; readonly status: 'not supported' | 'partly'; readonly re: RegExp }[] = [
  { api: 'workbook.csv', status: 'not supported', re: /\.csv\s*\.\s*(?:read|write|readFile|writeFile|writeBuffer|load)\b/ },
  { api: 'addImage', status: 'not supported', re: /\.addImage\s*\(/ },
  { api: 'addTable', status: 'not supported', re: /\.addTable\s*\(/ },
  { api: 'protect with a password', status: 'partly', re: /\.protect\s*\(\s*(?!\)|''|""|undefined\b|null\b)/ },
  { api: 'ModelContainer', status: 'not supported', re: /\bModelContainer\b/ },
  { api: 'deep import (exceljs/...)', status: 'not supported', re: /['"]exceljs\/[^'"]+['"]/ },
  { api: 'sharedFormula values', status: 'partly', re: /\bsharedFormula\s*:/ },
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
      let info;
      try {
        info = await stat(path);
      } catch {
        continue;
      }
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
    let updated = text;
    for (const c of changes) {
      // Replace the value in place so the file keeps its formatting and key order.
      const escaped = c.from.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      updated = updated.replace(new RegExp(`("exceljs"\\s*:\\s*)"${escaped}"`), `$1"${ALIAS}"`);
    }
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
        `Node.js: require('exceljs') of @xlvane/exceljs needs Node.js 22.12 or later (package.json engines.node: ${typeof engines === 'string' ? `"${engines}"` : 'not set'}).`,
      );
      lines.push('', `Scanned ${plural(files.length, 'source file')}${limited ? ` (stopped at ${String(MAX_FILES)})` : ''} for exceljs API that @xlvane/exceljs does not fully support:`);
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
  summary: 'Switch a project from exceljs to @xlvane/exceljs',
  description:
    'Prepares a project for @xlvane/exceljs, which offers the commonly used exceljs API. It shows the change to package.json ("exceljs": "npm:@xlvane/exceljs@^1", an npm alias, so the code keeps importing "exceljs") and lists places in the source code that use exceljs API outside the supported subset. Only files that mention exceljs are scanned (node_modules and dot folders are skipped), and matching is textual. package.json is changed only with --force; nothing else is changed.',
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
