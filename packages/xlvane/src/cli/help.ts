// Welcome screen and `help <command>` (docs/PLAN.md section 5).
import { optionsOf } from './args';
import type { Output } from './output';
import type { CommandSpec } from './types';

/** Repository of the project; the docs live there until a website exists. */
export const DOCS_URL = 'https://github.com/DevXord/xlvane';

/** Commands shown on the welcome screen, with the example and its description. */
const COMMON: readonly (readonly [string, string])[] = [
  ['xlvane inspect report.xlsx', 'See sheets, columns and problems'],
  ['xlvane read report.xlsx --sheet Sales', 'Print a sheet as a table'],
  ['xlvane convert report.xlsx out.csv', 'Convert to CSV or JSON'],
  ['xlvane create data.json new-report.xlsx', 'Build a formatted report'],
  ['xlvane diff old.xlsx new.xlsx', 'Show what changed'],
];

/** The screen shown by `xlvane` without arguments. */
export function welcome(out: Output, version: string, commands: readonly CommandSpec[]): string {
  const width = Math.max(...COMMON.map(([c]) => c.length)) + 3;
  return [
    `${out.bold(`xlvane ${version}`)} — read, write and inspect Excel files`,
    '',
    'Common tasks:',
    ...COMMON.map(([c, d]) => `  ${out.cyan(c.padEnd(width))}${d}`),
    '',
    `Commands: ${commands.map((c) => c.name).join(', ')}`,
    `More:  xlvane help <command>   ·   Docs: ${DOCS_URL}`,
    'For AI agents and scripts: add --json to any command (docs/cli-json.md)',
    '',
  ].join('\n');
}

/** Wraps a paragraph at `width` columns. */
function wrap(text: string, width: number, indent: string): string {
  const lines: string[] = [];
  let line = '';
  for (const word of text.split(/\s+/)) {
    if (line !== '' && indent.length + line.length + 1 + word.length > width) {
      lines.push(indent + line);
      line = word;
    } else line = line === '' ? word : `${line} ${word}`;
  }
  if (line !== '') lines.push(indent + line);
  return lines.join('\n');
}

/** The help of one command. */
export function commandHelp(out: Output, spec: CommandSpec): string {
  const options = optionsOf(spec);
  const labels = Object.entries(options).map(([name, o]) => {
    const long = `--${name}${o.type === 'string' ? ` <${o.value ?? 'value'}>` : ''}`;
    return [o.short ? `-${o.short}, ${long}` : `    ${long}`, o.description + (o.multiple ? ' (repeatable)' : '')] as const;
  });
  const width = Math.max(...labels.map(([l]) => l.length)) + 2;
  const exit = spec.checks ? `0 success, 1 ${spec.checks}, 2 usage error, 3 input or file error` : '0 success, 2 usage error, 3 input or file error';
  return [
    `${out.bold(`xlvane ${spec.name}`)} — ${spec.summary}`,
    '',
    'Usage:',
    `  xlvane ${spec.name} ${spec.usage}`,
    '',
    wrap(spec.description, 80, ''),
    '',
    'Options:',
    ...labels.map(([l, d]) => `  ${l.padEnd(width)}${d}`),
    '',
    'Examples:',
    ...spec.examples.flatMap((e) => [`  ${out.cyan(e.command)}`, `      ${e.description}`]),
    '',
    `Exit codes: ${exit}.`,
    `Related: ${spec.related.map((r) => `xlvane ${r}`).join(', ')}`,
    '',
  ].join('\n');
}
