// Welcome screen and `help <command>` (docs/PLAN.md section 5).
import { GLOBAL_OPTIONS, optionsOf } from './args';
import { wrap, type Output } from './output';
import type { CommandSpec } from './types';

/** Repository of the project; the docs live there until a website exists. */
export const DOCS_URL = 'https://github.com/DevXord/xlvane';

/** Where the documentation is after `npm install xlvane`, for references in help texts. */
export const DOCS_IN_PACKAGE = 'node_modules/xlvane/docs';

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
    `${out.bold(`xlvane ${version}`)} - read, write and inspect Excel files`,
    '',
    out.bold('Common tasks:'),
    ...COMMON.map(([c, d]) => `  ${out.cyan(c.padEnd(width))}${d}`),
    '',
    `${out.bold('Commands:')} ${commands.map((c) => c.name).join(', ')}`,
    `${out.bold('More:')}  xlvane help <command>   ·   Docs: ${DOCS_URL}`,
    `For AI agents and scripts: add --json to any command (${DOCS_IN_PACKAGE}/cli-json.md)`,
    '',
  ].join('\n');
}

/** Options that act on the command itself rather than on its data (help groups them as "Actions"). */
const ACTIONS: Readonly<Record<string, readonly string[]>> = {
  edit: ['set', 'append', 'insert-rows', 'delete-rows'],
};

/** Option descriptions that differ from the shared text for one command. */
const OPTION_TEXT: Readonly<Record<string, Readonly<Record<string, string>>>> = {
  migrate: { force: 'Change package.json (without it, only the report is printed)', 'dry-run': 'Show the change to package.json without writing it' },
};

/** The help of one command. */
export function commandHelp(out: Output, spec: CommandSpec): string {
  const width = Math.max(60, Math.min(out.width, 100));
  const options = optionsOf(spec);
  const label = (name: string): string => {
    const o = options[name]!;
    const long = `--${name}${o.type === 'string' ? ` <${o.value ?? 'value'}>` : ''}`;
    return o.short ? `-${o.short}, ${long}` : `    ${long}`;
  };
  const names = Object.keys(options);
  const actions = ACTIONS[spec.name] ?? [];
  const globals = Object.keys(GLOBAL_OPTIONS);
  const column = Math.min(Math.max(...names.map((n) => label(n).length)) + 4, 40);
  const describe = (name: string): string => {
    const o = options[name]!;
    const text = (OPTION_TEXT[spec.name]?.[name] ?? o.description) + (o.multiple ? ' (repeatable)' : '');
    const head = `  ${label(name)}`;
    return head.length + 2 > column ? `${head}\n${wrap(text, width, ' '.repeat(column))}` : wrap(text, width, ' '.repeat(column), head.padEnd(column));
  };
  const block = (title: string, list: readonly string[]): string[] => (list.length === 0 ? [] : [out.bold(title), ...list.map(describe), '']);
  const exit = spec.checks ? `0 success, 1 ${spec.checks}, 2 usage error, 3 input or file error` : '0 success, 2 usage error, 3 input or file error';
  const related = spec.related.filter((r) => r !== 'help');
  return [
    `${out.bold(`xlvane ${spec.name}`)} - ${spec.summary}`,
    '',
    out.bold('Usage:'),
    wrap(`xlvane ${spec.name} ${spec.usage}`, width, '      ', '  '),
    '',
    wrap(spec.description, width),
    '',
    ...block('Actions:', actions),
    ...block('Options:', names.filter((n) => !actions.includes(n) && !globals.includes(n))),
    ...block('Global options:', globals),
    out.bold('Examples:'),
    ...spec.examples.flatMap((e) => [`  ${out.cyan(e.command)}`, `      ${e.description}`]),
    '',
    `${out.bold('Exit codes:')} ${exit}.`,
    ...(related.length > 0 ? [`${out.bold('Related:')} ${related.map((r) => `xlvane ${r}`).join(', ')}`] : []),
    '',
  ].join('\n');
}
