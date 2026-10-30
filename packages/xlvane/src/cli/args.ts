// Argument parsing on util.parseArgs (F-NODE-09). parseArgs runs with `strict: false` and
// `tokens: true`, and the tokens are checked here, so that every mistake gets an XLV_E_USAGE
// error with a "Did you mean" suggestion instead of Node's ERR_PARSE_ARGS_* text (D-603).
import { parseArgs, type ParseArgsOptionConfig } from 'node:util';
import { XlvError } from '../errors';
import { suggest } from '../util/suggest';
import type { CommandSpec, OptionSpec, ParsedArgs } from './types';

/** Options every command accepts. */
export const GLOBAL_OPTIONS: Readonly<Record<string, OptionSpec>> = {
  json: { type: 'boolean', description: 'Print one JSON document to stdout (docs/cli-json.md)' },
  quiet: { type: 'boolean', short: 'q', description: 'Print only results and errors' },
  verbose: { type: 'boolean', description: 'Print more detail, including stack traces of internal errors' },
  help: { type: 'boolean', short: 'h', description: 'Show help for the command' },
};

/** Options of commands that write files. */
export const WRITE_OPTIONS: Readonly<Record<string, OptionSpec>> = {
  force: { type: 'boolean', description: 'Replace the output file if it exists' },
  'dry-run': { type: 'boolean', description: 'Show what would be written, without writing' },
};

/** All options of a command, in help order. */
export function optionsOf(spec: CommandSpec): Record<string, OptionSpec> {
  return { ...spec.options, ...(spec.writes ? WRITE_OPTIONS : {}), ...GLOBAL_OPTIONS };
}

/** A usage error (exit code 2). */
export function usageError(message: string, hint: string, details: Record<string, string | readonly string[]> = {}): XlvError {
  return new XlvError('XLV_E_USAGE', message, { hint, details });
}

/** Parses the arguments after the command name. */
export function parseCommandArgs(spec: CommandSpec, args: readonly string[]): ParsedArgs {
  const options = optionsOf(spec);
  const config: Record<string, ParseArgsOptionConfig> = {};
  for (const [name, o] of Object.entries(options)) {
    config[name] = { type: o.type, ...(o.short ? { short: o.short } : {}), ...(o.multiple ? { multiple: true } : {}) };
  }
  const { tokens } = parseArgs({ args: [...args], options: config, strict: false, allowPositionals: true, tokens: true });
  const help = `Run "xlvane help ${spec.name}" for the options and examples.`;
  const positionals: string[] = [];
  const values: Record<string, string | boolean | string[]> = {};
  for (const token of tokens) {
    if (token.kind === 'positional') {
      positionals.push(token.value);
      continue;
    }
    if (token.kind !== 'option') continue;
    const o = options[token.name];
    if (o === undefined) {
      const known = Object.keys(options).map((n) => `--${n}`);
      const typed = token.rawName.startsWith('--') ? token.rawName : `--${token.name}`;
      const close = token.rawName.startsWith('--') ? suggest(typed, known, 1) : [];
      throw usageError(
        `Unknown option "${token.rawName}" for "xlvane ${spec.name}".`,
        (close.length > 0 ? `Did you mean "${close[0]!}"? ` : '') + help,
        { command: spec.name, option: token.rawName, suggestions: close },
      );
    }
    if (o.type === 'boolean') {
      if (token.inlineValue) {
        throw usageError(`The option "--${token.name}" does not take a value.`, `Write "--${token.name}" alone. ${help}`, { command: spec.name, option: `--${token.name}` });
      }
      values[token.name] = true;
      continue;
    }
    const value = token.value;
    if (value === undefined || (!token.inlineValue && value.startsWith('-') && value.length > 1 && !/^-\d/.test(value))) {
      throw usageError(
        `The option "${token.rawName}" needs a value.`,
        `Write it as "${token.rawName} <${o.value ?? 'value'}>". ${help}`,
        { command: spec.name, option: token.rawName },
      );
    }
    if (o.multiple) values[token.name] = [...((values[token.name] as string[] | undefined) ?? []), value];
    else values[token.name] = value;
  }
  return { positionals, values };
}

/** Checks the number of positional arguments of a command. */
export function checkPositionals(spec: CommandSpec, args: ParsedArgs): void {
  const { min, max, names } = spec.positionals;
  const n = args.positionals.length;
  if (n < min) {
    const missing = names.slice(n, min).map((m) => `<${m}>`).join(' ');
    throw usageError(`"xlvane ${spec.name}" needs ${missing}.`, `Usage: xlvane ${spec.name} ${spec.usage}. Run "xlvane help ${spec.name}" for examples.`, {
      command: spec.name,
      missing: names.slice(n, min),
    });
  }
  if (n > max) {
    const extra = args.positionals.slice(max);
    throw usageError(
      `Too many arguments for "xlvane ${spec.name}": ${extra.map((e) => `"${e}"`).join(' ')}.`,
      `Usage: xlvane ${spec.name} ${spec.usage}. Quote values that contain spaces. Run "xlvane help ${spec.name}" for examples.`,
      { command: spec.name, extra },
    );
  }
}

/** A string option. */
export function str(args: ParsedArgs, name: string): string | undefined {
  const v = args.values[name];
  return typeof v === 'string' ? v : undefined;
}

/** A boolean option. */
export function flag(args: ParsedArgs, name: string): boolean {
  return args.values[name] === true;
}

/** A repeatable string option. */
export function list(args: ParsedArgs, name: string): readonly string[] {
  const v = args.values[name];
  return Array.isArray(v) ? (v as readonly string[]) : typeof v === 'string' ? [v] : [];
}

/** A positive integer option. */
export function positiveInt(args: ParsedArgs, command: string, name: string): number | undefined {
  const v = str(args, name);
  if (v === undefined) return undefined;
  if (!/^\d+$/.test(v) || Number(v) < 1 || !Number.isSafeInteger(Number(v))) {
    throw usageError(`The option "--${name}" must be a whole number of 1 or more, not "${v}".`, `For example --${name} 20.`, { command, option: `--${name}` });
  }
  return Number(v);
}

/** The --header option of read, convert and validate: "auto", "none" or a row number. */
export function headerOption(args: ParsedArgs, command: string): 'auto' | number | false {
  const v = str(args, 'header');
  if (v === undefined || v === 'auto') return 'auto';
  if (v === 'none') return false;
  if (/^\d+$/.test(v) && Number(v) >= 1) return Number(v);
  throw usageError(`The option "--header" must be "auto", "none" or a row number, not "${v}".`, 'For example --header 2 when the column names are in row 2.', {
    command,
    option: '--header',
  });
}
