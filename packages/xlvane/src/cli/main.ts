// The `xlvane` command line: dispatch, `--json` envelope, errors in three parts and exit codes
// (docs/PLAN.md section 5; docs/cli-json.md; docs/decisions.md D-601 … D-605).
import { XlvError, type XlvErrorJSON } from '../errors';
import { suggest } from '../util/suggest';
import { checkPositionals, parseCommandArgs, usageError } from './args';
import { COMMANDS } from './commands/index';
import { withPath } from './files';
import { commandHelp, welcome } from './help';
import { commandLine, Output, useColor } from './output';
import type { CliIO, CommandContext, CommandSpec } from './types';

/** Version of the `--json` envelope (docs/cli-json.md). */
export const JSON_VERSION = 1;

/** Exit codes (docs/cli-json.md). */
export const EXIT = { ok: 0, found: 1, usage: 2, failure: 3 } as const;

/** The `--json` document for an error. */
export interface ErrorEnvelope {
  readonly xlvane: 1;
  readonly command: string | null;
  readonly ok: false;
  readonly error: XlvErrorJSON;
  readonly try: string | null;
}

/** Arguments before "--". */
function optionArgs(argv: readonly string[]): readonly string[] {
  const end = argv.indexOf('--');
  return end === -1 ? argv : argv.slice(0, end);
}

/**
 * Runs the CLI with the arguments after `xlvane` and returns the exit code. Nothing here touches
 * `process`; the entry point (bin.ts) passes the real streams and environment.
 */
export async function run(argv: readonly string[], io: CliIO): Promise<number> {
  const out = new Output(useColor(io.env, io.stdoutIsTTY), io.stdoutIsTTY ? (io.columns ?? 100) : 100);
  const err = new Output(useColor(io.env, io.stderrIsTTY), 100);
  const opts = optionArgs(argv);
  const json = opts.includes('--json');
  const [first, ...rest] = argv;
  let spec: CommandSpec | undefined;
  const state: CommandContext['state'] = { lastFile: null };
  try {
    if (first === undefined || (first === '--json' && rest.length === 0)) {
      io.stdout(welcome(out, io.version, COMMANDS));
      return EXIT.ok;
    }
    if (first === '--version' || first === '-v') {
      io.stdout(`${io.version}\n`);
      return EXIT.ok;
    }
    if (first === 'help' || first === '--help' || first === '-h') {
      const topic = rest.find((a) => !a.startsWith('-'));
      if (topic === undefined) {
        io.stdout(welcome(out, io.version, COMMANDS));
        return EXIT.ok;
      }
      io.stdout(commandHelp(out, findCommand(topic)));
      return EXIT.ok;
    }
    if (first.startsWith('-')) {
      throw usageError(`"${first}" is not a command.`, 'Put the command first, for example "xlvane inspect report.xlsx --json". Run "xlvane" to see the commands.', {
        option: first,
      });
    }
    spec = findCommand(first);
    if (optionArgs(rest).some((a) => a === '--help' || a === '-h')) {
      io.stdout(commandHelp(out, spec));
      return EXIT.ok;
    }
    const args = parseCommandArgs(spec, rest);
    checkPositionals(spec, args);
    const quiet = args.values['quiet'] === true;
    const ctx: CommandContext = {
      io,
      args,
      argv: rest,
      json,
      quiet,
      verbose: args.values['verbose'] === true,
      state,
      note: (text) => {
        if (!json && !quiet) io.stderr(`${err.dim(text)}\n`);
      },
    };
    const outcome = await spec.run(ctx);
    if (json) {
      io.stdout(`${JSON.stringify({ xlvane: JSON_VERSION, command: spec.name, ok: true, result: outcome.result }, null, 2)}\n`);
    } else {
      const text = outcome.render(out);
      if (text !== '') io.stdout(text.endsWith('\n') ? text : `${text}\n`);
    }
    return outcome.exitCode;
  } catch (thrown) {
    let error = asXlvError(thrown);
    // Errors from reading a file name it, unless they come from the command line itself.
    if (state.lastFile !== null && error.code !== 'XLV_E_USAGE' && error.code !== 'XLV_E_INTERNAL') error = withPath(error, state.lastFile) as XlvError;
    const tryCommand = suggestCommand(error, first, rest, spec);
    if (json) {
      const envelope: ErrorEnvelope = { xlvane: JSON_VERSION, command: spec?.name ?? null, ok: false, error: error.toJSON(), try: tryCommand };
      io.stdout(`${JSON.stringify(envelope, null, 2)}\n`);
    } else {
      io.stderr(formatError(err, error, tryCommand));
      if (error.code === 'XLV_E_INTERNAL' && optionArgs(argv).includes('--verbose') && thrown instanceof Error && thrown.stack) {
        io.stderr(`${thrown.stack}\n`);
      }
    }
    return error.code === 'XLV_E_USAGE' ? EXIT.usage : EXIT.failure;
  }
}

function findCommand(name: string): CommandSpec {
  const spec = COMMANDS.find((c) => c.name === name);
  if (spec) return spec;
  const close = suggest(name, COMMANDS.map((c) => c.name), 1);
  throw usageError(
    `Unknown command "${name}".`,
    (close.length > 0 ? `Did you mean "xlvane ${close[0]!}"? ` : '') + `Commands: ${COMMANDS.map((c) => c.name).join(', ')}. Run "xlvane" for examples.`,
    { command: name, suggestions: close },
  );
}

function asXlvError(thrown: unknown): XlvError {
  if (thrown instanceof XlvError) return thrown;
  const name = thrown instanceof Error ? thrown.name : typeof thrown;
  const message = thrown instanceof Error ? thrown.message : String(thrown);
  return new XlvError('XLV_E_INTERNAL', `Unexpected error: ${message}`, {
    hint: 'This is a bug in xlvane. Please report it with the command you ran; --verbose prints the stack trace.',
    details: { name },
    cause: thrown,
  });
}

/** Replaces the value of an option (--name value or --name=value) or appends it. */
function withOption(args: readonly string[], name: string, value: string): string[] {
  const out = [...args];
  for (let i = 0; i < out.length; i++) {
    if (out[i] === `--${name}` && i + 1 < out.length) {
      out[i + 1] = value;
      return out;
    }
    if (out[i]!.startsWith(`--${name}=`)) {
      out[i] = `--${name}=${value}`;
      return out;
    }
  }
  return [...out, `--${name}`, value];
}

/** A ready-to-run command that fixes the error, when there is an obvious one. */
function suggestCommand(error: XlvError, first: string | undefined, rest: readonly string[], spec: CommandSpec | undefined): string | null {
  const d = error.details;
  const suggestions = Array.isArray(d['suggestions']) ? (d['suggestions'] as readonly string[]) : [];
  if (error.code === 'XLV_E_USAGE' && suggestions.length > 0) {
    if (spec === undefined && first === 'help') return commandLine(['help', suggestions[0]!]);
    if (spec === undefined && first !== undefined) return commandLine([suggestions[0]!, ...rest]);
    if (spec !== undefined && typeof d['option'] === 'string') {
      const option = d['option'];
      return commandLine([spec.name, ...rest.map((a) => (a === option ? suggestions[0]! : a.startsWith(`${option}=`) ? suggestions[0]! + a.slice(option.length) : a))]);
    }
  }
  if (spec === undefined) return null;
  if (error.code === 'XLV_E_FILE_EXISTS') return commandLine([spec.name, ...rest, '--force']);
  if (error.code === 'XLV_E_SHEET_NOT_FOUND' && Array.isArray(d['available']) && typeof error.location.sheet === 'string') {
    const close = suggest(error.location.sheet, d['available'] as readonly string[], 1);
    if (close.length > 0) return commandLine([spec.name, ...withOption(rest, 'sheet', close[0]!)]);
  }
  if (error.code === 'XLV_E_FILE_READ' && suggestions.length > 0 && typeof d['path'] === 'string') {
    const path = d['path'];
    if (rest.includes(path)) return commandLine([spec.name, ...rest.map((a) => (a === path ? suggestions[0]! : a))]);
  }
  return null;
}

/** Where the problem is, as one line. */
function where(error: XlvError): string | null {
  const parts: string[] = [];
  const d = error.details;
  if (typeof d['path'] === 'string') parts.push(d['path']);
  const l = error.location;
  if (l.part) parts.push(`part ${l.part}`);
  if (l.sheet) parts.push(`sheet "${l.sheet}"`);
  if (l.cell) parts.push(`cell ${l.cell}`);
  if (l.line !== undefined) parts.push(`line ${String(l.line)}${l.column !== undefined ? `, column ${String(l.column)}` : ''}`);
  if (parts.length === 0 && error.code === 'XLV_E_USAGE') {
    if (typeof d['option'] === 'string') parts.push(`option ${d['option']}`);
    else if (typeof d['command'] === 'string') parts.push(`command line`);
  }
  return parts.length > 0 ? parts.join(', ') : null;
}

/** An error in three parts: what happened, where, what to do next. */
export function formatError(err: Output, error: XlvError, tryCommand: string | null): string {
  const place = where(error);
  return [
    `${err.red(err.bold('Error:'))} ${error.message}`,
    ...(place ? [`Where: ${place}`] : []),
    `Fix:   ${error.hint}`,
    ...(tryCommand ? [`Try:   ${err.cyan(tryCommand)}`] : []),
    err.dim(`Code:  ${error.code} (${error.docsUrl})`),
    '',
  ].join('\n');
}
