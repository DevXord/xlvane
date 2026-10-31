// Types shared by the `xlvane` command line (docs/decisions.md D-601).
import type { Output } from './output';

/** Everything the CLI needs from its environment; tests pass their own (D-601). */
export interface CliIO {
  readonly stdout: (text: string) => void;
  readonly stderr: (text: string) => void;
  readonly env: Readonly<Record<string, string | undefined>>;
  /** Directory that relative paths are resolved against. */
  readonly cwd: string;
  readonly stdoutIsTTY: boolean;
  readonly stderrIsTTY: boolean;
  /** Terminal width, when stdout is a terminal. */
  readonly columns: number | undefined;
  /** Version of the xlvane package. */
  readonly version: string;
}

/** A command-line option. */
export interface OptionSpec {
  readonly type: 'string' | 'boolean';
  readonly short?: string;
  readonly multiple?: boolean;
  /** Placeholder for the value in help, e.g. "name". */
  readonly value?: string;
  readonly description: string;
}

/** An example in help: a full command line and what it does. */
export interface Example {
  readonly command: string;
  readonly description: string;
}

/** Parsed arguments of one command. */
export interface ParsedArgs {
  readonly positionals: readonly string[];
  readonly values: Readonly<Record<string, string | boolean | readonly string[] | undefined>>;
}

/** What a command returns: the `--json` result, the text for people, and the exit code. */
export interface CommandOutcome {
  readonly result: unknown;
  readonly render: (out: Output) => string;
  /** 0 success, 1 a check found something (diff, validate). */
  readonly exitCode: 0 | 1;
}

/** What a command gets when it runs. */
export interface CommandContext {
  readonly io: CliIO;
  readonly args: ParsedArgs;
  /** The arguments as typed, after the command name (to build "Try:" commands). */
  readonly argv: readonly string[];
  readonly json: boolean;
  readonly quiet: boolean;
  readonly verbose: boolean;
  /** The input file read last, so that errors without a path can name it. */
  readonly state: { lastFile: string | null };
  /** Text output to stderr (progress, notes); nothing with --json or --quiet. */
  readonly note: (text: string) => void;
}

/** A command of the CLI: its help and its handler. */
export interface CommandSpec {
  readonly name: string;
  /** One line, used in the welcome screen and as the help title. */
  readonly summary: string;
  /** A short paragraph for `help <command>`. */
  readonly description: string;
  /** Usage after `xlvane <name>`, e.g. "<file> [options]". */
  readonly usage: string;
  readonly positionals: { readonly min: number; readonly max: number; readonly names: readonly string[] };
  readonly options: Readonly<Record<string, OptionSpec>>;
  /** The command writes files: adds --force and --dry-run. */
  readonly writes?: boolean;
  /** Exit code 1 means "a check found something" for this command. */
  readonly checks?: string;
  readonly examples: readonly Example[];
  readonly related: readonly string[];
  readonly run: (ctx: CommandContext) => Promise<CommandOutcome>;
}
