// The MCP server: options, the server factory and the command line (docs/decisions.md D-701 …).
import { parseArgs } from 'node:util';
import { McpServer } from '@modelcontextprotocol/server';
import { serveStdio } from '@modelcontextprotocol/server/stdio';
import { isXlvError } from 'xlvane';
import { resolveRoots, type Roots } from './paths';
import { registerTools } from './tools';

/** Version reported in server info; kept equal to package.json by tests/mcp/server.test.ts. */
export const SERVER_VERSION = '0.0.0';

/** Default largest file read or written by the tools: 64 MiB. */
export const DEFAULT_MAX_FILE_BYTES: number = 64 * 1024 * 1024;

/** Default largest number of rows returned by one read_table call. */
export const DEFAULT_MAX_ROWS = 1000;

/**
 * Options of {@link createServer}.
 *
 * @example
 * ```ts
 * import type { ServerOptions } from '@xlvane/mcp';
 *
 * const options: ServerOptions = { roots: ['./data'], allowWrite: false };
 * ```
 */
export interface ServerOptions {
  /** Directories the tools may use. Relative paths in tool calls are resolved against the first one. Required, at least one. */
  readonly roots: readonly string[];
  /** Register the tools that write files (write_table, edit_cells, convert). Default false. */
  readonly allowWrite?: boolean;
  /** Largest file read or written, in bytes. Default 64 MiB. */
  readonly maxFileBytes?: number;
  /** Largest number of rows one read_table call returns. Default 1000. */
  readonly maxRows?: number;
  /** Directory that relative roots are resolved against. Default: the current directory. */
  readonly cwd?: string;
}

/** Facts about the configuration, shown to the client as server instructions. */
function instructions(roots: Roots, allowWrite: boolean, maxRows: number): string {
  return [
    `Tools for XLSX workbooks (xlvane). Files are read only inside: ${roots.dirs.join(', ')}. Relative paths are resolved against ${roots.dirs[0]!}.`,
    allowWrite
      ? 'Writing is enabled: write_table, edit_cells and convert create new files and replace an existing file only with "overwrite": true.'
      : 'Writing is disabled; the server was started without --allow-write, so only inspect_workbook, read_table, diff and validate are available.',
    `Files larger than ${String(roots.maxFileBytes)} bytes are refused; read_table returns at most ${String(maxRows)} rows per call.`,
    'Errors include a stable code (XLV_E_*), the location and a hint for fixing the call.',
  ].join(' ');
}

function build(roots: Roots, allowWrite: boolean, maxRows: number): McpServer {
  const server = new McpServer({ name: 'xlvane', version: SERVER_VERSION }, { capabilities: { tools: {} }, instructions: instructions(roots, allowWrite, maxRows) });
  registerTools(server, roots, allowWrite, { maxRows });
  return server;
}

async function prepare(options: ServerOptions): Promise<{ roots: Roots; allowWrite: boolean; maxRows: number }> {
  const maxRows = options.maxRows ?? DEFAULT_MAX_ROWS;
  const roots = await resolveRoots(options.roots, options.cwd ?? '.', options.maxFileBytes ?? DEFAULT_MAX_FILE_BYTES);
  return { roots, allowWrite: options.allowWrite === true, maxRows };
}

/**
 * Creates the xlvane MCP server with its tools. Connect it to any MCP transport; `runCli` serves
 * it over stdio. The roots are resolved (symbolic links followed) once, here.
 *
 * @example
 * ```ts
 * import { createServer } from '@xlvane/mcp';
 * import { StdioServerTransport } from '@modelcontextprotocol/server/stdio';
 *
 * const server = await createServer({ roots: ['./data'] });
 * await server.connect(new StdioServerTransport());
 * ```
 */
export async function createServer(options: ServerOptions): Promise<McpServer> {
  const p = await prepare(options);
  return build(p.roots, p.allowWrite, p.maxRows);
}

/** Where runCli writes text and how it serves; replaced in tests. */
export interface CliIO {
  readonly stdout: (text: string) => void;
  readonly stderr: (text: string) => void;
  readonly cwd: string;
  /** Serves the server until the client disconnects. */
  readonly serve: (factory: () => McpServer) => Promise<void>;
}

const HELP = `xlvane-mcp — MCP server for XLSX files (xlvane)

Usage:
  xlvane-mcp --root <dir> [--root <dir>...] [options]
  xlvane mcp --root <dir> [options]

Options:
  --root <dir>            Directory the tools may use (repeatable, at least one)
  --allow-write           Also offer write_table, edit_cells and convert
  --max-file-size <bytes> Largest file read or written (default ${String(DEFAULT_MAX_FILE_BYTES)})
  --max-rows <n>          Largest number of rows per read_table call (default ${String(DEFAULT_MAX_ROWS)})
  -h, --help              Show this help

The server speaks MCP over stdin and stdout; messages for people go to stderr.
Example client configuration: {"command": "npx", "args": ["@xlvane/mcp", "--root", "/path/to/data"]}
`;

function positive(value: string | undefined, name: string): number | undefined {
  if (value === undefined) return undefined;
  if (!/^\d+$/.test(value) || Number(value) < 1) throw new Error(`--${name} must be a whole number of 1 or more, not "${value}".`);
  return Number(value);
}

/** runCli with its environment passed in (tests). */
export async function runWith(args: readonly string[], io: CliIO): Promise<number> {
  let values: Record<string, unknown>;
  try {
    values = parseArgs({
      args: [...args],
      options: { root: { type: 'string', multiple: true }, 'allow-write': { type: 'boolean' }, 'max-file-size': { type: 'string' }, 'max-rows': { type: 'string' }, help: { type: 'boolean', short: 'h' } },
      strict: true,
      allowPositionals: false,
      tokens: true,
    }).values;
  } catch (error) {
    io.stderr(`xlvane-mcp: ${error instanceof Error ? error.message : String(error)}\nRun "xlvane-mcp --help" for the options.\n`);
    return 2;
  }
  if (values['help'] === true) {
    io.stdout(HELP);
    return 0;
  }
  let prepared;
  try {
    const roots = (values['root'] as string[] | undefined) ?? [];
    if (roots.length === 0) {
      io.stderr('xlvane-mcp: no directory was given.\nStart the server with the directories it may use, for example: xlvane-mcp --root ./data\n');
      return 2;
    }
    const maxFileBytes = positive(values['max-file-size'] as string | undefined, 'max-file-size');
    const maxRows = positive(values['max-rows'] as string | undefined, 'max-rows');
    prepared = await prepare({
      roots,
      cwd: io.cwd,
      allowWrite: values['allow-write'] === true,
      ...(maxFileBytes !== undefined ? { maxFileBytes } : {}),
      ...(maxRows !== undefined ? { maxRows } : {}),
    });
  } catch (error) {
    io.stderr(`xlvane-mcp: ${error instanceof Error ? error.message : String(error)}\n${isXlvError(error) ? `${error.hint}\n` : ''}`);
    return isXlvError(error) ? 3 : 2;
  }
  const { roots, allowWrite, maxRows } = prepared;
  io.stderr(`xlvane-mcp ${SERVER_VERSION}: serving ${roots.dirs.join(', ')} over stdio${allowWrite ? ' (writing enabled)' : ' (read-only)'}.\n`);
  await io.serve(() => build(roots, allowWrite, maxRows));
  return 0;
}

/**
 * Runs the server from the command line: parses `--root`, `--allow-write`, `--max-file-size`
 * and `--max-rows`, then serves MCP over stdio until the client closes stdin. Returns the exit
 * code. `xlvane mcp` calls this function.
 *
 * @example
 * ```ts
 * import { runCli } from '@xlvane/mcp';
 *
 * const code = await runCli(['--root', './data', '--allow-write']);
 * ```
 */
export async function runCli(args: readonly string[]): Promise<number> {
  const { default: process } = await import('node:process');
  return runWith(args, {
    stdout: (text) => {
      process.stdout.write(text);
    },
    stderr: (text) => {
      process.stderr.write(text);
    },
    cwd: process.cwd(),
    serve: (factory) =>
      new Promise<void>((resolve) => {
        serveStdio(factory);
        process.stdin.once('end', () => resolve());
        process.stdin.once('close', () => resolve());
      }),
  });
}
