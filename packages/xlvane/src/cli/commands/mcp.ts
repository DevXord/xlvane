import { XlvError } from '../../errors';
import { usageError } from '../args';
import type { CommandSpec } from '../types';

/** The package that holds the MCP server; xlvane does not depend on it (docs/PLAN.md section 3). */
export const MCP_PACKAGE = '@xlvane/mcp';

/** What `xlvane mcp` expects from @xlvane/mcp (D-607). */
interface McpModule {
  readonly runCli?: (args: readonly string[]) => Promise<number>;
}

/**
 * Whether loading failed because @xlvane/mcp itself is missing. A missing dependency of an
 * installed @xlvane/mcp has the same error code but names that dependency (D-I04).
 */
function isModuleNotFound(error: unknown): boolean {
  const code = typeof error === 'object' && error !== null && 'code' in error ? (error as { code: unknown }).code : undefined;
  const message = error instanceof Error ? error.message : '';
  return (code === 'ERR_MODULE_NOT_FOUND' || code === 'MODULE_NOT_FOUND') && (message.includes(`'${MCP_PACKAGE}'`) || message.includes(`"${MCP_PACKAGE}"`));
}

/** Options of the xlvane command line itself, which the server does not take. */
const OWN_OPTIONS = new Set(['--quiet', '-q', '--verbose']);

export const mcpCommand: CommandSpec = {
  name: 'mcp',
  summary: 'Start the MCP server for AI agents',
  description:
    'Starts the Model Context Protocol server from the separate package @xlvane/mcp over stdin and stdout, so that an AI agent can inspect, read, convert and compare spreadsheets through tools instead of code. The arguments after "mcp" are passed to the server: --root <dir> (required, repeatable) names a directory the tools may use, --allow-write adds the tools that write files, and --max-file-size <bytes>, --max-rows <n> and --max-response-chars <n> set its limits. --json is not available, because stdout carries the protocol. If @xlvane/mcp is not installed, the command prints how to install it. The exit code is the server\'s: 0 when the client closed the connection, 2 for wrong arguments, 3 for a directory that cannot be used.',
  usage: '[server options]',
  positionals: { min: 0, max: Number.POSITIVE_INFINITY, names: [] },
  passThrough: true,
  options: {},
  examples: [
    { command: 'xlvane mcp --root ./data', description: 'Serve the files in ./data, read-only' },
    { command: 'xlvane mcp --root ./data --allow-write', description: 'Also allow writing files there' },
  ],
  related: ['inspect', 'read'],
  async run(ctx) {
    if (ctx.json) {
      throw usageError('"xlvane mcp" does not take --json: the server uses stdout for the protocol.', 'Run "xlvane mcp --root <dir>" without --json; the server reports its own errors as MCP results.', { command: 'mcp', option: '--json' });
    }
    let mod: McpModule;
    try {
      mod = (await ctx.io.importModule(MCP_PACKAGE)) as McpModule;
    } catch (error) {
      if (!isModuleNotFound(error)) throw error;
      throw new XlvError('XLV_E_NOT_INSTALLED', `The MCP server is in the separate package ${MCP_PACKAGE}, which is not installed.`, {
        hint: `Install it next to xlvane with "npm install ${MCP_PACKAGE}", then run "xlvane mcp" again.`,
        details: { package: MCP_PACKAGE },
        cause: error,
      });
    }
    if (typeof mod.runCli !== 'function') {
      throw new XlvError('XLV_E_NOT_SUPPORTED', `The installed ${MCP_PACKAGE} does not contain the MCP server yet.`, {
        hint: `Update ${MCP_PACKAGE} to a version that provides the server.`,
        details: { package: MCP_PACKAGE, feature: 'runCli' },
      });
    }
    const code = await mod.runCli(ctx.argv.filter((a) => !OWN_OPTIONS.has(a)));
    // The server has already printed what went wrong; its exit code is kept (D-I04).
    return { result: { exitCode: code }, exitCode: code === 0 || code === 1 || code === 2 ? code : 3, render: () => '' };
  },
};
