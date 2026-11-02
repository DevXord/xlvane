import { XlvError } from '../../errors';
import type { CommandSpec } from '../types';

/** The package that holds the MCP server; xlvane does not depend on it (docs/PLAN.md section 3). */
export const MCP_PACKAGE = '@xlvane/mcp';

/** What `xlvane mcp` expects from @xlvane/mcp (D-607). */
interface McpModule {
  readonly runCli?: (args: readonly string[]) => Promise<number>;
}

function isModuleNotFound(error: unknown): boolean {
  const code = typeof error === 'object' && error !== null && 'code' in error ? (error as { code: unknown }).code : undefined;
  return code === 'ERR_MODULE_NOT_FOUND' || code === 'MODULE_NOT_FOUND';
}

export const mcpCommand: CommandSpec = {
  name: 'mcp',
  summary: 'Start the MCP server for AI agents',
  description:
    'Starts the Model Context Protocol server from the separate package @xlvane/mcp, so that an AI agent can inspect, read, convert and compare spreadsheets through tools instead of code. The arguments after "mcp" are passed to the server. If @xlvane/mcp is not installed, the command prints how to install it.',
  usage: '[server options]',
  positionals: { min: 0, max: Number.POSITIVE_INFINITY, names: [] },
  passThrough: true,
  options: {},
  examples: [
    { command: 'xlvane mcp', description: 'Start the server, or show how to install it' },
    { command: 'xlvane help mcp', description: 'This help' },
  ],
  related: ['inspect', 'read'],
  async run(ctx) {
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
    const code = await mod.runCli(ctx.argv);
    return { result: { exitCode: code }, exitCode: code === 0 ? 0 : 1, render: () => '' };
  },
};
