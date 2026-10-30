// The built server over stdio, as an MCP client starts it: `xlvane-mcp --root <dir>` and
// `xlvane mcp --root <dir>`. Both packages are built into their own directory (no shared dist).
import { mkdir, mkdtemp, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/client';
import { StdioClientTransport } from '@modelcontextprotocol/client/stdio';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { buildPackage, REPO } from '../helpers/build-package';
import { withData } from './helpers';

let modules = '';

beforeAll(async () => {
  await mkdir(join(REPO, 'node_modules/.cache'), { recursive: true });
  // Under the repository, so that fflate and the MCP SDK resolve from its node_modules.
  modules = join(await mkdtemp(join(REPO, 'node_modules/.cache/xlvane-mcp-')), 'node_modules');
  await buildPackage('xlvane', modules);
  await buildPackage('@xlvane/mcp', modules);
}, 300_000);

afterAll(async () => {
  if (modules) await rm(join(modules, '..'), { recursive: true, force: true });
});

async function session(args: string[], cwd: string): Promise<{ tools: string[]; sheets: unknown; close: () => Promise<void> }> {
  const transport = new StdioClientTransport({ command: process.execPath, args, cwd, stderr: 'pipe' });
  const client = new Client({ name: 'xlvane-stdio-test', version: '0' });
  await client.connect(transport);
  const tools = (await client.listTools()).tools.map((t) => t.name);
  const r = await client.callTool({ name: 'inspect_workbook', arguments: { path: 'report.xlsx' } });
  return { tools, sheets: (r.structuredContent as { sheetCount?: number } | undefined)?.sheetCount, close: () => client.close() };
}

describe('stdio', () => {
  it('xlvane-mcp serves the tools of the given directory', { timeout: 60_000 }, async () => {
    await withData(async (dir) => {
      const s = await session([join(modules, '@xlvane/mcp/dist/cli.js'), '--root', '.'], dir);
      expect(s.tools).toEqual(['inspect_workbook', 'read_table', 'diff', 'validate']);
      expect(s.sheets).toBe(2);
      await s.close();
    });
  });

  it('xlvane mcp starts the same server', { timeout: 60_000 }, async () => {
    await withData(async (dir) => {
      const s = await session([join(modules, 'xlvane/dist/cli.js'), 'mcp', '--root', dir, '--allow-write'], dir);
      expect(s.tools).toHaveLength(7);
      expect(s.sheets).toBe(2);
      await s.close();
    });
  });
});
