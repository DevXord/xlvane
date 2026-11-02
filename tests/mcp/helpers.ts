// Connects an in-memory MCP client (official SDK) to the xlvane server.
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/client';
import { InMemoryTransport } from '@modelcontextprotocol/server';
import { createServer, type ServerOptions } from '../../packages/mcp/src/server';
import { writeFixtures } from '../cli/helpers';

export interface Connected {
  readonly client: Client;
  /** Calls a tool and returns its structured result and whether it is an error. */
  call(name: string, args: Record<string, unknown>): Promise<{ isError: boolean; result: Record<string, unknown>; text: string }>;
  close(): Promise<void>;
}

export async function connect(options: ServerOptions): Promise<Connected> {
  const server = await createServer(options);
  const [serverSide, clientSide] = InMemoryTransport.createLinkedPair();
  await server.connect(serverSide);
  const client = new Client({ name: 'xlvane-tests', version: '0' });
  await client.connect(clientSide);
  return {
    client,
    async call(name, args) {
      const r = await client.callTool({ name, arguments: args });
      const text = (r.content as { type: string; text?: string }[]).map((c) => c.text ?? '').join('');
      return { isError: r.isError === true, result: (r.structuredContent ?? {}) as Record<string, unknown>, text };
    },
    async close() {
      await client.close();
      await server.close();
    },
  };
}

/** A temporary directory with the CLI fixtures (report.xlsx, old.xlsx, new.xlsx, data.json, ...). */
export async function withData<T>(fn: (dir: string) => Promise<T>): Promise<T> {
  const dir = await mkdtemp(join(tmpdir(), 'xlvane-mcp-'));
  try {
    await writeFixtures(dir);
    return await fn(dir);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}
