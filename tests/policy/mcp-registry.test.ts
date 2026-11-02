// Metadata for the official MCP Registry (docs/decisions.md D-I06, F-MCP-03): packages/mcp/server.json
// describes @xlvane/mcp, and package.json carries "mcpName", which the registry compares with the
// server name to verify ownership of the npm package. Limits from the schema
// https://static.modelcontextprotocol.io/schemas/2025-12-11/server.schema.json (read 2026-11-02).
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const read = (path: string): Record<string, unknown> => JSON.parse(readFileSync(path, 'utf8')) as Record<string, unknown>;

describe('MCP Registry metadata', () => {
  const pkg = read('packages/mcp/package.json');
  const server = read('packages/mcp/server.json');
  const packages = server['packages'] as Record<string, unknown>[];

  it('server.json names the package and its version, and package.json has the same mcpName', () => {
    expect(server['$schema']).toBe('https://static.modelcontextprotocol.io/schemas/2025-12-11/server.schema.json');
    // The registry compares names with case: GitHub login grants io.github.DevXord/* (D-I07).
    expect(server['name']).toBe('io.github.DevXord/xlvane');
    expect(pkg['mcpName']).toBe(server['name']);
    expect(server['version']).toBe(pkg['version']);
    expect(packages).toHaveLength(1);
    expect(packages[0]).toMatchObject({ registryType: 'npm', identifier: '@xlvane/mcp', version: pkg['version'], transport: { type: 'stdio' } });
  });

  it('stays within the schema limits', () => {
    expect(String(server['name'])).toMatch(/^[a-zA-Z0-9.-]+\/[a-zA-Z0-9._-]+$/);
    for (const field of ['description', 'title']) {
      expect(String(server[field]).length, field).toBeGreaterThan(0);
      expect(String(server[field]).length, field).toBeLessThanOrEqual(100);
    }
    const args = packages[0]!['packageArguments'] as Record<string, unknown>[];
    expect(args).toEqual([expect.objectContaining({ type: 'named', name: '--root', isRequired: true, isRepeated: true, format: 'filepath' })]);
  });
});
