// Acceptance criterion: `--json` output matches the documented schema (docs/cli-json.md,
// docs/cli-json/*.schema.json), for successes and errors of every command.
import { readdirSync, readFileSync } from 'node:fs';
import { writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { COMMANDS } from '../../packages/xlvane/src/cli/commands/index';
import { SchemaSet } from '../helpers/json-schema';
import { cli, withFixtures } from './helpers';

const ROOT = fileURLToPath(new URL('../..', import.meta.url));
const DIR = join(ROOT, 'docs/cli-json');
const schemas = new SchemaSet(DIR);

/** Runs with --json and returns the parsed document; stdout must hold exactly one JSON value. */
async function json(argv: readonly string[], dir: string, importModule?: (s: string) => Promise<unknown>): Promise<{ code: number; doc: Record<string, unknown>; stderr: string }> {
  const r = await cli([...argv, '--json'], { cwd: dir, env: { FORCE_COLOR: '1' }, tty: true, ...(importModule ? { importModule } : {}) });
  return { code: r.code, doc: JSON.parse(r.stdout) as Record<string, unknown>, stderr: r.stderr };
}

const SUCCESS: readonly (readonly [string, readonly string[], number])[] = [
  ['inspect', ['inspect', 'report.xlsx'], 0],
  ['inspect', ['inspect', 'report.xlsx', '--sheet', 'Costs'], 0],
  ['read', ['read', 'report.xlsx'], 0],
  ['read', ['read', 'report.xlsx', '--header', 'none', '--range', 'A1:C2', '--limit', '1'], 0],
  ['convert', ['convert', 'report.xlsx', 'out.json'], 0],
  ['convert', ['convert', 'sales.csv', 'out.xlsx', '--dry-run'], 0],
  ['create', ['create', 'data.json', 'out.xlsx', '--columns', 'name,score'], 0],
  ['edit', ['edit', 'report.xlsx', '--set', 'Sales!B2=5', '--append', 'rows.json', '--sheet', 'Sales', '-o', 'x.xlsx'], 0],
  ['diff', ['diff', 'old.xlsx', 'new.xlsx', '--styles'], 1],
  ['diff', ['diff', 'old.xlsx', 'old.xlsx'], 0],
  ['validate', ['validate', 'report.xlsx', '--sheet', 'Sales', '--columns', 'Region:string,Revenue:number,Date:date,Missing:any'], 1],
  ['validate', ['validate', 'report.xlsx'], 0],
  ['migrate', ['migrate', 'my-app'], 0],
];

describe('--json', () => {
  it('has a schema for every command, listed in docs/cli-json.md', () => {
    const files = readdirSync(DIR).filter((f) => f.endsWith('.schema.json')).sort();
    expect(files).toEqual(['common.schema.json', ...COMMANDS.map((c) => `${c.name}.schema.json`)].sort());
    const doc = readFileSync(join(ROOT, 'docs/cli-json.md'), 'utf8');
    for (const f of files) expect(doc, f).toContain(`cli-json/${f}`);
  });

  it('success documents match the schemas', { timeout: 60_000 }, async () => {
    await withFixtures(async (dir) => {
      for (const [command, argv, code] of SUCCESS) {
        const r = await json(argv, dir);
        expect(r.code, argv.join(' ')).toBe(code);
        expect(r.doc['ok'], argv.join(' ')).toBe(true);
        expect(schemas.validate(`${command}.schema.json`, r.doc), argv.join(' ')).toEqual([]);
        // No colour codes and no progress output with --json, even on a terminal with FORCE_COLOR.
        expect(JSON.stringify(r.doc)).not.toMatch(/\\u001b\[/);
        expect(r.stderr).toBe('');
      }
    });
  });

  it('error documents match the schemas, with the same exit codes', async () => {
    await withFixtures(async (dir) => {
      await writeFile(join(dir, 'bad.json'), '{');
      const cases: (readonly [string, readonly string[], number])[] = [
        ['inspect', ['inspect', 'missing.xlsx'], 3],
        ['read', ['read', 'report.xlsx', '--sheet', 'Salse'], 3],
        ['read', ['read', 'report.xlsx', '--shet', 'x'], 2],
        ['convert', ['convert', 'report.xlsx', 'sales.csv'], 3],
        ['create', ['create', 'bad.json', 'x.xlsx'], 3],
        ['edit', ['edit', 'report.xlsx', '--set', 'Sales!B2=1'], 3],
        ['diff', ['diff', 'old.xlsx'], 2],
        ['validate', ['validate', 'data.json'], 3],
        ['migrate', ['migrate', 'nowhere'], 3],
        ['mcp', ['mcp'], 3],
      ];
      for (const [command, argv, code] of cases) {
        const r = await json(argv, dir);
        expect(r.code, argv.join(' ')).toBe(code);
        expect(r.doc['ok']).toBe(false);
        expect(r.doc['command']).toBe(command);
        expect(schemas.validate(`${command}.schema.json`, r.doc), argv.join(' ')).toEqual([]);
        expect(r.stderr).toBe('');
      }
      const unknown = await json(['frobnicate'], dir);
      expect(unknown.code).toBe(2);
      expect(unknown.doc['command']).toBeNull();
      expect(schemas.validate('common.schema.json', unknown.doc)).toEqual([]);
      expect(schemas.validate('mcp.schema.json', unknown.doc)).toEqual([]);
    });
  });

  it('the suggested fix is part of the error document', async () => {
    await withFixtures(async (dir) => {
      const r = await json(['read', 'report.xlsx', '--sheet', 'Salse'], dir);
      expect(r.doc['try']).toBe('xlvane read report.xlsx --sheet Sales --json');
      expect((r.doc['error'] as { details: { available: string[] } }).details.available).toEqual(['Sales', 'Costs']);
    });
  });
});

describe('the schema validator', () => {
  const set = new SchemaSet(DIR);
  it('rejects documents that break the schema', () => {
    const ok = { xlvane: 1, command: 'mcp', ok: true, result: { exitCode: 0 } };
    expect(set.validate('mcp.schema.json', ok)).toEqual([]);
    expect(set.validate('mcp.schema.json', { ...ok, extra: 1 })).not.toEqual([]);
    expect(set.validate('mcp.schema.json', { ...ok, result: { exitCode: 'x' } })).not.toEqual([]);
    expect(set.validate('mcp.schema.json', { ...ok, command: 'read' })).not.toEqual([]);
    expect(set.validate('mcp.schema.json', { xlvane: 1, command: null, ok: false, try: null, error: { name: 'XlvError', code: 'nope', message: '', location: {}, hint: '', docsUrl: '', details: {} } })).not.toEqual([]);
  });
});
