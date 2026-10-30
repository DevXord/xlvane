// Every CLI error message is a snapshot: what happened, where, what to do next (docs/PLAN.md
// section 5). Paths are replaced by <dir> so the snapshot does not depend on the temp directory.
import { writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { cli, withFixtures } from './helpers';

const CASES: readonly (readonly [string, readonly string[], number])[] = [
  ['unknown command with suggestion', ['inspcet', 'report.xlsx'], 2],
  ['unknown command without suggestion', ['frobnicate'], 2],
  ['option before the command', ['--sheet', 'Sales'], 2],
  ['unknown option with suggestion', ['read', 'report.xlsx', '--shet', 'Sales'], 2],
  ['unknown short option', ['read', 'report.xlsx', '-x'], 2],
  ['option without its value', ['read', 'report.xlsx', '--sheet'], 2],
  ['boolean option with a value', ['diff', 'old.xlsx', 'new.xlsx', '--styles=yes'], 2],
  ['missing argument', ['convert', 'report.xlsx'], 2],
  ['too many arguments', ['inspect', 'report.xlsx', 'extra'], 2],
  ['invalid number', ['read', 'report.xlsx', '--limit', 'ten'], 2],
  ['invalid header option', ['read', 'report.xlsx', '--header', 'first'], 2],
  ['unknown help topic', ['help', 'raed'], 2],
  ['file not found with suggestion', ['inspect', 'reprot.xlsx'], 3],
  ['file not found', ['inspect', 'missing.xlsx'], 3],
  ['not a workbook', ['inspect', 'data.json'], 3],
  ['sheet not found', ['read', 'report.xlsx', '--sheet', 'Salse'], 3],
  ['output exists', ['convert', 'report.xlsx', 'sales.csv'], 3],
  ['edit in place without --force', ['edit', 'report.xlsx', '--set', 'Sales!B2=1'], 3],
  ['edit without changes', ['edit', 'report.xlsx', '-o', 'x.xlsx'], 2],
  ['invalid --set', ['edit', 'report.xlsx', '--set', 'B2', '-o', 'x.xlsx'], 2],
  ['--set without a sheet', ['edit', 'report.xlsx', '--set', 'B2=1', '-o', 'x.xlsx'], 2],
  ['unknown output format', ['convert', 'report.xlsx', 'out.pdf'], 2],
  ['same format on both sides', ['convert', 'sales.csv', 'out.json'], 2],
  ['invalid column rule', ['validate', 'report.xlsx', '--columns', 'Revenue:money'], 2],
  ['invalid JSON input', ['create', 'bad.json', 'out.xlsx'], 3],
  ['migrate without package.json', ['migrate', 'nowhere'], 3],
  ['mcp not installed', ['mcp'], 3],
];

describe('error messages', () => {
  it('are stable (snapshot) and use the documented exit codes', async () => {
    const blocks: string[] = [];
    await withFixtures(async (dir) => {
      await writeFile(join(dir, 'bad.json'), '[{"a": 1},');
      for (const [title, argv, code] of CASES) {
        const r = await cli(argv, { cwd: dir });
        expect(r.code, title).toBe(code);
        expect(r.stdout, title).toBe('');
        // Three parts: what (Error:), where (Where:, when known) and what to do (Fix:).
        expect(r.stderr, title).toMatch(/^Error: .+\n(Where: .+\n)?Fix: {3}.+\n/);
        // JSON.parse messages come from V8 and differ between Node.js versions.
        const stderr = r.stderr.split(dir).join('<dir>').replace(/(is not valid JSON: ).*/, '$1<parser message>');
        blocks.push(`## ${title}\n$ xlvane ${argv.join(' ')}\n${stderr}`);
      }
    });
    await expect(blocks.join('\n')).toMatchFileSnapshot('__snapshots__/errors.txt');
  });
});
