// Snapshot tests of every help screen (docs/PLAN.md section 5), and the acceptance criterion
// "every example from help works": each example line runs in a directory with fixture files.
import { describe, expect, it } from 'vitest';
import { COMMANDS } from '../../packages/xlvane/src/cli/commands/index';
import { cli, splitCommand, withFixtures } from './helpers';

const tmp = '/nonexistent-cwd';

describe('help screens', () => {
  it('welcome screen without arguments', async () => {
    const r = await cli([], { cwd: tmp });
    expect(r.code).toBe(0);
    expect(r.stderr).toBe('');
    await expect(r.stdout).toMatchFileSnapshot('__snapshots__/welcome.txt');
    expect((await cli(['help'], { cwd: tmp })).stdout).toBe(r.stdout);
    expect((await cli(['--help'], { cwd: tmp })).stdout).toBe(r.stdout);
  });

  it('--version', async () => {
    expect(await cli(['--version'], { cwd: tmp })).toEqual({ code: 0, stdout: '0.0.0-test\n', stderr: '' });
    expect((await cli(['-v'], { cwd: tmp })).stdout).toBe('0.0.0-test\n');
  });

  it('has at most 9 commands, each a verb with 2–4 examples and related commands', () => {
    expect(COMMANDS.length).toBeLessThanOrEqual(9);
    expect(COMMANDS.map((c) => c.name)).toEqual(['inspect', 'read', 'convert', 'create', 'edit', 'diff', 'validate', 'migrate', 'mcp']);
    for (const c of COMMANDS) {
      expect(c.examples.length, c.name).toBeGreaterThanOrEqual(2);
      expect(c.examples.length, c.name).toBeLessThanOrEqual(4);
      expect(c.related.length, c.name).toBeGreaterThan(0);
      for (const e of c.examples) expect(e.command.startsWith(`xlvane ${c.name}`) || e.command.startsWith('xlvane help'), e.command).toBe(true);
    }
  });

  for (const c of COMMANDS) {
    it(`help ${c.name}`, async () => {
      const r = await cli(['help', c.name], { cwd: tmp });
      expect(r.code).toBe(0);
      await expect(r.stdout).toMatchFileSnapshot(`__snapshots__/help-${c.name}.txt`);
      expect((await cli([c.name, '--help'], { cwd: tmp })).stdout).toBe(r.stdout);
      expect((await cli([c.name, '-h'], { cwd: tmp })).stdout).toBe(r.stdout);
    });
  }
});

/** Example command lines shown on a screen: lines starting with "xlvane " after indentation. */
function exampleLines(screen: string): string[] {
  return screen
    .split('\n')
    .map((l) => /^ {2}(xlvane \S.*?)(?: {2,}.*)?$/.exec(l)?.[1])
    .filter((l): l is string => l !== undefined && !l.includes('<') && !/\[[a-z -]+\]/.test(l)); // not the usage line
}

describe('every example from help works', () => {
  const screens = async (): Promise<string[]> => [
    (await cli([], { cwd: tmp })).stdout,
    ...(await Promise.all(COMMANDS.map(async (c) => (await cli(['help', c.name], { cwd: tmp })).stdout))),
  ];

  it('finds the examples', async () => {
    const all = (await screens()).flatMap(exampleLines);
    expect(all.length).toBe(5 + COMMANDS.reduce((n, c) => n + c.examples.length, 0));
  });

  it('runs each one in a fresh directory', { timeout: 60_000 }, async () => {
    const failures: string[] = [];
    for (const line of (await screens()).flatMap(exampleLines)) {
      await withFixtures(async (dir) => {
        const [, ...args] = splitCommand(line);
        const name = args[0]!;
        const spec = COMMANDS.find((c) => c.name === name);
        const r = await cli(args, { cwd: dir });
        // diff and validate exit with 1 when they find something (documented in their help).
        const allowed = spec?.checks ? [0, 1] : [0];
        // xlvane mcp needs the separate package; without it the documented result is exit 3.
        if (name === 'mcp' && !line.includes('help')) allowed.push(3);
        if (!allowed.includes(r.code)) failures.push(`${line}\n  exit ${String(r.code)}\n${r.stderr}${r.stdout}`);
        if (args.includes('--json')) JSON.parse(r.stdout);
      });
    }
    expect(failures, failures.join('\n')).toEqual([]);
  });
});
