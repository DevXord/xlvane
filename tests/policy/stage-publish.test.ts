// scripts/stage-publish.mjs stages each package of a release (docs/decisions.md D-I07): a version
// already on npm is skipped, a package that does not exist yet gets a warning, and a failed package
// does not stop the others, so running the workflow again finishes what is missing. npm is replaced
// by a script that records its calls (XLV_STAGE_NPM, tests only).
import { spawnSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';

async function run(state: Record<string, 'published' | 'exists' | 'missing' | 'fails'>): Promise<{ status: number | null; calls: string[]; stdout: string }> {
  const dir = await mkdtemp(join(tmpdir(), 'xlvane-stage-'));
  try {
    const log = join(dir, 'calls.txt');
    const fake = join(dir, 'npm.mjs');
    await writeFile(
      fake,
      `import { appendFileSync } from 'node:fs';
const args = process.argv.slice(2);
appendFileSync(${JSON.stringify(log)}, args.join(' ') + '\\n');
const state = ${JSON.stringify(state)};
if (args[0] === 'view') {
  const [name, version] = args[1].startsWith('@') ? ['@' + args[1].slice(1).split('@')[0], args[1].slice(1).split('@')[1]] : args[1].split('@');
  const s = state[name];
  if (s === 'missing') process.exit(1);
  if (version !== undefined) { if (s === 'published') console.log(version); process.exit(0); }
  console.log(name);
  process.exit(0);
}
if (args[0] === 'stage') {
  const dir = args[3];
  const name = { 'packages/xlvane': 'xlvane', 'packages/compat': '@xlvane/compat', 'packages/mcp': '@xlvane/mcp' }[dir];
  process.exit(state[name] === 'fails' ? 1 : 0);
}
process.exit(2);
`,
    );
    const r = spawnSync(process.execPath, ['scripts/stage-publish.mjs', 'v9.9.9'], { encoding: 'utf8', env: { ...process.env, XLV_STAGE_NPM: fake } });
    const calls = (await readFile(log, 'utf8').catch(() => '')).trim().split('\n').filter(Boolean);
    return { status: r.status, calls, stdout: r.stdout + r.stderr };
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

const staged = (calls: string[]) => calls.filter((c) => c.startsWith('stage publish')).map((c) => c.split(' ')[3]);

describe('stage-publish', () => {
  it('stages every package that exists and does not have the version yet', async () => {
    const r = await run({ xlvane: 'exists', '@xlvane/compat': 'exists', '@xlvane/mcp': 'exists' });
    expect(r.status).toBe(0);
    expect(staged(r.calls)).toEqual(['packages/xlvane', 'packages/compat', 'packages/mcp']);
  });

  it('skips versions already on npm, so a second run finishes the rest', async () => {
    const r = await run({ xlvane: 'published', '@xlvane/compat': 'published', '@xlvane/mcp': 'exists' });
    expect(r.status).toBe(0);
    expect(staged(r.calls)).toEqual(['packages/mcp']);
    expect(r.stdout).toContain('xlvane 9.9.9 is already on npm');
  });

  it('warns about a package that is not on npm yet and stages the others', async () => {
    const r = await run({ xlvane: 'exists', '@xlvane/compat': 'exists', '@xlvane/mcp': 'missing' });
    expect(r.status).toBe(0);
    expect(staged(r.calls)).toEqual(['packages/xlvane', 'packages/compat']);
    expect(r.stdout).toContain('::warning::@xlvane/mcp is not on npm yet');
  });

  it('a failed package does not stop the others, and the run fails at the end', async () => {
    const r = await run({ xlvane: 'exists', '@xlvane/compat': 'fails', '@xlvane/mcp': 'exists' });
    expect(r.status).toBe(1);
    expect(staged(r.calls)).toEqual(['packages/xlvane', 'packages/compat', 'packages/mcp']);
    expect(r.stdout).toContain('::error::@xlvane/compat');
  });

  it('the release workflow stages through this script only', async () => {
    const workflow = await readFile('.github/workflows/release.yml', 'utf8');
    expect(workflow).toContain('node scripts/stage-publish.mjs "$GITHUB_REF_NAME"');
    expect(workflow.match(/run:\s*npm stage publish/g) ?? []).toEqual([]);
  });
});
