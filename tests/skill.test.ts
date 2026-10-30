// The Agent Skill (skills/xlvane/SKILL.md): its frontmatter follows the Agent Skills
// specification (F-SKILL-01), and it installs with one command of the `skills` CLI from the
// repository directory, offline (telemetry and security-audit requests off, DISABLE_TELEMETRY).
// Its ```ts blocks are compiled by tests/policy/tsdoc-examples.test.ts.
import { spawnSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { binOf } from './helpers/build-package';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const SKILL = join(ROOT, 'skills/xlvane/SKILL.md');

/** The frontmatter as flat key/value pairs (the skill uses only single-line values). */
function frontmatter(text: string): Record<string, string> {
  const m = /^---\n([\s\S]*?)\n---\n/.exec(text);
  if (!m) throw new Error('SKILL.md has no YAML frontmatter');
  const out: Record<string, string> = {};
  for (const line of m[1]!.split('\n')) {
    const kv = /^([a-z-]+):\s*(.*)$/.exec(line);
    if (!kv) throw new Error(`Unexpected frontmatter line: ${line}`);
    out[kv[1]!] = kv[2]!;
  }
  return out;
}

describe('SKILL.md', () => {
  it('follows the Agent Skills specification', async () => {
    const fm = frontmatter(await readFile(SKILL, 'utf8'));
    expect(Object.keys(fm).sort()).toEqual(['compatibility', 'description', 'license', 'name']);
    // name: 1–64 characters, a-z, 0-9 and hyphens, no leading/trailing/double hyphen, equal to the directory name.
    expect(fm['name']).toMatch(/^(?!-)(?!.*--)[a-z0-9-]{1,64}(?<!-)$/);
    expect(fm['name']).toBe('xlvane');
    expect(fm['description']!.length).toBeGreaterThan(0);
    expect(fm['description']!.length).toBeLessThanOrEqual(1024);
    expect(fm['compatibility']!.length).toBeLessThanOrEqual(500);
    expect(fm['license']).toBe('MIT');
  });

  it('installs with one command (npx skills add <path>)', { timeout: 120_000 }, async () => {
    const project = await mkdtemp(join(tmpdir(), 'xlvane-skill-'));
    try {
      await writeFile(join(project, 'package.json'), '{"name":"skill-test","private":true}\n');
      const r = spawnSync(process.execPath, [binOf('skills', 'skills'), 'add', join(ROOT, 'skills'), '--skill', 'xlvane', '--agent', 'claude-code', '--yes', '--copy'], {
        cwd: project,
        encoding: 'utf8',
        timeout: 100_000,
        env: { PATH: process.env['PATH'], HOME: project, USERPROFILE: project, DISABLE_TELEMETRY: '1', DO_NOT_TRACK: '1', CI: '1' },
      });
      expect(r.status, `${r.stdout}\n${r.stderr}`).toBe(0);
      expect(await readFile(join(project, '.claude/skills/xlvane/SKILL.md'), 'utf8')).toBe(await readFile(SKILL, 'utf8'));
    } finally {
      await rm(project, { recursive: true, force: true });
    }
  });
});
