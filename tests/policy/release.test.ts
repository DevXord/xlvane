// Release policy (docs/PLAN.md phase 10; docs/decisions.md D-1001 … D-1004): published packages
// are public with provenance and the same version, the release workflow publishes through npm
// trusted publishing (OIDC) with staged publishing and no secrets, every action is pinned to a
// commit SHA, and no shipped text still says that the packages are not published.
import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const pkg = (dir: string) => JSON.parse(readFileSync(`${dir}/package.json`, 'utf8')) as Record<string, unknown> & { version: string; name: string };
const xlvane = pkg('packages/xlvane');
const exceljs = pkg('packages/compat');
const workflows = ['ci.yml', 'release.yml'].map((f) => [f, readFileSync(`.github/workflows/${f}`, 'utf8')] as const);

describe('published packages', () => {
  it('are public, with provenance, at the same version', () => {
    for (const p of [xlvane, exceljs]) {
      expect(p['private'], p.name).toBeUndefined();
      expect(p['publishConfig'], p.name).toEqual({ access: 'public', provenance: true });
      expect(p['repository'], p.name).toMatchObject({ url: 'git+https://github.com/DevXord/xlvane.git' });
    }
    expect(exceljs.version).toBe(xlvane.version);
    expect((exceljs['dependencies'] as Record<string, string>)['xlvane']).toBe(`^${xlvane.version}`);
    expect(pkg('.')['private']).toBe(true);
  });

  it('the version check accepts the matching tag only', () => {
    const run = (tag: string) => spawnSync('node', ['scripts/check-release-version.mjs', tag], { encoding: 'utf8' }).status;
    expect(run(`v${xlvane.version}`)).toBe(0);
    expect(run('v0.0.1')).toBe(1);
    expect(run(xlvane.version)).toBe(1);
  });

  it('no shipped text says the packages are unpublished', () => {
    const files = ['README.md', 'packages/xlvane/README.md', 'packages/compat/README.md', 'skills/xlvane/SKILL.md', 'docs/exceljs-compat.md', 'docs/errors.md', 'docs/cli-json.md', 'packages/xlvane/src/cli/commands/migrate.ts'];
    for (const f of files) expect(readFileSync(f, 'utf8'), f).not.toMatch(/not (?:yet )?published|until (?:it|they) (?:is|are) published/i);
  });
});

describe('workflows', () => {
  it('install the same LibreOffice packages in CI and before a release', () => {
    // The LibreOffice tests need python3-uno (UNO fact scripts) and Carlito (default row heights).
    const installs = workflows.map(([f, text]) => [f, text.match(/apt-get install[^\n]*libreoffice-calc[^\n]*/)?.[0]] as const);
    for (const [f, line] of installs) {
      expect(line, f).toContain('python3-uno');
      expect(line, f).toContain('fonts-crosextra-carlito');
    }
  });

  it('pin every action to a full commit SHA', () => {
    for (const [name, text] of workflows) {
      for (const m of text.matchAll(/uses:\s*(\S+)/g)) expect(m[1], name).toMatch(/^[\w.-]+\/[\w.-]+@[0-9a-f]{40}$/);
    }
  });

  it('use no secrets and grant id-token only to the publish job', () => {
    for (const [name, text] of workflows) {
      expect(text, name).not.toMatch(/secrets\.|NODE_AUTH_TOKEN|NPM_TOKEN/);
      expect(text, name).toMatch(/^permissions:\n {2}contents: read$/m);
    }
    const release = workflows[1]![1];
    expect(release.match(/id-token: write/g)).toHaveLength(1);
    const publish = release.slice(release.indexOf('\n  publish:'));
    expect(publish).toContain('id-token: write');
    expect(publish).toContain('needs: verify');
    expect(publish).toContain('npm stage publish -w packages/xlvane');
    expect(publish).toContain('npm stage publish -w packages/compat');
    expect(publish).not.toMatch(/npm publish/);
    expect(release).toMatch(/tags: \['v\*'\]/);
  });
});
