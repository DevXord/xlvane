// The publish job holds the OIDC token (id-token: write) used for npm provenance, so it must not
// run install scripts of dependencies (pre-release security audit, docs/decisions.md D-R03).
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

describe('release workflow', () => {
  // `npm install -g npm@…` in that job updates npm itself and is not a dependency install.
  it('installs the dependencies without lifecycle scripts in the job that can publish', () => {
    const text = readFileSync('.github/workflows/release.yml', 'utf8');
    const jobs = text.split(/\n {2}(?=[\w-]+:\n)/);
    const publishing = jobs.filter((job) => /id-token:\s*write/.test(job));
    expect(publishing.length).toBeGreaterThan(0);
    for (const job of publishing) {
      const installs = job.match(/run:\s*npm ci\b[^\n]*/g) ?? [];
      expect(installs.length).toBeGreaterThan(0);
      for (const line of installs) expect(line).toContain('--ignore-scripts');
    }
  });
});
