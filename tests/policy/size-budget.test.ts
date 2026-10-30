// The size budget is enforced by scripts/check-pkg.mjs and published in the package README;
// this test keeps the two numbers equal (docs/decisions.md D-208).
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

describe('size budget', () => {
  it('is the same in scripts/check-pkg.mjs and packages/xlvane/README.md', () => {
    const script = readFileSync('scripts/check-pkg.mjs', 'utf8');
    const readme = readFileSync('packages/xlvane/README.md', 'utf8');
    const budget = /const SIZE_BUDGET_BYTES = ([\d_]+);/.exec(script)?.[1]?.replace(/_/g, '');
    expect(budget).toBeDefined();
    const kb = Number(budget) / 1000;
    expect(readme).toContain(`**Size budget:** ${String(kb)} kB`);
  });
});
