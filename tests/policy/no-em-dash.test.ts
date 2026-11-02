// The maintainer's rule: no em dash (U+2014) anywhere in the repository's text; "-" is used
// instead. Checks every tracked text file except third-party corpus files.
import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const ROOT = fileURLToPath(new URL('../..', import.meta.url));
const TEXT = /\.(?:ts|mts|mjs|js|json|md|yml|yaml|txt|html|css|svg)$/;

describe('em dash', () => {
  it('does not occur in tracked text files', () => {
    const files = spawnSync('git', ['ls-files'], { cwd: ROOT, encoding: 'utf8' })
      .stdout.split('\n')
      .filter((f) => TEXT.test(f) && !f.startsWith('tests/corpus/') && f !== 'package-lock.json');
    expect(files.length).toBeGreaterThan(100);
    const found = files.filter((f) => readFileSync(join(ROOT, f), 'utf8').includes('—'));
    expect(found).toEqual([]);
  });
});
