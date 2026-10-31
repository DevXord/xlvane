// The portable smoke bundle in Node.js (the CI matrix runs this on Node.js 22, 24 and 26). The
// other runtimes run the same bundle through `npm run test:runtimes` (D-904).
import { mkdir, mkdtemp, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
// @ts-expect-error -- plain JavaScript build script without type declarations
import { buildSmoke } from '../../scripts/build-smoke.mjs';
import type { SmokeResult } from './smoke';

const ROOT = fileURLToPath(new URL('../..', import.meta.url));
let dir = '';
let bundle = '';

beforeAll(async () => {
  await mkdir(join(ROOT, 'node_modules/.cache'), { recursive: true });
  dir = await mkdtemp(join(ROOT, 'node_modules/.cache/xlvane-smoke-'));
  bundle = (await buildSmoke(dir)) as string;
}, 120_000);

afterAll(async () => {
  if (dir) await rm(dir, { recursive: true, force: true });
});

describe('portable smoke bundle', () => {
  it('passes every check in Node.js', { timeout: 120_000 }, async () => {
    const { runSmoke } = (await import(pathToFileURL(bundle).href)) as { runSmoke: () => Promise<SmokeResult[]> };
    const results = await runSmoke();
    expect(results.length).toBeGreaterThan(10);
    expect(results.filter((r) => !r.ok)).toEqual([]);
  });
});
