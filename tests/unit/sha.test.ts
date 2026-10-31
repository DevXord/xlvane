// The SHA implementations used for password verifiers (packages/xlvane/src/util/sha.ts, D-C08)
// against node:crypto, for empty, short, block-boundary and long inputs.
import { createHash } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import { sha, SHA_ALGORITHMS } from '../../packages/xlvane/src/util/sha';

const NODE: Record<string, string> = { 'SHA-1': 'sha1', 'SHA-256': 'sha256', 'SHA-384': 'sha384', 'SHA-512': 'sha512' };

function hex(b: Uint8Array): string {
  return [...b].map((x) => x.toString(16).padStart(2, '0')).join('');
}

describe('sha', () => {
  it('matches node:crypto for every length from 0 to 300 bytes and for long inputs', () => {
    const lengths = [...Array.from({ length: 301 }, (_, i) => i), 1000, 4096, 100_003];
    let seed = 1;
    for (const n of lengths) {
      const data = new Uint8Array(n);
      for (let i = 0; i < n; i++) data[i] = (seed = (seed * 1103515245 + 12345) & 0x7fffffff) & 0xff;
      for (const alg of SHA_ALGORITHMS) {
        expect(hex(sha(alg, data)), `${alg} ${String(n)}`).toBe(createHash(NODE[alg]!).update(data).digest('hex'));
      }
    }
  });
});
