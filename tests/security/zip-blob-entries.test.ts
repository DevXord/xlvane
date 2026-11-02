// Opening a ZIP archive given as a Blob loads the local header of every entry into a list of
// segments. Each addition sorted the whole list and each lookup searched it linearly, so 10 000
// entries took quadratic time (1.1.1 security review, D-H01). Segments are now inserted in order
// and found by binary search over the largest end so far.
import { describe, expect, it } from 'vitest';
import { openWorkbookStream } from '../../packages/xlvane/src/stream-read';
import { Bytes, NotLoaded } from '../../packages/xlvane/src/zip/reader';
import { buildXlsx } from '../helpers/xlsx-builder';

describe('ZIP archives from a Blob', () => {
  it('opens a file with 9 990 entries in linear time', async () => {
    const extraParts: Record<string, string> = {};
    for (let i = 0; i < 9990; i++) extraParts[`x/${String(i)}.bin`] = '';
    const bytes = buildXlsx({ sheets: [{ name: 'S', data: '' }], extraParts });
    const start = performance.now();
    await openWorkbookStream(new Blob([bytes]));
    expect(performance.now() - start).toBeLessThan(3000);
  }, 60_000);

  it('finds exactly the segments a naive search finds, for overlapping segments added in any order', () => {
    let seed = 12345;
    const random = (n: number): number => {
      seed = (seed * 1103515245 + 12345) & 0x7fffffff;
      return seed % n;
    };
    for (let round = 0; round < 50; round++) {
      const size = 1000;
      const b = new Bytes(size, []);
      const added: [number, number][] = [];
      for (let k = 0; k < 40; k++) {
        const start = random(size);
        const length = 1 + random(k % 5 === 0 ? 400 : 20);
        const end = Math.min(size, start + length);
        const data = new Uint8Array(end - start);
        for (let i = 0; i < data.length; i++) data[i] = (start + i) & 0xff;
        b.add(start, data);
        added.push([start, end]);
        for (let q = 0; q < 20; q++) {
          const s = random(size);
          const e = Math.min(size, s + 1 + random(30));
          const covered = added.some(([a, z]) => a <= s && e <= z);
          expect(b.has(s, e)).toBe(covered);
          if (covered) {
            const got = b.slice(s, e);
            for (let i = 0; i < got.length; i++) expect(got[i]).toBe((s + i) & 0xff);
          } else expect(() => b.slice(s, e)).toThrow(NotLoaded);
        }
      }
    }
  });
});
