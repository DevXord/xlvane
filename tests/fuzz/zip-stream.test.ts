// Property tests for the streaming ZIP writer and reader: any entries and any chunking round-trip,
// and arbitrary bytes never throw anything but XlvError.
import fc from 'fast-check';
import { describe, expect, it } from 'vitest';
import { isXlvError } from '../../packages/xlvane/src/errors';
import { DEFAULT_READ_LIMITS } from '../../packages/xlvane/src/limits';
import { openZip } from '../../packages/xlvane/src/zip/reader';
import { createZipStream, openZipSource } from '../../packages/xlvane/src/zip/stream';
import { NUM_RUNS } from './runs';

const RUNS = { numRuns: Math.max(1, Math.floor(NUM_RUNS / 20)) };

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  for await (const chunk of stream) chunks.push(chunk);
  const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
  let at = 0;
  for (const c of chunks) {
    out.set(c, at);
    at += c.length;
  }
  return out;
}

const content = fc.oneof(
  fc.uint8Array({ maxLength: 3000 }),
  fc.tuple(fc.string({ maxLength: 20 }), fc.integer({ min: 0, max: 400 })).map(([s, n]) => new TextEncoder().encode(s.repeat(n))),
);
const entry = fc.record({ data: content, cuts: fc.array(fc.integer({ min: 1, max: 1000 }), { maxLength: 6 }) });

describe('streaming ZIP properties', () => {
  it('round-trips any entries with any chunking', async () => {
    await fc.assert(
      fc.asyncProperty(fc.array(entry, { maxLength: 6 }), async (list) => {
        const entries = list.map((e, i) => {
          const parts: Uint8Array[] = [];
          let at = 0;
          for (const cut of e.cuts) {
            parts.push(e.data.subarray(at, at + cut));
            at += cut;
          }
          parts.push(e.data.subarray(at));
          return { name: `part${String(i)}.bin`, data: parts };
        });
        const bytes = await collect(createZipStream(entries));
        const sync = openZip(bytes, DEFAULT_READ_LIMITS);
        const blob = await openZipSource(new Blob([bytes as BlobPart]), DEFAULT_READ_LIMITS);
        list.forEach((e, i) => expect(sync.read(`part${String(i)}.bin`)).toEqual(e.data));
        for (const [i, e] of list.entries()) expect(await collect(blob.stream(`part${String(i)}.bin`))).toEqual(e.data);
      }),
      RUNS,
    );
  });

  it('never throws anything but XlvError on arbitrary input', async () => {
    await fc.assert(
      fc.asyncProperty(fc.uint8Array({ maxLength: 400 }), async (bytes) => {
        try {
          const zip = await openZipSource(new Blob([bytes as BlobPart]), DEFAULT_READ_LIMITS);
          for (const e of zip.entries) await zip.read(e);
        } catch (error) {
          expect(isXlvError(error)).toBe(true);
        }
      }),
      { numRuns: NUM_RUNS },
    );
  });
});
