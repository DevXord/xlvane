// Property tests: the streaming reader returns the same cells as readWorkbook, and the streaming
// writer gives the same values as the batch writer.
import fc from 'fast-check';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { formatCellRef } from '../../packages/xlvane/src/sml/cell-ref';
import { openWorkbookStream } from '../../packages/xlvane/src/stream-read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook, type CellInput } from '../../packages/xlvane/src/write';
import { firstLoneSurrogate } from '../../packages/xlvane/src/xml/writer';
import { NUM_RUNS } from './runs';

const RUNS = { numRuns: Math.max(1, Math.floor(NUM_RUNS / 50)) };
const text = fc.string({ maxLength: 12 }).filter((s) => firstLoneSurrogate(s) === -1);
const value: fc.Arbitrary<CellInput> = fc.oneof(
  text,
  fc.double({ noNaN: true, noDefaultInfinity: true }),
  fc.boolean(),
  fc.date({ min: new Date(Date.UTC(1900, 0, 1)), max: new Date(Date.UTC(9999, 11, 31)), noInvalidDate: true }),
  fc.constant(null),
);
const grid = fc.array(fc.array(value, { maxLength: 6 }), { minLength: 1, maxLength: 30 });

async function bytesOf(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

describe('streaming properties', () => {
  it('stream reading equals batch reading, and stream writing equals batch writing', async () => {
    await fc.assert(
      fc.asyncProperty(grid, fc.boolean(), async (rows, shared) => {
        const batch = createWorkbook();
        const s = batch.addSheet('S');
        rows.forEach((row, r) => row.forEach((v, c) => s.set(formatCellRef(r + 1, c + 1), v, {})));
        const batchBytes = writeWorkbook(batch);
        const streamBytes = await bytesOf(writeWorkbookStream({ sharedStrings: shared, sheets: [{ name: 'S', rows }] }));
        const expected = [...readWorkbook(batchBytes).getSheet('S').cells()].map((c) => [c.ref, c.value]);
        for (const bytes of [batchBytes, streamBytes]) {
          const book = await openWorkbookStream(bytes);
          const got: unknown[] = [];
          for await (const row of book.rows('S')) for (const c of row.cells) got.push([c.ref, c.value]);
          expect(got).toEqual(expected);
        }
      }),
      RUNS,
    );
  });
});
