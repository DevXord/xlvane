// Phase 4 acceptance criterion: streaming read (and write) needs constant memory. Retained heap
// is measured after forced garbage collection while rows are processed; reading 8 times as many
// rows must not retain noticeably more memory.
import { describe, expect, it } from 'vitest';
import { openWorkbookStream } from '../../packages/xlvane/src/stream-read';
import { writeWorkbookStream, type StreamRowInput } from '../../packages/xlvane/src/stream-write';

const gc = (globalThis as { gc?: () => void }).gc;
const MB = 1024 * 1024;

function* rows(n: number): Generator<StreamRowInput> {
  for (let i = 1; i <= n; i++) yield [i, `text ${String(i % 1000)}`, i * 0.5, i % 2 === 0, new Date(Date.UTC(2024, 0, (i % 28) + 1))];
}

function retained(): number {
  gc?.();
  return process.memoryUsage().heapUsed;
}

/** Writes `n` rows; returns the file and the most heap retained above the start while writing. */
async function write(n: number): Promise<{ bytes: Uint8Array; peak: number }> {
  const base = retained();
  let peak = 0;
  function* sampled(): Generator<StreamRowInput> {
    let i = 0;
    for (const r of rows(n)) {
      if (++i % 25_000 === 0) peak = Math.max(peak, retained() - base);
      yield r;
    }
  }
  const chunks: Uint8Array[] = [];
  let size = 0;
  const reader = writeWorkbookStream({ sheets: [{ name: 'S', rows: sampled() }] }).getReader();
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    size += value.length;
  }
  // The collected output is not counted: it is what the caller keeps.
  const bytes = new Uint8Array(size);
  let at = 0;
  for (const c of chunks) {
    bytes.set(c, at);
    at += c.length;
  }
  return { bytes, peak };
}

async function read(bytes: Uint8Array): Promise<{ cells: number; peak: number }> {
  const base = retained();
  let peak = 0;
  let cells = 0;
  let i = 0;
  const book = await openWorkbookStream(bytes);
  for await (const row of book.rows('S')) {
    cells += row.cells.length;
    if (++i % 25_000 === 0) peak = Math.max(peak, retained() - base);
  }
  return { cells, peak };
}

describe('streaming memory', () => {
  it('runs with --expose-gc', () => {
    expect(typeof gc).toBe('function');
  });

  it('retains about the same memory for 50 000 and 400 000 rows', { timeout: 300_000 }, async () => {
    const small = await write(50_000);
    const large = await write(400_000);
    const smallRead = await read(small.bytes);
    const largeRead = await read(large.bytes);
    expect(smallRead.cells).toBe(250_000);
    expect(largeRead.cells).toBe(2_000_000);
    const report = { writeSmall: small.peak / MB, writeLarge: large.peak / MB, readSmall: smallRead.peak / MB, readLarge: largeRead.peak / MB };
    console.log('retained heap (MB):', JSON.stringify(report));
    // A batch read of 2 000 000 cells retains about 180 bytes per cell (D-115), i.e. ~350 MB.
    expect(largeRead.peak - smallRead.peak).toBeLessThan(16 * MB);
    expect(large.peak - small.peak).toBeLessThan(16 * MB);
    expect(largeRead.peak).toBeLessThan(48 * MB);
  });
});
