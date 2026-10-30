import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { DEFAULT_READ_LIMITS } from '../../packages/xlvane/src/limits';
import { openZip } from '../../packages/xlvane/src/zip/reader';
import { createZipStream, openZipSource, type ZipStreamEntry } from '../../packages/xlvane/src/zip/stream';
import { writeZip } from '../../packages/xlvane/src/zip/writer';
import { expectXlvErrorAsync } from '../helpers/expect-error';
import { buildZip } from '../helpers/raw-zip';

const enc = new TextEncoder();
const dec = new TextDecoder();

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

function* pieces(text: string, size: number): Generator<Uint8Array> {
  const bytes = enc.encode(text);
  for (let i = 0; i < bytes.length; i += size) yield bytes.subarray(i, i + size);
}

const SAMPLE: ZipStreamEntry[] = [
  { name: '[Content_Types].xml', data: enc.encode('<Types/>') },
  { name: 'xl/worksheets/sheet1.xml', data: pieces('<row/>'.repeat(20_000), 777) },
  { name: 'empty.bin', data: new Uint8Array(0) },
];

describe('createZipStream', () => {
  it('writes archives that the strict reader, fflate and openZipSource accept', async () => {
    const bytes = await collect(createZipStream(SAMPLE));
    const zip = openZip(bytes, DEFAULT_READ_LIMITS);
    expect(zip.entries.map((e) => e.name)).toEqual(['[Content_Types].xml', 'xl/worksheets/sheet1.xml', 'empty.bin']);
    expect(dec.decode(zip.read('xl/worksheets/sheet1.xml'))).toBe('<row/>'.repeat(20_000));
    expect(zip.read('empty.bin')).toEqual(new Uint8Array(0));
    const files = unzipSync(bytes);
    expect(dec.decode(files['[Content_Types].xml'])).toBe('<Types/>');
    const source = await openZipSource(bytes, DEFAULT_READ_LIMITS);
    expect(dec.decode(await source.read('xl/worksheets/sheet1.xml'))).toBe('<row/>'.repeat(20_000));
  });

  it('uses data descriptors with a signature and no flags but bit 3', async () => {
    const bytes = await collect(createZipStream([{ name: 'a.txt', data: enc.encode('hello') }]));
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    expect(view.getUint32(0, true)).toBe(0x04034b50);
    expect(view.getUint16(6, true)).toBe(0x0008);
    expect(view.getUint32(14, true)).toBe(0); // CRC in the descriptor
    expect(view.getUint16(12, true)).toBe(0x21); // fixed date 1980-01-01
    const zip = openZip(bytes, DEFAULT_READ_LIMITS);
    const entry = zip.entries[0]!;
    const descriptor = entry.dataOffset + entry.compressedSize;
    expect(view.getUint32(descriptor, true)).toBe(0x08074b50);
    expect(view.getUint32(descriptor + 4, true)).toBe(entry.crc32);
  });

  it('is deterministic for the same input and chunking', async () => {
    const make = () => createZipStream([{ name: 'x.xml', data: pieces('abc'.repeat(5000), 100) }]);
    expect(await collect(make())).toEqual(await collect(make()));
  });

  it('accepts async iterables of entries and of data', async () => {
    async function* data(): AsyncGenerator<Uint8Array> {
      yield enc.encode('one ');
      yield enc.encode('two');
    }
    async function* entries(): AsyncGenerator<ZipStreamEntry> {
      yield { name: 'a.txt', data: data() };
      yield { name: 'b.txt', data: [enc.encode('b')] };
    }
    const zip = openZip(await collect(createZipStream(entries())), DEFAULT_READ_LIMITS);
    expect(dec.decode(zip.read('a.txt'))).toBe('one two');
    expect(dec.decode(zip.read('b.txt'))).toBe('b');
  });

  it('asks for the next entry only after the previous entry is fully consumed', async () => {
    const log: string[] = [];
    function* data(name: string): Generator<Uint8Array> {
      for (let i = 0; i < 3; i++) {
        log.push(`${name}:${String(i)}`);
        yield enc.encode(`${name}${String(i)}`);
      }
      log.push(`${name}:done`);
    }
    function* entries(): Generator<ZipStreamEntry> {
      log.push('entry a');
      yield { name: 'a', data: data('a') };
      log.push('entry b');
      yield { name: 'b', data: data('b') };
      log.push('entries done');
    }
    await collect(createZipStream(entries()));
    expect(log).toEqual(['entry a', 'a:0', 'a:1', 'a:2', 'a:done', 'entry b', 'b:0', 'b:1', 'b:2', 'b:done', 'entries done']);
  });

  it('pulls data lazily (backpressure)', async () => {
    let produced = 0;
    function* data(): Generator<Uint8Array> {
      for (let i = 0; i < 1000; i++) {
        produced++;
        yield new Uint8Array(1024).fill(i % 251);
      }
    }
    const reader = createZipStream([{ name: 'big.bin', data: data() }]).getReader();
    await reader.read();
    expect(produced).toBeLessThan(50);
    await reader.cancel();
  });

  it('rejects invalid and duplicate names as stream errors', async () => {
    await expectXlvErrorAsync(() => collect(createZipStream([{ name: 'ż.xml', data: new Uint8Array(1) }])), 'XLV_E_INVALID_INPUT');
    await expectXlvErrorAsync(
      () => collect(createZipStream([{ name: 'a.xml', data: new Uint8Array(1) }, { name: 'A.XML', data: new Uint8Array(1) }])),
      'XLV_E_INVALID_INPUT',
    );
  });

  it('refuses more entries than a ZIP without ZIP64 can hold', { timeout: 120_000 }, async () => {
    function* many(): Generator<ZipStreamEntry> {
      for (let i = 0; i <= 0xfffe; i++) yield { name: `f${String(i)}`, data: new Uint8Array(0) };
    }
    await expectXlvErrorAsync(() => collect(createZipStream(many())), 'XLV_E_LIMIT_EXCEEDED');
  });
});

describe('openZipSource', () => {
  const archive = writeZip([
    { name: 'a.txt', data: enc.encode('stored?') },
    { name: 'dir/b.xml', data: enc.encode('<b>deflated</b>'.repeat(1000)) },
  ]);

  it('reads Uint8Array, ArrayBuffer and Blob sources alike', async () => {
    const sources = [archive, archive.slice().buffer, new Blob([archive as BlobPart])];
    for (const source of sources) {
      const zip = await openZipSource(source, DEFAULT_READ_LIMITS);
      expect(zip.entries.map((e) => e.name)).toEqual(['a.txt', 'dir/b.xml']);
      expect(zip.find('DIR/B.XML')?.name).toBe('dir/b.xml');
      expect(dec.decode(await zip.read('a.txt'))).toBe('stored?');
      expect(dec.decode(await collect(zip.stream('dir/b.xml')))).toBe('<b>deflated</b>'.repeat(1000));
    }
  });

  it('gives the same entries as openZip, including data descriptors and ZIP64', async () => {
    const raw = buildZip(
      [
        { name: 'x.xml', data: 'x'.repeat(3000), dataDescriptor: true },
        { name: 'y.xml', data: 'yyy', method: 0, zip64: true },
      ],
      { zip64: true },
    );
    const sync = openZip(raw, DEFAULT_READ_LIMITS);
    const blob = await openZipSource(new Blob([raw as BlobPart]), DEFAULT_READ_LIMITS);
    expect(blob.entries).toEqual(sync.entries);
    for (const e of sync.entries) expect(await blob.read(e)).toEqual(sync.read(e));
  });

  it('reports missing entries', async () => {
    const zip = await openZipSource(archive, DEFAULT_READ_LIMITS);
    await expectXlvErrorAsync(() => zip.read('nope'), 'XLV_E_PART_MISSING');
  });

  it('reads a large Blob entry in small pieces', async () => {
    let seed = 12345;
    const noise = new Uint8Array(20_000_000).map(() => {
      seed = (seed * 1103515245 + 12345) >>> 0;
      return seed >>> 24;
    });
    const big = writeZip([{ name: 'big.bin', data: noise }]);
    let largest = 0;
    let requested = 0;
    class CountingBlob extends Blob {
      override slice(start?: number, end?: number, type?: string): Blob {
        const part = super.slice(start, end, type);
        requested += part.size;
        return part;
      }
    }
    const zip = await openZipSource(new CountingBlob([big as BlobPart]), { ...DEFAULT_READ_LIMITS });
    const before = requested;
    expect(before).toBeLessThan(200_000);
    const reader = zip.stream('big.bin').getReader();
    let total = 0;
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      largest = Math.max(largest, value.length);
      total += value.length;
    }
    expect(total).toBe(20_000_000);
    // fflate emits inflated data per decoded block, so chunks are block-sized, never the whole entry.
    expect(largest).toBeLessThan(2_000_000);
  });
});
