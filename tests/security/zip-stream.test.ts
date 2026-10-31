// Hostile archives through the streaming ZIP reader: the same structural rules as openZip, and
// limits enforced on the bytes actually produced while streaming.
import { describe, expect, it } from 'vitest';
import { DEFAULT_READ_LIMITS, type ReadLimits } from '../../packages/xlvane/src/limits';
import { openZipSource } from '../../packages/xlvane/src/zip/stream';
import { expectXlvErrorAsync } from '../helpers/expect-error';
import { buildZip, deflate } from '../helpers/raw-zip';

const MiB = 1024 * 1024;
const limits = (overrides: Partial<ReadLimits> = {}): ReadLimits => ({ ...DEFAULT_READ_LIMITS, ...overrides });
const zeros = (n: number): Uint8Array => new Uint8Array(n);
const asBlob = (bytes: Uint8Array): Blob => new Blob([bytes as BlobPart]);

async function drain(stream: ReadableStream<Uint8Array>): Promise<number> {
  let total = 0;
  for await (const chunk of stream) total += chunk.length;
  return total;
}

const structural: [string, Uint8Array, string, Record<string, unknown>?][] = [
  ['text', new TextEncoder().encode('a,b,c\n1,2,3\n'), 'XLV_E_NOT_ZIP'],
  ['empty', new Uint8Array(0), 'XLV_E_NOT_ZIP'],
  ['trailing data', buildZip([{ name: 'a', data: 'x' }], { suffix: new Uint8Array([1, 2, 3]) }), 'XLV_E_ZIP_CORRUPT'],
  ['leading data', buildZip([{ name: 'a', data: 'x' }], { prefix: new Uint8Array(16) }), 'XLV_E_ZIP_CORRUPT'],
  ['overlap', buildZip([{ name: 'a.xml', data: 'aaaa' }, { name: 'b.xml', data: 'bbbb', localOffset: 0, localName: 'a.xml' }]), 'XLV_E_ZIP_CORRUPT', { reason: 'overlapping-entries' }],
  ['duplicate', buildZip([{ name: 'xl/a.xml', data: '1' }, { name: 'XL/A.xml', data: '2' }]), 'XLV_E_ZIP_CORRUPT', { reason: 'duplicate-name' }],
  ['local name', buildZip([{ name: 'a.xml', localName: 'b.xml', data: 'x' }]), 'XLV_E_ZIP_CORRUPT'],
  ['encrypted', buildZip([{ name: 'a.xml', data: 'x', flags: 1 }]), 'XLV_E_ZIP_UNSUPPORTED'],
  ['bad count', buildZip([{ name: 'a', data: 'x' }], { entryCount: 5 }), 'XLV_E_ZIP_CORRUPT'],
  ['multi-disk', buildZip([{ name: 'a', data: 'x' }], { diskNumber: 1 }), 'XLV_E_ZIP_UNSUPPORTED'],
  ['bad local offset', buildZip([{ name: 'a', data: 'x', localOffset: 9_999_999 }]), 'XLV_E_ZIP_CORRUPT'],
];

describe('streaming reader: structure', () => {
  for (const [label, bytes, code, details] of structural) {
    it(`rejects ${label} like openZip (bytes and Blob)`, async () => {
      await expectXlvErrorAsync(() => openZipSource(bytes, limits()), code as never, details);
      await expectXlvErrorAsync(() => openZipSource(asBlob(bytes), limits()), code as never, details);
    });
  }

  it('recognises CFB containers and enforces maxInputBytes on Blobs', async () => {
    const cfb = new Uint8Array(512);
    cfb.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    await expectXlvErrorAsync(() => openZipSource(asBlob(cfb), limits()), 'XLV_E_UNSUPPORTED_CONTAINER');
    const zip = buildZip([{ name: 'a', data: 'x' }]);
    await expectXlvErrorAsync(() => openZipSource(asBlob(zip), limits({ maxInputBytes: 10 })), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxInputBytes' });
  });

  it('enforces maxEntries', async () => {
    const entries = Array.from({ length: 11 }, (_, i) => ({ name: `f${String(i)}`, data: 'x' }));
    await expectXlvErrorAsync(() => openZipSource(asBlob(buildZip(entries)), limits({ maxEntries: 10 })), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxEntries' });
  });
});

describe('streaming reader: zip bombs', () => {
  it('stops an entry that exceeds maxEntryBytes although its header lies', async () => {
    const zip = await openZipSource(asBlob(buildZip([{ name: 'bomb.xml', data: zeros(4 * MiB), declaredSize: 100 }])), limits({ maxEntryBytes: 1 * MiB }));
    await expectXlvErrorAsync(() => drain(zip.stream('bomb.xml')), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxEntryBytes' });
  });

  it('rejects an entry whose declared size exceeds maxEntryBytes before inflating', async () => {
    const zip = await openZipSource(buildZip([{ name: 'big.xml', data: 'small', declaredSize: 2 * MiB }]), limits({ maxEntryBytes: 1 * MiB }));
    await expectXlvErrorAsync(() => zip.read('big.xml'), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxEntryBytes' });
  });

  it('stops a 200 MB bomb near the limit without producing it', async () => {
    // 200 MiB of zeros deflate to about 200 KB; the header claims 1 byte. maxEntryBytes stops it.
    const compressed = deflate(zeros(200 * MiB), 9);
    const bomb = buildZip([{ name: 'bomb.xml', data: new Uint8Array(1), compressed, declaredSize: 1 }]);
    const zip = await openZipSource(asBlob(bomb), limits({ maxEntryBytes: 4 * MiB }));
    let produced = 0;
    const reader = zip.stream('bomb.xml').getReader();
    const error = await (async () => {
      try {
        for (;;) {
          const { done, value } = await reader.read();
          if (done) return null;
          produced += value.length;
        }
      } catch (e) {
        return e;
      }
    })();
    expect((error as { code?: string } | null)?.code).toBe('XLV_E_LIMIT_EXCEEDED');
    expect(produced).toBeLessThan(8 * MiB);
  });

  it('enforces maxTotalBytes across streamed entries', async () => {
    const zip = await openZipSource(
      buildZip([
        { name: 'a.xml', data: zeros(600 * 1024) },
        { name: 'b.xml', data: zeros(600 * 1024) },
      ]),
      limits({ maxTotalBytes: 1 * MiB, maxCompressionRatio: Number.POSITIVE_INFINITY }),
    );
    expect(await drain(zip.stream('a.xml'))).toBe(600 * 1024);
    await expectXlvErrorAsync(() => drain(zip.stream('b.xml')), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxTotalBytes' });
  });

  it('detects size and CRC mismatches at the end of the stream', async () => {
    const lie = await openZipSource(buildZip([{ name: 'lie.xml', data: 'x'.repeat(5000), declaredSize: 100 }]), limits());
    await expectXlvErrorAsync(() => drain(lie.stream('lie.xml')), 'XLV_E_ZIP_CORRUPT');
    const crc = await openZipSource(asBlob(buildZip([{ name: 'c.xml', data: 'hello', crc: 1234 }])), limits());
    await expectXlvErrorAsync(() => drain(crc.stream('c.xml')), 'XLV_E_ZIP_CORRUPT', { reason: 'crc-mismatch' });
  });

  it('rejects broken Deflate data and unsupported methods', async () => {
    const broken = await openZipSource(buildZip([{ name: 'x.xml', data: 'abc', compressed: new Uint8Array([0xff, 0xff, 0xff]) }]), limits());
    await expectXlvErrorAsync(() => drain(broken.stream('x.xml')), 'XLV_E_ZIP_CORRUPT');
    const method = await openZipSource(buildZip([{ name: 'x.xml', data: 'abc', method: 12 }]), limits());
    await expectXlvErrorAsync(() => method.read('x.xml'), 'XLV_E_ZIP_UNSUPPORTED');
  });
});
