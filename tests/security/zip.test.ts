// Hostile ZIP archives: zip bombs, lying headers, overlapping entries, ambiguous names.
import { describe, expect, it } from 'vitest';
import { DEFAULT_READ_LIMITS, type ReadLimits } from '../../packages/xlvane/src/limits';
import { openZip } from '../../packages/xlvane/src/zip/reader';
import { expectXlvError } from '../helpers/expect-error';
import { buildZip, deflate } from '../helpers/raw-zip';
import { expectPrototypesUntouched } from '../helpers/prototypes';

const MiB = 1024 * 1024;
const limits = (overrides: Partial<ReadLimits> = {}): ReadLimits => ({ ...DEFAULT_READ_LIMITS, ...overrides });
const zeros = (n: number): Uint8Array => new Uint8Array(n);

describe('container detection', () => {
  it('rejects data that is not a ZIP archive', () => {
    expectXlvError(() => openZip(new TextEncoder().encode('a,b,c\n1,2,3\n'), limits()), 'XLV_E_NOT_ZIP');
    expectXlvError(() => openZip(new Uint8Array(0), limits()), 'XLV_E_NOT_ZIP');
  });

  it('recognises Compound File Binary containers (.xls or encrypted .xlsx)', () => {
    const cfb = new Uint8Array(512);
    cfb.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    const error = expectXlvError(() => openZip(cfb, limits()), 'XLV_E_UNSUPPORTED_CONTAINER');
    expect(error.hint).toMatch(/\.xls|password|encrypt/i);
  });

  it('enforces maxInputBytes before parsing', () => {
    expectXlvError(() => openZip(buildZip([{ name: 'a', data: 'x' }]), limits({ maxInputBytes: 10 })), 'XLV_E_LIMIT_EXCEEDED', {
      limit: 'maxInputBytes',
    });
  });

  it('rejects a truncated archive', () => {
    const zip = buildZip([{ name: 'a.txt', data: 'x'.repeat(1000), method: 0 }]);
    expectXlvError(() => openZip(zip.subarray(0, zip.length - 10), limits()), 'XLV_E_ZIP_CORRUPT');
    expectXlvError(() => openZip(zip.subarray(0, 200), limits()), 'XLV_E_ZIP_CORRUPT');
  });

  it('rejects data after the end of central directory record', () => {
    expectXlvError(() => openZip(buildZip([{ name: 'a', data: 'x' }], { suffix: new Uint8Array([1, 2, 3]) }), limits()), 'XLV_E_ZIP_CORRUPT');
  });

  it('rejects data before the first entry (offsets would be wrong)', () => {
    expectXlvError(() => openZip(buildZip([{ name: 'a', data: 'x' }], { prefix: new Uint8Array(16) }), limits()), 'XLV_E_ZIP_CORRUPT');
  });
});

describe('entry count', () => {
  it('enforces maxEntries before reading the central directory', () => {
    const entries = Array.from({ length: 11 }, (_, i) => ({ name: `f${i}`, data: 'x' }));
    expectXlvError(() => openZip(buildZip(entries), limits({ maxEntries: 10 })), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxEntries' });
    expect(openZip(buildZip(entries), limits({ maxEntries: 11 })).entries).toHaveLength(11);
  });

  it('rejects an entry count that does not match the central directory', () => {
    expectXlvError(() => openZip(buildZip([{ name: 'a', data: 'x' }], { entryCount: 5 }), limits()), 'XLV_E_ZIP_CORRUPT');
  });
});

describe('zip bombs', () => {
  it('stops decompressing when an entry exceeds maxEntryBytes, even if the header lies', () => {
    const real = zeros(4 * MiB);
    const zip = buildZip([{ name: 'bomb.xml', data: real, declaredSize: 100 }]);
    const reader = openZip(zip, limits({ maxEntryBytes: 1 * MiB }));
    expectXlvError(() => reader.read('bomb.xml'), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxEntryBytes' });
  });

  it('rejects an entry whose declared size exceeds maxEntryBytes without decompressing it', () => {
    const zip = buildZip([{ name: 'big.xml', data: 'small', declaredSize: 2 * MiB }]);
    const reader = openZip(zip, limits({ maxEntryBytes: 1 * MiB }));
    expectXlvError(() => reader.read('big.xml'), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxEntryBytes' });
  });

  it('rejects an entry whose real size differs from the declared size', () => {
    const zip = buildZip([{ name: 'lie.xml', data: 'x'.repeat(5000), declaredSize: 100 }]);
    expectXlvError(() => openZip(zip, limits()).read('lie.xml'), 'XLV_E_ZIP_CORRUPT');
  });

  it('enforces maxTotalBytes across entries', () => {
    const zip = buildZip([
      { name: 'a.xml', data: zeros(600 * 1024) },
      { name: 'b.xml', data: zeros(600 * 1024) },
    ]);
    const reader = openZip(zip, limits({ maxTotalBytes: 1 * MiB, maxCompressionRatio: Number.POSITIVE_INFINITY }));
    reader.read('a.xml');
    expectXlvError(() => reader.read('b.xml'), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxTotalBytes' });
  });

  it('enforces maxCompressionRatio once an entry produced compressionRatioMinBytes', () => {
    const zip = buildZip([
      { name: 'small.xml', data: zeros(512 * 1024) },
      { name: 'large.xml', data: zeros(8 * MiB) },
    ]);
    const reader = openZip(zip, limits());
    expect(reader.read('small.xml')).toHaveLength(512 * 1024);
    expectXlvError(() => reader.read('large.xml'), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxCompressionRatio' });
  });

  it('does not count a stored entry towards the compression ratio', () => {
    const zip = buildZip([{ name: 'stored.bin', data: zeros(2 * MiB), method: 0 }]);
    expect(openZip(zip, limits()).read('stored.bin')).toHaveLength(2 * MiB);
  });
});

describe('ambiguous or overlapping entries', () => {
  it('rejects two central directory entries that point at the same data', () => {
    const zip = buildZip([
      { name: 'a.xml', data: 'aaaa' },
      { name: 'b.xml', data: 'bbbb', localOffset: 0, localName: 'a.xml' },
    ]);
    expectXlvError(() => openZip(zip, limits()), 'XLV_E_ZIP_CORRUPT', { reason: 'overlapping-entries' });
  });

  it('rejects an entry whose local header lies inside the data of another entry', () => {
    const zip = buildZip([
      { name: 'a.xml', data: 'x'.repeat(200), method: 0 },
      { name: 'b.xml', data: 'y', localOffset: 40 },
    ]);
    expectXlvError(() => openZip(zip, limits()), 'XLV_E_ZIP_CORRUPT');
  });

  it('rejects duplicate names, also when they differ only in letter case', () => {
    expectXlvError(
      () => openZip(buildZip([{ name: 'xl/a.xml', data: '1' }, { name: 'xl/a.xml', data: '2' }]), limits()),
      'XLV_E_ZIP_CORRUPT',
      { reason: 'duplicate-name' },
    );
    expectXlvError(
      () => openZip(buildZip([{ name: 'xl/a.xml', data: '1' }, { name: 'XL/A.xml', data: '2' }]), limits()),
      'XLV_E_ZIP_CORRUPT',
      { reason: 'duplicate-name' },
    );
  });

  it('rejects a local header whose name differs from the central directory', () => {
    const zip = buildZip([{ name: 'a.xml', localName: 'b.xml', data: 'x' }]);
    expectXlvError(() => openZip(zip, limits()).read('a.xml'), 'XLV_E_ZIP_CORRUPT');
  });
});

describe('unsupported or broken entries', () => {
  it('rejects encrypted entries', () => {
    const zip = buildZip([{ name: 'a.xml', data: 'x', flags: 0x0001 }]);
    expectXlvError(() => openZip(zip, limits()), 'XLV_E_ZIP_UNSUPPORTED', { reason: 'encrypted' });
  });

  it('rejects compression methods other than stored and deflated', () => {
    const zip = buildZip([{ name: 'a.xml', data: 'x', method: 12, compressed: new Uint8Array([1, 2, 3]) }]);
    expectXlvError(() => openZip(zip, limits()).read('a.xml'), 'XLV_E_ZIP_UNSUPPORTED', { reason: 'compression-method' });
  });

  it('rejects multi-disk archives', () => {
    expectXlvError(() => openZip(buildZip([{ name: 'a', data: 'x' }], { diskNumber: 1 }), limits()), 'XLV_E_ZIP_UNSUPPORTED', {
      reason: 'multi-disk',
    });
  });

  it('rejects a CRC-32 mismatch', () => {
    const zip = buildZip([{ name: 'a.xml', data: 'content', crc: 0x12345678 }]);
    expectXlvError(() => openZip(zip, limits()).read('a.xml'), 'XLV_E_ZIP_CORRUPT', { reason: 'crc-mismatch' });
  });

  it('rejects a stored entry whose sizes disagree', () => {
    const zip = buildZip([{ name: 'a.xml', data: 'content', method: 0, declaredSize: 3 }]);
    expectXlvError(() => openZip(zip, limits()).read('a.xml'), 'XLV_E_ZIP_CORRUPT');
  });

  it('rejects an invalid deflate stream', () => {
    const zip = buildZip([{ name: 'a.xml', data: 'content', compressed: new Uint8Array([0xff, 0xff, 0xff, 0xff]) }]);
    expectXlvError(() => openZip(zip, limits()).read('a.xml'), 'XLV_E_ZIP_CORRUPT');
  });

  it('rejects a deflate stream that ends early', () => {
    const full = deflate('x'.repeat(10_000) + 'y'.repeat(10_000));
    const zip = buildZip([{ name: 'a.xml', data: 'x'.repeat(10_000) + 'y'.repeat(10_000), compressed: full.subarray(0, 8) }]);
    expectXlvError(() => openZip(zip, limits()).read('a.xml'), 'XLV_E_ZIP_CORRUPT');
  });

  it('rejects a central directory that points outside the file', () => {
    expectXlvError(() => openZip(buildZip([{ name: 'a', data: 'x' }], { centralOffset: 1_000_000 }), limits()), 'XLV_E_ZIP_CORRUPT');
  });

  it('rejects a local header offset outside the file', () => {
    expectXlvError(() => openZip(buildZip([{ name: 'a', data: 'x', localOffset: 999_999 }]), limits()), 'XLV_E_ZIP_CORRUPT');
  });

  it('leaves prototypes untouched', () => {
    openZip(buildZip([{ name: '__proto__', data: '{"polluted":true}' }]), limits()).read('__proto__');
    expectPrototypesUntouched();
  });
});
