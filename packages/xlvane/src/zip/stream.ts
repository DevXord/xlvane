// Streaming ZIP access (docs/decisions.md D-402).
//
// Reading: `openZipSource` validates an archive held in memory or in a Blob with exactly the
// rules of `openZip` (D-102; the parsing functions are shared through `Bytes`, which can hold
// only some segments of the file). With a Blob only the tail, the central directory and the local
// headers are read up front; entry data is read and inflated piece by piece, and the limits apply
// to the bytes actually produced, as in reader.ts.
//
// Writing: `createZipStream` is pull-based. It asks the entry iterable for the next entry only
// when the previous one is complete, and each entry's data iterable for the next chunk only when
// the consumer wants more output. Entries are deflated incrementally with flag bit 3 and a data
// descriptor that carries the optional signature (Z-15); the central directory follows the last
// entry. Timestamps are fixed (D-202). ZIP64 is not written.
import { Deflate, Inflate } from 'fflate';
import { XlvError } from '../errors';
import type { ReadLimits } from '../limits';
import { crc32 } from './crc32';
import {
  Bytes,
  CORRUPT_HINT,
  HEAD_BYTES,
  LOCAL_SIZE,
  METHOD_DEFLATE,
  METHOD_STORED,
  NotLoaded,
  TAIL_BYTES,
  checkSource,
  corrupt,
  limitExceeded,
  locateCentralDirectory,
  readCentralDirectory,
  unsupported,
  validateLocalHeaders,
  type RawEntry,
  type ZipEntry,
} from './reader';

/** An archive opened by {@link openZipSource}; entries are read asynchronously. */
export interface StreamZip {
  readonly entries: readonly ZipEntry[];
  /** Finds an entry by name; names are compared case-insensitively (E-58). */
  find(name: string): ZipEntry | undefined;
  /**
   * Reads a whole entry, with the same limits and checks as `ZipReader.read`. `maxBytes`, when
   * given, replaces `maxEntryBytes` for this entry (parts held in memory use a lower bound than
   * worksheets that are streamed).
   */
  read(entry: ZipEntry | string, maxBytes?: number): Promise<Uint8Array>;
  /**
   * Inflates an entry piece by piece. Limits apply to the bytes produced so far; size and CRC-32
   * are checked at the end. Problems are reported as an `XlvError` through the stream.
   */
  stream(entry: ZipEntry | string, maxBytes?: number): ReadableStream<Uint8Array>;
}

/**
 * Compressed bytes handed to the inflater at a time. Limits are checked on every output chunk
 * fflate emits, and the input is cut small so that a highly compressed piece cannot produce much
 * output before the next check.
 */
const PIECE = 4 * 1024;

/** Random access to the input. */
interface Source {
  readonly size: number;
  read(start: number, end: number): Promise<Uint8Array>;
  /** The bytes `[start, end)` in pieces of at most {@link PIECE} bytes. */
  pieces(start: number, end: number): AsyncIterator<Uint8Array>;
}

function memorySource(data: Uint8Array): Source {
  return {
    size: data.length,
    read: (start, end) => Promise.resolve(data.subarray(start, end)),
    pieces: async function* (start, end) {
      for (let at = start; at < end; at += PIECE) yield data.subarray(at, Math.min(end, at + PIECE));
    },
  };
}

function blobSource(blob: Blob): Source {
  return {
    size: blob.size,
    read: async (start, end) => new Uint8Array(await blob.slice(start, end).arrayBuffer()),
    pieces: async function* (start, end) {
      const reader = blob.slice(start, end).stream().getReader();
      try {
        for (;;) {
          const { done, value } = await reader.read();
          if (done) return;
          for (let at = 0; at < value.length; at += PIECE) yield value.subarray(at, at + PIECE);
        }
      } finally {
        reader.releaseLock();
      }
    },
  };
}

function toSource(source: unknown): Source {
  if (source instanceof Uint8Array) return memorySource(source);
  if (source instanceof ArrayBuffer) return memorySource(new Uint8Array(source));
  if (ArrayBuffer.isView(source)) return memorySource(new Uint8Array(source.buffer, source.byteOffset, source.byteLength));
  if (typeof Blob !== 'undefined' && source instanceof Blob) return blobSource(source);
  throw new XlvError('XLV_E_INVALID_INPUT', `Expected the file as a Uint8Array, ArrayBuffer or Blob, got ${source === null ? 'null' : typeof source}.`, {
    hint: 'Pass the bytes of the file, or a Blob (in Node.js: await fs.openAsBlob(path)).',
  });
}

/** Runs `fn`, loading missing byte ranges and retrying; the parsing functions are synchronous. */
async function withLoading<T>(b: Bytes, source: Source, fn: () => T): Promise<T> {
  for (let attempt = 0; ; attempt++) {
    try {
      return fn();
    } catch (error) {
      if (!(error instanceof NotLoaded) || attempt > 16) throw error;
      const start = Math.max(0, Math.min(error.start, source.size));
      const end = Math.min(source.size, Math.max(error.end, start + 1024));
      if (end <= start) throw error;
      b.add(start, await source.read(start, end));
    }
  }
}

/**
 * Opens a ZIP archive from memory or a Blob and validates its structure with the rules of
 * `openZip`. Throws (rejects with) `XLV_E_NOT_ZIP`, `XLV_E_UNSUPPORTED_CONTAINER`,
 * `XLV_E_ZIP_CORRUPT`, `XLV_E_ZIP_UNSUPPORTED` or `XLV_E_LIMIT_EXCEEDED`.
 */
export async function openZipSource(source: Uint8Array | ArrayBuffer | Blob, limits: ReadLimits): Promise<StreamZip> {
  const src = toSource(source);
  const size = src.size;
  const head = await src.read(0, Math.min(size, HEAD_BYTES));
  checkSource(size, head, limits);

  let b: Bytes;
  if (typeof Blob !== 'undefined' && source instanceof Blob) {
    b = new Bytes(size, [{ start: 0, data: head }]);
    const tailStart = Math.max(0, size - TAIL_BYTES);
    if (size > 0) b.add(tailStart, await src.read(tailStart, size));
  } else {
    b = new Bytes(size, [{ start: 0, data: await src.read(0, size) }]);
  }

  const dir = await withLoading(b, src, () => locateCentralDirectory(b, limits));
  if (!b.has(dir.centralOffset, dir.centralLimit) && dir.centralLimit > dir.centralOffset) {
    b.add(dir.centralOffset, await src.read(dir.centralOffset, dir.centralLimit));
  }
  const entries: RawEntry[] = await withLoading(b, src, () => readCentralDirectory(b, dir.centralOffset, dir.centralLimit, dir.entryCount));
  for (const entry of entries) {
    const at = entry.headerOffset;
    if (at + LOCAL_SIZE > dir.centralOffset || b.has(at, at + LOCAL_SIZE)) continue;
    const fixed = await src.read(at, at + LOCAL_SIZE);
    b.add(at, fixed);
    const nameLength = (fixed[26] as number) | ((fixed[27] as number) << 8);
    const extraLength = (fixed[28] as number) | ((fixed[29] as number) << 8);
    const end = Math.min(dir.centralOffset, at + LOCAL_SIZE + nameLength + extraLength);
    if (end > at + LOCAL_SIZE) b.add(at + LOCAL_SIZE, await src.read(at + LOCAL_SIZE, end));
  }
  await withLoading(b, src, () => validateLocalHeaders(b, entries, dir.centralOffset));
  return new SourceZip(src, entries, limits);
}

class SourceZip implements StreamZip {
  readonly entries: readonly ZipEntry[];
  private readonly byName: Map<string, ZipEntry>;
  /** Decompressed bytes produced so far by all entries (maxTotalBytes). */
  private produced = 0;

  constructor(
    private readonly source: Source,
    entries: readonly RawEntry[],
    private readonly limits: ReadLimits,
  ) {
    this.entries = entries.map(({ nameBytes: _nameBytes, ...entry }) => entry);
    this.byName = new Map(this.entries.map((e) => [e.name.toLowerCase(), e]));
  }

  find(name: string): ZipEntry | undefined {
    return this.byName.get(name.toLowerCase());
  }

  async read(entry: ZipEntry | string, maxBytes?: number): Promise<Uint8Array> {
    const chunks: Uint8Array[] = [];
    let length = 0;
    const reader = this.stream(entry, maxBytes).getReader();
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
      length += value.length;
    }
    if (chunks.length === 1) return chunks[0] as Uint8Array;
    const out = new Uint8Array(length);
    let at = 0;
    for (const chunk of chunks) {
      out.set(chunk, at);
      at += chunk.length;
    }
    return out;
  }

  /** Checks that apply before any byte of `entry` is read, in the order of `ZipReader.read`. */
  private precheck(entryOrName: ZipEntry | string, maxEntryBytes: number): ZipEntry {
    const entry = typeof entryOrName === 'string' ? this.find(entryOrName) : entryOrName;
    if (entry === undefined) {
      throw new XlvError('XLV_E_PART_MISSING', `The archive has no entry "${String(entryOrName)}".`, {
        location: { part: `/${String(entryOrName)}` },
        hint: 'The workbook is incomplete. Re-save it in a spreadsheet application.',
      });
    }
    const { limits } = this;
    if (entry.uncompressedSize > maxEntryBytes) {
      throw limitExceeded('maxEntryBytes', maxEntryBytes, `ZIP entry "${entry.name}" declares ${entry.uncompressedSize} bytes, more than the limit.`, entry.name);
    }
    if (this.produced + entry.uncompressedSize > limits.maxTotalBytes) {
      throw limitExceeded('maxTotalBytes', limits.maxTotalBytes, 'The decompressed parts of the file are larger than the limit.', entry.name);
    }
    if (entry.method === METHOD_STORED) {
      if (entry.compressedSize !== entry.uncompressedSize) {
        throw corrupt(`Stored ZIP entry "${entry.name}" has inconsistent sizes.`, 'size-mismatch', entry.headerOffset, { name: entry.name });
      }
    } else if (entry.method !== METHOD_DEFLATE) {
      throw unsupported(
        `ZIP entry "${entry.name}" uses compression method ${entry.method}.`,
        'compression-method',
        'Only stored and Deflate-compressed entries are supported. Re-save the workbook in a spreadsheet application.',
        entry.name,
      );
    }
    return entry;
  }

  stream(entryOrName: ZipEntry | string, maxBytes?: number): ReadableStream<Uint8Array> {
    let entry: ZipEntry;
    let pieces: AsyncIterator<Uint8Array> | null = null;
    let inflater: Inflate | null = null;
    const out: Uint8Array[] = [];
    let produced = 0;
    let crc = 0;
    let ended = false;
    let finished = false;
    const zip = this;
    const { limits } = this;
    const maxEntryBytes = Math.min(maxBytes ?? limits.maxEntryBytes, limits.maxEntryBytes);

    const take = (chunk: Uint8Array): void => {
      produced += chunk.length;
      zip.produced += chunk.length;
      if (produced > maxEntryBytes) {
        throw limitExceeded('maxEntryBytes', maxEntryBytes, `ZIP entry "${entry.name}" decompresses to more than the limit.`, entry.name);
      }
      if (zip.produced > limits.maxTotalBytes) {
        throw limitExceeded('maxTotalBytes', limits.maxTotalBytes, 'The decompressed parts of the file are larger than the limit.', entry.name);
      }
      if (entry.method === METHOD_DEFLATE && produced > limits.compressionRatioMinBytes && produced / Math.max(1, entry.compressedSize) > limits.maxCompressionRatio) {
        throw limitExceeded(
          'maxCompressionRatio',
          limits.maxCompressionRatio,
          `ZIP entry "${entry.name}" expands more than ${String(limits.maxCompressionRatio)} times, which is typical of a zip bomb.`,
          entry.name,
        );
      }
      if (chunk.length > 0) {
        crc = crc32(chunk, crc);
        out.push(chunk);
      }
    };

    const finish = (): void => {
      if (entry.method === METHOD_DEFLATE && !ended) throw corrupt(`ZIP entry "${entry.name}" ends early.`, 'deflate', entry.headerOffset, { name: entry.name });
      if (produced !== entry.uncompressedSize) {
        throw corrupt(`ZIP entry "${entry.name}" is ${produced} bytes, but its header says ${entry.uncompressedSize}.`, 'size-mismatch', entry.headerOffset, {
          name: entry.name,
        });
      }
      if (crc !== entry.crc32) throw corrupt(`ZIP entry "${entry.name}" failed its CRC-32 check.`, 'crc-mismatch', entry.headerOffset, { name: entry.name });
      finished = true;
    };

    const inflateError = (error: unknown): XlvError =>
      error instanceof XlvError
        ? error
        : new XlvError('XLV_E_ZIP_CORRUPT', `ZIP entry "${entry.name}" is not a valid Deflate stream.`, {
            location: { part: `/${entry.name}`, offset: entry.headerOffset },
            hint: CORRUPT_HINT,
            details: { reason: 'deflate', name: entry.name },
            cause: error,
          });

    return new ReadableStream<Uint8Array>({
      start: () => {
        entry = zip.precheck(entryOrName, maxEntryBytes);
        pieces = zip.source.pieces(entry.dataOffset, entry.dataOffset + entry.compressedSize);
        if (entry.method === METHOD_DEFLATE) {
          inflater = new Inflate();
          inflater.ondata = (chunk: Uint8Array, final: boolean) => {
            take(chunk);
            if (final) ended = true;
          };
        }
      },
      pull: async (controller) => {
        try {
          while (out.length === 0 && !finished) {
            const next = await (pieces as AsyncIterator<Uint8Array>).next();
            if (next.done === true) {
              if (inflater && !ended) inflater.push(new Uint8Array(0), true);
              finish();
            } else if (inflater) {
              if (!ended) inflater.push(next.value, false);
            } else {
              take(next.value);
            }
          }
          while (out.length > 0) controller.enqueue(out.shift() as Uint8Array);
          if (finished) controller.close();
        } catch (error) {
          await pieces?.return?.();
          throw entry === undefined ? error : inflateError(error);
        }
      },
      cancel: async () => {
        await pieces?.return?.();
      },
    });
  }
}

// ---------------------------------------------------------------------------------------------
// Writing

/** One file for {@link createZipStream}. */
export interface ZipStreamEntry {
  /** Item name; printable ASCII, no leading slash. */
  readonly name: string;
  readonly data: Uint8Array | Iterable<Uint8Array> | AsyncIterable<Uint8Array>;
}

/** Entries a ZIP without ZIP64 can hold (D-202). */
const MAX_ENTRIES = 0xfffe;
const MAX_32 = 0xfffffffe;
const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const SIG_DESCRIPTOR = 0x08074b50;
const VERSION = 20;
const FLAG_DESCRIPTOR = 0x0008;
const DOS_DATE = (0 << 9) | (1 << 5) | 1;
const DOS_TIME = 0;

const encoder = new TextEncoder();

function tooLarge(what: string): XlvError {
  return new XlvError('XLV_E_LIMIT_EXCEEDED', `The workbook is too large to write: ${what}.`, {
    hint: 'Split the data into several workbooks. Writing ZIP64 archives is not supported yet.',
    details: { limit: 'zip32' },
  });
}

function checkName(name: unknown, seen: Set<string>): string {
  if (typeof name !== 'string' || !/^[\x21-\x7e][\x20-\x7e]*$/.test(name) || name.startsWith('/') || name.length > 0xffff) {
    const shown = String(name).slice(0, 60);
    throw new XlvError('XLV_E_INVALID_INPUT', `"${shown}" is not a valid ZIP item name.`, {
      hint: 'Item names must be printable ASCII without a leading slash. This is an internal error; please report it.',
      details: { name: shown },
    });
  }
  const key = name.toLowerCase();
  if (seen.has(key)) {
    throw new XlvError('XLV_E_INVALID_INPUT', `The ZIP item name "${name.slice(0, 60)}" is used twice.`, {
      hint: 'Item names must be unique ignoring case. This is an internal error; please report it.',
      details: { name: name.slice(0, 60) },
    });
  }
  seen.add(key);
  return name;
}

function iteratorOf<T>(value: Iterable<T> | AsyncIterable<T>): Iterator<T> | AsyncIterator<T> {
  const asyncFactory = (value as AsyncIterable<T>)[Symbol.asyncIterator];
  if (typeof asyncFactory === 'function') return asyncFactory.call(value);
  return (value as Iterable<T>)[Symbol.iterator]();
}

interface OpenEntry {
  readonly nameBytes: Uint8Array;
  readonly offset: number;
  readonly data: Iterator<Uint8Array> | AsyncIterator<Uint8Array>;
  readonly deflater: Deflate;
  crc: number;
  size: number;
  compressed: number;
}

/**
 * Writes a ZIP archive as a stream. Entries are requested from `entries` one at a time, only
 * after the previous entry's data is fully consumed, and data chunks only when the consumer
 * reads; memory therefore stays bounded by the chunk sizes. Problems (invalid or duplicate names,
 * archives that would need ZIP64) are reported as an `XlvError` through the stream.
 */
export function createZipStream(entries: Iterable<ZipStreamEntry> | AsyncIterable<ZipStreamEntry>, level = 6): ReadableStream<Uint8Array> {
  let iterator: Iterator<ZipStreamEntry> | AsyncIterator<ZipStreamEntry> | null = null;
  let current: OpenEntry | null = null;
  const central: Uint8Array[] = [];
  const seen = new Set<string>();
  const pending: Uint8Array[] = [];
  let offset = 0;
  let count = 0;
  let closed = false;

  const emit = (chunk: Uint8Array): void => {
    if (chunk.length === 0) return;
    offset += chunk.length;
    if (offset > MAX_32) throw tooLarge('more than 4 GiB');
    pending.push(chunk);
  };

  const startEntry = (entry: ZipStreamEntry): void => {
    count++;
    if (count > MAX_ENTRIES) throw tooLarge(`more than ${String(MAX_ENTRIES)} parts`);
    const name = checkName(entry.name, seen);
    const nameBytes = encoder.encode(name);
    const local = new Uint8Array(30 + nameBytes.length);
    const v = new DataView(local.buffer);
    v.setUint32(0, SIG_LOCAL, true);
    v.setUint16(4, VERSION, true);
    v.setUint16(6, FLAG_DESCRIPTOR, true);
    v.setUint16(8, METHOD_DEFLATE, true);
    v.setUint16(10, DOS_TIME, true);
    v.setUint16(12, DOS_DATE, true);
    // CRC-32 and sizes (14..25) stay 0: they follow in the data descriptor (Z-09, Z-15).
    v.setUint16(26, nameBytes.length, true);
    local.set(nameBytes, 30);
    const headerOffset = offset;
    emit(local);
    const data = entry.data;
    const dataIterator = data instanceof Uint8Array ? [data][Symbol.iterator]() : iteratorOf(data);
    // Level 0 gives Deflate stored blocks (the entry method stays 8, Deflate).
    const deflater = new Deflate({ level: level as 6 });
    const open: OpenEntry = { nameBytes, offset: headerOffset, data: dataIterator, deflater, crc: 0, size: 0, compressed: 0 };
    deflater.ondata = (chunk: Uint8Array) => {
      open.compressed += chunk.length;
      emit(chunk);
    };
    current = open;
  };

  const finishEntry = (open: OpenEntry): void => {
    open.deflater.push(new Uint8Array(0), true);
    if (open.size > MAX_32 || open.compressed > MAX_32) throw tooLarge('a part larger than 4 GiB');
    const descriptor = new Uint8Array(16);
    const d = new DataView(descriptor.buffer);
    d.setUint32(0, SIG_DESCRIPTOR, true);
    d.setUint32(4, open.crc, true);
    d.setUint32(8, open.compressed, true);
    d.setUint32(12, open.size, true);
    emit(descriptor);
    const cd = new Uint8Array(46 + open.nameBytes.length);
    const c = new DataView(cd.buffer);
    c.setUint32(0, SIG_CENTRAL, true);
    c.setUint16(4, VERSION, true);
    c.setUint16(6, VERSION, true);
    c.setUint16(8, FLAG_DESCRIPTOR, true);
    c.setUint16(10, METHOD_DEFLATE, true);
    c.setUint16(12, DOS_TIME, true);
    c.setUint16(14, DOS_DATE, true);
    c.setUint32(16, open.crc, true);
    c.setUint32(20, open.compressed, true);
    c.setUint32(24, open.size, true);
    c.setUint16(28, open.nameBytes.length, true);
    c.setUint32(42, open.offset, true);
    cd.set(open.nameBytes, 46);
    central.push(cd);
    current = null;
  };

  const finishArchive = (): void => {
    const cdOffset = offset;
    for (const cd of central) emit(cd);
    const cdSize = offset - cdOffset;
    const eocd = new Uint8Array(22);
    const e = new DataView(eocd.buffer);
    e.setUint32(0, SIG_EOCD, true);
    e.setUint16(8, central.length, true);
    e.setUint16(10, central.length, true);
    e.setUint32(12, cdSize, true);
    e.setUint32(16, cdOffset, true);
    emit(eocd);
    closed = true;
  };

  return new ReadableStream<Uint8Array>({
    pull: async (controller) => {
      try {
        // Work until there is output: a pull that enqueues nothing would not be called again.
        while (pending.length === 0 && !closed) {
          const open = current as OpenEntry | null;
          if (open === null) {
            iterator ??= iteratorOf(entries);
            const next = await iterator.next();
            if (next.done === true) finishArchive();
            else startEntry(next.value);
            continue;
          }
          const next = await open.data.next();
          if (next.done === true) {
            finishEntry(open);
            continue;
          }
          const chunk = next.value;
          if (!(chunk instanceof Uint8Array)) {
            throw new XlvError('XLV_E_INVALID_INPUT', 'ZIP entry data must be Uint8Array chunks.', { hint: 'This is an internal error; please report it.' });
          }
          open.crc = crc32(chunk, open.crc);
          open.size += chunk.length;
          if (open.size > MAX_32) throw tooLarge('a part larger than 4 GiB');
          open.deflater.push(chunk, false);
        }
        while (pending.length > 0) controller.enqueue(pending.shift() as Uint8Array);
        if (closed) controller.close();
      } catch (error) {
        await (current as OpenEntry | null)?.data.return?.();
        await iterator?.return?.();
        throw error;
      }
    },
    cancel: async () => {
      await (current as OpenEntry | null)?.data.return?.();
      await iterator?.return?.();
    },
  });
}
