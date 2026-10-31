import { XlvError } from './errors';
import { suggest } from './util/suggest';

/**
 * Resource limits applied while reading a file. Every limit protects against a class of
 * malicious or broken input (zip bombs, deeply nested XML, huge sheets). Exceeding a limit
 * throws `XLV_E_LIMIT_EXCEEDED`; the error's `details.limit` names the limit.
 * Any limit can be set to `Infinity` to switch it off.
 *
 * @example
 * ```ts
 * import { readWorkbook, type ReadLimits } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const limits: Partial<ReadLimits> = { maxCells: 50_000_000, maxInputBytes: Infinity };
 * const workbook = readWorkbook(bytes, { limits });
 * console.log(workbook.sheets.length);
 * ```
 */
export interface ReadLimits {
  /** Maximum size of the input file in bytes. Default 256 MiB. */
  readonly maxInputBytes: number;
  /** Maximum number of entries in the ZIP central directory. Default 10 000. */
  readonly maxEntries: number;
  /** Maximum decompressed size of one ZIP entry, measured while decompressing. Default 256 MiB. */
  readonly maxEntryBytes: number;
  /** Maximum decompressed size of all entries read from one file. Default 512 MiB. */
  readonly maxTotalBytes: number;
  /** Maximum ratio of decompressed to compressed size of one entry. Default 200. */
  readonly maxCompressionRatio: number;
  /** The compression ratio is checked only once an entry has produced this many bytes. Default 1 MiB. */
  readonly compressionRatioMinBytes: number;
  /** Maximum element nesting depth in an XML part. Default 128. */
  readonly maxXmlDepth: number;
  /** Maximum number of attributes on one XML element. Default 256. */
  readonly maxXmlAttributes: number;
  /** Maximum length of an XML element or attribute name. Default 1 024 characters. */
  readonly maxXmlNameLength: number;
  /** Maximum length of one XML text node or attribute value. Default 8 Mi characters. */
  readonly maxXmlTextLength: number;
  /** Maximum number of elements plus attributes in a part that is parsed into a tree (workbook, styles, relationships). Default 1 000 000. */
  readonly maxXmlTreeElements: number;
  /** Maximum number of cells in the workbook. Default 5 000 000 (about 0.9 GB of memory; see the README). */
  readonly maxCells: number;
  /**
   * Maximum total length, in UTF-16 code units, of the shared strings that cells refer to, each
   * reference counted again. Protects against a small file that points thousands of cells at
   * one huge shared string (a few MB that become many GB once values are copied, e.g. to CSV).
   * Default 512 Mi.
   */
  readonly maxSharedStringChars: number;
}

const MiB = 1024 * 1024;

/**
 * The default {@link ReadLimits}. Frozen; pass `{ limits: { ... } }` to change individual values.
 *
 * @example
 * ```ts
 * import { DEFAULT_READ_LIMITS } from 'xlvane';
 * console.log(DEFAULT_READ_LIMITS.maxEntries); // 10000
 * ```
 */
export const DEFAULT_READ_LIMITS: Readonly<ReadLimits> = Object.freeze({
  maxInputBytes: 256 * MiB,
  maxEntries: 10_000,
  maxEntryBytes: 256 * MiB,
  maxTotalBytes: 512 * MiB,
  maxCompressionRatio: 200,
  compressionRatioMinBytes: 1 * MiB,
  maxXmlDepth: 128,
  maxXmlAttributes: 256,
  maxXmlNameLength: 1024,
  maxXmlTextLength: 8 * MiB,
  maxXmlTreeElements: 1_000_000,
  maxCells: 5_000_000,
  maxSharedStringChars: 512 * MiB,
});

const LIMIT_NAMES = Object.keys(DEFAULT_READ_LIMITS) as (keyof ReadLimits)[];

/**
 * Validates user-supplied limits and merges them over the defaults. Only own, known keys are
 * read, so keys such as `__proto__` cannot change anything; unknown keys are rejected.
 */
export function resolveLimits(input: unknown): ReadLimits {
  const limits: Record<keyof ReadLimits, number> = { ...DEFAULT_READ_LIMITS };
  if (input === undefined) return limits;
  if (input === null || typeof input !== 'object' || Array.isArray(input)) {
    throw new XlvError('XLV_E_INVALID_OPTION', `Option "limits" must be an object, got ${describe(input)}.`, {
      hint: 'Pass an object such as { limits: { maxEntries: 10000 } }, or leave "limits" out to use the defaults.',
    });
  }
  for (const key of Object.keys(input)) {
    if (!(LIMIT_NAMES as string[]).includes(key)) {
      const close = suggest(key, LIMIT_NAMES);
      throw new XlvError('XLV_E_INVALID_OPTION', `Unknown limit "${key}".`, {
        hint:
          close.length > 0
            ? `Did you mean ${close.map((c) => `"${c}"`).join(' or ')}? Known limits: ${LIMIT_NAMES.join(', ')}.`
            : `Known limits: ${LIMIT_NAMES.join(', ')}.`,
        details: { option: key, known: LIMIT_NAMES },
      });
    }
  }
  for (const name of LIMIT_NAMES) {
    if (!Object.hasOwn(input, name)) continue;
    const value: unknown = (input as Record<string, unknown>)[name];
    limits[name] = validate(name, value);
  }
  return limits;
}

function validate(name: keyof ReadLimits, value: unknown): number {
  const isRatio = name === 'maxCompressionRatio';
  const valid =
    typeof value === 'number' &&
    (value === Number.POSITIVE_INFINITY ||
      (isRatio ? Number.isFinite(value) && value >= 1 : Number.isSafeInteger(value) && value > 0));
  if (!valid) {
    throw new XlvError(
      'XLV_E_INVALID_OPTION',
      isRatio
        ? `Limit "${name}" must be a number of at least 1 or Infinity, got ${describe(value)}.`
        : `Limit "${name}" must be a whole number greater than 0 or Infinity, got ${describe(value)}.`,
      {
        hint: `Pass for example { limits: { ${name}: ${String(DEFAULT_READ_LIMITS[name])} } }, or leave it out to use the default.`,
        details: { option: name },
      },
    );
  }
  return value;
}

function describe(value: unknown): string {
  if (value === null) return 'null';
  if (Array.isArray(value)) return 'an array';
  if (typeof value === 'string') return JSON.stringify(value.length > 40 ? `${value.slice(0, 40)}…` : value);
  if (typeof value === 'object' || typeof value === 'function') return `a value of type ${typeof value}`;
  return String(value);
}
