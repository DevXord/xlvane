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
  /** Maximum length of one XML text node, attribute value, comment or processing instruction. Default 8 Mi characters. */
  readonly maxXmlTextLength: number;
  /** Maximum number of elements plus attributes in a part that is parsed into a tree (workbook, styles, relationships). Default 1 000 000. */
  readonly maxXmlTreeElements: number;
  /**
   * Maximum number of cells in the workbook. Default 5 000 000 (about 0.9 GB of memory; see the
   * README). `readWorkbook` reads charts after the cells of every sheet, and their chart groups and
   * series (each counted once per chart part) count against what the cells leave of this limit;
   * charts over it are skipped with `XLV_W_LIMIT_REACHED`.
   */
  readonly maxCells: number;
  /**
   * Maximum total length, in UTF-16 code units, of the shared strings that cells refer to, each
   * reference counted again. Protects against a small file that points thousands of cells at
   * one huge shared string (a few MB that become many GB once values are copied, e.g. to CSV).
   * The same limit applies, counted separately, to shared formulas: the length of the formula of
   * each cell that shares one, as moved to that cell. Default 512 Mi.
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

/**
 * How many positions (rows × columns) a table or CSV made from a sheet may have per cell read
 * from it, beyond `maxCells`: a few cells far apart must not make the output a whole sheet
 * (D-H01).
 */
export const OUTPUT_PER_CELL = 16;

/**
 * Checks the number of positions an output built from a sheet would have against
 * max(`maxCells`, OUTPUT_PER_CELL × cells read). Throws `XLV_E_LIMIT_EXCEEDED` above it.
 * `hint` replaces the default hint, which suits outputs read from a file.
 */
export function checkOutputPositions(positions: number, cells: number, maxCells: number, what: string, sheet: string, hint?: string): void {
  const budget = Math.max(maxCells, OUTPUT_PER_CELL * cells);
  if (positions <= budget) return;
  throw new XlvError('XLV_E_LIMIT_EXCEEDED', `${what} of sheet "${sheet}" would have ${String(positions)} positions (rows × columns) for ${String(cells)} cells, more than the limit of ${String(budget)}.`, {
    location: { sheet },
    hint: hint ?? `The cells are spread over a far larger area than they fill. Pass a smaller range (for example { range: "A1:H1000" } or --range), or raise limits.maxCells if the file is trusted.`,
    details: { limit: 'maxCells', value: budget },
  });
}

/**
 * Every row object repeats the keys of all columns, so a long header makes a small file give
 * gigabytes of keys (and JSON.stringify of the rows fail). The key text of all rows is bounded
 * by `limits.maxSharedStringChars`, the bound on text that cells repeat (D-H01).
 */
export function checkKeyText(columns: readonly { readonly key: string }[], rows: number, limit: number, sheet: string): void {
  let perRow = 0;
  for (const c of columns) perRow += c.key.length;
  if (perRow * rows <= limit) return;
  throw new XlvError('XLV_E_LIMIT_EXCEEDED', `The rows of the table of sheet "${sheet}" would repeat ${String(perRow)} characters of column names in each of ${String(rows)} rows, more than the limit of ${String(limit)} characters.`, {
    location: { sheet },
    hint: `Each row object holds the text of every column name. Read fewer columns or rows with the range option (for example { range: "A1:H1000" } or --range), or raise limits.maxSharedStringChars if the file is trusted.`,
    details: { limit: 'maxSharedStringChars', value: limit },
  });
}
