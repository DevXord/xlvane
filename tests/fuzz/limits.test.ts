// Property tests for resolveLimits: any input yields a complete, valid set of limits or
// XLV_E_INVALID_OPTION, and keys such as __proto__ never reach a prototype.
import fc from 'fast-check';
import type { Arbitrary } from 'fast-check';
import { describe, expect, it } from 'vitest';
import { XlvError } from '../../packages/xlvane/src/errors';
import { DEFAULT_READ_LIMITS, resolveLimits, type ReadLimits } from '../../packages/xlvane/src/limits';
import { expectPrototypesUntouched } from '../helpers/prototypes';
import { NUM_RUNS } from './runs';

const LIMIT_NAMES = Object.keys(DEFAULT_READ_LIMITS) as (keyof ReadLimits)[];
const DEFAULTS_SNAPSHOT = { ...DEFAULT_READ_LIMITS };
const HOSTILE_KEYS = ['__proto__', 'constructor', 'prototype', 'toString', 'valueOf', 'hasOwnProperty', 'polluted', ''];

const isLimitName = (key: string): key is keyof ReadLimits => (LIMIT_NAMES as string[]).includes(key);

/** The documented contract: Infinity, a ratio of at least 1, or a whole number greater than 0. */
function isValidLimit(name: keyof ReadLimits, value: unknown): value is number {
  if (typeof value !== 'number') return false;
  if (value === Number.POSITIVE_INFINITY) return true;
  return name === 'maxCompressionRatio' ? Number.isFinite(value) && value >= 1 : Number.isSafeInteger(value) && value > 0;
}

function attempt(input: unknown): ReadLimits | XlvError {
  try {
    return resolveLimits(input);
  } catch (error) {
    if (error instanceof XlvError) return error;
    throw error;
  }
}

/** A complete limits object: a fresh plain object with exactly the known keys, all valid. */
function expectCompleteLimits(limits: ReadLimits): void {
  expect(limits).not.toBe(DEFAULT_READ_LIMITS);
  expect(Object.getPrototypeOf(limits)).toBe(Object.prototype);
  expect(Object.keys(limits).sort()).toEqual([...LIMIT_NAMES].sort());
  for (const name of LIMIT_NAMES) expect(isValidLimit(name, limits[name])).toBe(true);
}

function expectInvalidOption(error: XlvError): void {
  expect(error.code).toBe('XLV_E_INVALID_OPTION');
  expect(error.hint.length).toBeGreaterThan(0);
}

const limitKey = fc.oneof(
  { arbitrary: fc.constantFrom(...LIMIT_NAMES), weight: 4 },
  { arbitrary: fc.constantFrom(...HOSTILE_KEYS), weight: 2 },
  { arbitrary: fc.string(), weight: 1 },
);

const limitValue: Arbitrary<unknown> = fc.oneof(
  fc.integer({ min: 1, max: 1_000_000 }),
  fc.constantFrom(
    Number.POSITIVE_INFINITY,
    Number.NEGATIVE_INFINITY,
    Number.NaN,
    0,
    -0,
    -1,
    0.5,
    1,
    1.5,
    Number.EPSILON,
    Number.MAX_SAFE_INTEGER,
    Number.MAX_SAFE_INTEGER + 1,
    Number.MAX_VALUE,
  ),
  fc.double(),
  fc.anything(),
);

describe('resolveLimits', () => {
  it('returns complete limits or throws XLV_E_INVALID_OPTION for any value', () => {
    const input = fc.anything({
      key: limitKey,
      values: [limitValue],
      withBigInt: true,
      withBoxedValues: true,
      withDate: true,
      withMap: true,
      withSet: true,
      withNullPrototype: true,
      withObjectString: true,
      withSparseArray: true,
      withTypedArray: true,
      withUnicodeString: true,
    });
    fc.assert(
      fc.property(input, (value) => {
        const result = attempt(value);
        if (result instanceof XlvError) {
          expectInvalidOption(result);
        } else {
          expectCompleteLimits(result);
          for (const name of LIMIT_NAMES) {
            const own = typeof value === 'object' && value !== null && Object.hasOwn(value, name);
            if (!own) expect(result[name]).toBe(DEFAULT_READ_LIMITS[name]);
          }
        }
        expect(DEFAULT_READ_LIMITS).toEqual(DEFAULTS_SNAPSHOT);
        expectPrototypesUntouched();
      }),
      { numRuns: NUM_RUNS },
    );
  });

  it('merges own known keys over the defaults and rejects unknown keys and invalid values', () => {
    const input = fc.record({
      entries: fc.array(fc.tuple(limitKey, limitValue), { maxLength: 6 }),
      prototype: fc.constantFrom('object', 'null', 'inherited'),
    });
    fc.assert(
      fc.property(input, ({ entries, prototype }) => {
        // Inherited limits must be ignored; only own keys count.
        const target: object =
          prototype === 'null'
            ? (Object.create(null) as object)
            : prototype === 'inherited'
              ? (Object.create({ maxEntries: 1, polluted: true }) as object)
              : {};
        // defineProperty makes "__proto__" an own key instead of changing the prototype.
        for (const [key, value] of entries) {
          Object.defineProperty(target, key, { value, enumerable: true, configurable: true, writable: true });
        }
        const own = new Map(entries);

        const unknown = Object.keys(target).find((key) => !isLimitName(key));
        const invalid = LIMIT_NAMES.find((name) => own.has(name) && !isValidLimit(name, own.get(name)));
        const result = attempt(target);
        if (unknown !== undefined || invalid !== undefined) {
          expect(result).toBeInstanceOf(XlvError);
          const error = result as XlvError;
          expectInvalidOption(error);
          // Unknown keys are reported first, in key order; then invalid values, in limit order.
          expect(error.details['option']).toBe(unknown ?? invalid);
        } else {
          expect(result).not.toBeInstanceOf(XlvError);
          const limits = result as ReadLimits;
          expectCompleteLimits(limits);
          const expected: Record<string, unknown> = { ...DEFAULT_READ_LIMITS };
          for (const [key, value] of own) expected[key] = value;
          expect(limits).toEqual(expected);
        }
        expectPrototypesUntouched();
      }),
      { numRuns: NUM_RUNS },
    );
  });
});
