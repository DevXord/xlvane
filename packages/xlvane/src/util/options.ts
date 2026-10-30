// Reading option objects given by callers: own enumerable properties only, each read once,
// unknown names rejected with a "Did you mean" hint (the same rules as D-201).
import { XlvError } from '../errors';
import type { ReadLimits } from '../limits';
import type { Workbook } from '../model';
import { readWorkbook } from '../read';
import { suggest } from './suggest';

/** Reads `value` as an options object with the given known names. `undefined` gives an empty map. */
export function readOptions(value: unknown, known: readonly string[], what: string): Map<string, unknown> {
  const out = new Map<string, unknown>();
  if (value === undefined) return out;
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new XlvError('XLV_E_INVALID_OPTION', `${what} must be an object.`, {
      hint: `Pass an object with some of these properties: ${known.join(', ')}, or leave it out.`,
    });
  }
  for (const key of Object.keys(value)) {
    if (!known.includes(key)) {
      const close = suggest(key, known);
      throw new XlvError('XLV_E_INVALID_OPTION', `Unknown option "${key.slice(0, 40)}" in ${what}.`, {
        hint: `${close.length > 0 ? `Did you mean "${close[0] ?? ''}"? ` : ''}Known options: ${known.join(', ')}.`,
        details: { option: key.slice(0, 40), known },
      });
    }
    const v = (value as Record<string, unknown>)[key];
    if (v !== undefined) out.set(key, v);
  }
  return out;
}

/** Throws `XLV_E_INVALID_OPTION` for an option value outside the allowed set. */
export function invalidOptionValue(name: string, value: unknown, allowed: string): XlvError {
  const shown = typeof value === 'string' ? JSON.stringify(value.slice(0, 40)) : typeof value === 'object' ? typeof value : String(value);
  return new XlvError('XLV_E_INVALID_OPTION', `The option "${name}" has the value ${shown}.`, {
    hint: `Use ${allowed}.`,
    details: { option: name },
  });
}

/** Returns `value` if it is one of `allowed`, otherwise throws `XLV_E_INVALID_OPTION`. */
export function oneOfOption<T extends string>(name: string, value: unknown, allowed: readonly T[], fallback: T): T {
  if (value === undefined) return fallback;
  if (typeof value === 'string' && (allowed as readonly string[]).includes(value)) return value as T;
  throw invalidOptionValue(name, value, allowed.map((a) => `"${a}"`).join(', '));
}

/** Returns `value` if it is a boolean, otherwise throws `XLV_E_INVALID_OPTION`. */
export function booleanOption(name: string, value: unknown, fallback: boolean): boolean {
  if (value === undefined) return fallback;
  if (typeof value === 'boolean') return value;
  throw invalidOptionValue(name, value, 'true or false');
}

/** `true` for a workbook returned by `readWorkbook` (duck-typed, as in `inspect`). */
export function isWorkbook(input: unknown): input is Workbook {
  return (
    input !== null &&
    typeof input === 'object' &&
    !ArrayBuffer.isView(input) &&
    !(input instanceof ArrayBuffer) &&
    typeof (input as Workbook).getSheet === 'function' &&
    Array.isArray((input as Workbook).sheets)
  );
}

/** Accepts bytes or an already read workbook. */
export function openInput(input: unknown, limits: unknown): Workbook {
  if (isWorkbook(input)) return input;
  return readWorkbook(input as Uint8Array, limits === undefined ? undefined : { limits: limits as Partial<ReadLimits> });
}
