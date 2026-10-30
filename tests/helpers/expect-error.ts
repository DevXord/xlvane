import { expect } from 'vitest';
import { XlvError, type XlvErrorCode } from '../../packages/xlvane/src/errors';

/** Runs `fn`, expects it to throw an XlvError with `code`, and returns the error. */
export function expectXlvError(fn: () => unknown, code: XlvErrorCode, details?: Record<string, unknown>): XlvError {
  let thrown: unknown;
  try {
    fn();
  } catch (error) {
    thrown = error;
  }
  expect(thrown, `expected ${code} to be thrown`).toBeInstanceOf(XlvError);
  const error = thrown as XlvError;
  expect({ code: error.code, message: error.message }).toMatchObject({ code });
  if (details) expect(error.details).toMatchObject(details);
  expect(error.hint.length).toBeGreaterThan(0);
  return error;
}

/** Async variant of {@link expectXlvError}. */
export async function expectXlvErrorAsync(
  fn: () => Promise<unknown>,
  code: XlvErrorCode,
  details?: Record<string, unknown>,
): Promise<XlvError> {
  let thrown: unknown;
  try {
    await fn();
  } catch (error) {
    thrown = error;
  }
  expect(thrown, `expected ${code} to be thrown`).toBeInstanceOf(XlvError);
  const error = thrown as XlvError;
  expect(error.code).toBe(code);
  if (details) expect(error.details).toMatchObject(details);
  return error;
}
