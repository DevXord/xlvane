import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { ERROR_CODES, WARNING_CODES, XlvError, isXlvError } from '../../packages/xlvane/src/errors';
import { suggest } from '../../packages/xlvane/src/util/suggest';

describe('XlvError', () => {
  const error = new XlvError('XLV_E_SHEET_NOT_FOUND', 'Sheet "Q4" does not exist.', {
    location: { sheet: 'Q4' },
    hint: 'Available sheets: Sales, Costs.',
    details: { available: ['Sales', 'Costs'] },
  });

  it('is an Error with a stable code, location, hint and docs link', () => {
    expect(error).toBeInstanceOf(Error);
    expect(error.name).toBe('XlvError');
    expect(error.code).toBe('XLV_E_SHEET_NOT_FOUND');
    expect(error.message).toBe('Sheet "Q4" does not exist.');
    expect(error.location).toEqual({ sheet: 'Q4' });
    expect(error.hint).toBe('Available sheets: Sales, Costs.');
    expect(error.docsUrl).toBe(
      'https://github.com/DevXord/xlvane/blob/main/docs/errors.md#xlv_e_sheet_not_found',
    );
  });

  it('serialises to the same data with toJSON()', () => {
    const json = {
      name: 'XlvError',
      code: 'XLV_E_SHEET_NOT_FOUND',
      message: 'Sheet "Q4" does not exist.',
      location: { sheet: 'Q4' },
      hint: 'Available sheets: Sales, Costs.',
      docsUrl: 'https://github.com/DevXord/xlvane/blob/main/docs/errors.md#xlv_e_sheet_not_found',
      details: { available: ['Sales', 'Costs'] },
    };
    expect(error.toJSON()).toEqual(json);
    expect(JSON.parse(JSON.stringify(error))).toEqual(json);
  });

  it('keeps the original error as cause', () => {
    const cause = new RangeError('inner');
    const wrapped = new XlvError('XLV_E_ZIP_CORRUPT', 'Broken.', { hint: 'Re-save the file.', cause });
    expect(wrapped.cause).toBe(cause);
    expect(wrapped.location).toEqual({});
    expect(wrapped.details).toEqual({});
  });

  it('isXlvError recognises only XlvError instances', () => {
    expect(isXlvError(error)).toBe(true);
    expect(isXlvError(new Error('x'))).toBe(false);
    expect(isXlvError({ code: 'XLV_E_ZIP_CORRUPT' })).toBe(false);
    expect(isXlvError(undefined)).toBe(false);
  });
});

describe('error catalogue', () => {
  it('documents every error and warning code in docs/errors.md', async () => {
    const doc = await readFile(fileURLToPath(new URL('../../docs/errors.md', import.meta.url)), 'utf8');
    const missing = [...ERROR_CODES, ...WARNING_CODES].filter((code) => !doc.includes(`\n### ${code}\n`));
    expect(missing).toEqual([]);
  });

  it('uses unique codes with the documented prefixes', () => {
    expect(new Set(ERROR_CODES).size).toBe(ERROR_CODES.length);
    expect(new Set(WARNING_CODES).size).toBe(WARNING_CODES.length);
    for (const code of ERROR_CODES) expect(code).toMatch(/^XLV_E_[A-Z0-9_]+$/);
    for (const code of WARNING_CODES) expect(code).toMatch(/^XLV_W_[A-Z0-9_]+$/);
  });
});

describe('suggest', () => {
  it('returns close matches, best first', () => {
    expect(suggest('Salse', ['Sales', 'Costs', 'Summary'])).toEqual(['Sales']);
    expect(suggest('maxEntrys', ['maxEntries', 'maxEntryBytes', 'maxCells'])).toEqual([
      'maxEntries',
      'maxEntryBytes',
    ]);
  });

  it('matches case-insensitively and returns nothing when nothing is close', () => {
    expect(suggest('sales', ['Sales'])).toEqual(['Sales']);
    expect(suggest('Revenue', ['Sales', 'Costs'])).toEqual([]);
  });

  it('stays fast on long inputs', () => {
    const start = performance.now();
    suggest('x'.repeat(100_000), ['y'.repeat(100_000), 'Sales']);
    expect(performance.now() - start).toBeLessThan(200);
  });
});
