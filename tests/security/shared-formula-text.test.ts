// A small file can share one huge formula with thousands of cells: 13 KB with a 1 MiB master
// formula and 2 000 dependent cells became 2 Gi characters (about 20 s and 2 GB once the texts
// were read). The expanded text of shared formulas is limited by maxSharedStringChars, counted
// apart from shared strings: the master's length, once per cell that shares it (1.1.1 security
// review, D-H01).
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { openWorkbookStream } from '../../packages/xlvane/src/stream-read';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx } from '../helpers/xlsx-builder';

function sharedFormula(master: string, cells: number): Uint8Array {
  let data = `<row r="1"><c r="B1"><f t="shared" ref="B1:B${String(cells)}" si="0">${master}</f><v>1</v></c></row>`;
  for (let r = 2; r <= cells; r++) data += `<row r="${String(r)}"><c r="B${String(r)}"><f t="shared" si="0"/><v>1</v></c></row>`;
  return buildXlsx({ sheets: [{ name: 'S', data }] });
}

describe('expanded shared formulas', () => {
  const huge = `A1+${'1+'.repeat(1 << 19)}1`;

  it('readWorkbook refuses 2 000 cells sharing a 1 MiB formula, quickly', () => {
    const bytes = sharedFormula(huge, 2000);
    const start = performance.now();
    const error = expectXlvError(() => readWorkbook(bytes), 'XLV_E_LIMIT_EXCEEDED');
    expect(performance.now() - start).toBeLessThan(3000);
    expect(error.details?.limit).toBe('maxSharedStringChars');
    expect(error.location.sheet).toBe('S');
  }, 30_000);

  it('openWorkbookStream refuses it too', async () => {
    const book = await openWorkbookStream(sharedFormula(huge, 2000));
    let error: unknown = null;
    try {
      for await (const _ of book.rows('S'));
    } catch (e) {
      error = e;
    }
    expect((error as { code?: string }).code).toBe('XLV_E_LIMIT_EXCEEDED');
  }, 30_000);

  it('counts the length of each moved formula, apart from shared strings', () => {
    const limits = { maxSharedStringChars: 100 };
    // 9 dependents of a 10-character formula, moved to A2…A10: 8 × 10 + 11 = 91 characters, within the limit.
    const ok = readWorkbook(sharedFormula('A1+1000000', 10), { limits });
    expect([...(ok.sheets[0]?.cells() ?? [])].map((c) => c.formula?.text)).toContain('A10+1000000');
    // 10 dependents: 8 × 10 + 2 × 11 = 102 characters.
    expectXlvError(() => readWorkbook(sharedFormula('A1+1000000', 11), { limits }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxSharedStringChars' });
    expectXlvError(() => readWorkbook(sharedFormula('A1+1000000', 12), { limits }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxSharedStringChars' });
  });
});

describe('the budget counts the moved text (independent review of the patch)', () => {
  // "1:1" moved a million rows down is "1048567:1048567": counting the master's length let
  // the expanded text reach about four times the limit.
  const master = `SUM(${Array.from({ length: 2000 }, () => '1:1').join(',')})`;
  function far(first: number, count: number): Uint8Array {
    let data = `<row r="1"><c r="A1"><f t="shared" ref="A1:A1048576" si="0">${master}</f><v>1</v></c></row>`;
    for (let r = first; r < first + count; r++) data += `<row r="${String(r)}"><c r="A${String(r)}"><f t="shared" si="0"/><v>1</v></c></row>`;
    return buildXlsx({ sheets: [{ name: 'S', data }] });
  }
  const limits = { maxSharedStringChars: 10 * master.length };

  it('refuses ten cells far down whose moved text is four times the master', () => {
    expectXlvError(() => readWorkbook(far(1048567, 10), { limits }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxSharedStringChars' });
  });

  it('accepts eight cells near the master, whose moved text is as long as the master', () => {
    const wb = readWorkbook(far(2, 8), { limits });
    expect([...(wb.sheets[0]?.cells() ?? [])][8]?.formula?.text?.length).toBe(master.length);
  });
});
