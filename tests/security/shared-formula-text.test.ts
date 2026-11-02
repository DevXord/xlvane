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

  it('counts the master length once per dependent cell, apart from shared strings', () => {
    const limits = { maxSharedStringChars: 100 };
    // 10 dependents of a 10-character formula: 100 characters, within the limit.
    const ok = readWorkbook(sharedFormula('A1+1000000', 11), { limits });
    expect([...(ok.sheets[0]?.cells() ?? [])].map((c) => c.formula?.text)).toContain('A11+1000000');
    expectXlvError(() => readWorkbook(sharedFormula('A1+1000000', 12), { limits }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxSharedStringChars' });
  });
});
