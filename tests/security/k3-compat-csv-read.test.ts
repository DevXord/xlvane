// @xlvane/compat csv.read after the K3 review of the K2 fixes: csv.read had limits on rows and on
// fields per row, but none on the fields in total, so 10 MB of "1,1,…" (5 000 000 cells) took
// 885 MB. At most DEFAULT_READ_LIMITS.maxCells (5 000 000) fields are added to the worksheet;
// more throw XLV_E_LIMIT_EXCEEDED (details.limit "maxCells"), found before any row is added.
import { DEFAULT_READ_LIMITS } from 'xlvane';
import { describe, expect, it } from 'vitest';
import ExcelJS from '../../packages/compat/src/index';

const enc = new TextEncoder();

async function read(text: string, options: Record<string, unknown> = {}): Promise<{ error: unknown; ms: number; wb: InstanceType<typeof ExcelJS.Workbook> }> {
  const wb = new ExcelJS.Workbook();
  const start = performance.now();
  let error: unknown = null;
  try {
    await wb.csv.read([enc.encode(text)], options as never);
  } catch (e) {
    error = e;
  }
  return { error, ms: performance.now() - start, wb };
}

/** 1 000 rows of 5 000 fields and one more row with `last` fields: 5 000 000 + `last` fields. */
function wide(last: number): string {
  return `${`1${',1'.repeat(4999)}\n`.repeat(1000)}${Array.from({ length: last }, () => '1').join(',')}`;
}

describe('K3 item 5: a total cell budget for csv.read', () => {
  it('refuses more than 5 000 000 fields before adding any row', async () => {
    expect(DEFAULT_READ_LIMITS.maxCells).toBe(5_000_000);
    const { error, ms, wb } = await read(wide(1));
    expect(error).toMatchObject({ code: 'XLV_E_LIMIT_EXCEEDED', details: { limit: 'maxCells', value: 5_000_000, line: 1001 }, location: { line: 1001 } });
    expect((error as { hint?: string }).hint).toMatch(/split/i);
    expect(wb.worksheets[0]!.rowCount).toBe(0);
    expect(ms).toBeLessThan(5000);
  }, 60_000);

  it('counts only the fields of rows that are added', async () => {
    // skipRows and maxRows leave all but a few rows out.
    const skipped = await read(wide(3), { parserOptions: { skipRows: 1000 } });
    expect(skipped.error).toBeNull();
    expect(skipped.wb.worksheets[0]!.getRow(1).values).toEqual([undefined, 1, 1, 1]);
    const first = await read(wide(1), { parserOptions: { maxRows: 2 } });
    expect(first.error).toBeNull();
    expect(first.wb.worksheets[0]!.rowCount).toBe(2);
  }, 60_000);

  it('reads small files as before', async () => {
    const { error, wb } = await read('a,b\n1,2\n');
    expect(error).toBeNull();
    expect(wb.worksheets[0]!.getSheetValues().slice(1).map((r) => (r as unknown[]).slice(1))).toEqual([['a', 'b'], [1, 2]]);
  });
});
