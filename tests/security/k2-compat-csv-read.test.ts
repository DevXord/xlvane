// @xlvane/compat csv.read after the K2 penetration test: every row of the input was collected
// before addRow refused row 1 048 577 (20 MB of line feeds took 2 GB), and quoted fields were
// built one character at a time (an 80 MB quoted field ran out of memory). Parsing now stops at
// the sheet's 1 048 576 rows and 16 384 columns, rows are added as they are parsed, and quoted
// fields are built from slices of the input.
import { describe, expect, it } from 'vitest';
import ExcelJS from '../../packages/compat/src/index';

const enc = new TextEncoder();
const MAX_ROWS = 1_048_576;
const MAX_COLS = 16_384;

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

describe('K2 item 6: csv.read limits', () => {
  it('stops at row 1 048 577 before parsing the rest of the input', async () => {
    // The text after the limit is malformed: it was reached before (XLV_E_CSV_MALFORMED), now
    // parsing stops at the row the sheet cannot hold.
    const { error } = await read(`${'1\n'.repeat(MAX_ROWS + 1)}"never closed`);
    expect(error).toMatchObject({ code: 'XLV_E_LIMIT_EXCEEDED', details: { limit: 'maxRows', value: MAX_ROWS } });
    expect((error as { hint?: string }).hint).toMatch(/split/i);
    expect((error as { location?: { line?: number } }).location?.line).toBe(MAX_ROWS + 1);
  }, 60_000);

  it('accepts exactly 1 048 576 rows', async () => {
    const { error, wb } = await read('\n'.repeat(MAX_ROWS));
    expect(error).toBeNull();
    expect(wb.worksheets[0]!.rowCount).toBe(MAX_ROWS);
  }, 60_000);

  it('refuses a row of more than 16 384 fields without collecting them', async () => {
    const { error, ms } = await read(`a${',x'.repeat(4_000_000)}`);
    expect(error).toMatchObject({ code: 'XLV_E_LIMIT_EXCEEDED', details: { limit: 'maxColumns', value: MAX_COLS } });
    expect((error as { location?: { line?: number } }).location?.line).toBe(1);
    expect(ms).toBeLessThan(8000);
  }, 30_000);

  it('accepts a row of exactly 16 384 fields', async () => {
    const { error, wb } = await read(`a${',x'.repeat(MAX_COLS - 1)}`);
    expect(error).toBeNull();
    expect(wb.worksheets[0]!.getRow(1).cellCount).toBe(MAX_COLS);
  });

  it('counts only the rows that are added: skipped or empty rows and the header do not count', async () => {
    // 3 million empty lines with ignoreEmpty, then a header that is skipped and two data rows.
    const wide = `h${',h'.repeat(MAX_COLS)}`;
    const { error, ms, wb } = await read(`${'\n'.repeat(3_000_000)}${wide}\n1,2\n3,4\n`, { parserOptions: { ignoreEmpty: true, headers: true } });
    expect(error).toBeNull();
    expect(wb.worksheets[0]!.getSheetValues().filter(Boolean)).toHaveLength(2);
    expect(ms).toBeLessThan(8000);
  }, 60_000);

  it('builds a long quoted field in linear time', async () => {
    const n = 30_000_000;
    const { error, ms, wb } = await read(`"${'a'.repeat(n / 2)}""${'b'.repeat(n / 2)}"`);
    expect(error).toBeNull();
    const value = wb.worksheets[0]!.getCell('A1').value as string;
    expect(value.length).toBe(n + 1);
    expect(value[n / 2]).toBe('"');
    expect(ms).toBeLessThan(2000);
  }, 60_000);

  it('keeps fast-csv parsing: quotes, escapes, skipRows, maxRows, comments, trim', async () => {
    const text = '# note\nh1,h2\n"a ""q"" b", c \n"x\ny",z\n1,2\n3,4\n';
    const { wb } = await read(text, { parserOptions: { comment: '#', headers: true, skipRows: 1, maxRows: 2, trim: true } });
    expect(wb.worksheets[0]!.getSheetValues().slice(1).map((r) => (r as unknown[]).slice(1))).toEqual([['x\ny', 'z'], [1, 2]]);
    const escaped = await read('"a\\"b",c', { parserOptions: { escape: '\\' } });
    expect(escaped.wb.worksheets[0]!.getCell('A1').value).toBe('a"b');
    const open = await read('a\n"b\nc');
    expect(open.error).toMatchObject({ code: 'XLV_E_CSV_MALFORMED', location: { line: 2 } });
  });
});
