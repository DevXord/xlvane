// Small files whose few cells span a huge rectangle must not make toCSV, readTable, toJSON or
// readTableStream produce a value for every position of it (1.1.1 security review, D-H01): the
// positions written are limited to maxCells or 16 times the cells read, whichever is larger.
import { describe, expect, it } from 'vitest';
import { toCSV } from '../../packages/xlvane/src/csv';
import { readTableStream } from '../../packages/xlvane/src/stream-read';
import { readTable, toJSON } from '../../packages/xlvane/src/table';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx } from '../helpers/xlsx-builder';
import { cli, withFixtures } from '../cli/helpers';
import { writeFile } from 'node:fs/promises';
import { join } from 'node:path';

function timed<T>(fn: () => T): { value: T | Error; ms: number } {
  const start = performance.now();
  let value: T | Error;
  try {
    value = fn();
  } catch (error) {
    value = error as Error;
  }
  return { value, ms: performance.now() - start };
}

/** A1 and one far cell: two cells spanning the whole sheet (1.5 KB). */
const farApart = (far: string): Uint8Array =>
  buildXlsx({ sheets: [{ name: 'S', data: `<row r="1"><c r="A1"><v>1</v></c></row><row r="${far.replace(/^[A-Z]+/, '')}"><c r="${far}"><v>2</v></c></row>` }] });

/** A header in A1 and XFD1, then one value per row in column A. */
function wide(rows: number): Uint8Array {
  let data = '<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c><c r="XFD1" t="inlineStr"><is><t>z</t></is></c></row>';
  for (let r = 2; r <= rows + 1; r++) data += `<row r="${String(r)}"><c r="A${String(r)}"><v>${String(r)}</v></c></row>`;
  return buildXlsx({ sheets: [{ name: 'S', data }] });
}

describe('output positions are limited', () => {
  it('toCSV refuses A1 and XFD1048576 at once', () => {
    const { value, ms } = timed(() => toCSV(farApart('XFD1048576')));
    expect(ms).toBeLessThan(2000);
    expect((value as { code?: string }).code).toBe('XLV_E_LIMIT_EXCEEDED');
    expect((value as { details?: { limit?: string } }).details?.limit).toBe('maxCells');
  });

  it('toCSV still writes a sparse sheet within the budget, and a range of a huge one', () => {
    expect(toCSV(farApart('C3'))).toBe('1,,\r\n,,\r\n,,2\r\n');
    expect(toCSV(farApart('XFD1048576'), { range: 'A1:B2' })).toBe('1,\r\n,\r\n');
  });

  it('readTable and toJSON refuse a header 16 384 columns wide over 1 000 rows of one cell', () => {
    const bytes = wide(1000);
    for (const run of [(): unknown => readTable(bytes), (): unknown => toJSON(bytes)]) {
      const { value, ms } = timed(run);
      expect(ms).toBeLessThan(3000);
      expect((value as { code?: string }).code).toBe('XLV_E_LIMIT_EXCEEDED');
    }
    expectXlvError(() => readTable(bytes), 'XLV_E_LIMIT_EXCEEDED');
    // A narrow range of the same sheet reads normally.
    expect(readTable(bytes, { range: 'A1:B5' }).rows).toHaveLength(4);
  });

  it('readTableStream stops instead of building 16 384 properties per input cell', async () => {
    const bytes = wide(2000);
    const start = performance.now();
    let error: unknown = null;
    let rows = 0;
    try {
      for await (const _ of readTableStream(new Blob([bytes as BlobPart]))) rows++;
    } catch (e) {
      error = e;
    }
    expect(performance.now() - start).toBeLessThan(5000);
    expect((error as { code?: string } | null)?.code).toBe('XLV_E_LIMIT_EXCEEDED');
    expect(rows).toBeLessThan(2000);
  });

  it('readTableStream without a header handles 20 rows of 16 384 cells (no RangeError)', async () => {
    let data = '';
    for (let r = 1; r <= 20; r++) {
      data += `<row r="${String(r)}">`;
      for (let c = 1; c <= 16384; c += 1) data += `<c><v>${String(c)}</v></c>`;
      data += '</row>';
    }
    const bytes = buildXlsx({ sheets: [{ name: 'S', data }] });
    let error: unknown = null;
    try {
      for await (const _ of readTableStream(new Blob([bytes as BlobPart]), { header: false })) break;
    } catch (e) {
      error = e;
    }
    expect(error).toBeNull();
  }, 60_000);

  it('xlvane read ends with the limit error on the wide file', async () => {
    await withFixtures(async (dir) => {
      await writeFile(join(dir, 'wide.xlsx'), wide(1000));
      const start = performance.now();
      const r = await cli(['read', 'wide.xlsx'], { cwd: dir });
      expect(performance.now() - start).toBeLessThan(5000);
      expect(r.code).toBe(3);
      expect(r.stderr).toContain('XLV_E_LIMIT_EXCEEDED');
    });
  });
});

describe('readTableStream and the caller\'s maxCells (independent review of the patch)', () => {
  it('uses limits.maxCells for the output budget, as the hint says', async () => {
    let error: unknown = null;
    try {
      for await (const _ of readTableStream(new Blob([wide(100) as BlobPart]), { limits: { maxCells: 1000 } }));
    } catch (e) {
      error = e;
    }
    expect((error as { code?: string } | null)?.code).toBe('XLV_E_LIMIT_EXCEEDED');
  });
});

describe('repeated header names (independent review of the patch)', () => {
  // One row of 16 384 cells all holding "x", and one data row: naming the columns x, x_2, …
  // searched from x_2 again for every column (quadratic; 9.9 s before).
  const data = `<row r="1">${'<c t="inlineStr"><is><t>x</t></is></c>'.repeat(16384)}</row><row r="2"><c r="A2"><v>1</v></c></row>`;
  const bytes = buildXlsx({ sheets: [{ name: 'S', data }] });

  it('readTable names 16 384 columns called "x" in linear time, with the same names', () => {
    const start = performance.now();
    const table = readTable(bytes);
    expect(performance.now() - start).toBeLessThan(2000);
    const keys = Object.keys(table.rows[0] ?? {});
    expect(keys.slice(0, 3)).toEqual(['x', 'x_2', 'x_3']);
    expect(keys[16383]).toBe('x_16384');
  });

  it('readTableStream does too', async () => {
    const start = performance.now();
    let keys: string[] = [];
    for await (const item of readTableStream(new Blob([bytes as BlobPart]))) keys = Object.keys(item.value ?? {});
    expect(performance.now() - start).toBeLessThan(2000);
    expect(keys[16383]).toBe('x_16384');
  });

  it('a header that is itself "x_2" keeps the earlier naming', () => {
    const t = readTable(buildXlsx({ sheets: [{ name: 'S', data: '<row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c><c r="B1" t="inlineStr"><is><t>x_2</t></is></c><c r="C1" t="inlineStr"><is><t>x</t></is></c><c r="D1" t="inlineStr"><is><t>x</t></is></c></row><row r="2"><c r="A2"><v>1</v></c></row>' }] }));
    expect(Object.keys(t.rows[0] ?? {})).toEqual(['x', 'x_2', 'x_3', 'x_4']);
  });
});
