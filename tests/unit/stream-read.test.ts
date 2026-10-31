import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { openWorkbookStream, readTableStream, type StreamRow } from '../../packages/xlvane/src/stream-read';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { z } from 'zod';
import { expectXlvErrorAsync } from '../helpers/expect-error';
import { buildXlsx, row } from '../helpers/xlsx-builder';

function sample(): Uint8Array {
  const wb = createWorkbook();
  const s = wb.addSheet('Data');
  s.set('A1', 'name').set('B1', 'total').set('C1', 'when');
  for (let r = 2; r <= 200; r++) s.set(`A${String(r)}`, `n${String(r)}`).set(`B${String(r)}`, r * 1.5).set(`C${String(r)}`, new Date(Date.UTC(2024, 0, (r % 28) + 1)));
  s.set('B201', { formula: 'SUM(B2:B200)', result: 1 }).row(3, { height: 22 }).row(5, { hidden: true });
  s.merge('A202:C202').freeze({ rows: 1 });
  wb.addSheet('Other', { state: 'hidden' }).set('A1', 'x');
  return writeWorkbook(wb);
}

async function all(iter: AsyncIterable<StreamRow>): Promise<StreamRow[]> {
  const out: StreamRow[] = [];
  for await (const r of iter) out.push(r);
  return out;
}

describe('openWorkbookStream', () => {
  it('lists sheets and streams the same cells as readWorkbook', async () => {
    const bytes = sample();
    const book = await openWorkbookStream(bytes);
    expect(book.sheets.map((s) => [s.name, s.state, s.kind])).toEqual([
      ['Data', 'visible', 'worksheet'],
      ['Other', 'hidden', 'worksheet'],
    ]);
    expect(book.dateSystem).toBe('1900');
    const rows = await all(book.rows('Data'));
    const batch = readWorkbook(bytes).getSheet('Data');
    expect(rows.flatMap((r) => r.cells)).toEqual([...batch.cells()]);
    expect(rows.map((r) => r.number).slice(0, 3)).toEqual([1, 2, 3]);
    expect(rows[2]).toMatchObject({ number: 3, height: 22, hidden: false });
    expect(rows[4]).toMatchObject({ number: 5, hidden: true });
    expect(rows[0]?.height).toBeNull();
  });

  it('reads from a Blob and an ArrayBuffer', async () => {
    const bytes = sample();
    for (const source of [new Blob([bytes as BlobPart]), bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer]) {
      const book = await openWorkbookStream(source);
      expect((await all(book.rows('Other')))[0]?.cells[0]?.value).toBe('x');
    }
  });

  it('can be read as a ReadableStream, and sheets can be read again', async () => {
    const book = await openWorkbookStream(sample());
    const reader = book.rowStream('Data').getReader();
    const first = await reader.read();
    expect(first.value?.cells.map((c) => c.value)).toEqual(['name', 'total', 'when']);
    await reader.cancel();
    expect((await all(book.rows('Data'))).length).toBe(201);
  });

  it('expands shared formulas whose master comes first', async () => {
    const sheet = row(1, '<c r="A1"><f t="shared" ref="A1:A3" si="0">B1*2</f><v>2</v></c>') + row(2, '<c r="A2"><f t="shared" si="0"/><v>4</v></c>');
    const book = await openWorkbookStream(buildXlsx({ sheets: [{ name: 'S', data: sheet }] }));
    const rows = await all(book.rows('S'));
    expect(rows[1]?.cells[0]?.formula?.text).toBe('B2*2');
    expect(book.warnings).toEqual([]);
  });

  it('applies limits other than maxCells', async () => {
    const bytes = sample();
    const book = await openWorkbookStream(bytes, { limits: { maxCells: 1 } });
    expect((await all(book.rows('Data'))).length).toBe(201);
    await expectXlvErrorAsync(async () => all((await openWorkbookStream(bytes, { limits: { maxXmlTextLength: 2 } })).rows('Data')), 'XLV_E_LIMIT_EXCEEDED');
  });

  it('reports a missing sheet and bad input', async () => {
    const book = await openWorkbookStream(sample());
    expect(() => book.rows('data')).toThrow(/no sheet named "data"/);
    await expectXlvErrorAsync(() => openWorkbookStream(new Uint8Array([1, 2, 3])), 'XLV_E_NOT_ZIP');
    await expectXlvErrorAsync(() => openWorkbookStream('file.xlsx' as never), 'XLV_E_INVALID_INPUT');
  });
});

describe('readTableStream', () => {
  it('yields row objects one by one, with schema validation', async () => {
    const schema = z.object({ name: z.string(), total: z.number().max(100) });
    const valid: unknown[] = [];
    const issues: string[] = [];
    for await (const item of readTableStream(sample(), { sheet: 'Data', schema })) {
      if (item.issues.length === 0) valid.push(item.value);
      else issues.push(...item.issues.map((i) => i.cell ?? `row ${String(i.row)}`));
    }
    expect(valid[0]).toEqual({ name: 'n2', total: 3 });
    expect(valid).toHaveLength(65); // total = r * 1.5 <= 100 for r = 2 … 66; row 201 has no name
    expect(issues).toContain('B67');
  });

  it('gives the same rows as readTable without a schema', async () => {
    const { readTable } = await import('../../packages/xlvane/src/table');
    const rows: unknown[] = [];
    for await (const item of readTableStream(sample())) rows.push(item.value);
    expect(rows).toEqual(readTable(sample()).rows);
  });
});
