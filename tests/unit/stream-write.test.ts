import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { writeTableStream, writeWorkbookStream, type StreamRowInput } from '../../packages/xlvane/src/stream-write';
import { readTable } from '../../packages/xlvane/src/table';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvErrorAsync } from '../helpers/expect-error';

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  const reader = stream.getReader();
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
  }
  const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
  let at = 0;
  for (const c of chunks) {
    out.set(c, at);
    at += c.length;
  }
  return out;
}

function* rows(n: number): Generator<StreamRowInput> {
  yield { values: ['id', 'name', 'when', 'ok'], style: { font: { bold: true } } };
  for (let i = 1; i <= n; i++) yield [i, `name ${String(i)}`, new Date(Date.UTC(2024, 0, (i % 28) + 1)), i % 2 === 0];
}

describe('writeWorkbookStream', () => {
  it('writes rows pulled from an iterable, readable by the batch reader', async () => {
    const bytes = await collect(
      writeWorkbookStream({
        sheets: [
          { name: 'Data', rows: rows(1000), columns: [{ width: 8 }, { width: 20 }], freeze: { rows: 1 }, autoFilter: true },
          { name: 'Notes', state: 'hidden', rows: [['x', { formula: 'Data!A2*2', result: 2 }, { error: '#N/A' }]] },
        ],
      }),
    );
    const wb = readWorkbook(bytes);
    const s = wb.getSheet('Data');
    expect(s.cellCount).toBe(4 * 1001);
    expect(s.cell('B501')?.value).toBe('name 500');
    expect(s.cell('C2')?.value).toEqual(new Date(Date.UTC(2024, 0, 2)));
    expect(s.cell('C2')?.numFmt).toBe('yyyy-mm-dd');
    expect(s.cell('A1')?.style.font?.bold).toBe(true);
    expect(s.freeze).toEqual({ rows: 1, cols: 0 });
    expect(s.autoFilter).toBe('A1:D1001');
    expect(s.columns[1]).toMatchObject({ min: 2, width: 20 });
    const notes = wb.getSheet('Notes');
    expect(notes.state).toBe('hidden');
    expect(notes.cell('B1')?.formula?.text).toBe('Data!A2*2');
    expect(notes.cell('C1')?.value).toBe('#N/A');
    expect(wb.warnings).toEqual([]);
  });

  it('writes inline strings by default and a shared string table on request', async () => {
    const inline = unzipSync(await collect(writeWorkbookStream({ sheets: [{ name: 'S', rows: [['a', 'a']] }] })));
    expect(Object.keys(inline)).not.toContain('xl/sharedStrings.xml');
    expect(new TextDecoder().decode(inline['xl/worksheets/sheet1.xml'])).toContain('t="inlineStr"');
    const shared = await collect(writeWorkbookStream({ sharedStrings: true, sheets: [{ name: 'S', rows: [['a', 'a']] }] }));
    expect(new TextDecoder().decode(unzipSync(shared)['xl/sharedStrings.xml'])).toContain('uniqueCount="1"');
    expect(readWorkbook(shared).getSheet('S').cell('B1')?.value).toBe('a');
  });

  it('accepts async iterables, row heights and per-cell styles', async () => {
    async function* gen(): AsyncGenerator<StreamRowInput> {
      yield { values: [1, 2], height: 30, hidden: false, styles: [{ numFmt: '0.00' }, undefined] };
      await Promise.resolve();
      yield { values: [], hidden: true };
      yield [3];
    }
    const s = readWorkbook(await collect(writeWorkbookStream({ dateSystem: '1904', sheets: [{ name: 'S', rows: gen() }] }))).getSheet('S');
    expect(s.cell('A1')?.numFmt).toBe('0.00');
    expect(s.cell('B1')?.numFmt).toBeNull();
    expect(s.rowInfo(1)).toEqual({ height: 30, hidden: false });
    expect(s.rowInfo(2)).toEqual({ height: null, hidden: true });
    expect(s.cell('A3')?.value).toBe(3);
  });

  it('pulls rows only as the consumer reads (backpressure)', async () => {
    let produced = 0;
    function* lazy(): Generator<StreamRowInput> {
      for (let i = 0; i < 100_000; i++) {
        produced++;
        yield [i, 'x'.repeat(50)];
      }
    }
    const reader = writeWorkbookStream({ sheets: [{ name: 'S', rows: lazy() }] }).getReader();
    await reader.read();
    await reader.read();
    expect(produced).toBeLessThan(100_000);
    await reader.cancel();
  });

  it('gives the same workbook as the batch writer', async () => {
    const batch = createWorkbook();
    const s = batch.addSheet('S');
    s.set('A1', 'h').set('B1', 2.5).set('A2', true).set('B2', new Date(Date.UTC(2020, 5, 1, 12)));
    const streamed = await collect(
      writeWorkbookStream({
        sharedStrings: true,
        sheets: [{ name: 'S', rows: [['h', 2.5], [true, new Date(Date.UTC(2020, 5, 1, 12))]] }],
      }),
    );
    expect(readTable(streamed, { header: false }).rows).toEqual(readTable(writeWorkbook(batch), { header: false }).rows);
  });

  it('reports invalid values and specs as stream errors with the cell', async () => {
    const error = await expectXlvErrorAsync(() => collect(writeWorkbookStream({ sheets: [{ name: 'S', rows: [[1], [Number.NaN]] }] })), 'XLV_E_INVALID_VALUE');
    expect(error.location).toMatchObject({ sheet: 'S', cell: 'A2' });
    expect(() => writeWorkbookStream({ sheets: [] })).toThrow(/no sheets/);
    expect(() => writeWorkbookStream({ sheets: [{ name: 'a/b', rows: [] }] })).toThrow();
    expect(() => writeWorkbookStream({ sheets: [{ name: 'S', rows: [] }], bogus: 1 } as never)).toThrow(/bogus/);
  });
});

describe('writeTableStream', () => {
  it('writes objects as a report table', async () => {
    async function* objects(): AsyncGenerator<Record<string, unknown>> {
      for (let i = 1; i <= 3; i++) yield { id: i, name: `n${String(i)}`, extra: i === 2 ? 'x' : undefined };
    }
    const bytes = await collect(writeTableStream(objects(), { columns: ['id', 'name', 'extra'], sheet: 'T' }));
    const t = readTable(bytes);
    expect(t.rows).toEqual([
      { id: 1, name: 'n1', extra: null },
      { id: 2, name: 'n2', extra: 'x' },
      { id: 3, name: 'n3', extra: null },
    ]);
    const s = readWorkbook(bytes).getSheet('T');
    expect(s.freeze).toEqual({ rows: 1, cols: 0 });
    expect(s.autoFilter).toBe('A1:C4');
    expect(s.cell('A1')?.style.font?.bold).toBe(true);
  });

  it('needs the columns up front', () => {
    expect(() => writeTableStream([] as never[], {} as never)).toThrow(/columns/);
  });
});
