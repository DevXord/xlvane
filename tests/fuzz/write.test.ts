// Property tests for the writer: any workbook built through the public API is written to a
// package that our strict reader accepts and that reads back with the same content.
import fc from 'fast-check';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { formatCellRef } from '../../packages/xlvane/src/sml/cell-ref';
import { createWorkbook, writeWorkbook, type CellInput } from '../../packages/xlvane/src/write';
import { isXlvError } from '../../packages/xlvane/src/errors';
import type { ChartInput, ChartType } from '../../packages/xlvane/src/charts';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { firstLoneSurrogate } from '../../packages/xlvane/src/xml/writer';
import { NUM_RUNS } from './runs';

const RUNS = { numRuns: Math.max(1, Math.floor(NUM_RUNS / 20)) };

const tricky = fc.constantFrom('_x0041_', '_x005F_', '\u0001', '\r', '\r\n', '\t', ' ', '<', '&', ']]>', '😀', '￿', 'ż');
const text = fc
  .oneof(fc.string({ unit: 'binary', maxLength: 30 }), fc.array(tricky, { maxLength: 8 }).map((a) => a.join('')))
  .filter((s) => firstLoneSurrogate(s) === -1);
const number = fc.double({ noNaN: true, noDefaultInfinity: true });
const date = fc.date({ min: new Date(Date.UTC(1900, 0, 1)), max: new Date(Date.UTC(9999, 11, 31)), noInvalidDate: true });
const value: fc.Arbitrary<CellInput> = fc.oneof(text, number, fc.boolean(), date);

const sheetName = fc
  .string({ minLength: 1, maxLength: 31, unit: fc.constantFrom('a', 'B', ' ', '_', '1', 'ż', '😀', '&', '<', '-') })
  .filter((s) => s.length <= 31);

const cells = fc.array(
  fc.record({ row: fc.integer({ min: 1, max: 60 }), col: fc.integer({ min: 1, max: 30 }), value }),
  { maxLength: 40 },
);

function expected(v: CellInput): unknown {
  if (typeof v === 'number') return Object.is(v, -0) ? 0 : v;
  if (v instanceof Date) {
    // Stored with millisecond precision in a double serial; compare to the millisecond.
    return v.getTime();
  }
  return v;
}

describe('writer properties', () => {
  it('reads back every value written', () => {
    fc.assert(
      fc.property(fc.uniqueArray(sheetName, { minLength: 1, maxLength: 3, selector: (s) => s.toLowerCase() }), cells, (names, list) => {
        const wb = createWorkbook();
        for (const name of names) wb.addSheet(name);
        const sheet = wb.sheets[0]!;
        const last = new Map<string, CellInput>();
        for (const { row, col, value: v } of list) {
          const ref = formatCellRef(row, col);
          // An explicit empty style replaces the cell's style; without it a later number would keep
          // an earlier date's number format (set() keeps the current style, as Excel does).
          sheet.set(ref, v, {});
          last.set(ref, v);
        }
        const read = readWorkbook(writeWorkbook(wb));
        expect(read.sheets.map((s) => s.name)).toEqual(names);
        const rs = read.getSheet(names[0]!);
        for (const [ref, v] of last) {
          const got = rs.cell(ref)?.value;
          if (v instanceof Date) expect(Math.abs((got as Date).getTime() - v.getTime())).toBeLessThanOrEqual(1);
          else expect(got).toEqual(expected(v));
        }
        expect(rs.cellCount).toBe(last.size);
      }),
      RUNS,
    );
  });
});

/** A chart over sheets of the workbook, as the writer accepts it (D-K01). */
const chartInput = (names: readonly string[]): fc.Arbitrary<ChartInput> => {
  const range = fc.record({ sheet: fc.constantFrom(...names), col: fc.integer({ min: 1, max: 30 }), row: fc.integer({ min: 1, max: 60 }), length: fc.integer({ min: 1, max: 20 }), down: fc.boolean() }).map(({ sheet, col, row, length, down }) => {
    const end = down ? formatCellRef(row + length - 1, col) : formatCellRef(row, col + length - 1);
    return `'${sheet.replace(/'/g, "''")}'!${formatCellRef(row, col)}:${end}`;
  });
  const label = fc.string({ minLength: 1, maxLength: 20, unit: fc.constantFrom('a', 'Z', ' ', '_', '&', '<', '"', 'ż', '😀', "'") });
  const series = fc.record({ values: range, categories: range, name: label }, { requiredKeys: ['values'] });
  return fc
    .record({
      type: fc.constantFrom<ChartType>('column', 'bar', 'line', 'pie', 'scatter', 'area'),
      series: fc.array(series, { minLength: 1, maxLength: 5 }),
      title: fc.option(label, { nil: undefined }),
      legend: fc.constantFrom<ChartInput['legend']>('right', 'bottom', 'top', 'left', false, undefined),
      stacked: fc.constantFrom<ChartInput['stacked']>(true, false, 'percent', undefined),
    })
    .map((c) => {
      const pie = c.type === 'pie';
      const chart: ChartInput = { type: c.type, anchor: 'A1:F10', series: pie ? c.series.slice(0, 1) : c.series };
      return { ...chart, ...(c.title === undefined ? {} : { title: c.title }), ...(c.legend === undefined ? {} : { legend: c.legend }), ...(pie || c.type === 'scatter' || c.stacked === undefined ? {} : { stacked: c.stacked }) };
    });
};

describe('chart properties', () => {
  it('reads back the types, titles and references of every chart the writer accepts', () => {
    fc.assert(
      fc.property(
        fc.uniqueArray(sheetName, { minLength: 1, maxLength: 3, selector: (s) => s.toLowerCase() }).chain((names) => fc.tuple(fc.constant(names), fc.array(chartInput(names), { minLength: 1, maxLength: 3 }))),
        ([names, charts]) => {
          const wb = createWorkbook();
          for (const name of names) wb.addSheet(name);
          // Labels have no "!", so no name reads as a sheet-qualified cell reference.
          for (const c of charts) wb.sheets[0]!.addChart(c);
          const read = readWorkbook(writeWorkbook(wb));
          expect(read.warnings).toEqual([]);
          const got = read.sheets[0]!.charts;
          expect(got.map((c) => c.types)).toEqual(charts.map((c) => [c.type]));
          expect(got.map((c) => c.title)).toEqual(charts.map((c) => (c.title === undefined ? null : { text: c.title })));
          expect(got.map((c) => c.series.length)).toEqual(charts.map((c) => c.series.length));
        },
      ),
      RUNS,
    );
  });

  it('reads hostile values in chart parts without throwing anything but XlvError', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 1).addChart({ type: 'line', anchor: 'B2:F9', title: 'T', series: [{ name: 'S!A1', categories: 'S!A1:A2', values: 'S!A1:A2' }], axes: { y: { min: 0, max: 5 } } });
    const files = unzipSync(writeWorkbook(wb));
    const original = strFromU8(files['xl/charts/chart1.xml']!);
    const hostile = fc.constantFrom('', ' ', '__proto__', 'constructor', '1e999', '-1', 'NaN', 'bar', 'percentStacked', 'tr', '<x/>', '&amp;', '_x0000_', '😀', 'r'.repeat(300));
    fc.assert(
      fc.property(fc.array(fc.tuple(fc.nat(), hostile), { maxLength: 6 }), fc.nat(), (edits, drop) => {
        // Replace some attribute values and drop one element (with its content).
        let xml = original;
        const values = [...xml.matchAll(/val="[^"]*"/g)];
        for (const [i, v] of edits) {
          const m = values[i % values.length]!;
          xml = xml.slice(0, m.index) + `val="${v.replace(/[<&"]/g, '')}"` + xml.slice(m.index + m[0].length);
        }
        const tags = [...xml.matchAll(/<c:([a-zA-Z]+)[ >][^]*?<\/c:\1>/g)];
        const t = tags[drop % tags.length];
        if (t !== undefined) xml = xml.slice(0, t.index) + xml.slice(t.index + t[0].length);
        try {
          const charts = readWorkbook(zipSync({ ...files, 'xl/charts/chart1.xml': strToU8(xml) })).getSheet('S').charts;
          expect(Array.isArray(charts)).toBe(true);
        } catch (error) {
          if (!isXlvError(error)) throw error;
        }
      }),
      RUNS,
    );
  });
});
