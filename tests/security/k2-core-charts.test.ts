// Charts found by the 1.3.0 penetration test of the core (D-K01): the text of chart caches written by
// writeWorkbook had no budget (20 000 cells of 32 767 characters threw a plain RangeError, smaller
// sizes wrote chart parts the reader refused), every chart part read built all its series (K parts of
// a million <ser/> built K million objects from a file of a few hundred kilobytes), the sheets of the
// charts were looked up with a map built per chart (quadratic in sheets and charts), and inspect cut
// chart titles inside a surrogate pair.
import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { inspect } from '../../packages/xlvane/src/inspect';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const CHART_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

const rels = (list: readonly (readonly [string, string, string])[]): string =>
  `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${PR}">${list.map(([id, type, target]) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`).join('')}</Relationships>`;

/** A two-cell chart frame that shows the chart of relationship `rid`. */
const frame = (rid: string, name = 'Chart'): string =>
  `<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;

/** One worksheet "S" with one cell and a drawing whose frames show the given chart parts. */
function chartBook(charts: readonly string[]): Uint8Array {
  const files: Record<string, Uint8Array> = {};
  const overrides = [
    ['/xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'],
    ['/xl/worksheets/sheet1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'],
    ['/xl/drawings/drawing1.xml', 'application/vnd.openxmlformats-officedocument.drawing+xml'],
    ...charts.map((_, i) => [`/xl/charts/chart${String(i + 1)}.xml`, CHART_TYPE]),
  ];
  files['[Content_Types].xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.map(([p, t]) => `<Override PartName="${p!}" ContentType="${t!}"/>`).join('')}</Types>`,
  );
  files['_rels/.rels'] = strToU8(rels([['rId1', 'officeDocument', 'xl/workbook.xml']]));
  files['xl/workbook.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rS1"/></sheets></workbook>`);
  files['xl/_rels/workbook.xml.rels'] = strToU8(rels([['rS1', 'worksheet', 'worksheets/sheet1.xml']]));
  files['xl/worksheets/sheet1.xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><drawing r:id="rD"/></worksheet>`,
  );
  files['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(rels([['rD', 'drawing', '../drawings/drawing1.xml']]));
  files['xl/drawings/drawing1.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8"?><xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${R}">${charts.map((_, i) => frame(`rC${String(i)}`)).join('')}</xdr:wsDr>`);
  files['xl/drawings/_rels/drawing1.xml.rels'] = strToU8(rels(charts.map((_, i) => [`rC${String(i)}`, 'chart', `../charts/chart${String(i + 1)}.xml`] as const)));
  charts.forEach((xml, i) => (files[`xl/charts/chart${String(i + 1)}.xml`] = strToU8(xml)));
  return zipSync(files, { level: 9 });
}

const chartPart = (plotArea: string): string => `<?xml version="1.0" encoding="UTF-8"?><c:chartSpace xmlns:c="${C}"><c:chart><c:plotArea>${plotArea}</c:plotArea></c:chart></c:chartSpace>`;
const bars = (series: number): string => `<c:barChart>${'<c:ser/>'.repeat(series)}</c:barChart>`;

describe('writeWorkbook: the text of chart caches has a budget', () => {
  it('refuses cached text of more than 8 Mi characters in a workbook with XLV_E_INVALID_CHART', () => {
    // 300 cells of 32 767 characters: 9.8 million characters of cached text (the penetration test
    // used 20 000 such cells, which threw a RangeError while the string was built).
    const wb = createWorkbook();
    const s = wb.addSheet('D');
    const big = 'x'.repeat(32_767);
    for (let r = 1; r <= 300; r++) s.set(`A${String(r)}`, big).set(`B${String(r)}`, r);
    s.addChart({ type: 'column', anchor: 'D2:J12', series: [{ categories: 'D!A1:A300', values: 'D!B1:B300' }] });
    const error = expectXlvError(() => writeWorkbook(wb), 'XLV_E_INVALID_CHART', { option: 'series', limit: 8 * 1024 * 1024 });
    expect(error.location.sheet).toBe('D');
    expect(error.hint).toContain('writeWorkbookStream');
    expect(error.message).toContain('characters');
  });

  it('counts the text after escaping', () => {
    // 60 cells of 32 767 "&": 1.97 million characters, 9.8 million once written as "&amp;".
    const wb = createWorkbook();
    const s = wb.addSheet('D');
    const amp = '&'.repeat(32_767);
    for (let r = 1; r <= 60; r++) s.set(`A${String(r)}`, amp).set(`B${String(r)}`, r);
    s.addChart({ type: 'column', anchor: 'D2:J12', series: [{ categories: 'D!A1:A60', values: 'D!B1:B60' }] });
    expectXlvError(() => writeWorkbook(wb), 'XLV_E_INVALID_CHART', { limit: 8 * 1024 * 1024 });
  });

  it('writes cached text below the budget, and the streaming writer writes no caches', async () => {
    const wb = createWorkbook();
    const s = wb.addSheet('D');
    const big = 'x'.repeat(32_767);
    for (let r = 1; r <= 200; r++) s.set(`A${String(r)}`, big).set(`B${String(r)}`, r);
    s.addChart({ type: 'column', anchor: 'D2:J12', series: [{ categories: 'D!A1:A200', values: 'D!B1:B200' }] });
    // The repeated text compresses far beyond the reader's default ratio.
    const loose = { limits: { maxCompressionRatio: Number.POSITIVE_INFINITY } };
    const back = readWorkbook(writeWorkbook(wb), loose);
    expect(back.getSheet('D').charts[0]?.series[0]?.categories).toBe('D!$A$1:$A$200');

    const rows = Array.from({ length: 300 }, (_, i) => [big, i + 1]);
    const stream = writeWorkbookStream({ sheets: [{ name: 'D', rows, charts: [{ type: 'column', anchor: 'D2:J12', series: [{ categories: 'D!A1:A300', values: 'D!B1:B300' }] }] }] });
    const chunks: Uint8Array[] = [];
    for await (const c of stream) chunks.push(c);
    const bytes = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
    let at = 0;
    for (const c of chunks) {
      bytes.set(c, at);
      at += c.length;
    }
    expect(readWorkbook(bytes, loose).getSheet('D').charts).toHaveLength(1);
  }, 60_000);
});

describe('reading charts: series and chart groups count against maxCells', () => {
  it('skips the charts read after the budget is spent with XLV_W_LIMIT_REACHED', () => {
    const bytes = chartBook([chartPart(bars(400)), chartPart(bars(400)), chartPart(bars(400))]);
    expect(readWorkbook(bytes).getSheet('S').charts).toHaveLength(3);
    // One cell, then 401 series and groups per chart: the third chart is over 1000.
    const wb = readWorkbook(bytes, { limits: { maxCells: 1000 } });
    expect(wb.getSheet('S').cell('A1')?.value).toBe(1);
    expect(wb.getSheet('S').charts.map((c) => c.series.length)).toEqual([400, 400]);
    expect(wb.warnings.map((w) => [w.code, w.location.part, w.location.sheet])).toEqual([['XLV_W_LIMIT_REACHED', '/xl/charts/chart3.xml', 'S']]);
    expect(wb.warnings[0]!.message).toContain('maxCells');
  });

  it('counts chart groups too', () => {
    const groups = chartPart('<c:lineChart/>'.repeat(600));
    const wb = readWorkbook(chartBook([groups, groups.replace('<c:plotArea>', '<c:plotArea><c:areaChart/>')]), { limits: { maxCells: 1000 } });
    expect(wb.getSheet('S').charts.map((c) => c.types.length)).toEqual([600]);
    expect(wb.warnings.map((w) => w.code)).toEqual(['XLV_W_LIMIT_REACHED']);
  });

  it('does not parse further chart parts once the budget is spent', () => {
    // Ten distinct parts of 300 000 series each (the penetration test: 2 parts of a million series
    // built two million objects, 484 MB of heap). With maxCells 1000 only the cell and nothing of
    // the large parts is kept, and the parts after the first are not parsed.
    const big = chartPart(bars(300_000));
    const bytes = chartBook(Array.from({ length: 10 }, (_, i) => big.replace('<c:plotArea>', `<c:plotArea><!--${String(i)}-->`)));
    const t0 = performance.now();
    const wb = readWorkbook(bytes, { limits: { maxCells: 1000, maxCompressionRatio: Number.POSITIVE_INFINITY } });
    const ms = performance.now() - t0;
    expect(wb.getSheet('S').charts).toEqual([]);
    expect(wb.warnings.map((w) => w.code)).toEqual(Array.from({ length: 10 }, () => 'XLV_W_LIMIT_REACHED'));
    expect(ms).toBeLessThan(5000);
  }, 60_000);

  it('gives charts what the cells of every sheet leave, so a chart never makes cells fail', () => {
    // Before, cells read after a chart would have had less budget: here the chart comes first.
    const wb = createWorkbook();
    wb.addSheet('A').set('A1', 1).addChart({ type: 'column', anchor: 'D2:J12', series: [{ values: 'A!A1:A1' }, { values: 'A!A1:A1' }] });
    const b = wb.addSheet('B');
    for (let r = 1; r <= 9; r++) b.set(`A${String(r)}`, r);
    const bytes = writeWorkbook(wb);
    // 10 cells and 3 series and groups: maxCells 10 reads every cell and skips the chart.
    const read = readWorkbook(bytes, { limits: { maxCells: 10 } });
    expect(read.getSheet('B').cell('A9')?.value).toBe(9);
    expect(read.getSheet('A').charts).toEqual([]);
    expect(read.warnings.map((w) => w.code)).toEqual(['XLV_W_LIMIT_REACHED']);
    expect(readWorkbook(bytes, { limits: { maxCells: 13 } }).getSheet('A').charts).toHaveLength(1);
  });
});

describe('writing charts: sheets are looked up once per workbook', () => {
  const build = (n: number, charts: boolean) => {
    const wb = createWorkbook();
    for (let i = 0; i < n; i++) {
      const s = wb.addSheet(`S${String(i)}x`).set('B1', 1);
      if (charts) s.addChart({ type: 'line', anchor: 'C2:J12', series: [{ values: `S${String(i)}x!B1:B2` }] });
    }
    return wb;
  };
  const time = (fn: () => unknown): number => {
    const t0 = performance.now();
    fn();
    return performance.now() - t0;
  };

  it('writes 3000 sheets with one chart each in time linear in the sheets', async () => {
    // Before: 2000 sheets took 2.1 s with charts and 0.46 s without.
    const n = 3000;
    const plain = build(n, false);
    const charted = build(n, true);
    const plainMs = time(() => writeWorkbook(plain));
    const chartMs = time(() => writeWorkbook(charted));
    expect(chartMs).toBeLessThan(3 * plainMs + 1000);

    const specs = (withCharts: boolean) => Array.from({ length: n }, (_, i) => ({ name: `S${String(i)}x`, rows: [[1]], ...(withCharts ? { charts: [{ type: 'line' as const, anchor: 'C2:J12', series: [{ values: `S${String(i)}x!B1:B2` }] }] } : {}) }));
    const drain = async (withCharts: boolean): Promise<number> => {
      const t0 = performance.now();
      for await (const chunk of writeWorkbookStream({ sheets: specs(withCharts) })) void chunk;
      return performance.now() - t0;
    };
    const streamPlain = await drain(false);
    const streamCharts = await drain(true);
    expect(streamCharts).toBeLessThan(3 * streamPlain + 1000);
  }, 120_000);
});

describe('inspect: titles are cut on a code point boundary', () => {
  it('does not leave half of a surrogate pair at the cut', () => {
    const wb = createWorkbook();
    const title = `${'T'.repeat(99)}${'\u{1F600}'.repeat(10)}`;
    wb.addSheet('D').set('A1', 1).addChart({ type: 'column', anchor: 'D2:J12', title: title.slice(0, 255), series: [{ values: 'D!A1:A1' }] });
    wb.defineName('N', `"${'x'.repeat(98)}\u{1F600}"&"y"`);
    const result = inspect(writeWorkbook(wb));
    const shown = result.sheets[0]!.charts[0]!.title;
    expect(shown).toEqual({ text: `${'T'.repeat(99)}…` });
    expect(result.definedNames[0]!.refersTo).toBe(`"${'x'.repeat(98)}…`);
  });
});
