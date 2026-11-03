// LibreOffice opens the charts xlvane writes (D-K01; docs/facts/charts.md CH-20, CH-33): re-saved
// as XLSX they keep their types, stacking, titles and series references, and drawn as PNG they
// are not blank. Skipped when `soffice` is not installed, unless XLV_REQUIRE_LIBREOFFICE=1 (CI).
import { spawnSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { unzlibSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import type { ChartInput } from '../../packages/xlvane/src/charts';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';

const SOFFICE = process.env['XLV_SOFFICE'] ?? 'soffice';
const available = spawnSync(SOFFICE, ['--version'], { encoding: 'utf8', timeout: 60_000 }).status === 0;
const required = process.env['XLV_REQUIRE_LIBREOFFICE'] === '1';

const S = "'Sales Data'";
const series = (n: number): ChartInput['series'] =>
  ['B', 'C', 'D'].slice(0, n).map((c) => ({ name: `${S}!${c}1`, categories: `${S}!$A$2:$A$7`, values: `${S}!${c}2:${c}7` }));
const CHARTS: Record<string, Omit<ChartInput, 'anchor'>> = {
  column: { type: 'column', title: 'Column', series: series(3), axes: { x: { title: 'Month' }, y: { title: 'Units', min: 0, max: 40 } } },
  columnStacked: { type: 'column', title: 'Column stacked', series: series(3), stacked: true, legend: 'bottom' },
  barPercent: { type: 'bar', title: 'Bar percent', series: series(2), stacked: 'percent', legend: 'top' },
  line: { type: 'line', title: 'Line', series: series(3), legend: 'left' },
  lineStacked: { type: 'line', series: series(2), stacked: true, legend: false },
  pie: { type: 'pie', title: 'Pie', series: series(1) },
  scatter: { type: 'scatter', title: 'Scatter', series: [{ name: 'X and Y', categories: `${S}!F2:F7`, values: `${S}!G2:G7` }] },
  area: { type: 'area', title: 'Area', series: series(2) },
  areaStacked: { type: 'area', series: series(3), stacked: true },
};

/** Data in A1:G7 of "Sales Data" and the charts below it, two side by side. */
function book(charts: Record<string, Omit<ChartInput, 'anchor'>>, withCharts = true): Uint8Array {
  const wb = createWorkbook();
  const data = wb.addSheet('Sales Data').set('A1', 'Month').set('B1', 'North').set('C1', 'South').set('D1', 'West').set('F1', 'X').set('G1', 'Y');
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'].forEach((m, i) => {
    const r = String(i + 2);
    data.set(`A${r}`, m).set(`B${r}`, 10 + i * 3).set(`C${r}`, 25 - i * 2).set(`D${r}`, 12 + ((i * 7) % 9)).set(`F${r}`, i * 1.5).set(`G${r}`, (i * i) / 2 + 1);
  });
  if (withCharts) Object.values(charts).forEach((c, i) => data.addChart({ ...c, anchor: `${'AJ'[i % 2]!}${String(Math.floor(i / 2) * 22 + 10)}:${'HQ'[i % 2]!}${String(Math.floor(i / 2) * 22 + 29)}` }));
  return writeWorkbook(wb);
}

/**
 * Counts the pixels of a PNG whose channels differ by more than 40 (coloured, not white, grey or
 * black), decoding 8-bit RGB or RGBA images without interlacing (PNG filters 0-4).
 */
function colouredPixels(png: Uint8Array): number {
  const view = new DataView(png.buffer, png.byteOffset, png.byteLength);
  let pos = 8;
  let width = 0;
  let height = 0;
  let channels = 0;
  const idat: Uint8Array[] = [];
  while (pos < png.length) {
    const length = view.getUint32(pos);
    const type = String.fromCharCode(...png.subarray(pos + 4, pos + 8));
    const data = png.subarray(pos + 8, pos + 8 + length);
    if (type === 'IHDR') {
      width = view.getUint32(pos + 8);
      height = view.getUint32(pos + 12);
      expect([data[8], data[10], data[11]]).toEqual([8, 0, 0]);
      channels = data[9] === 2 ? 3 : data[9] === 6 ? 4 : 0;
      expect(channels).toBeGreaterThan(0);
    } else if (type === 'IDAT') idat.push(data);
    pos += 12 + length;
  }
  const raw = unzlibSync(new Uint8Array(idat.flatMap((d) => [...d])));
  const stride = width * channels;
  let prev = new Uint8Array(stride);
  let count = 0;
  for (let y = 0; y < height; y++) {
    const filter = raw[y * (stride + 1)]!;
    const line = raw.slice(y * (stride + 1) + 1, (y + 1) * (stride + 1));
    for (let i = 0; i < stride; i++) {
      const a = i >= channels ? line[i - channels]! : 0;
      const b = prev[i]!;
      const c = i >= channels ? prev[i - channels]! : 0;
      const p = a + b - c;
      const pa = Math.abs(p - a);
      const pb = Math.abs(p - b);
      const pc = Math.abs(p - c);
      const predictor = [0, a, b, (a + b) >> 1, pa <= pb && pa <= pc ? a : pb <= pc ? b : c][filter]!;
      line[i] = (line[i]! + predictor) & 0xff;
    }
    for (let x = 0; x < width; x++) {
      const [r, g, bl] = [line[x * channels]!, line[x * channels + 1]!, line[x * channels + 2]!];
      if (Math.max(r, g, bl) - Math.min(r, g, bl) > 40) count++;
    }
    prev = line;
  }
  return count;
}

function convert(dir: string, to: string, files: string[]): void {
  const profile = pathToFileURL(join(dir, 'profile')).href;
  const result = spawnSync(SOFFICE, [`-env:UserInstallation=${profile}`, '--headless', '--convert-to', to, '--outdir', join(dir, 'out'), ...files], { encoding: 'utf8', timeout: 300_000 });
  expect(result.status, result.stderr).toBe(0);
}

describe.skipIf(!available && !required)('LibreOffice and charts', () => {
  it('re-saves every chart type with its types, stacking, titles and series references', { timeout: 400_000 }, async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-lo-charts-'));
    try {
      const file = join(dir, 'charts.xlsx');
      await writeFile(file, book(CHARTS));
      const stream = writeWorkbookStream({ sheets: [{ name: 'Data', rows: [['Jan', 1], ['Feb', 2]], charts: [{ type: 'column', anchor: 'D2:J15', series: [{ categories: 'Data!A1:A2', values: 'Data!B1:B2' }] }] }] });
      const streamed = join(dir, 'streamed.xlsx');
      await writeFile(streamed, new Uint8Array(await new Response(stream).arrayBuffer()));
      convert(dir, 'xlsx:Calc MS Excel 2007 XML', [file, streamed]);
      const back = readWorkbook(await readFile(join(dir, 'out', 'charts.xlsx'))).getSheet('Sales Data').charts;
      const expected = Object.values(CHARTS);
      expect(back.map((c) => c.types)).toEqual(expected.map((c) => [c.type]));
      expect(back.map((c) => c.stacked)).toEqual(expected.map((c) => (c.type === 'pie' || c.type === 'scatter' ? null : (c.stacked ?? false))));
      expect(back.map((c) => (c.title !== null && 'text' in c.title ? c.title.text : null))).toEqual(expected.map((c) => c.title ?? null));
      expect(back.map((c) => c.legend)).toEqual(expected.map((c) => (c.legend === false ? null : (c.legend ?? 'right'))));
      // LibreOffice writes references absolute and keeps the quotes a sheet name needs (CH-18); it
      // writes a literal series name back as a strRef whose f is a quoted string (CH-36).
      const abs = (r: string | undefined): string | null => (r === undefined ? null : r.replace(/\$/g, '').replace(/([A-Z]+)(\d+)/g, '$$$1$$$2'));
      expect(back.map((c) => c.series.map((s) => [s.type, s.name, s.categories, s.values]))).toEqual(
        expected.map((c) => c.series.map((s) => [c.type, s.name === undefined ? null : s.name.includes('!') ? { ref: abs(s.name) } : { ref: `"${s.name}"` }, abs(s.categories), abs(s.values)])),
      );
      expect(back[0]!.axes).toEqual({ x: { title: { text: 'Month' }, min: null, max: null }, y: { title: { text: 'Units' }, min: 0, max: 40 } });
      // The streamed chart has no caches; LibreOffice reads it from the cells (CH-20).
      const s = readWorkbook(await readFile(join(dir, 'out', 'streamed.xlsx'))).getSheet('Data').charts;
      expect(s.map((c) => [c.types, c.series[0]?.values])).toEqual([[['column'], 'Data!$B$1:$B$2']]);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it('draws every chart type: the PNG of each has coloured pixels, the sheet without a chart none', { timeout: 400_000 }, async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-lo-chart-png-'));
    try {
      const files: string[] = [];
      for (const [name, chart] of Object.entries({ ...CHARTS, empty: CHARTS['column']! })) {
        const file = join(dir, `${name}.xlsx`);
        await writeFile(file, book({ [name]: chart }, name !== 'empty'));
        files.push(file);
      }
      convert(dir, 'png', files);
      const coloured = async (name: string): Promise<number> => colouredPixels(await readFile(join(dir, 'out', `${name}.png`)));
      expect(await coloured('empty')).toBe(0);
      // Bars, areas and slices fill thousands of pixels; six scatter markers of size 5 a few hundred.
      for (const name of Object.keys(CHARTS)) expect(await coloured(name), name).toBeGreaterThan(name === 'scatter' ? 200 : 1000);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
