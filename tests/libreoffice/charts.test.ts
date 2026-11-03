// LibreOffice opens the charts xlvane writes (D-K01; docs/facts/charts.md CH-20, CH-33): re-saved
// as XLSX they keep their types, stacking, titles and series references, and drawn as PNG they
// are not blank. Skipped when `soffice` is not installed, unless XLV_REQUIRE_LIBREOFFICE=1 (CI).
import { spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, join } from 'node:path';
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
  const run = (list: string[]): string => {
    const result = spawnSync(SOFFICE, [`-env:UserInstallation=${profile}`, '--headless', '--convert-to', to, '--outdir', join(dir, 'out'), ...list], { encoding: 'utf8', timeout: 300_000 });
    expect(result.status, result.stderr).toBe(0);
    return `${result.stdout}${result.stderr}`;
  };
  run(files);
  // In CI a conversion has finished with status 0 but without the output of a file; such files are
  // converted again one at a time (up to three times), and a file still without output fails the
  // test with what LibreOffice printed.
  const extension = to.split(':')[0]!;
  const output = (file: string): string => join(dir, 'out', `${basename(file).replace(/\.[^.]+$/, '')}.${extension}`);
  for (const file of files) {
    let printed = '';
    for (let attempt = 0; attempt < 3 && !existsSync(output(file)); attempt++) printed = run([file]);
    expect(existsSync(output(file)), `LibreOffice wrote no ${extension} for ${basename(file)}:\n${printed}`).toBe(true);
  }
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

  it('writes again a chart that LibreOffice re-saved with an unquoted cell-like sheet name (D-K01)', { timeout: 400_000 }, async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-lo-chart-s1-'));
    try {
      const wb = createWorkbook();
      wb.addSheet('S1').set('A1', 'n').set('A2', 1).set('A3', 2).addChart({ type: 'column', anchor: 'C2:H12', series: [{ name: "'S1'!A1", values: "'S1'!A2:A3" }] });
      const file = join(dir, 's1.xlsx');
      await writeFile(file, writeWorkbook(wb));
      convert(dir, 'xlsx:Calc MS Excel 2007 XML', [file]);
      const [chart] = readWorkbook(await readFile(join(dir, 'out', 's1.xlsx'))).getSheet('S1').charts;
      const values = chart!.series[0]!.values!;
      const name = chart!.series[0]!.name;
      expect(values.replace(/'/g, '')).toBe('S1!$A$2:$A$3');
      const again = createWorkbook();
      again.addSheet('S1').set('A1', 'n').set('A2', 1).set('A3', 2).addChart({ type: 'column', anchor: 'C2:H12', series: [{ ...(name !== null && 'ref' in name ? { name: name.ref } : {}), values }] });
      expect(readWorkbook(writeWorkbook(again)).getSheet('S1').charts[0]?.series[0]?.values).toBe("'S1'!$A$2:$A$3");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it('draws every chart type: the PNG of each has coloured pixels, the sheet without a chart none', { timeout: 400_000 }, async (ctx) => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-lo-chart-png-'));
    try {
      // The CI image installs libreoffice-calc only, whose LibreOffice has no PNG export filter
      // ("Error: no export filter"); the drawing is then not checked there (the re-save test above
      // still checks every chart), but it is wherever the filter exists.
      const probe = join(dir, 'probe.xlsx');
      await writeFile(probe, book({}, false));
      const profile = pathToFileURL(join(dir, 'profile')).href;
      const tried = spawnSync(SOFFICE, [`-env:UserInstallation=${profile}`, '--headless', '--convert-to', 'png', '--outdir', join(dir, 'probe'), probe], { encoding: 'utf8', timeout: 300_000 });
      if (/no export filter/i.test(`${tried.stdout}${tried.stderr}`)) ctx.skip();
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
