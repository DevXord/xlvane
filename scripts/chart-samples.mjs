// Writes one sample workbook per chart type and variant (D-K01) into the directory given as the
// first argument, for opening in Excel or another spreadsheet application.
// Usage: npm run build -w packages/xlvane && node scripts/chart-samples.mjs <directory>
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { createWorkbook, writeWorkbook, writeWorkbookStream } from '../packages/xlvane/dist/index.js';

const dir = process.argv[2];
if (!dir) {
  console.error('Usage: node scripts/chart-samples.mjs <directory>');
  process.exit(1);
}
mkdirSync(dir, { recursive: true });

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
/** Sheet "Sales Data": months in A, three series in B:D with names in row 1; x/y pairs in F:G. */
function dataRows() {
  const rows = [['Month', 'North', 'South', 'West', '', 'X', 'Y']];
  MONTHS.forEach((m, i) => rows.push([m, 10 + i * 3, 25 - i * 2, 12 + ((i * 7) % 9), null, i * 1.5, (i * i) / 2 + 1]));
  return rows;
}

const S = "'Sales Data'";
const three = [
  { name: `${S}!B1`, categories: `${S}!A2:A7`, values: `${S}!B2:B7` },
  { name: `${S}!C1`, categories: `${S}!A2:A7`, values: `${S}!C2:C7` },
  { name: `${S}!D1`, categories: `${S}!A2:A7`, values: `${S}!D2:D7` },
];
const samples = {
  'column-clustered': { type: 'column', title: 'Column, clustered', series: three, axes: { x: { title: 'Month' }, y: { title: 'Units' } } },
  'column-stacked': { type: 'column', title: 'Column, stacked', series: three, stacked: true },
  'column-percent': { type: 'column', title: 'Column, 100 % stacked', series: three, stacked: 'percent' },
  'bar-clustered': { type: 'bar', title: 'Bar, clustered', series: three, legend: 'bottom', axes: { x: { title: 'Month' }, y: { title: 'Units' } } },
  'bar-stacked': { type: 'bar', title: 'Bar, stacked', series: three, stacked: true },
  'bar-percent': { type: 'bar', title: 'Bar, 100 % stacked', series: three, stacked: 'percent' },
  'line-standard': { type: 'line', title: 'Line', series: three, legend: 'top', axes: { y: { title: 'Units', min: 0, max: 40 } } },
  'line-stacked': { type: 'line', title: 'Line, stacked', series: three, stacked: true },
  'line-percent': { type: 'line', title: 'Line, 100 % stacked', series: three, stacked: 'percent' },
  pie: { type: 'pie', title: 'Pie', series: [three[0]], legend: 'right' },
  scatter: { type: 'scatter', title: 'Scatter', series: [{ name: 'Y values', categories: `${S}!F2:F7`, values: `${S}!G2:G7`, color: 'C00000' }], axes: { x: { title: 'X' }, y: { title: 'Y' } } },
  'area-standard': { type: 'area', title: 'Area', series: three.slice(0, 2), legend: 'left' },
  'area-stacked': { type: 'area', title: 'Area, stacked', series: three, stacked: true },
  'area-percent': { type: 'area', title: 'Area, 100 % stacked', series: three, stacked: 'percent', legend: false },
};

for (const [name, chart] of Object.entries(samples)) {
  const wb = createWorkbook();
  const sheet = wb.addSheet('Sales Data');
  dataRows().forEach((row, r) => row.forEach((v, c) => v !== null && v !== '' && sheet.set(`${'ABCDEFG'[c]}${String(r + 1)}`, v)));
  sheet.addChart({ anchor: 'I2:P20', ...chart });
  writeFileSync(join(dir, `${name}.xlsx`), writeWorkbook(wb));
}

// The same column chart from the streaming writer, which writes no caches (PROVISIONAL, R6-4).
const stream = writeWorkbookStream({ sheets: [{ name: 'Sales Data', rows: dataRows().map((row) => row.map((v) => (v === '' ? null : v))), charts: [{ anchor: 'I2:P20', ...samples['column-clustered'], title: 'Column, streamed without caches' }] }] });
writeFileSync(join(dir, 'column-stream-no-cache.xlsx'), new Uint8Array(await new Response(stream).arrayBuffer()));
console.log(`Wrote ${String(Object.keys(samples).length + 1)} files to ${dir}`);
