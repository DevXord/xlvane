// Benchmark runner: `npm run bench` (builds xlvane, installs bench/ dependencies, runs this).
// Each variant runs in a fresh child process BENCH_RUNS times (default 5); the median is
// reported. Environment: BENCH_ROWS (default 100000), BENCH_COLS (default 20), BENCH_RUNS.
// Method and API sources: bench/README.md.
import { spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';
import { samplePositions, value, same, describe } from './data.mjs';

const here = fileURLToPath(new URL('.', import.meta.url));
const root = join(here, '..', '..');
const rows = Number(process.env.BENCH_ROWS ?? 100_000);
const cols = Number(process.env.BENCH_COLS ?? 20);
const runs = Number(process.env.BENCH_RUNS ?? 5);
const tmp = join(here, 'tmp');
const dist = join(root, 'packages/xlvane/dist/index.js');

const WRITE = ['xlvane-write', 'xlvane-write-stream', 'exceljs-write', 'exceljs-write-stream', 'sheetjs-write'];
const READ = ['xlvane-read', 'xlvane-read-stream', 'exceljs-read', 'exceljs-read-stream', 'sheetjs-read'];

if (!existsSync(dist)) {
  console.error('packages/xlvane/dist is missing; run `npm run build` first (npm run bench does).');
  process.exit(1);
}
const { readWorkbook } = await import(dist);

function runVariant(name, file) {
  const result = spawnSync(process.execPath, [join(here, 'variant.mjs'), name, String(rows), String(cols), file], {
    encoding: 'utf8',
    env: { ...process.env, TZ: 'UTC' },
    maxBuffer: 16 * 1024 * 1024,
  });
  if (result.status !== 0) throw new Error(`${name} failed:\n${result.stderr}`);
  return JSON.parse(result.stdout.trim().split('\n').pop());
}

/** Reads a written file with xlvane and checks the cell count and the sample values. */
async function validateFile(name, file) {
  const sheet = readWorkbook(new Uint8Array(await readFile(file))).getSheet('Data');
  if (sheet.cellCount !== rows * cols) throw new Error(`${name}: wrote ${sheet.cellCount} cells, expected ${rows * cols}`);
  for (const [r, c] of samplePositions(rows, cols)) {
    const ref = `${letters(c)}${r + 1}`;
    const got = sheet.cell(ref)?.value;
    const expected = value(r, c, cols);
    if (!same(got, expected)) throw new Error(`${name}: ${ref} is ${describe(got)}, expected ${describe(expected)}`);
  }
}

function letters(col) {
  let out = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) out = String.fromCharCode(65 + ((n - 1) % 26)) + out;
  return out;
}

const median = (xs) => {
  const s = [...xs].sort((a, b) => a - b);
  const m = Math.floor(s.length / 2);
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
};

async function measure(name, file, validate) {
  const samples = [];
  for (let i = 0; i < runs; i++) {
    const r = runVariant(name, file);
    if (r.unavailable) return { variant: name, unavailable: r.unavailable };
    if (i === 0 && validate) await validateFile(name, file);
    samples.push(r);
    process.stdout.write('.');
  }
  return {
    variant: name,
    runs: samples.length,
    medianMs: median(samples.map((s) => s.ms)),
    minMs: Math.min(...samples.map((s) => s.ms)),
    maxMs: Math.max(...samples.map((s) => s.ms)),
    medianMaxRssBytes: median(samples.map((s) => s.maxRssBytes)),
    medianPeakHeapBytes: median(samples.map((s) => s.peakHeapBytes)),
    samples,
  };
}

async function pkgVersion(path) {
  return JSON.parse(await readFile(path, 'utf8')).version;
}

await rm(tmp, { recursive: true, force: true });
await mkdir(tmp, { recursive: true });
console.log(`Benchmark: ${rows} rows x ${cols} columns, ${runs} runs per variant`);

const input = join(tmp, 'input.xlsx');
runVariant('xlvane-write', input);
await validateFile('input', input);

const results = [];
console.log('baseline');
results.push({ ...(await measure('baseline', input, false)), kind: 'baseline' });
for (const name of WRITE) {
  console.log(`\n${name}`);
  results.push({ ...(await measure(name, join(tmp, `${name}.xlsx`), true)), kind: 'write' });
}
for (const name of READ) {
  console.log(`\n${name}`);
  results.push({ ...(await measure(name, input, false)), kind: 'read' });
}
console.log();

const cpus = os.cpus();
const environment = {
  date: new Date().toISOString(),
  cpu: cpus[0]?.model ?? 'unknown',
  cores: cpus.length,
  memoryBytes: os.totalmem(),
  os: `${os.type()} ${os.release()} (${os.arch()})`,
  node: process.version,
  versions: {
    xlvane: await pkgVersion(join(root, 'packages/xlvane/package.json')),
    exceljs: await pkgVersion(join(here, '..', 'node_modules/exceljs/package.json')),
    xlsx: await pkgVersion(join(here, '..', 'node_modules/xlsx/package.json')),
  },
  rows,
  cols,
  runs,
};

const host = os.hostname().replace(/[^A-Za-z0-9_-]/g, '_').slice(0, 40);
await mkdir(join(here, 'results'), { recursive: true });
const jsonName = `${environment.date.slice(0, 10)}-${host}.json`;
await writeFile(join(here, 'results', jsonName), `${JSON.stringify({ environment, input: { writtenBy: 'xlvane-write', bytes: (await readFile(input)).length }, results }, null, 2)}\n`);

const mb = (b) => (b / 1024 / 1024).toFixed(0);
const s = (ms) => (ms / 1000).toFixed(2);
const baseline = results.find((r) => r.kind === 'baseline');
const table = (kind) =>
  results
    .filter((r) => r.kind === kind)
    .map((r) =>
      r.unavailable
        ? `| ${r.variant} | not available (${r.unavailable}) | | | |`
        : `| ${r.variant} | ${s(r.medianMs)} | ${s(r.minMs)}–${s(r.maxMs)} | ${mb(r.medianMaxRssBytes)} | ${mb(r.medianPeakHeapBytes)} |`,
    )
    .join('\n');

const md = `# Benchmark results

${rows.toLocaleString('en-US')} rows × ${cols} columns (no header row), ${runs} runs per variant, median reported. Raw data: \`bench/perf/results/${jsonName}\`.

## Environment

| | |
|---|---|
| Date | ${environment.date} |
| CPU | ${environment.cpu}, ${environment.cores} logical cores |
| Memory | ${(environment.memoryBytes / 1024 ** 3).toFixed(1)} GiB |
| OS | ${environment.os} |
| Node.js | ${environment.node} |
| xlvane | ${environment.versions.xlvane} (this repository, unminified build) |
| exceljs | ${environment.versions.exceljs} |
| xlsx (SheetJS) | ${environment.versions.xlsx} (npm) |

## Write

| Variant | Time, s (median) | Range, s | Peak RSS, MiB | Peak heap, MiB |
|---|---|---|---|---|
${table('write')}

## Read

Every read variant reads the same file, written by \`xlvane-write\` (${((await readFile(input)).length / 1024 / 1024).toFixed(1)} MiB), and visits every cell.

| Variant | Time, s (median) | Range, s | Peak RSS, MiB | Peak heap, MiB |
|---|---|---|---|---|
${table('read')}

Baseline (Node.js with the three libraries imported, no work): peak RSS ${baseline && !baseline.unavailable ? mb(baseline.medianMaxRssBytes) : '?'} MiB.

## Method

- Each run is a fresh \`node\` process (\`bench/perf/variant.mjs\`) with \`TZ=UTC\`; the time covers the library work only (building and writing the workbook, or reading it and visiting every cell), not process start-up or imports.
- Peak RSS is \`process.resourceUsage().maxRSS\` (kibibytes per the Node.js documentation) at the end of the run, so it includes Node.js itself and the imported libraries (see the baseline).
- Peak heap is \`process.memoryUsage().heapUsed\` sampled every 10 ms and once at the end of the timed work. Synchronous work blocks the sampling timer, so for synchronous variants this is the heap at the end of the work (with the result still alive), a lower bound of the true peak.
- Every write variant's file is read back with xlvane and checked (cell count and sample values of every column type); every read variant checks the cell count and the same sample values.
- The data is deterministic (\`bench/perf/data.mjs\`): a unique id column, then integers, floats, short texts, dates and booleans.
- Streaming variants: xlvane streams to/from a file through Web Streams (\`Readable.fromWeb\`, \`fs.openAsBlob\`); exceljs uses its streaming WorkbookWriter (without shared strings) and WorkbookReader.

## Caveats

- SheetJS is the npm release 0.18.5 (March 2022). That version has CVE-2023-30533 (prototype pollution when reading crafted files); newer releases are distributed only from the SheetJS CDN, which this environment could not reach. The benchmark reads only files it generated itself.
- Results depend on the machine and on the data mix; they are one run on one machine, not a general ranking.
`;
await writeFile(join(here, 'RESULTS.md'), md);
console.log(md);
await rm(tmp, { recursive: true, force: true });
