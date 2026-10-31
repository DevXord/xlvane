// Runs one benchmark variant in this process and prints one JSON line with the result:
//   node bench/perf/variant.mjs <variant> <rows> <cols> <file>
// Write variants write <file>; read variants read it. Timing covers only the library work
// (building and writing, or reading and visiting every cell), not the imports.
// Library APIs used are listed with their sources in bench/README.md.
import { readFile, writeFile } from 'node:fs/promises';
import { createWriteStream } from 'node:fs';
import { Readable } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import { performance } from 'node:perf_hooks';
import { rowValues, samplePositions, value, same, describe } from './data.mjs';

const [variant, rowsArg, colsArg, file] = process.argv.slice(2);
const rows = Number(rowsArg);
const cols = Number(colsArg);

const xlvanePath = new URL('../../packages/xlvane/dist/index.js', import.meta.url);

function letters(col) {
  let s = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

// Peak heap: sampled every 10 ms while the event loop is free, and once more at the end of the
// timed work (synchronous work cannot be sampled while it runs; see bench/README.md).
let peakHeap = 0;
const sample = () => {
  const used = process.memoryUsage().heapUsed;
  if (used > peakHeap) peakHeap = used;
};
const timer = setInterval(sample, 10);

/** Collects sampled values from a read variant: [row, col, value]. */
function checkSamples(get) {
  for (const [r, c] of samplePositions(rows, cols)) {
    const expected = value(r, c, cols);
    const got = get(r, c);
    if (!same(got, expected)) throw new Error(`${variant}: cell ${letters(c)}${r + 1} is ${describe(got)}, expected ${describe(expected)}`);
  }
}

const variants = {
  // ----- write -------------------------------------------------------------------------------
  async 'xlvane-write'() {
    const { createWorkbook, writeWorkbook } = await import(xlvanePath);
    const start = performance.now();
    const wb = createWorkbook();
    const sheet = wb.addSheet('Data');
    const refs = Array.from({ length: cols }, (_, c) => letters(c));
    for (let r = 0; r < rows; r++) {
      const values = rowValues(r, cols);
      for (let c = 0; c < cols; c++) sheet.set(refs[c] + (r + 1), values[c]);
    }
    const bytes = writeWorkbook(wb);
    sample();
    const ms = performance.now() - start;
    await writeFile(file, bytes);
    return { ms };
  },

  async 'xlvane-write-stream'() {
    const xlvane = await import(xlvanePath);
    if (typeof xlvane.writeWorkbookStream !== 'function') return { unavailable: 'writeWorkbookStream is not exported by this build' };
    const start = performance.now();
    function* source() {
      for (let r = 0; r < rows; r++) yield rowValues(r, cols);
    }
    const stream = xlvane.writeWorkbookStream({ sheets: [{ name: 'Data', rows: source() }] });
    await pipeline(Readable.fromWeb(stream), createWriteStream(file));
    sample();
    return { ms: performance.now() - start };
  },

  async 'exceljs-write'() {
    const { default: ExcelJS } = await import('exceljs');
    const start = performance.now();
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Data');
    for (let r = 0; r < rows; r++) ws.addRow(rowValues(r, cols));
    const buffer = await wb.xlsx.writeBuffer();
    sample();
    const ms = performance.now() - start;
    await writeFile(file, buffer);
    return { ms };
  },

  async 'exceljs-write-stream'() {
    const { default: ExcelJS } = await import('exceljs');
    const start = performance.now();
    const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: file, useStyles: true, useSharedStrings: false });
    const ws = wb.addWorksheet('Data');
    for (let r = 0; r < rows; r++) ws.addRow(rowValues(r, cols)).commit();
    ws.commit();
    await wb.commit();
    sample();
    return { ms: performance.now() - start };
  },

  async 'sheetjs-write'() {
    const { default: XLSX } = await import('xlsx');
    const start = performance.now();
    const data = Array.from({ length: rows }, (_, r) => rowValues(r, cols));
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    sample();
    const ms = performance.now() - start;
    await writeFile(file, buffer);
    return { ms };
  },

  // ----- read --------------------------------------------------------------------------------
  async 'xlvane-read'() {
    const { readWorkbook } = await import(xlvanePath);
    const bytes = new Uint8Array(await readFile(file));
    const start = performance.now();
    const sheet = readWorkbook(bytes).getSheet('Data');
    let cells = 0;
    const kept = new Map();
    const wanted = new Set(samplePositions(rows, cols).map(([r, c]) => `${r},${c}`));
    for (const row of sheet.rows()) {
      for (const cell of row) {
        cells++;
        const key = `${cell.row - 1},${cell.col - 1}`;
        if (wanted.has(key)) kept.set(key, cell.value);
      }
    }
    sample();
    const ms = performance.now() - start;
    checkSamples((r, c) => kept.get(`${r},${c}`));
    return { ms, cells };
  },

  async 'xlvane-read-stream'() {
    const xlvane = await import(xlvanePath);
    if (typeof xlvane.openWorkbookStream !== 'function') return { unavailable: 'openWorkbookStream is not exported by this build' };
    const { openAsBlob } = await import('node:fs');
    const blob = await openAsBlob(file);
    const start = performance.now();
    const book = await xlvane.openWorkbookStream(blob);
    let cells = 0;
    const kept = new Map();
    const wanted = new Set(samplePositions(rows, cols).map(([r, c]) => `${r},${c}`));
    for await (const row of book.rows('Data')) {
      for (const cell of row.cells) {
        cells++;
        const key = `${cell.row - 1},${cell.col - 1}`;
        if (wanted.has(key)) kept.set(key, cell.value);
      }
    }
    sample();
    const ms = performance.now() - start;
    checkSamples((r, c) => kept.get(`${r},${c}`));
    return { ms, cells };
  },

  async 'exceljs-read'() {
    const { default: ExcelJS } = await import('exceljs');
    const buffer = await readFile(file);
    const start = performance.now();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.getWorksheet('Data');
    let cells = 0;
    const kept = new Map();
    const wanted = new Set(samplePositions(rows, cols).map(([r, c]) => `${r},${c}`));
    ws.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        cells++;
        const key = `${rowNumber - 1},${colNumber - 1}`;
        if (wanted.has(key)) kept.set(key, cell.value);
      });
    });
    sample();
    const ms = performance.now() - start;
    checkSamples((r, c) => kept.get(`${r},${c}`));
    return { ms, cells };
  },

  async 'exceljs-read-stream'() {
    const { default: ExcelJS } = await import('exceljs');
    const start = performance.now();
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(file, { sharedStrings: 'cache', styles: 'cache', worksheets: 'emit', hyperlinks: 'ignore' });
    let cells = 0;
    const kept = new Map();
    const wanted = new Set(samplePositions(rows, cols).map(([r, c]) => `${r},${c}`));
    for await (const worksheet of reader) {
      for await (const row of worksheet) {
        row.eachCell((cell, colNumber) => {
          cells++;
          const key = `${row.number - 1},${colNumber - 1}`;
          if (wanted.has(key)) kept.set(key, cell.value);
        });
      }
    }
    sample();
    const ms = performance.now() - start;
    checkSamples((r, c) => kept.get(`${r},${c}`));
    return { ms, cells };
  },

  async 'sheetjs-read'() {
    const { default: XLSX } = await import('xlsx');
    const buffer = await readFile(file);
    const start = performance.now();
    const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
    const ws = wb.Sheets['Data'];
    let cells = 0;
    for (const key of Object.keys(ws)) if (key[0] !== '!') cells++;
    sample();
    const ms = performance.now() - start;
    checkSamples((r, c) => ws[letters(c) + (r + 1)]?.v);
    return { ms, cells };
  },

  // ----- baseline ----------------------------------------------------------------------------
  async baseline() {
    // Node.js itself, with all three libraries imported: the floor under every RSS number.
    await import(xlvanePath);
    await import('exceljs');
    await import('xlsx');
    sample();
    return { ms: 0 };
  },
};

const run = variants[variant];
if (!run) {
  console.error(`Unknown variant "${variant}". Known: ${Object.keys(variants).join(', ')}`);
  process.exit(2);
}
try {
  const result = await run();
  clearInterval(timer);
  if (result.cells !== undefined && result.cells !== rows * cols) {
    throw new Error(`${variant}: read ${result.cells} cells, expected ${rows * cols}`);
  }
  // maxRSS is in kibibytes (Node.js process.resourceUsage(), ru_maxrss).
  const maxRssBytes = process.resourceUsage().maxRSS * 1024;
  console.log(JSON.stringify({ variant, ...result, maxRssBytes, peakHeapBytes: peakHeap }));
} catch (error) {
  clearInterval(timer);
  console.error(error && error.stack ? error.stack : String(error));
  process.exit(1);
}
