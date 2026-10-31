// A portable smoke suite for every supported runtime (Node.js, Bun, Deno, browsers, Cloudflare
// Workers; docs/decisions.md D-904). It uses only the core API (no Node.js modules, no Buffer, no
// file system): every input is generated here. scripts/build-smoke.mjs bundles it into one ESM
// file that each runtime runner in tests/runtimes/ imports and runs.
import ExcelJS from '@xlvane/compat';
import {
  appendRows,
  createWorkbook,
  diff,
  editCells,
  fromCSV,
  hashPassword,
  inspect,
  openWorkbookStream,
  readTable,
  readWorkbook,
  toCSV,
  verifyPassword,
  writeTable,
  writeWorkbook,
  writeWorkbookStream,
  XlvError,
} from 'xlvane';

/** The result of one check. */
export interface SmokeResult {
  readonly name: string;
  readonly ok: boolean;
  readonly error?: string;
}

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(message);
}

function eq(actual: unknown, expected: unknown, what: string): void {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  assert(a === e, `${what}: expected ${e}, got ${a}`);
}

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

/** A small report used by several checks. */
function report(): Uint8Array {
  const wb = createWorkbook();
  const s = wb.addSheet('Sales');
  const head = { font: { bold: true }, fill: { color: 'D9E1F2' }, border: { bottom: { style: 'thin' } } } as const;
  s.set('A1', 'Region', head).set('B1', 'Revenue', head).set('C1', 'Date', head);
  s.set('A2', 'North').set('B2', 1250.5, { numFmt: '#,##0.00' }).set('C2', new Date(Date.UTC(2024, 0, 31)));
  s.set('A3', 'Południe ✓').set('B3', 980).set('C3', new Date(Date.UTC(2024, 1, 29, 13, 45)));
  s.set('A4', 'Total').set('B4', { formula: 'SUM(B2:B3)', result: 2230.5 });
  s.merge('A6:C6').set('A6', 'merged');
  s.freeze({ rows: 1 }).autoFilter('A1:C3');
  wb.addSheet('Hidden', { state: 'hidden' }).set('A1', { error: '#N/A' });
  return writeWorkbook(wb);
}

function letters(col: number): string {
  let s = '';
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

/** Generated workbooks of different shapes, standing in for corpus files. */
function variedInputs(): [string, Uint8Array][] {
  const out: [string, Uint8Array][] = [];
  const wide = createWorkbook();
  const w = wide.addSheet('Wide');
  for (let c = 1; c <= 300; c++) w.set(`${letters(c)}1`, c);
  out.push(['300 columns', writeWorkbook(wide)]);
  const tall = createWorkbook();
  const t = tall.addSheet('Tall');
  for (let r = 1; r <= 5000; r++) t.set(`A${String(r)}`, r).set(`B${String(r)}`, `row ${String(r)}`);
  out.push(['5000 rows', writeWorkbook(tall)]);
  const d1904 = createWorkbook({ dateSystem: '1904' });
  d1904.addSheet('D').set('A1', new Date(Date.UTC(2024, 5, 1)));
  out.push(['1904 dates', writeWorkbook(d1904)]);
  const many = createWorkbook();
  for (let i = 1; i <= 20; i++) many.addSheet(`S${String(i)}`).set('A1', i);
  out.push(['20 sheets', writeWorkbook(many)]);
  const text = createWorkbook();
  text.addSheet('Text').set('A1', 'line 1\nline 2').set('A2', '  spaces  ').set('A3', 'zażółć 😀').set('A4', '=not a formula').set('A5', 'x'.repeat(32_767));
  out.push(['text edge cases', writeWorkbook(text)]);
  out.push(['writeTable report', writeTable([{ id: 1, when: new Date(Date.UTC(2020, 0, 1)), ok: true }, { id: 2, when: null, ok: false }])]);
  out.push(['fromCSV', fromCSV('a,b\r\n1,2024-01-31\r\n"x,y",TRUE\r\n')]);
  return out;
}

function checks(): (readonly [string, () => Promise<void> | void])[] {
  return [
    [
      'write and read back values, styles, merges and layout',
      () => {
        const wb = readWorkbook(report());
        const s = wb.getSheet('Sales');
        eq(s.cell('B2')?.value, 1250.5, 'B2');
        eq(s.cell('B2')?.numFmt, '#,##0.00', 'B2 format');
        eq((s.cell('C2')?.value as Date).toISOString(), '2024-01-31T00:00:00.000Z', 'C2 date');
        eq((s.cell('C3')?.value as Date).toISOString(), '2024-02-29T13:45:00.000Z', 'C3 date and time');
        eq(s.cell('A3')?.value, 'Południe ✓', 'unicode text');
        eq(s.cell('B4')?.formula?.text, 'SUM(B2:B3)', 'formula');
        eq(s.cell('B4')?.value, 2230.5, 'cached result');
        eq(s.cell('A1')?.style.font?.bold, true, 'bold header');
        eq(s.merges, ['A6:C6'], 'merges');
        eq(s.freeze, { rows: 1, cols: 0 }, 'frozen panes');
        eq(s.autoFilter, 'A1:C3', 'autofilter');
        eq(wb.getSheet('Hidden').state, 'hidden', 'hidden sheet');
        eq(wb.getSheet('Hidden').cell('A1')?.value, '#N/A', 'error value');
        eq(wb.warnings, [], 'warnings');
      },
    ],
    [
      'writing is deterministic',
      () => {
        const a = report();
        const b = report();
        assert(a.length === b.length && a.every((v, i) => v === b[i]), 'two writes of the same workbook differ');
      },
    ],
    [
      'inspect',
      () => {
        const r = inspect(report());
        eq(r.sheetCount, 2, 'sheet count');
        const sales = r.sheets[0]!;
        eq(sales.headerRow, { row: 1, confidence: 'high' }, 'header row');
        eq(sales.columns.map((c) => c.dominantType), ['string', 'number', 'date'], 'column types');
      },
    ],
    [
      'readTable',
      () => {
        const t = readTable(report(), { sheet: 'Sales', range: 'A1:C3' });
        eq(t.rows.length, 2, 'rows');
        eq(t.rows[0]?.['Region'], 'North', 'first row');
        assert(t.rows[1]?.['Date'] instanceof Date, 'dates are Date objects');
      },
    ],
    [
      'toCSV escapes formula-like text and fromCSV reads it back',
      () => {
        const wb = createWorkbook();
        wb.addSheet('S').set('A1', 'h').set('A2', '=cmd|calc').set('A3', 12.5);
        const csv = toCSV(writeWorkbook(wb));
        eq(csv, `h\r\n"'=cmd|calc"\r\n12.5\r\n`, 'csv');
        const back = readTable(fromCSV(csv));
        eq(back.rows, [{ h: "'=cmd|calc" }, { h: 12.5 }], 'csv round trip');
      },
    ],
    [
      'editCells and appendRows keep the rest, diff finds the change',
      () => {
        const before = report();
        const edited = editCells(before, [{ sheet: 'Sales', cell: 'B3', value: 1000 }]).bytes;
        const appended = appendRows(edited, 'Sales', [['East', 1, new Date(Date.UTC(2024, 2, 1))]]).bytes;
        const s = readWorkbook(appended).getSheet('Sales');
        eq(s.cell('B3')?.value, 1000, 'edited value');
        eq(s.cell('A1')?.style.font?.bold, true, 'style kept');
        eq(s.cell('B4')?.formula?.text, 'SUM(B2:B3)', 'formula kept');
        const d = diff(before, edited);
        eq(d.summary.cellsChanged, 1, 'one change');
        eq(d.sheets[0]?.cells[0]?.ref, 'B3', 'changed cell');
      },
    ],
    [
      'Web Streams: writeWorkbookStream and openWorkbookStream',
      async () => {
        async function* rows() {
          yield ['id', 'name'];
          for (let i = 1; i <= 2000; i++) yield [i, `n${String(i)}`];
        }
        const bytes = await collect(writeWorkbookStream({ sheets: [{ name: 'S', rows: rows(), freeze: { rows: 1 } }] }));
        const book = await openWorkbookStream(bytes);
        eq(book.sheets.map((s) => s.name), ['S'], 'sheets');
        let count = 0;
        let last: unknown = null;
        for await (const row of book.rows('S')) {
          count++;
          last = row.cells[1]?.value;
        }
        eq(count, 2001, 'streamed rows');
        eq(last, 'n2000', 'last value');
        const blob = await openWorkbookStream(new Blob([bytes as Uint8Array<ArrayBuffer>]));
        eq(blob.sheets.length, 1, 'Blob input');
      },
    ],
    [
      '@xlvane/compat without Node.js Buffer',
      async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Data');
        ws.columns = [
          { header: 'Name', key: 'name', width: 20 },
          { header: 'Total', key: 'total' },
        ];
        ws.addRow({ name: 'Tea', total: 3.5 });
        ws.getCell('B2').numFmt = '0.00';
        const buffer = await wb.xlsx.writeBuffer();
        const copy = await new ExcelJS.Workbook().xlsx.load(buffer);
        const sheet = copy.getWorksheet('Data');
        eq(sheet?.getCell('A2').value, 'Tea', 'value');
        eq(sheet?.getCell('B2').numFmt, '0.00', 'number format');
        eq(sheet?.getColumn(1).width, 20, 'column width');
      },
    ],
    [
      'password verifiers: fixed-salt hash and random salt (D-C08)',
      async () => {
        const hash = hashPassword('secret', { salt: 'c2FsdHNhbHRzYWx0c2FsdA==', spinCount: 1000 });
        eq(verifyPassword(hash, 'secret'), true, 'verify');
        eq(verifyPassword(hash, 'Secret'), false, 'reject');
        const random = hashPassword('secret', { spinCount: 10 });
        eq(random.saltValue.length, 24, 'random 16-byte salt');
        const wb = createWorkbook();
        wb.addSheet('S').set('A1', 1).protect({ password: hash });
        eq(verifyPassword(readWorkbook(writeWorkbook(wb)).getSheet('S').protection!, 'secret'), true, 'round trip');
        const xl = new ExcelJS.Workbook();
        const ws = xl.addWorksheet('P');
        await ws.protect('secret', { spinCount: 10 });
        eq(verifyPassword(readWorkbook(new Uint8Array(await xl.xlsx.writeBuffer())).getSheet('P').protection!, 'secret'), true, 'exceljs layer');
      },
    ],
    [
      'untrusted input: errors with codes, no crash',
      () => {
        const cases: [string, Uint8Array, string][] = [
          ['empty', new Uint8Array(0), 'XLV_E_NOT_ZIP'],
          ['text', new TextEncoder().encode('Region,Revenue\nNorth,1'), 'XLV_E_NOT_ZIP'],
          ['truncated', report().slice(0, 300), 'XLV_E_ZIP_CORRUPT'],
          ['CFB container', new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, ...new Array<number>(504).fill(0)]), 'XLV_E_UNSUPPORTED_CONTAINER'],
        ];
        for (const [what, bytes, code] of cases) {
          try {
            readWorkbook(bytes);
            throw new Error(`${what}: no error`);
          } catch (error) {
            assert(error instanceof XlvError, `${what}: not an XlvError: ${String(error)}`);
            eq(error.code, code, `${what} code`);
          }
        }
      },
    ],
    ...variedInputs().map(
      ([name, bytes]) =>
        [
          `read generated file: ${name}`,
          () => {
            const wb = readWorkbook(bytes);
            assert(wb.sheets.length > 0, 'no sheets');
            inspect(wb);
            const first = wb.sheets.find((s) => s.state === 'visible')!;
            let n = 0;
            for (const _ of first.cells()) n++;
            assert(n > 0, 'no cells');
            toCSV(wb, { sheet: first.name });
          },
        ] as const,
    ),
  ];
}

/** Runs every check and returns one result per check (a failure does not stop the others). */
export async function runSmoke(): Promise<SmokeResult[]> {
  const results: SmokeResult[] = [];
  let list: ReturnType<typeof checks>;
  try {
    list = checks();
  } catch (error) {
    return [{ name: 'generate inputs', ok: false, error: error instanceof Error ? `${error.name}: ${error.message}` : String(error) }];
  }
  for (const [name, check] of list) {
    try {
      await check();
      results.push({ name, ok: true });
    } catch (error) {
      results.push({ name, ok: false, error: error instanceof Error ? `${error.name}: ${error.message}` : String(error) });
    }
  }
  return results;
}
