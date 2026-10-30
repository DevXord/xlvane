// Opens files written by xlvane in LibreOffice (headless) and checks what it sees:
// 1. CSV export of every sheet must contain the expected values;
// 2. LibreOffice's own XLSX re-save must keep merges, frozen panes, filters, widths and styles.
// Skipped when `soffice` is not installed, unless XLV_REQUIRE_LIBREOFFICE=1 (set in CI).
// CSV filter options: LibreOffice help "CSV Filter Options" (docs/decisions.md F-LO-01).
import { spawnSync } from 'node:child_process';
import { mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { createWorkbook, writeWorkbook, type NewWorkbook } from '../../packages/xlvane/src/write';

const SOFFICE = process.env['XLV_SOFFICE'] ?? 'soffice';
const available = spawnSync(SOFFICE, ['--version'], { encoding: 'utf8', timeout: 60_000 }).status === 0;
const required = process.env['XLV_REQUIRE_LIBREOFFICE'] === '1';
if (!available && !required) console.warn(`LibreOffice tests skipped: "${SOFFICE}" was not found. Set XLV_SOFFICE to its path.`);

/** Field separator 44 (,), text delimiter 34 ("), charset 76 (UTF-8), first line 1, no column
 * formats, language 0, quoted as text false, numbers as numbers true, contents as shown true,
 * formulas false, trim false, all sheets -1. */
const CSV_FILTER = 'csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,true,false,false,-1';

function fixtures(): Record<string, NewWorkbook> {
  const values = createWorkbook();
  const data = values.addSheet('Data');
  data.set('A1', 'Text', { font: { bold: true } }).set('B1', 'Number').set('C1', 'Date').set('D1', 'Bool').set('E1', 'Formula');
  data
    .set('A2', 'zażółć "q", 😀')
    .set('B2', 1234.5, { numFmt: '0.00' })
    .set('C2', new Date(Date.UTC(2024, 0, 31)))
    .set('D2', true)
    .set('E2', { formula: 'B2*2', result: 2469 });
  data
    .set('A3', 'line\nbreak')
    .set('B3', -0.001)
    .set('C3', new Date(Date.UTC(2024, 1, 29, 13, 45, 30)))
    .set('D3', false)
    .set('E3', { error: '#N/A' });
  data.set('A4', '_x0041_ <&>').set('B4', 1e20, { numFmt: '0.00E+00' }).set('C4', new Date(Date.UTC(1900, 2, 1))).set('E4', { formula: 'A2&"!"', result: 'zażółć "q", 😀!' });
  data.merge('A6:C6').freeze({ rows: 1, cols: 1 }).autoFilter('A1:E4').column('A', { width: 30 }).column('F', { hidden: true }).row(1, { height: 25 });
  values.addSheet('Second & <x>').set('A1', { formula: 'Data!B2+1', result: 1235.5 });
  values.addSheet('Hidden', { state: 'hidden' }).set('A1', 'secret');

  const dates1904 = createWorkbook({ dateSystem: '1904' });
  dates1904.addSheet('S').set('A1', new Date(Date.UTC(1904, 0, 1))).set('A2', new Date(Date.UTC(2030, 5, 15, 6)));

  const styles = createWorkbook();
  const st = styles.addSheet('Styles');
  st.set('A1', 'bold red', { font: { bold: true, color: 'FF0000', size: 14, name: 'Arial' } });
  st.set('A2', 'fill', { fill: { color: 'FFFF00' } });
  st.set('A3', 'border', { border: { top: { style: 'thin' }, bottom: { style: 'double', color: '0000FF' } } });
  st.set('A4', 'wrapped centre', { alignment: { horizontal: 'center', vertical: 'top', wrapText: true } });
  st.set('A5', 0.256, { numFmt: '0.0%' });
  st.style('B1:B3', { fill: { pattern: 'darkGrid', fgColor: '112233', bgColor: 'FFFFFF' } });
  return { values, dates1904, styles };
}

const EXPECTED_CSV: Record<string, string[][]> = {
  'values-Data': [
    ['Text', 'Number', 'Date', 'Bool', 'Formula'],
    ['zażółć "q", 😀', '1234.50', '2024-01-31', 'TRUE', '2469'],
    ['line\nbreak', '-0.001', '2024-02-29 13:45:30', 'FALSE', '#N/A'],
    ['_x0041_ <&>', '1.00E+20', '1900-03-01', '', 'zażółć "q", 😀!'],
  ],
  'values-Second & <x>': [['1235.5']],
  'values-Hidden': [['secret']],
  'dates1904-S': [['1904-01-01'], ['2030-06-15 06:00:00']],
  'styles-Styles': [['bold red', ''], ['fill', ''], ['border', ''], ['wrapped centre'], ['25.6%']],
};

/** Minimal RFC 4180 parser: quoted fields, doubled quotes, line breaks inside quotes. */
function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  let i = 0;
  let fieldStarted = false;
  while (i < text.length) {
    const ch = text[i]!;
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') {
        field += '"';
        i += 2;
        continue;
      }
      if (ch === '"') quoted = false;
      else field += ch;
      i++;
      continue;
    }
    if (ch === '"') {
      quoted = true;
      fieldStarted = true;
    } else if (ch === ',') {
      row.push(field);
      field = '';
      fieldStarted = true;
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      if (fieldStarted || field !== '' || row.length > 0) row.push(field);
      rows.push(row);
      row = [];
      field = '';
      fieldStarted = false;
    } else {
      field += ch;
      fieldStarted = true;
    }
    i++;
  }
  if (fieldStarted || field !== '' || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

let dir = '';

function soffice(args: string[]): void {
  const profile = pathToFileURL(join(dir, 'profile')).href;
  const result = spawnSync(SOFFICE, [`-env:UserInstallation=${profile}`, '--headless', ...args], { encoding: 'utf8', timeout: 180_000 });
  if (result.status !== 0) throw new Error(`soffice failed (${String(result.status)}): ${result.stderr}${result.stdout}`);
}

describe.skipIf(!available && !required)('LibreOffice opens files written by xlvane', () => {
  const names = Object.keys(fixtures());

  beforeAll(async () => {
    dir = await mkdtemp(join(tmpdir(), 'xlvane-lo-'));
    const files: string[] = [];
    for (const [name, wb] of Object.entries(fixtures())) {
      const file = join(dir, `${name}.xlsx`);
      await writeFile(file, writeWorkbook(wb));
      files.push(file);
    }
    soffice(['--convert-to', CSV_FILTER, '--outdir', join(dir, 'csv'), ...files]);
    soffice(['--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', join(dir, 'resaved'), ...files]);
  }, 400_000);

  afterAll(async () => {
    if (dir !== '') await rm(dir, { recursive: true, force: true });
  });

  it('exports one CSV file per sheet', async () => {
    const produced = (await readdir(join(dir, 'csv'))).map((f) => f.replace(/\.csv$/, '')).sort();
    expect(produced).toEqual(Object.keys(EXPECTED_CSV).sort());
    expect(names.length).toBe(3);
  });

  for (const [file, expected] of Object.entries(EXPECTED_CSV)) {
    it(`shows the expected values in ${file}`, async () => {
      const csv = await readFile(join(dir, 'csv', `${file}.csv`), 'utf8');
      const rows = parseCsv(csv);
      // LibreOffice pads rows to the used width; compare the non-empty prefix of each row.
      const trim = (r: string[]): string[] => {
        const out = [...r];
        while (out.length > 0 && out[out.length - 1] === '') out.pop();
        return out;
      };
      expect(rows.map(trim)).toEqual(expected.map(trim));
    });
  }

  it('keeps layout and sheet states when LibreOffice re-saves the file', async () => {
    const wb = readWorkbook(await readFile(join(dir, 'resaved', 'values.xlsx')));
    expect(wb.sheets.map((s) => [s.name, s.state])).toEqual([
      ['Data', 'visible'],
      ['Second & <x>', 'visible'],
      ['Hidden', 'hidden'],
    ]);
    const s = wb.getSheet('Data');
    expect(s.merges).toEqual(['A6:C6']);
    expect(s.freeze).toEqual({ rows: 1, cols: 1 });
    expect(s.autoFilter).toBe('A1:E4');
    expect(s.columns.find((c) => c.min <= 1 && c.max >= 1)?.width).toBeCloseTo(30, 0);
    expect(s.columns.find((c) => c.min <= 6 && c.max >= 6)?.hidden).toBe(true);
    expect(s.rowInfo(1)?.height).toBeCloseTo(25, 0);
    expect(s.cell('A1')?.style.font?.bold).toBe(true);
    expect(s.cell('C2')?.value).toEqual(new Date(Date.UTC(2024, 0, 31)));
    expect(s.cell('E2')?.formula?.text).toBe('B2*2');
  });

  it('keeps fonts, fills, borders, alignment and number formats', async () => {
    const s = readWorkbook(await readFile(join(dir, 'resaved', 'styles.xlsx'))).getSheet('Styles');
    expect(s.cell('A1')?.style.font).toMatchObject({ name: 'Arial', size: 14, bold: true, color: 'FFFF0000' });
    expect(s.cell('A2')?.style.fill).toMatchObject({ pattern: 'solid', fgColor: 'FFFFFF00' });
    expect(s.cell('A3')?.style.border).toMatchObject({ top: { style: 'thin' }, bottom: { style: 'double', color: 'FF0000FF' } });
    expect(s.cell('A4')?.style.alignment).toMatchObject({ horizontal: 'center', vertical: 'top', wrapText: true });
    expect(s.cell('A5')?.numFmt).toBe('0.0%');
  });

  it('keeps the 1904 date system', async () => {
    const wb = readWorkbook(await readFile(join(dir, 'resaved', 'dates1904.xlsx')));
    expect(wb.dateSystem).toBe('1904');
    expect(wb.getSheet('S').cell('A1')?.value).toEqual(new Date(Date.UTC(1904, 0, 1)));
  });
});
