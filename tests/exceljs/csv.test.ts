// xlvane tests for workbook.csv in @xlvane/exceljs (docs/decisions.md D-E03, docs/facts/exceljs.md
// X4-23, X4-24): exceljs 4.4.0 options and output, CSV-injection escaping, and the documented
// fixes of exceljs bugs. The expected text of exceljs itself was observed with exceljs 4.4.0
// (bench/node_modules) on the same input.
import { describe, expect, it } from 'vitest';
import { csvValue, formatDate, parseDate } from '../../packages/exceljs/src/csv';
import { ExcelJS } from './helpers';

const text = async (wb: any, options?: unknown): Promise<string> => new TextDecoder('utf-8', { ignoreBOM: true }).decode(await wb.csv.writeBuffer(options));
const book = () => {
  const wb = new ExcelJS.Workbook();
  return { wb, ws: wb.addWorksheet('sheet1') };
};
const rows = (ws: any): unknown[][] => {
  const out: unknown[][] = [];
  ws.eachRow({ includeEmpty: true }, (row: any) => out.push(row.values.slice(1)));
  return out;
};

describe('csv.write', () => {
  it('writes the fields exceljs writes (quotes, links, JSON, errors, booleans, dates)', async () => {
    const { wb, ws } = book();
    ws.getCell('A1').value = 7;
    ws.getCell('B1').value = 'Hello, World!';
    ws.getCell('C1').value = new Date('2016-08-16T16:23:59.720Z');
    ws.getCell('D1').value = { formula: 'A1', result: 7 };
    ws.getCell('E1').value = { formula: 'A2' };
    ws.getCell('F1').value = { hyperlink: 'http://www.link.com', text: 'www.link.com' };
    ws.getCell('G1').value = '<a href="x">H&</a>';
    ws.getCell('H1').value = { num: 7 };
    ws.getCell('I1').value = true;
    ws.getCell('J1').value = false;
    ws.getCell('K1').value = { error: '#N/A' };
    ws.getCell('L1').value = 'a|b';
    // Dates are written in local time with the offset; the test runs in any time zone.
    const date = formatDate(new Date('2016-08-16T16:23:59.720Z'), undefined, false);
    expect(await text(wb)).toBe(`7,"Hello, World!",${date},7,,http://www.link.com,"<a href=""x"">H&</a>","{""num"":7}",true,false,#N/A,"a|b"`);
  });

  it('writes one blank line per missing row, also before the first row (exceljs writes one too few there)', async () => {
    const { wb, ws } = book();
    ws.getCell('A3').value = 1;
    ws.getCell('A6').value = 2;
    expect(await text(wb)).toBe('\n\n1\n\n\n2');
    expect(await text(wb, { includeEmptyRows: false })).toBe('1\n2');
  });

  it('writes a formula result of 0 or false, and the text of an error result (exceljs writes an empty field)', async () => {
    const { wb, ws } = book();
    ws.getCell('A1').value = { formula: '1-1', result: 0 };
    ws.getCell('B1').value = { formula: 'FALSE()', result: false };
    ws.getCell('C1').value = { formula: 'NA()', result: { error: '#N/A' } };
    expect(await text(wb)).toBe('0,false,#N/A');
  });

  it('escapes formula-like text by default and writes it unchanged with formulaInjection: "off"', async () => {
    const { wb, ws } = book();
    ws.addRow(['=cmd|calc', '+1', '@x', '-5', -5, 'ok']);
    expect(await text(wb)).toBe(`"'=cmd|calc","'+1","'@x","'-5",-5,ok`);
    expect(await text(wb, { formulaInjection: 'off' })).toBe('"=cmd|calc",+1,@x,-5,-5,ok');
    expect(await text(wb, { formatterOptions: { quote: false } })).toBe("'=cmd|calc,'+1,'@x,'-5,-5,ok");
    await expect(wb.csv.writeBuffer({ formulaInjection: 'maybe' })).rejects.toMatchObject({ code: 'XLV_E_INVALID_OPTION' });
  });

  it('follows dateFormat, dateUTC, map and the fast-csv formatter options', async () => {
    const { wb, ws } = book();
    ws.addRow(['a', new Date(Date.UTC(2024, 0, 31, 8, 5, 9, 7)), 'x y']);
    ws.addRow([1, 2]);
    expect(await text(wb, { dateFormat: 'DD/MM/YYYY HH:mm:ss.SSS ddd MMM A', dateUTC: true })).toBe('a,31/01/2024 08:05:09.007 Wed Jan AM,x y\n1,2');
    expect(await text(wb, { dateUTC: true })).toBe('a,2024-01-31T08:05:09Z,x y\n1,2');
    expect(await text(wb, { map: (v: unknown, i: number) => `${String(i)}:${String(v instanceof Date ? 'D' : v)}` })).toBe('0:a,1:D,2:x y\n0:1,1:2');
    expect(await text(wb, { formatterOptions: { delimiter: ';', rowDelimiter: '\r\n', quoteColumns: true, includeEndRowDelimiter: true, writeBOM: true }, dateUTC: true })).toBe(
      '\uFEFF"a";"2024-01-31T08:05:09Z";"x y"\r\n"1";"2"\r\n',
    );
    expect(await text(wb, { formatterOptions: { headers: ['h1', 'h2'] }, dateUTC: true })).toBe('h1,h2\na,2024-01-31T08:05:09Z\n1,2');
  });

  it('picks the worksheet by sheetName or sheetId, and names a missing one (exceljs writes an empty file)', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('One').addRow([1]);
    wb.addWorksheet('Two').addRow([2]);
    expect(await text(wb)).toBe('1');
    expect(await text(wb, { sheetName: 'Two' })).toBe('2');
    expect(await text(wb, { sheetId: 2 })).toBe('2');
    await expect(wb.csv.writeBuffer({ sheetName: 'Three' })).rejects.toMatchObject({ code: 'XLV_E_SHEET_NOT_FOUND' });
    expect(await text(new ExcelJS.Workbook())).toBe('');
  });

  it('writes to a Node.js-style stream and to a Web WritableStream', async () => {
    const { wb, ws } = book();
    ws.addRow(['a', 1]);
    // The two methods of a Node.js writable stream that csv.write uses.
    const chunks: Uint8Array[] = [];
    const node = {
      write(chunk: Uint8Array) {
        chunks.push(chunk);
        return true;
      },
      end(cb: () => void) {
        cb();
      },
    };
    await wb.csv.write(node);
    expect(chunks.map((c) => new TextDecoder().decode(c)).join('')).toBe('a,1');
    let web = '';
    await wb.csv.write(
      new WritableStream<Uint8Array>({
        write(chunk) {
          web += new TextDecoder().decode(chunk);
        },
      }),
    );
    expect(web).toBe('a,1');
  });
});

describe('csv.read', () => {
  it('adds a worksheet and converts fields as exceljs does', async () => {
    const wb = new ExcelJS.Workbook();
    // An async iterable of chunks, as a Node.js readable stream is.
    async function* chunks() {
      yield '1, 2 ,0x10,';
      yield new TextEncoder().encode('1e3,-Infinity,Infinity,TRUE,true,#N/A,x\n');
    }
    const ws = await wb.csv.read(chunks());
    expect(ws.name).toBe('sheet1');
    expect(rows(ws)).toEqual([[1, 2, 16, 1000, '-Infinity', 'Infinity', 'TRUE', true, { error: '#N/A' }, 'x']]);
  });

  it('keeps white space as text (exceljs reads " " as 0)', async () => {
    const ws = await new ExcelJS.Workbook().csv.read(['a, ,b\n']);
    expect(rows(ws)).toEqual([['a', ' ', 'b']]);
  });

  it('reads the default date formats strictly, in local time', async () => {
    const local = new Date(2016, 7, 16, 16, 23, 59);
    const ws = await new ExcelJS.Workbook().csv.read([`${formatDate(local, undefined, false)},2016-08-16T16:23:59,08-16-2016,2016-08-16,2016-8-16,2016-02-30\n`]);
    expect(rows(ws)).toEqual([[local, local, new Date(2016, 7, 16), new Date(2016, 7, 16), '2016-8-16', '2016-02-30']]);
  });

  it('handles quotes, empty lines and line ends as fast-csv does', async () => {
    const ws = await new ExcelJS.Workbook().csv.read(['a,b\r\n\r\n"q""x"," sp "\r\n,\n  \n"a" ,b\n"multi\nline",z']);
    expect(rows(ws)).toEqual([['a', 'b'], [], ['q"x', ' sp '], [], [], ['a', 'b'], ['multi\nline', 'z']]);
    await expect(new ExcelJS.Workbook().csv.read(['"open\n'])).rejects.toMatchObject({ code: 'XLV_E_CSV_MALFORMED' });
  });

  it('follows the fast-csv parser options', async () => {
    const read = async (input: string, parserOptions: Record<string, unknown>) => rows(await new ExcelJS.Workbook().csv.read([input], { parserOptions }));
    expect(await read('a;b\n', { delimiter: ';' })).toEqual([['a', 'b']]);
    expect(await read('a,b\n\n,,\n1,2', { ignoreEmpty: true })).toEqual([['a', 'b'], [1, 2]]);
    expect(await read('#c\na,b\n', { comment: '#' })).toEqual([['a', 'b']]);
    expect(await read('a\nb\nc\nd', { skipRows: 1, maxRows: 2 })).toEqual([['b'], ['c']]);
    expect(await read('x\ny\nz', { skipLines: 1 })).toEqual([['y'], ['z']]);
    expect(await read(' a , b \n', { trim: true })).toEqual([['a', 'b']]);
    expect(await read("'a,b',c\n", { quote: "'" })).toEqual([['a,b', 'c']]);
    expect(await read('"a,b",c\n', { quote: null })).toEqual([['"a', 'b"', 'c']]);
    // exceljs crashes with headers: true ("data.map is not a function"); xlvane skips the header line.
    expect(await read('a,b\n1,2\n', { headers: true })).toEqual([[1, 2]]);
    await expect(read('a', { delimiter: ';;' })).rejects.toMatchObject({ code: 'XLV_E_INVALID_OPTION' });
  });

  it('uses map, dateFormats and sheetName', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = await wb.csv.read(['31/01/2024 08:15,7\n'], { sheetName: 'Imported', dateFormats: ['DD/MM/YYYY HH:mm'], map: (v: string, i: number) => (i === 1 ? `#${v}` : csvValue(v, ['DD/MM/YYYY HH:mm'])) });
    expect(ws.name).toBe('Imported');
    expect(rows(ws)).toEqual([[new Date(2024, 0, 31, 8, 15), '#7']]);
  });

  it('reads Web ReadableStreams, byte arrays, a byte order mark and other encodings', async () => {
    const bytes = new TextEncoder().encode('\uFEFFzażółć,1\n');
    const ws = await new ExcelJS.Workbook().csv.read(new Blob([bytes]).stream());
    expect(rows(ws)).toEqual([['zażółć', 1]]);
    const latin1 = Uint8Array.from([0x63, 0x61, 0x66, 0xe9]);
    expect(rows(await new ExcelJS.Workbook().csv.read([latin1], { parserOptions: { encoding: 'latin1' } }))).toEqual([['café']]);
  });

  it('rejects createInputStream as exceljs 4 does', () => {
    expect(() => new ExcelJS.Workbook().csv.createInputStream()).toThrow(/deprecated/);
  });
});

describe('dates of the CSV reader', () => {
  it('parse strictly, with offsets, and give null when the text does not match', () => {
    expect(parseDate('2016-08-16T16:23:59+00:00', 'YYYY-MM-DD[T]HH:mm:ssZ')?.getTime()).toBe(
      // Strict mode compares with the local formatting, so an offset other than the local one does not match.
      formatDate(new Date('2016-08-16T16:23:59Z'), 'YYYY-MM-DD[T]HH:mm:ssZ', false) === '2016-08-16T16:23:59+00:00' ? Date.UTC(2016, 7, 16, 16, 23, 59) : undefined,
    );
    expect(parseDate('2016-08-16T16:23:59Z', 'YYYY-MM-DD[T]HH:mm:ssZ')).toBeNull();
    expect(parseDate('31 Jan 2024 10:30 pm', 'DD MMM YYYY hh:mm a')).toEqual(new Date(2024, 0, 31, 22, 30));
    expect(parseDate('99-01-02', 'YY-MM-DD')).toEqual(new Date(1999, 0, 2));
    expect(parseDate('2024-13-01', 'YYYY-MM-DD')).toBeNull();
  });
});

describe('review fixes (D-E06)', () => {
  it('parses dayjs formats with unbracketed letters and day names as exceljs does', () => {
    // Observed with exceljs 4.4.0 / dayjs 1.11.23: each of these gives a Date at 2020-01-02 (local).
    const local = (d: Date | null): string | null => (d === null ? null : `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()} ${d.getHours()}:${d.getMinutes()}:${d.getSeconds()}`);
    expect(local(parseDate('2020-01-02T03:04:05', 'YYYY-MM-DDTHH:mm:ss'))).toBe('2020-1-2 3:4:5');
    expect(local(parseDate('Thursday 2020-01-02', 'dddd YYYY-MM-DD'))).toBe('2020-1-2 0:0:0');
    expect(local(parseDate('Thu 2020-01-02', 'ddd YYYY-MM-DD'))).toBe('2020-1-2 0:0:0');
    // dayjs 1.11.23 without the matching plugins: invalid in strict mode.
    expect(parseDate('01/02/2020', 'L')).toBeNull();
    expect(parseDate('2 2020', 'Q YYYY')).toBeNull();
    expect(parseDate('2nd 2020-01', 'Do YYYY-MM')).toBeNull();
  });

  it('formats dates with the default pattern when dateFormat is an empty string', () => {
    const d = new Date(Date.UTC(2020, 0, 2, 3, 4, 5));
    expect(formatDate(d, '', true)).toBe(formatDate(d, undefined, true));
    expect(formatDate(d, '', true)).toBe('2020-01-02T03:04:05Z');
  });

  it('writes one field per header for array rows, as fast-csv does', async () => {
    const { wb, ws } = book();
    ws.addRow([1, 2, 3]);
    ws.addRow([4]);
    // exceljs 4.4.0 gives "a,b\n1,2\n4,".
    expect(await text(wb, { formatterOptions: { headers: ['a', 'b'] } })).toBe('a,b\n1,2\n4,');
    expect(await text(wb, { formatterOptions: { headers: ['=h', 'b'] } })).toBe(`"'=h",b\n1,2\n4,`);
  });

  it('names the line where an unclosed quote opens and refuses long quote and escape characters', async () => {
    const read = async (csv: string, parserOptions?: unknown) => {
      const { wb } = book();
      return wb.csv.read([new TextEncoder().encode(csv)], { parserOptions });
    };
    await expect(read('"abc\ndef\n')).rejects.toThrow(/CSV line 1:/);
    await expect(read('a,"b\nc')).rejects.toThrow(/CSV line 1:/);
    await expect(read('x\n"y')).rejects.toThrow(/CSV line 2:/);
    await expect(read('a', { quote: '""' })).rejects.toMatchObject({ code: 'XLV_E_INVALID_OPTION' });
    await expect(read('a', { escape: '\\\\' })).rejects.toMatchObject({ code: 'XLV_E_INVALID_OPTION' });
  });
});
