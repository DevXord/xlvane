import { describe, expect, it } from 'vitest';
import { fromCSV, toCSV } from '../../packages/xlvane/src/csv';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { readTable } from '../../packages/xlvane/src/table';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';

function book(): Uint8Array {
  const wb = createWorkbook();
  const s = wb.addSheet('Data');
  s.set('A1', 'name').set('B1', 'amount').set('C1', 'when').set('D1', 'ok').set('E1', 'note');
  s.set('A2', 'Alice').set('B2', 10.5).set('C2', new Date(Date.UTC(2024, 0, 31))).set('D2', true).set('E2', 'say "hi", ok');
  s.set('A3', 'Bob').set('B3', -3).set('C3', new Date(Date.UTC(2024, 1, 29, 13, 45, 30))).set('D3', false).set('E3', 'two\nlines');
  s.set('B5', { formula: 'B2+B3', result: 7.5 }).set('D5', { error: '#N/A' });
  wb.addSheet('Other').set('B2', 'x');
  return writeWorkbook(wb);
}

describe('toCSV', () => {
  it('writes the first visible sheet from A1 with RFC 4180 quoting', () => {
    expect(toCSV(book())).toBe(
      [
        'name,amount,when,ok,note',
        'Alice,10.5,2024-01-31,TRUE,"say ""hi"", ok"',
        'Bob,-3,2024-02-29T13:45:30,FALSE,"two\nlines"',
        ',,,,',
        ',7.5,,#N/A,',
        '',
      ].join('\r\n'),
    );
  });

  it('keeps empty leading rows and columns so positions match the sheet', () => {
    expect(toCSV(book(), { sheet: 'Other', newline: '\n' })).toBe(',\n,x\n');
  });

  it('supports other delimiters, a byte order mark and a range', () => {
    const csv = toCSV(book(), { delimiter: ';', bom: true, range: 'A1:B2' });
    expect(csv).toBe('﻿name;amount\r\nAlice;10.5\r\n');
    expect(toCSV(book(), { delimiter: '\t', range: 'E2' })).toBe('"say ""hi"", ok"\r\n'); // quotes force quoting
  });

  it('escapes text that a spreadsheet would run as a formula', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('S');
    const risky = ['=1+2', '+1', '-1', '@SUM(A1)', '\t=x', '\r=x', '\n=x', '＝1', '＋1', '－1', '＠x', '=1+2";=1+2'];
    risky.forEach((text, i) => s.set(`A${String(i + 1)}`, text));
    s.set('B1', -5).set('B2', { formula: 'A1', result: '=evil()' });
    const csv = toCSV(writeWorkbook(wb), { newline: '\n' });
    const rows = csv.split('\n');
    expect(rows[0]).toBe(`"'=1+2",-5`);
    expect(rows[1]).toBe(`"'+1","'=evil()"`);
    expect(rows[2]).toBe(`"'-1",`);
    expect(rows[3]).toBe(`"'@SUM(A1)",`);
    expect(rows[4]).toBe(`"'\t=x",`);
    expect(csv).toContain(`\n"'\r=x",\n"'\n=x",\n"'＝1",\n"'＋1",\n"'－1",\n"'＠x",\n`);
    expect(csv).toContain(`\n"'=1+2"";=1+2",\n`);
    const off = toCSV(writeWorkbook(wb), { newline: '\n', formulaInjection: 'off' }).split('\n');
    expect(off[0]).toBe('=1+2,-5');
  });

  it('validates options', () => {
    expectXlvError(() => toCSV(book(), { delimiter: '"' }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => toCSV(book(), { delimiter: ',,' }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => toCSV(book(), { newline: '\r' as never }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => toCSV(book(), { formulaInjection: 'no' as never }), 'XLV_E_INVALID_OPTION');
  });
});

describe('fromCSV', () => {
  it('reads a header and infers numbers, booleans and dates', () => {
    const bytes = fromCSV('id,name,price,active,since,zip\r\n1,"Smith, J.",9.99,TRUE,2024-01-31,00501\r\n2,"say ""hi""",-1e3,false,2024-02-29T13:45:30,\r\n');
    const t = readTable(bytes);
    expect(t.rows).toEqual([
      { id: 1, name: 'Smith, J.', price: 9.99, active: true, since: new Date(Date.UTC(2024, 0, 31)), zip: '00501' },
      { id: 2, name: 'say "hi"', price: -1000, active: false, since: new Date(Date.UTC(2024, 1, 29, 13, 45, 30)), zip: null },
    ]);
  });

  it('keeps everything as text with types: "text"', () => {
    const t = readTable(fromCSV('a,b\n1,TRUE\n', { types: 'text' }));
    expect(t.rows).toEqual([{ a: '1', b: 'TRUE' }]);
  });

  it('never turns text into formulas', () => {
    const s = readWorkbook(fromCSV('a\n=HYPERLINK("http://x")\n')).getSheet('Sheet1');
    expect(s.cell('A2')?.value).toBe('=HYPERLINK("http://x")');
    expect(s.cell('A2')?.formula).toBeNull();
  });

  it('works without a header, with other delimiters, CRLF or LF, a BOM and bytes', () => {
    const s = readWorkbook(fromCSV(new TextEncoder().encode('﻿x;1\ny;2'), { header: false, delimiter: ';', sheet: 'Raw' })).getSheet('Raw');
    expect(s.cell('A1')?.value).toBe('x');
    expect(s.cell('B2')?.value).toBe(2);
    expect(s.cell('A1')?.style).toEqual({});
  });

  it('accepts ragged rows and quoted line breaks', () => {
    const t = readTable(fromCSV('a,b,c\n1\n"x\ny",2,3,4\n'), { header: 1 });
    expect(t.rows).toEqual([
      { a: 1, b: null, c: null, D: null },
      { a: 'x\ny', b: 2, c: 3, D: 4 },
    ]);
  });

  it('reports malformed CSV with its line and column', () => {
    expectXlvError(() => fromCSV('a,b\n"open,2\n'), 'XLV_E_CSV_MALFORMED', { line: 2 });
    expectXlvError(() => fromCSV('a,b\n"x"y,2\n'), 'XLV_E_CSV_MALFORMED', { line: 2, column: 4 });
    expectXlvError(() => fromCSV(new Uint8Array([0x61, 0xff, 0x0a])), 'XLV_E_CSV_MALFORMED');
  });

  it('round-trips with toCSV', () => {
    const csv = toCSV(fromCSV('h1,h2\r\n"a,b",1.5\r\n"line\nbreak",\r\n', { types: 'text' }), { formulaInjection: 'off' });
    expect(csv).toBe('h1,h2\r\n"a,b",1.5\r\n"line\nbreak",\r\n');
  });
});
