// @xlvane/compat CSV writing after the K2 penetration test: ways around the formula-injection
// escape. (1) A leading byte order mark: spreadsheet applications drop a BOM at the start of a
// file and then run the formula that follows (LibreOffice confirmed), and with latin1 or utf16le
// output a field can encode such a BOM ("ï»¿" is EF BB BF in latin1, U+FEFF is FF FE in
// UTF-16LE, U+FFFE gives the UTF-16BE BOM FE FF). (2) A `map` that returns no string (a String
// object, an array, an object with toString) was written as its text without being checked.
// (3) A comma (or ";", tab, "|") inside a field written with another delimiter: an application
// that splits on the comma gets a field starting with "=".
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import ExcelJS from '../../packages/compat/src/index';

const FORMULA = /^[=+\-@\t\r\n＝＋－＠﹦﹢﹣﹫]/;

/** RFC 4180 fields with the given delimiter (quotes doubled inside quoted fields). */
function parse(text: string, delimiter = ','): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i]!;
    if (quoted) {
      if (c === '"' && text[i + 1] === '"') {
        field += '"';
        i++;
      } else if (c === '"') quoted = false;
      else field += c;
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
  }
  if (field !== '' || row.length > 0) rows.push([...row, field]);
  return rows;
}

/**
 * The text a reader that honours byte order marks gets from the bytes: a UTF-8, UTF-16LE or
 * UTF-16BE BOM at the start selects that encoding and is dropped; otherwise `fallback`. Leading
 * U+FEFF characters left after that are dropped as well, as LibreOffice does.
 */
function readerText(bytes: Uint8Array, fallback: string): string {
  let text: string;
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) text = new TextDecoder('utf-8', { ignoreBOM: true }).decode(bytes.subarray(3));
  else if (bytes[0] === 0xff && bytes[1] === 0xfe) text = new TextDecoder('utf-16le', { ignoreBOM: true }).decode(bytes.subarray(2));
  else if (bytes[0] === 0xfe && bytes[1] === 0xff) text = new TextDecoder('utf-16be', { ignoreBOM: true }).decode(bytes.subarray(2));
  else text = new TextDecoder(fallback, { ignoreBOM: true }).decode(bytes);
  return text.replace(/^﻿+/, '');
}

/** Every field of every row a BOM-aware reader sees, checked against formula starts (also after a leading BOM). */
function expectNoFormula(bytes: Uint8Array, fallback = 'utf-8', delimiter = ','): string[][] {
  const rows = parse(readerText(bytes, fallback), delimiter);
  for (const row of rows) {
    for (const f of row) {
      expect(FORMULA.test(f), JSON.stringify(f)).toBe(false);
      expect(FORMULA.test(f.replace(/^[﻿￾]+/, '')), JSON.stringify(f)).toBe(false);
    }
  }
  return rows;
}

async function inDir<T>(fn: (dir: string) => Promise<T>): Promise<T> {
  const dir = await mkdtemp(join(tmpdir(), 'xlvane-k2-csv-'));
  try {
    return await fn(dir);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

function book(values: unknown[]): InstanceType<typeof ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('S').addRow(values as never);
  return wb;
}

async function buffer(wb: InstanceType<typeof ExcelJS.Workbook>, options: Record<string, unknown> = {}): Promise<Uint8Array> {
  return (await wb.csv.writeBuffer(options as never)) as Uint8Array;
}

async function file(wb: InstanceType<typeof ExcelJS.Workbook>, options: Record<string, unknown>): Promise<Uint8Array> {
  return inDir(async (dir) => {
    await wb.csv.writeFile(join(dir, 'a.csv'), options as never);
    return new Uint8Array(await readFile(join(dir, 'a.csv')));
  });
}

describe('K2 item 1: a leading byte order mark', () => {
  const BOM_VALUES = ['﻿=1+1', '￾=1+1', '﻿﻿@SUM(1)', '\u0000﻿=1+1'];

  for (const value of BOM_VALUES) {
    it(`writeBuffer escapes ${JSON.stringify(value)}`, async () => {
      const bytes = await buffer(book([value, 'x']));
      const rows = expectNoFormula(bytes);
      expect(rows[0]![0]!.startsWith("'")).toBe(true);
    });
  }

  it('writeBuffer with writeBOM escapes a field that starts with a BOM', async () => {
    expectNoFormula(await buffer(book(['﻿=1+1', 'x']), { formatterOptions: { writeBOM: true } }));
  });

  it('write(stream) escapes it', async () => {
    const chunks: Uint8Array[] = [];
    const stream = new WritableStream<Uint8Array>({ write: (c) => void chunks.push(c) });
    await book(['﻿=1+1', 'x']).csv.write(stream);
    expectNoFormula(new Uint8Array(chunks.flatMap((c) => [...c])));
  });

  it('writeFile (utf8) escapes it', async () => {
    expectNoFormula(await file(book(['﻿=1+1', 'x']), {}));
  });

  it('escapes a header that starts with a BOM', async () => {
    expectNoFormula(await buffer(book(['a']), { formatterOptions: { headers: ['﻿=1+1'] } }));
  });

  it('escapes hyperlink text that starts with a BOM', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S').getCell('A1').value = { text: 'x', hyperlink: '﻿=1+1' };
    expectNoFormula(await buffer(wb));
  });

  it('latin1: a field whose bytes are a UTF-8 BOM followed by "=" is escaped', async () => {
    const bytes = await file(book(['ï»¿=1+1', 'x']), { encoding: 'latin1' });
    expect([...bytes.subarray(0, 3)]).not.toEqual([0xef, 0xbb, 0xbf]);
    expectNoFormula(bytes, 'latin1');
  });

  for (const value of ['ÿþ=\u00001', 'þÿ\u0000=']) {
    it(`latin1: a field whose bytes are a UTF-16 BOM (${JSON.stringify(value)}) is escaped`, async () => {
      const bytes = await file(book([value, 'x']), { encoding: 'latin1' });
      expect(bytes[0]).toBe(0x22);
      expect(parse(new TextDecoder('latin1').decode(bytes))[0]).toEqual([`'${value.replace(/\0/g, '')}`, 'x']);
    });
  }

  it('latin1: a field whose bytes are UTF-8 for a full-width "=" is escaped', async () => {
    // EF BC 9D is U+FF1D in UTF-8, which readers that guess UTF-8 show as a formula sign.
    const bytes = await file(book(['ï¼\u009D1+1', 'x']), { encoding: 'latin1' });
    expect(parse(new TextDecoder('latin1').decode(bytes))[0]).toEqual(["'ï¼\u009D1+1", 'x']);
    expect(FORMULA.test(new TextDecoder('utf-8').decode(bytes))).toBe(false);
  });

  it('latin1 with writeBOM: a first field starting with "þ" (the BOM is written as FF) is escaped', async () => {
    const bytes = await file(book(['þ=1', 'x']), { encoding: 'latin1', formatterOptions: { writeBOM: true } });
    expect([...bytes.subarray(0, 2)]).not.toEqual([0xff, 0xfe]);
  });

  it('latin1 keeps ordinary text', async () => {
    const bytes = await file(book(['café', 'ÿes', 'þing']), { encoding: 'latin1' });
    expect(new TextDecoder('latin1').decode(bytes)).toBe('café,ÿes,þing');
  });

  it('utf16le: U+FEFF (FF FE) and U+FFFE (FE FF, the UTF-16BE BOM) at the start of a field are escaped', async () => {
    const le = await file(book(['﻿=1+1', 'x']), { encoding: 'utf16le' });
    expectNoFormula(le, 'utf-16le');
    expect(parse(new TextDecoder('utf-16le').decode(le))[0]).toEqual(["'\uFEFF=1+1", 'x']);
    // U+3D00 U+3100 … are "=1…" read as UTF-16BE after the FE FF of U+FFFE.
    const be = await file(book(['￾㴀㄀⬀㄀', 'x']), { encoding: 'utf16le' });
    expectNoFormula(be, 'utf-16le');
    expect([...be.subarray(0, 2)]).not.toEqual([0xfe, 0xff]);
    expect(parse(new TextDecoder('utf-16le').decode(be))[0]![0]!.startsWith("'\uFFFE")).toBe(true);
  });

  it('utf16le: code units whose bytes are a UTF-8 BOM followed by "=" are escaped', async () => {
    // U+BBEF U+3DBF U+2B31 U+0031: EF BB BF 3D 31 2B 31 00.
    const bytes = await file(book(['믯㶿⬱1', 'x']), { encoding: 'utf16le' });
    expect([...bytes.subarray(0, 3)]).not.toEqual([0xef, 0xbb, 0xbf]);
  });

  it('utf16le keeps CJK text', async () => {
    const bytes = await file(book(['中国', 'x']), { encoding: 'utf16le' });
    expect(new TextDecoder('utf-16le').decode(bytes)).toBe('中国,x');
  });

  it('formulaInjection "off" writes the BOM unchanged, as exceljs', async () => {
    expect(new TextDecoder('utf-8', { ignoreBOM: true }).decode(await buffer(book(['﻿=1+1', 'x']), { formulaInjection: 'off' }))).toBe('﻿=1+1,x');
  });
});

describe('K2 item 2: map returning values that are not strings', () => {
  const cases: [string, (v: unknown) => unknown][] = [
    ['a String object', () => new String('=1+1')],
    // Three items: a row whose first field is an array of two is a row of [header, value] pairs
    // for fast-csv (@fast-csv/format/build/src/formatter/RowFormatter.js L26-31), whose values
    // are written instead (F9, tests/exceljs/l1-fixes.test.ts).
    ['an array', () => ['=1+1', 'x', 'y']],
    ['an object with toString', () => ({ toString: () => '@cmd' })],
    ['a String object with a BOM', () => new String('﻿=1+1')],
  ];
  for (const [name, map] of cases) {
    it(`escapes ${name}`, async () => {
      const rows = expectNoFormula(await buffer(book(['a']), { map }));
      expect(rows[0]![0]!.startsWith("'")).toBe(true);
    });
  }

  it('latin1: an object whose text narrows to "=" is escaped', async () => {
    const bytes = await file(book(['a']), { encoding: 'latin1', map: () => ({ toString: () => 'Ľ1+1' }) });
    expectNoFormula(bytes, 'latin1');
    expect(new TextDecoder('latin1').decode(bytes)).toBe('?1+1');
  });

  it('calls toString once', async () => {
    let calls = 0;
    const text = new TextDecoder().decode(
      await buffer(book(['a']), {
        map: () => ({
          toString: () => (calls++ === 0 ? 'safe' : '=1+1'),
        }),
      }),
    );
    expect(text).toBe('safe');
    expect(calls).toBe(1);
  });

  it('keeps numbers, booleans and empty values as before (negative numbers are not escaped, CSV-11)', async () => {
    const values = [-5, true, null, undefined, 0];
    const text = new TextDecoder().decode(await buffer(book(['a', 'b', 'c', 'd', 'e']), { map: (_v: unknown, i: number) => values[i] }));
    expect(text).toBe('-5,true,,,0');
  });
});

describe('K2 item 3: a comma or another common delimiter inside a field', () => {
  for (const delimiter of [';', '\t', '|']) {
    it(`quotes "x,=1+1" with delimiter ${JSON.stringify(delimiter)}`, async () => {
      const text = new TextDecoder().decode(await buffer(book(['x,=1+1', 'y']), { formatterOptions: { delimiter } }));
      expect(text).toBe(`"x,=1+1"${delimiter}y`);
      // A reader that splits on commas, semicolons, tabs or bars gets no field starting with "=".
      for (const d of [',', ';', '\t', '|']) for (const row of parse(text, d)) for (const f of row) expect(FORMULA.test(f), `${d}: ${f}`).toBe(false);
    });
  }

  for (const inner of [';', '\t', '|']) {
    it(`quotes a field holding ${JSON.stringify(inner)} with the comma delimiter`, async () => {
      const text = new TextDecoder().decode(await buffer(book([`x${inner}=1+1`, 'y'])));
      expect(text).toBe(`"x${inner}=1+1",y`);
    });
  }

  for (const value of ['x,=1+1', 'x; @SUM(1)', 'x|\t-1', 'x,"=1+1"', 'x, " +1"', 'x;\uFEFF=1']) {
    it(`refuses ${JSON.stringify(value)} with quote: false`, async () => {
      await expect(buffer(book([value, 'y']), { formatterOptions: { delimiter: '\t', quote: false } })).rejects.toMatchObject({ code: 'XLV_E_INVALID_OPTION', details: { option: 'quote' } });
    });
  }

  it('writes other fields with these characters unquoted with quote: false', async () => {
    const text = new TextDecoder().decode(await buffer(book(['Hello, World!', 'a;b', '{"x":1,"y":"z"}', 'x,']), { formatterOptions: { delimiter: '\t', quote: false } }));
    expect(text).toBe('Hello, World!\ta;b\t{"x":1,"y":"z"}\tx,');
  });

  it('keeps fast-csv behaviour with formulaInjection "off"', async () => {
    const text = new TextDecoder().decode(await buffer(book(['x,=1+1', 'y']), { formulaInjection: 'off', formatterOptions: { delimiter: ';' } }));
    expect(text).toBe('x,=1+1;y');
    const plain = new TextDecoder().decode(await buffer(book(['x,y', 'z']), { formulaInjection: 'off', formatterOptions: { delimiter: ';', quote: false } }));
    expect(plain).toBe('x,y;z');
  });
});
