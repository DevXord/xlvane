// @xlvane/compat csv.writeFile after the K3 review of the K2 fixes. (1) utf16le output: LibreOffice
// 24.2.7.2 reads a UTF-16LE CSV file as 8-bit text and drops its NUL bytes, with or without the
// UTF-16LE byte order mark (FF FE; checked with the default import and with charset 76, UTF-8;
// only charset 65535 decodes UTF-16), so a cell "㴀㄀⬀㄀" (bytes 00 3D 00 31 00 2B 00 31)
// imported as the formula =1+1 in any field and row. The bytes of other characters can also give
// a line break, a delimiter or a quote there ("ਗ=1+1" is 17 0A then "=1+1"). While formulas are
// escaped, utf16le output now starts with FF FE (readers that honour it decode UTF-16), a field
// whose bytes without NULs start with "=" is escaped, and a field whose bytes hold a delimiter,
// line break or quote followed by "=" is refused. (3) The byte order mark check looked at one
// field at a time, so the bytes of a BOM could span a field boundary: with an empty delimiter
// (latin1 ["ï", "»¿=1+1"], utf16le ["믯", "㶿1"]) or a delimiter holding bytes of a BOM (latin1
// "ï»" with [null, "¿=1+1"]) LibreOffice ran the formula. An empty delimiter, and a delimiter or
// row delimiter with characters other than ASCII in latin1 or utf16le output, are now refused
// while formulas are escaped.
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import ExcelJS from '../../packages/compat/src/index';

type Book = InstanceType<typeof ExcelJS.Workbook>;

function book(rows: unknown[][]): Book {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  for (const row of rows) ws.addRow(row as never);
  return wb;
}

async function file(wb: Book, options: Record<string, unknown>): Promise<Uint8Array> {
  const dir = await mkdtemp(join(tmpdir(), 'xlvane-k3-csv-'));
  try {
    await wb.csv.writeFile(join(dir, 'a.csv'), options as never);
    return new Uint8Array(await readFile(join(dir, 'a.csv')));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

async function refusal(promise: Promise<unknown>): Promise<unknown> {
  return promise.then(
    () => null,
    (e: unknown) => e,
  );
}

const FORMULA_START = /^[=+\-@\uFF1D\uFF0B\uFF0D\uFF20\uFE66\uFE62\uFE63\uFE6B]/;

/**
 * The texts a reader can get from the bytes: a UTF-8, UTF-16LE or UTF-16BE byte order mark at the
 * very start selects that encoding (LibreOffice 24.2 acts on a mark only there); without one, the
 * bytes read as UTF-8 and as 8-bit text.
 */
function readerTexts(bytes: Uint8Array): string[] {
  const decode = (label: string, from: number): string => new TextDecoder(label, { ignoreBOM: true }).decode(bytes.subarray(from));
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return [decode('utf-8', 3)];
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return [decode('utf-16le', 2)];
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return [decode('utf-16be', 2)];
  return [decode('utf-8', 0), decode('latin1', 0)];
}

/**
 * The fields that LibreOffice 24.2.7.2 imports as formulas from these bytes read as 8-bit text with
 * a comma delimiter (its default import and charset 76): NUL bytes are dropped; an unquoted field
 * is a formula when it starts with "=" and has more; a field may open a quote after spaces, which
 * then closes only at a quote followed by a comma, a line break or the end (a quote never closed
 * is text, and the field ends at the next comma or line break), and is a formula when its text
 * starts with "=" (each rule checked with LibreOffice on hand-made files; on 47 generated files
 * this function found every formula LibreOffice made, and two more in quoted fields that
 * LibreOffice kept as text).
 */
function libreOfficeFormulas(bytes: Uint8Array): string[] {
  let text = new TextDecoder('latin1').decode(bytes.filter((b) => b !== 0));
  if (text.startsWith('ï»¿')) text = text.slice(3);
  const out: string[] = [];
  let i = 0;
  const n = text.length;
  const end = (j: number): boolean => j >= n || text[j] === ',' || text[j] === '\n' || text[j] === '\r';
  while (i <= n) {
    let j = i;
    while (text[j] === ' ') j++;
    let value = '';
    let closed = false;
    if (text[j] === '"') {
      j++;
      for (;;) {
        if (j >= n) break;
        if (text[j] === '"' && text[j + 1] === '"') {
          value += '"';
          j += 2;
        } else if (text[j] === '"' && end(j + 1)) {
          j++;
          closed = true;
          break;
        } else value += text[j++];
      }
      if (closed) while (!end(j)) value += text[j++];
    }
    // A quote that is never closed is text, and the field ends at the next comma or line break.
    if (!closed) {
      value = '';
      j = i;
      while (!end(j)) value += text[j++];
    }
    // "=" alone stays text.
    if (value.length > 1 && value.startsWith('=')) out.push(value);
    if (j >= n) break;
    i = text[j] === '\r' && text[j + 1] === '\n' ? j + 2 : j + 1;
  }
  return out;
}

/** The fields of the text after a UTF-16LE byte order mark, as RFC 4180 reads them, that start like a formula. */
function utf16Formulas(bytes: Uint8Array): string[] {
  expect([...bytes.subarray(0, 2)]).toEqual([0xff, 0xfe]);
  const text = new TextDecoder('utf-16le', { ignoreBOM: true }).decode(bytes.subarray(2));
  const out: string[] = [];
  let field = '';
  let quoted = false;
  let started = false;
  for (let i = 0; i <= text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === '"' && text[i + 1] === '"') {
        field += '"';
        i++;
      } else if (c === '"') quoted = false;
      else if (c !== undefined) field += c;
      continue;
    }
    if (c === '"' && !started) {
      quoted = true;
      started = true;
    } else if (c === undefined || c === ',' || c === '\n' || c === '\r') {
      if (FORMULA_START.test(field.replace(/^﻿+/, ''))) out.push(field);
      field = '';
      started = false;
    } else {
      field += c;
      started = true;
    }
  }
  return out;
}

describe('K3 item 1: utf16le output', () => {
  const CELL = '㴀㄀⬀㄀';

  it('the repro: LibreOffice would run "㴀㄀⬀㄀" in any field without the fix', async () => {
    const plain = await file(book([[CELL, 'x'], ['a', CELL]]), { encoding: 'utf16le', formulaInjection: 'off' });
    expect(libreOfficeFormulas(plain)).toEqual(['=1+1', '=1+1']);
  });

  it('starts with FF FE and escapes every field whose bytes without NULs start with "=" (or a quote and "=")', async () => {
    const bytes = await file(book([[CELL, 'x'], ['a', CELL], ['丽', ' 㴀1', '∀㴀1']]), { encoding: 'utf16le' });
    expect([...bytes.subarray(0, 2)]).toEqual([0xff, 0xfe]);
    expect(new TextDecoder('utf-16le', { ignoreBOM: true }).decode(bytes.subarray(2))).toBe(`"'${CELL}",x\na,"'${CELL}"\n"'丽", 㴀1,"'∀㴀1"`);
    expect(libreOfficeFormulas(bytes)).toEqual([]);
    expect(utf16Formulas(bytes)).toEqual([]);
  });

  for (const [name, rows] of [
    ['a line break byte (U+0A17) before "="', [['ਗ=1+1']]],
    ['a comma byte (U+2C0C) before "="', [['x', 'Ⰼ=1+1']]],
    ['a quote byte (U+2200) that closes a quoted field before ",="', [['a∀,=1+1', 'y']]],
    ['a quote byte in an earlier field that opens a quote in the reader', [['∀x', ',=1+1']]],
    ['a quote byte in an earlier row', [['∀x'], ['y', ',=1+1']]],
    ['a code unit whose bytes are a line break and "=" (U+3D0A)', [['x㴊1+1']]],
  ] as const) {
    it(`refuses a field with ${name}`, async () => {
      const wb = book(rows.map((r) => [...r]));
      expect(libreOfficeFormulas(await file(wb, { encoding: 'utf16le', formulaInjection: 'off' })).length).toBeGreaterThan(0);
      const error = await refusal(file(wb, { encoding: 'utf16le' }));
      expect(error).toMatchObject({ code: 'XLV_E_INVALID_OPTION', details: { option: 'encoding' }, location: { sheet: 'S' } });
      expect((error as { hint?: string }).hint).toContain('UTF-8');
    });
  }

  it('refuses a quoted first field with a delimiter or line break before "=", which the mark makes LibreOffice split', async () => {
    // The mark's bytes come before the opening quote of the first field when LibreOffice reads
    // the file as 8-bit text (FF FE " x , = ...), so the quote does not quote (checked).
    const marked = await file(book([['x,=1+1', 'y']]), { encoding: 'utf16le', formulaInjection: 'off', formatterOptions: { writeBOM: true } });
    expect(libreOfficeFormulas(marked)).toEqual(['=1+1"']);
    for (const value of ['x,=1+1', 'x\n=1+1', 'x|=1']) {
      const error = await refusal(file(book([[value, 'y']]), { encoding: 'utf16le' }));
      expect(error).toMatchObject({ code: 'XLV_E_INVALID_OPTION', details: { option: 'encoding', field: 1, row: 1 }, location: { sheet: 'S', cell: 'A1' } });
      expect((error as Error).message).toContain('byte order mark');
      const header = await refusal(file(book([['y']]), { encoding: 'utf16le', formatterOptions: { headers: [value] } }));
      expect(header).toMatchObject({ code: 'XLV_E_INVALID_OPTION', details: { option: 'encoding', field: 1, row: 0 } });
    }
    // Elsewhere the same text is quoted as usual.
    const bytes = await file(book([['y', 'x,=1+1'], ['x,=1+1']]), { encoding: 'utf16le' });
    expect(new TextDecoder('utf-16le', { ignoreBOM: true }).decode(bytes.subarray(2))).toBe('y,"x,=1+1"\n"x,=1+1"');
    expect(libreOfficeFormulas(bytes)).toEqual([]);
  });

  it('latin1 with writeBOM (the byte FF) refuses such a first field too', async () => {
    const plain = await file(book([['x,=1+1', 'y']]), { encoding: 'latin1', formulaInjection: 'off', formatterOptions: { writeBOM: true } });
    expect(libreOfficeFormulas(plain)).toEqual(['=1+1"']);
    const error = await refusal(file(book([['x,=1+1', 'y']]), { encoding: 'latin1', formatterOptions: { writeBOM: true } }));
    expect(error).toMatchObject({ code: 'XLV_E_INVALID_OPTION', details: { option: 'encoding' }, location: { sheet: 'S', cell: 'A1' } });
    expect((error as { hint?: string }).hint).toContain('writeBOM');
    // Without writeBOM, or in UTF-8 (whose mark EF BB BF LibreOffice drops), it is written.
    expect(new TextDecoder('latin1').decode(await file(book([['x,=1+1', 'y']]), { encoding: 'latin1' }))).toBe('"x,=1+1",y');
    expect(libreOfficeFormulas(await file(book([['x,=1+1', 'y']]), { formatterOptions: { writeBOM: true } }))).toEqual([]);
  });

  it('keeps CJK text and symbols whose bytes hold such characters, but no "=" after them', async () => {
    const row = ['中国', '上海', '丢失', '米中', '丬', '∀x ≤ 1', '−5', 'x = 1', 'a∀,b'];
    const bytes = await file(book([row, ['∀'], ['a,b', 'c=d']]), { encoding: 'utf16le' });
    expect(new TextDecoder('utf-16le', { ignoreBOM: true }).decode(bytes.subarray(2))).toBe(`${row.map((v) => (v.includes(',') ? `"${v}"` : v)).join(',')}\n∀\n"a,b",c=d`);
    expect(libreOfficeFormulas(bytes)).toEqual([]);
  });

  it('random fields are either refused or give LibreOffice and UTF-16 readers no formula', async () => {
    const POOL = ['a', ' ', ',', '"', '\n', '=', '㴀', '丽', '∀', 'ਗ', '上', '丬', '⨽', '㴊', '中', '−', 'Ⰼ'];
    let seed = 12345;
    const next = (k: number): number => {
      seed = (seed * 1103515245 + 12345) % 2147483648;
      return seed % k;
    };
    let written = 0;
    for (let t = 0; t < 400; t++) {
      const rows = Array.from({ length: 1 + next(3) }, () => Array.from({ length: 1 + next(3) }, () => Array.from({ length: next(5) }, () => POOL[next(POOL.length)]).join('')));
      let bytes: Uint8Array;
      try {
        bytes = await file(book(rows), { encoding: 'utf16le' });
      } catch (e) {
        expect(e, JSON.stringify(rows)).toMatchObject({ code: 'XLV_E_INVALID_OPTION', details: { option: 'encoding' } });
        continue;
      }
      written++;
      expect(libreOfficeFormulas(bytes), JSON.stringify(rows)).toEqual([]);
      expect(utf16Formulas(bytes), JSON.stringify(rows)).toEqual([]);
    }
    // The pool is mostly characters with such bytes; about half of the files are written.
    expect(written).toBeGreaterThan(150);
  });

  it('writes the mark once with writeBOM: true', async () => {
    const bytes = await file(book([['a']]), { encoding: 'utf16le', formatterOptions: { writeBOM: true } });
    expect([...bytes]).toEqual([0xff, 0xfe, 0x61, 0x00]);
  });

  it('writes the mark for an empty worksheet and for every utf16le spelling', async () => {
    for (const encoding of ['utf16le', 'utf-16le', 'UCS2', 'ucs-2']) {
      expect([...(await file(book([]), { encoding }))]).toEqual([0xff, 0xfe]);
    }
  });

  it('a field starting with U+FEFF after the mark is still escaped', async () => {
    const bytes = await file(book([['﻿=1+1']]), { encoding: 'utf16le' });
    expect(new TextDecoder('utf-16le', { ignoreBOM: true }).decode(bytes.subarray(2))).toBe('"\'﻿=1+1"');
  });

  it('keeps exceljs output (no mark, nothing escaped or refused) with formulaInjection "off", and utf8 and latin1 output unchanged', async () => {
    expect([...(await file(book([['a']]), { encoding: 'utf16le', formulaInjection: 'off' }))]).toEqual([0x61, 0x00]);
    expect(new TextDecoder('utf-16le').decode(await file(book([['ਗ=1', CELL]]), { encoding: 'utf16le', formulaInjection: 'off' }))).toBe(`ਗ=1,${CELL}`);
    expect([...(await file(book([['a']]), {}))]).toEqual([0x61]);
    expect([...(await file(book([['a']]), { encoding: 'latin1' }))]).toEqual([0x61]);
  });
});

describe('K3 item 3: byte order marks across field boundaries', () => {
  for (const [encoding, row] of [
    ['latin1', ['ï', '»¿=1+1']],
    ['utf16le', ['믯', '㶿1']],
    ['utf8', ['a', '=1']],
  ] as const) {
    it(`refuses an empty delimiter (${encoding})`, async () => {
      const error = await refusal(file(book([[...row]]), { encoding, formatterOptions: { delimiter: '' } }));
      expect(error).toMatchObject({ code: 'XLV_E_INVALID_OPTION', details: { option: 'delimiter' } });
      expect((error as { hint?: string }).hint).toContain('formulaInjection');
    });
  }

  it('refuses an empty delimiter in writeBuffer too', async () => {
    await expect(book([['a', 'b']]).csv.writeBuffer({ formatterOptions: { delimiter: '' } } as never)).rejects.toMatchObject({ code: 'XLV_E_INVALID_OPTION', details: { option: 'delimiter' } });
  });

  for (const [encoding, delimiter] of [
    ['latin1', 'ï»'],
    ['latin1', '§'],
    ['latin1', 'a '],
    ['utf16le', '中'],
    ['utf16le', '¦'],
  ] as const) {
    it(`refuses the delimiter ${JSON.stringify(delimiter)} with ${encoding} output`, async () => {
      const error = await refusal(file(book([[null, '¿=1+1']]), { encoding, formatterOptions: { delimiter } }));
      expect(error).toMatchObject({ code: 'XLV_E_INVALID_OPTION', details: { option: 'delimiter' } });
      expect((error as Error).message).toContain('ASCII');
      expect((error as { hint?: string }).hint).toMatch(/ASCII|utf8/);
    });
  }

  for (const encoding of ['latin1', 'utf16le']) {
    it(`refuses a row delimiter with characters other than ASCII with ${encoding} output`, async () => {
      const error = await refusal(file(book([['a'], ['b']]), { encoding, formatterOptions: { rowDelimiter: '\n\u0085' } }));
      expect(error).toMatchObject({ code: 'XLV_E_INVALID_OPTION', details: { option: 'rowDelimiter' } });
      expect((error as Error).message).toContain('ASCII');
    });
  }

  it('keeps ASCII delimiters with latin1 and utf16le, and other delimiters with utf8', async () => {
    expect(new TextDecoder('latin1').decode(await file(book([['a', 'b'], ['c', 'd']]), { encoding: 'latin1', formatterOptions: { delimiter: ';', rowDelimiter: '\r\n' } }))).toBe('a;b\r\nc;d');
    expect(new TextDecoder('utf-16le').decode(await file(book([['a', 'b']]), { encoding: 'utf16le', formatterOptions: { delimiter: '|' } }))).toBe('a|b');
    expect(new TextDecoder().decode(await file(book([['a', 'b']]), { formatterOptions: { delimiter: '§' } }))).toBe('a§b');
  });

  it('keeps fast-csv behaviour with formulaInjection "off"', async () => {
    const bytes = await file(book([['ï', '»¿=1+1']]), { encoding: 'latin1', formulaInjection: 'off', formatterOptions: { delimiter: '' } });
    expect([...bytes.subarray(0, 4)]).toEqual([0xef, 0xbb, 0xbf, 0x3d]);
    const odd = await file(book([[null, '¿=1+1']]), { encoding: 'latin1', formulaInjection: 'off', formatterOptions: { delimiter: 'ï»' } });
    expect([...odd.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
  });

  it('checks the first bytes of the file, whatever line or field they come from', async () => {
    // Each case is a first line or first field that leaves the file's first bytes to another
    // value: an empty header line, missing rows before the first one, an empty first field.
    const values = ['ï»¿=1', 'þÿ', 'ÿþ', 'ï¼\u009D1', '=1', '\uFEFF=1', '믯㶿1'];
    for (const encoding of ['latin1', 'utf16le'] as const) {
      for (const writeBOM of [false, true]) {
        for (const value of values) {
          const layouts: [Book, Record<string, unknown>][] = [];
          layouts.push([book([[value, 'x']]), {}]);
          layouts.push([book([['a', value]]), { headers: [value] }]);
          layouts.push([book([[value]]), { headers: [] }]);
          const late = new ExcelJS.Workbook();
          late.addWorksheet('S').getRow(3).values = [value];
          layouts.push([late, {}]);
          layouts.push([book([[null, value]]), {}]);
          layouts.push([book([['', value]]), { quoteColumns: true }]);
          for (const [wb, formatter] of layouts) {
            const bytes = await file(wb, { encoding, formatterOptions: { writeBOM, ...formatter } });
            const head = [...bytes.subarray(0, 4)];
            const label = `${encoding} ${String(writeBOM)} ${JSON.stringify(formatter)} ${JSON.stringify(value)}: ${head.join(' ')}`;
            if (encoding === 'utf16le') expect(head.slice(0, 2), label).toEqual([0xff, 0xfe]);
            for (const text of readerTexts(bytes)) expect(FORMULA_START.test(text.replace(/^\uFEFF+/, '')), label).toBe(false);
          }
        }
      }
    }
  });
});
