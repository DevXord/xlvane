// Ways around the CSV formula-injection escape found by the 1.3.0 penetration test (docs/facts/csv.md
// CSV-12, CSV-14): a field that starts with a byte order mark before a formula sign (spreadsheet
// applications drop a BOM at the start of the file, so the first field ran as a formula in
// LibreOffice 24.2.7.2), and a field with a comma written with another delimiter (an application that
// reads the file with commas splits the field, so text after the comma starts a field).
import { describe, expect, it } from 'vitest';
import { toCSV } from '../../packages/xlvane/src/csv';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/index';

function book(values: readonly string[]): Uint8Array {
  const wb = createWorkbook();
  const s = wb.addSheet('S');
  values.forEach((v, i) => s.set(`A${String(i + 1)}`, v));
  return writeWorkbook(wb);
}

/** Fields of RFC 4180 text with the given delimiter (quotes doubled inside quoted fields). */
function parse(text: string, delimiter: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  let wasQuoted = false;
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
    if (c === '"' && field === '' && !wasQuoted) quoted = wasQuoted = true;
    else if (c === delimiter || c === '\n' || c === '\r') {
      row.push(field);
      field = '';
      wasQuoted = false;
      if (c !== delimiter) {
        if (c === '\r' && text[i + 1] === '\n') i++;
        rows.push(row);
        row = [];
      }
    } else field += c;
  }
  if (field !== '' || row.length > 0) rows.push([...row, field]);
  return rows;
}

/** An unquoted field that starts like a formula, after the BOMs a spreadsheet application drops. */
const FORMULA = /^[\ufeff\ufffe]*[=+\-@\t\r\n＝＋－＠﹦﹢﹣﹫]/;

describe('toCSV: a leading byte order mark before a formula sign', () => {
  const values = ['﻿=1+1', '￾=1+1', '﻿﻿+1', '﻿@SUM(A1)', '﻿', '﻿plain'];

  it('escapes text that starts with U+FEFF or U+FFFE', () => {
    const csv = toCSV(book(values));
    expect(csv.split('\r\n').slice(0, -1)).toEqual(values.map((v) => `"'${v}"`));
    // No line starts with a BOM outside quotes, so no reader can drop it and see a formula.
    for (const line of csv.split('\r\n')) expect(FORMULA.test(line)).toBe(false);
  });

  it('escapes the first field of the file, which LibreOffice ran as a formula', () => {
    const csv = toCSV(book(['﻿=1+1']));
    expect(csv).toBe('"\'﻿=1+1"\r\n');
    expect(toCSV(book(['﻿=1+1']), { bom: true })).toBe('﻿"\'﻿=1+1"\r\n');
  });

  it('keeps the text unchanged with formulaInjection "off"', () => {
    expect(toCSV(book(['﻿=1+1']), { formulaInjection: 'off' })).toBe('﻿=1+1\r\n');
  });
});

describe('toCSV: separators other than the delimiter are quoted', () => {
  const wb = (() => {
    const w = createWorkbook();
    w.addSheet('S').set('A1', 'x,=1+1').set('B1', 'y;@SUM(1)').set('C1', 'z\t+1').set('D1', 'w|-1').set('E1', 'plain');
    return writeWorkbook(w);
  })();

  for (const delimiter of [',', ';', '\t', '|']) {
    it(`with delimiter ${JSON.stringify(delimiter)}, a field with , ; tab or | is quoted`, () => {
      const csv = toCSV(wb, { delimiter });
      expect(csv).toBe(['"x,=1+1"', '"y;@SUM(1)"', '"z\t+1"', '"w|-1"', 'plain'].join(delimiter) + '\r\n');
      expect(parse(csv, delimiter)).toEqual([['x,=1+1', 'y;@SUM(1)', 'z\t+1', 'w|-1', 'plain']]);
    });
  }

  it('in a one-column file, a reader that splits with another of the four separators keeps each field whole', () => {
    // The reproduction: x,=1+1 written with ";" was split by LibreOffice's default comma import into
    // "x" and the formula =1+1 (CSV-14).
    const one = book(['x,=1+1', 'y;@SUM(1)', 'z\t+1', 'w|-1']);
    for (const delimiter of [',', ';', '\t', '|']) {
      const csv = toCSV(one, { delimiter });
      for (const reader of [',', ';', '\t', '|']) expect(parse(csv, reader), `written with ${JSON.stringify(delimiter)}, read with ${JSON.stringify(reader)}`).toEqual([['x,=1+1'], ['y;@SUM(1)'], ['z\t+1'], ['w|-1']]);
    }
  });

  it('quotes such fields also with formulaInjection "off" (quoting does not change the value)', () => {
    expect(toCSV(wb, { delimiter: ';', formulaInjection: 'off' })).toBe('"x,=1+1";"y;@SUM(1)";"z\t+1";"w|-1";plain\r\n');
  });

  it('quotes a field with a comma for another delimiter such as ":"', () => {
    expect(toCSV(wb, { delimiter: ':', range: 'A1:A1' })).toBe('"x,=1+1"\r\n');
  });
});
