// The streaming reader on hostile files (1.2.2 security review, items K6-K8 and M4): one row
// element cannot buffer more cells than a row holds, parts the reader does not use are not
// inflated, part names map to ZIP items as in the batch reader, and the row objects of a table
// are bounded by the text they repeat.
import { strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { formatCellRef } from '../../packages/xlvane/src/sml/cell-ref';
import { openWorkbookStream, readTableStream } from '../../packages/xlvane/src/stream-read';
import { readTable, toJSON } from '../../packages/xlvane/src/table';
import { expectXlvError, expectXlvErrorAsync } from '../helpers/expect-error';
import { buildZip, type RawEntry } from '../helpers/raw-zip';
import { buildXlsx, inline, row } from '../helpers/xlsx-builder';

async function drain<T>(items: AsyncIterable<T>): Promise<T[]> {
  const out: T[] = [];
  for await (const item of items) out.push(item);
  return out;
}

describe('cells of one row element (K6)', () => {
  // Every cell of the row element names its own row.
  const scattered = (n: number): string => `<row r="1">${Array.from({ length: n }, (_, i) => `<c r="A${String(i + 1)}"><v>${String(i)}</v></c>`).join('')}</row>`;

  it('refuses a row element that holds more cells than a row can (rows)', async () => {
    const book = await openWorkbookStream(buildXlsx({ sheets: [{ name: 'S', data: scattered(20_000) }] }));
    const error = await expectXlvErrorAsync(() => drain(book.rows('S')), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxCells', value: 16_384 });
    expect(error.message).toMatch(/one row element/i);
    expect(error.hint.length).toBeGreaterThan(0);
    expect(error.location?.sheet).toBe('S');
  });

  it('refuses a row element that holds more cells than a row can (readTableStream)', async () => {
    await expectXlvErrorAsync(() => drain(readTableStream(buildXlsx({ sheets: [{ name: 'S', data: scattered(20_000) }] }))), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxCells' });
  });

  it('counts style-only cells too, and not a cell given twice', async () => {
    const styles = '<cellXfs count="2"><xf/><xf applyFont="1"/></cellXfs>';
    const styled = `<row r="1">${Array.from({ length: 16_385 }, (_, i) => `<c r="A${String(i + 1)}" s="1"/>`).join('')}</row>`;
    const book = await openWorkbookStream(buildXlsx({ sheets: [{ name: 'S', data: styled }], styles }));
    await expectXlvErrorAsync(() => drain(book.rows('S')), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxCells' });
    // The same 16 384 cells twice: a value, then style only (the last definition is used).
    const twice = `<row r="1">${Array.from({ length: 16_384 }, (_, i) => `<c r="${formatCellRef(1, i + 1)}"><v>1</v></c>`).join('')}${Array.from({ length: 16_384 }, (_, i) => `<c r="${formatCellRef(1, i + 1)}" s="1"/>`).join('')}</row>`;
    const again = await openWorkbookStream(buildXlsx({ sheets: [{ name: 'S', data: twice }], styles }));
    const rows = await drain(again.rows('S'));
    expect(rows[0]?.styledCells).toHaveLength(16_384);
  });

  it('streams a full row of 16 384 cells, and many rows', async () => {
    const full = `<row r="1">${Array.from({ length: 16_384 }, (_, i) => `<c r="${formatCellRef(1, i + 1)}"><v>${String(i)}</v></c>`).join('')}</row>`;
    const rest = Array.from({ length: 1000 }, (_, i) => `<row r="${String(i + 2)}"><c r="A${String(i + 2)}"><v>1</v></c></row>`).join('');
    const book = await openWorkbookStream(buildXlsx({ sheets: [{ name: 'S', data: full + rest }] }));
    const rows = await drain(book.rows('S'));
    expect(rows).toHaveLength(1001);
    expect(rows[0]?.cells).toHaveLength(16_384);
    // A row element may still hold cells of other rows, up to a row's worth.
    const book2 = await openWorkbookStream(buildXlsx({ sheets: [{ name: 'S', data: scattered(16_384) }] }));
    expect((await drain(book2.rows('S'))).flatMap((r) => r.cells)).toHaveLength(16_384);
  });

  it('leaves the batch reader unchanged', () => {
    expect([...readWorkbook(buildXlsx({ sheets: [{ name: 'S', data: scattered(20_000) }] })).getSheet('S').cells()]).toHaveLength(20_000);
  });
});

describe('relationships parts that the reader does not use (K7)', () => {
  // The .rels entries use a compression method xlvane does not support (LZMA, 14): reading one
  // would fail, so the file opens only if they are left alone, as readWorkbook leaves them.
  const withJunk = (count: number): Uint8Array => {
    const files = unzipSync(
      buildXlsx({
        sheets: [{ name: 'S', data: row(1, inline('A1', 'x')) + row(2, '<c r="A2"><v>2</v></c>') }],
        extraParts: {
          'xl/worksheets/_rels/sheet1.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
        },
      }),
    );
    const entries: RawEntry[] = Object.entries(files).map(([name, data]) => ({ name, data }));
    for (let i = 1; i <= count; i++) entries.push({ name: `junk/_rels/x${String(i)}.xml.rels`, data: strToU8('<Relationships/>'.repeat(1000)), method: 14 });
    // Also next to real parts: the relationships of a part that nothing refers to.
    entries.push({ name: 'xl/_rels/unused.xml.rels', data: strToU8('<x/>'), method: 14 });
    return buildZip(entries);
  };

  it('opens and streams a file whose unreferenced .rels entries are never read', async () => {
    const bytes = withJunk(6);
    expect(readWorkbook(bytes).getSheet('S').cell('A2')?.value).toBe(2);
    const book = await openWorkbookStream(bytes);
    const rows = await drain(book.rows('S'));
    expect(rows.map((r) => r.cells.map((c) => c.value))).toEqual([['x'], [2]]);
    expect(await book.notes('S')).toEqual([]);
    expect(await drain(readTableStream(bytes, { header: false }))).toHaveLength(2);
  });
});

describe('part names with percent-encoded characters (K8)', () => {
  // The workbook part is the ZIP item "xl/%C3%A9.xml", part name "/xl/é.xml" (Part 2 §7.3.5),
  // and the sheet "xl/worksheets/sh%C3%A9et1.xml".
  const encoded = (raw: boolean): Uint8Array => {
    const files = unzipSync(buildXlsx({ sheets: [{ name: 'S', data: row(1, '<c r="A1" t="s"><v>0</v></c>', '<c r="B1"><v>7</v></c>') }], sharedStrings: ['shared'] }));
    const wb = raw ? 'xl/\u00e9.xml' : 'xl/%C3%A9.xml';
    const sheet = raw ? 'sh\u00e9et1.xml' : 'sh%C3%A9et1.xml';
    const text = (name: string): string => new TextDecoder().decode(files[name]);
    const out: Record<string, Uint8Array> = { ...files };
    delete out['xl/workbook.xml'];
    delete out['xl/_rels/workbook.xml.rels'];
    delete out['xl/worksheets/sheet1.xml'];
    out[wb] = files['xl/workbook.xml']!;
    out[raw ? 'xl/_rels/\u00e9.xml.rels' : 'xl/_rels/%C3%A9.xml.rels'] = strToU8(text('xl/_rels/workbook.xml.rels').replace('worksheets/sheet1.xml', `worksheets/${sheet}`));
    out[`xl/worksheets/${sheet}`] = files['xl/worksheets/sheet1.xml']!;
    out['_rels/.rels'] = strToU8(text('_rels/.rels').replace('xl/workbook.xml', 'xl/%C3%A9.xml'));
    out['[Content_Types].xml'] = strToU8(text('[Content_Types].xml').replace('/xl/workbook.xml', '/xl/%C3%A9.xml').replace('/xl/worksheets/sheet1.xml', '/xl/worksheets/sh%C3%A9et1.xml'));
    return zipSync(out);
  };

  for (const raw of [false, true]) {
    it(`streams the same values as readWorkbook (${raw ? 'UTF-8 item names' : 'percent-encoded item names'})`, async () => {
      const bytes = encoded(raw);
      const batch = [...readWorkbook(bytes).getSheet('S').cells()].map((c) => [c.ref, c.value]);
      expect(batch).toEqual([
        ['A1', 'shared'],
        ['B1', 7],
      ]);
      const book = await openWorkbookStream(bytes);
      const streamed = (await drain(book.rows('S'))).flatMap((r) => r.cells.map((c) => [c.ref, c.value]));
      expect(streamed).toEqual(batch);
    });
  }
});

describe('header text repeated in every row object (M4)', () => {
  // A header cell of 4 000 000 characters, compressible but less than 200:1, and `rows` data rows.
  const longHeader = (rows: number): Uint8Array => {
    let seed = 7;
    const parts: string[] = [];
    for (let i = 0; i < 4_000_000 / 250; i++) {
      seed = (seed * 1_103_515_245 + 12_345) % 2_147_483_648;
      parts.push('a'.repeat(249) + String.fromCharCode(98 + (seed % 25)));
    }
    const data = row(1, inline('A1', parts.join(''))) + Array.from({ length: rows }, (_, i) => row(i + 2, `<c r="A${String(i + 2)}"><v>1</v></c>`)).join('');
    return buildXlsx({ sheets: [{ name: 'S', data }] });
  };
  const lowered = { maxSharedStringChars: 100_000_000 };

  it('refuses rows whose header keys add up to more than maxSharedStringChars (readTable, toJSON)', () => {
    const bytes = longHeader(100);
    expect(bytes.length).toBeLessThan(200_000);
    const error = expectXlvError(() => readTable(bytes, { limits: lowered }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxSharedStringChars', value: 100_000_000 });
    expect(error.location?.sheet).toBe('S');
    expect(error.message.length).toBeLessThan(400);
    expectXlvError(() => toJSON(bytes, { limits: lowered }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxSharedStringChars' });
    // 100 rows × 4 M = 400 M characters stay within the default of 512 Mi.
    expect(readTable(bytes).rows).toHaveLength(100);
    // 200 rows do not, so JSON.stringify can no longer fail on them.
    expectXlvError(() => readTable(longHeader(200)), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxSharedStringChars' });
  }, 60_000);

  it('streams the same rows one at a time: only the names of one row count there', async () => {
    expect(await drain(readTableStream(longHeader(100), { limits: lowered }))).toHaveLength(100);
    expect(await drain(readTableStream(longHeader(100)))).toHaveLength(100);
  }, 60_000);

  it('reads ordinary tables as before', async () => {
    const bytes = buildXlsx({ sheets: [{ name: 'S', data: row(1, inline('A1', 'Name'), inline('B1', 'Total')) + row(2, inline('A2', 'x'), '<c r="B2"><v>3</v></c>') }] });
    expect(readTable(bytes, { limits: { maxSharedStringChars: 9 } }).rows).toEqual([{ Name: 'x', Total: 3 }]);
    expectXlvError(() => readTable(bytes, { limits: { maxSharedStringChars: 8 } }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxSharedStringChars' });
    expect((await drain(readTableStream(bytes, { limits: { maxSharedStringChars: 9 } }))).map((i) => i.value)).toEqual([{ Name: 'x', Total: 3 }]);
  });
});

describe('readTableStream and repeated header names (independent review of the patch, D-J01)', () => {
  it('checks the header names per row, not over every row streamed', async () => {
    const { readTableStream } = await import('../../packages/xlvane/src/stream-read');
    const { createWorkbook: create, writeWorkbook: write } = await import('../../packages/xlvane/src/write');
    const wb = create();
    const s = wb.addSheet('S');
    for (let c = 1; c <= 10; c++) s.set(`${String.fromCharCode(64 + c)}1`, `${'h'.repeat(100)}${String(c)}`);
    for (let r = 2; r <= 201; r++) s.set(`A${String(r)}`, r);
    let n = 0;
    for await (const _row of readTableStream(new Blob([write(wb) as BlobPart]), { limits: { maxSharedStringChars: 10_000 } })) n++;
    expect(n).toBe(200);
  });
});
