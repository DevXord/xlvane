// Small hostile inputs must not make the core library take time that grows with the product of
// two counts from the file or from the caller (1.2.2 security review, items K1-K5, M3, M5, C3,
// C4, C6). Each case took from several seconds to a minute or more before the fix; the time
// budgets are generous so that slow machines pass.
import { describe, expect, it } from 'vitest';
import { diff } from '../../packages/xlvane/src/diff';
import { editCells } from '../../packages/xlvane/src/edit';
import { inspect } from '../../packages/xlvane/src/inspect';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { formatCellRef } from '../../packages/xlvane/src/sml/cell-ref';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { fromCSV } from '../../packages/xlvane/src/csv';
import { fromJSON, readTable, writeTable } from '../../packages/xlvane/src/table';
import { expectXlvError } from '../helpers/expect-error';
import { readTableStream } from '../../packages/xlvane/src/stream-read';
import type { StandardSchemaV1 } from '../../packages/xlvane/src/standard-schema';
import { buildXlsx } from '../helpers/xlsx-builder';

const PKG_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NSR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const rels = (items: readonly (readonly [string, string, string])[]): string =>
  `<?xml version="1.0"?><Relationships xmlns="${PKG_RELS}">${items
    .map(([id, t, tgt]) => `<Relationship Id="${id}" Type="${t.startsWith('http') ? t : `${NSR}/${t}`}" Target="${tgt}"${tgt.startsWith('http') ? ' TargetMode="External"' : ''}/>`)
    .join('')}</Relationships>`;

function timed<T>(fn: () => T): { ms: number; value: T } {
  const t0 = performance.now();
  const value = fn();
  return { ms: performance.now() - t0, value };
}

describe('notes and their VML boxes (K1)', () => {
  it('matches many notes with many VML parts in linear time', () => {
    const notes = 100_000;
    const parts = 8000;
    const comments = `<?xml version="1.0"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>a</author></authors><commentList>${Array.from(
      { length: notes },
      (_, i) => `<comment ref="A${String(i + 1)}" authorId="0"><text><t>n</t></text></comment>`,
    ).join('')}</commentList></comments>`;
    // Each VML part holds the box of one note.
    const vml = (row: number): string =>
      `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape><x:ClientData ObjectType="Note"><x:Row>${String(row)}</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>`;
    const extra: Record<string, string> = { 'xl/comments1.xml': comments };
    const items: [string, string, string][] = [['rIdC', 'comments', '../comments1.xml']];
    for (let i = 0; i < parts; i++) {
      extra[`xl/drawings/vmlDrawing${String(i + 1)}.vml`] = vml(i * 7);
      items.push([`rIdV${String(i + 1)}`, 'vmlDrawing', `../drawings/vmlDrawing${String(i + 1)}.vml`]);
    }
    extra['xl/worksheets/_rels/sheet1.xml.rels'] = rels(items);
    const bytes = buildXlsx({ sheets: [{ name: 'S', data: '<row r="1"><c r="A1"><v>1</v></c></row>' }], extraParts: extra });
    const run = timed(() => readWorkbook(bytes));
    const read = run.value.getSheet('S').notes;
    expect(read).toHaveLength(notes);
    expect(read[0]?.box).toBeDefined();
    expect(read[7]?.box).toBeDefined();
    expect(read[1]?.box).toBeUndefined();
    expect(run.ms).toBeLessThan(6000);
  }, 120_000);
});

describe('pictures of a drawing (K2)', () => {
  it('resolves the relationships of many pictures in linear time', () => {
    const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
    const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const pictures = 10_000;
    const others = 50_000;
    const items: [string, string, string][] = [];
    for (let i = 0; i < others; i++) items.push([`rIdX${String(i)}`, 'hyperlink', `https://example.com/${String(i)}`]);
    items.push(['rIdImg', 'image', '../media/image1.png'], ['rIdLink', 'hyperlink', 'https://example.com/picture']);
    const pic = (i: number): string =>
      `<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${String(i)}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="10" cy="10"/>` +
      `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${String(i + 2)}" name="P${String(i)}"><a:hlinkClick r:id="rIdLink"/></xdr:cNvPr><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImg"/></xdr:blipFill></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
    const drawing = `<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${NSR}">${Array.from({ length: pictures }, (_, i) => pic(i)).join('')}</xdr:wsDr>`;
    const bytes = buildXlsx({
      sheets: [{ name: 'S', data: '<row r="1"><c r="A1"><v>1</v></c></row>', after: '<drawing r:id="rId1"/>' }],
      extraParts: {
        'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', 'drawing', '../drawings/drawing1.xml']]),
        'xl/drawings/drawing1.xml': drawing,
        'xl/drawings/_rels/drawing1.xml.rels': rels(items),
        'xl/media/image1.png': new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      },
    });
    const run = timed(() => readWorkbook(bytes));
    const images = run.value.getSheet('S').images;
    expect(images).toHaveLength(pictures);
    expect(images[0]?.hyperlink).toBe('https://example.com/picture');
    expect(images[pictures - 1]?.image?.part).toBe('/xl/media/image1.png');
    expect(run.ms).toBeLessThan(6000);
  }, 120_000);

  it('takes the relationship of the right type for r:embed and hlinkClick', () => {
    const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
    const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const pic = (embed: string, link: string): string =>
      `<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="10" cy="10"/>` +
      `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="P"><a:hlinkClick r:id="${link}"/></xdr:cNvPr><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${embed}"/></xdr:blipFill></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
    const bytes = buildXlsx({
      sheets: [{ name: 'S', data: '', after: '<drawing r:id="rId1"/>' }],
      extraParts: {
        'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', 'drawing', '../drawings/drawing1.xml']]),
        'xl/drawings/drawing1.xml': `<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${NSR}">${pic('rIdImg', 'rIdLink')}${pic('rIdLink', 'rIdImg')}${pic('nope', 'nope')}</xdr:wsDr>`,
        'xl/drawings/_rels/drawing1.xml.rels': rels([['rIdImg', 'image', '../media/image1.png'], ['rIdLink', 'hyperlink', 'https://example.com/']]),
        'xl/media/image1.png': new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      },
    });
    const images = readWorkbook(bytes).getSheet('S').images;
    expect(images.map((i) => [i.image?.part ?? null, i.hyperlink])).toEqual([
      ['/xl/media/image1.png', 'https://example.com/'],
      [null, null],
      [null, null],
    ]);
  });
});

describe('header detection (K3)', () => {
  it('checks rows of years against the next row in linear time', () => {
    // Text in A, then 16 383 alternating years: every row is a header candidate whose years are
    // consecutive, and the next row holds only years, so each candidate is rejected late.
    const yearRow = (r: number): string => {
      let cells = `<c r="A${String(r)}" t="inlineStr"><is><t>h</t></is></c>`;
      for (let col = 2; col <= 16_384; col++) cells += `<c r="${formatCellRef(r, col)}"><v>${col % 2 === 0 ? '1900' : '1901'}</v></c>`;
      return `<row r="${String(r)}">${cells}</row>`;
    };
    const data = Array.from({ length: 21 }, (_, i) => yearRow(i + 1)).join('');
    const bytes = buildXlsx({ sheets: [1, 2, 3].map((n) => ({ name: `S${String(n)}`, data })) });
    const workbook = readWorkbook(bytes);
    const table = timed(() => readTable(workbook, { sheet: 'S1' }));
    expect(table.value.headerRow).toBeNull();
    const report = timed(() => inspect(workbook));
    expect(report.value.sheets[0]?.headerRow).toBeNull();
    expect(table.ms + report.ms).toBeLessThan(8000);
  }, 180_000);
});

describe('tables written from rows (K4)', () => {
  const wide = 16_384;
  const firstObject = (): Record<string, number> => {
    const o: Record<string, number> = {};
    for (let i = 0; i < wide; i++) o[`k${String(i)}`] = i;
    return o;
  };

  it('refuses one wide object followed by many empty ones quickly (fromJSON)', () => {
    const json = JSON.stringify([firstObject(), ...Array.from({ length: 40_000 }, () => ({}))]);
    const run = timed(() => {
      try {
        return fromJSON(json);
      } catch (error) {
        return error;
      }
    });
    expectXlvError(() => {
      throw run.value;
    }, 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxCells' });
    expect(run.ms).toBeLessThan(5000);
  }, 120_000);

  it('refuses one wide array followed by many short ones quickly (writeTable)', () => {
    const rows: unknown[][] = [Array.from({ length: wide }, (_, i) => i), ...Array.from({ length: 40_000 }, () => [1])];
    const run = timed(() => {
      try {
        return writeTable(rows, { preset: 'plain' });
      } catch (error) {
        return error;
      }
    });
    expectXlvError(() => {
      throw run.value;
    }, 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxCells' });
    expect(run.ms).toBeLessThan(5000);
  }, 120_000);

  it('refuses a long first CSV line followed by many short lines quickly (fromCSV)', () => {
    const csv = `${','.repeat(wide - 1)}\n${'a\n'.repeat(40_000)}`;
    const run = timed(() => {
      try {
        return fromCSV(csv);
      } catch (error) {
        return error;
      }
    });
    expectXlvError(() => {
      throw run.value;
    }, 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxCells' });
    expect(run.ms).toBeLessThan(5000);
  }, 120_000);

  it('writes wide sparse rows in time proportional to their content', () => {
    // 16 384 columns and 300 rows stay within the output budget (4.9 M positions).
    const rows = [firstObject(), ...Array.from({ length: 300 }, (_, i) => ({ [`k${String(i * 50)}`]: i }))];
    const run = timed(() => fromJSON(rows));
    const table = readTable(run.value);
    expect(table.rows).toHaveLength(301);
    expect(table.rows[2]?.['k50']).toBe(1);
    expect(table.rows[2]?.['k51']).toBeNull();
    expect(run.ms).toBeLessThan(5000);
  }, 120_000);

  it('keeps cell values, number formats and errors of mixed rows', () => {
    const out = readWorkbook(
      writeTable(
        [
          { b: 2, a: 1 },
          { c: 'x' },
          {},
        ],
        { preset: 'plain', columns: ['a', { key: 'b', numFmt: '0.00' }, 'c', { key: 'a', header: 'a again' }] },
      ),
    ).getSheet('Sheet1');
    const show = (sheet: ReturnType<ReturnType<typeof readWorkbook>['getSheet']>) => [
      [...sheet.cells()].map((c) => [c.ref, c.value, c.numFmt]),
      [...sheet.styledCells()].map((c) => [c.ref, c.numFmt]),
    ];
    expect(show(out)).toEqual([
      [
        ['A1', 'a', null],
        ['B1', 'b', null],
        ['C1', 'c', null],
        ['D1', 'a again', null],
        ['A2', 1, null],
        ['B2', 2, '0.00'],
        ['D2', 1, null],
        ['C3', 'x', null],
      ],
      [
        ['B3', '0.00'],
        ['B4', '0.00'],
      ],
    ]);
    const positional = readWorkbook(writeTable([[1, 2], [3], []], { preset: 'plain', columns: ['x', { key: 'y', numFmt: '0.0' }] })).getSheet('Sheet1');
    expect(show(positional)).toEqual([
      [
        ['A1', 'x', null],
        ['B1', 'y', null],
        ['A2', 1, null],
        ['B2', 2, '0.0'],
        ['A3', 3, null],
      ],
      [
        ['B3', '0.0'],
        ['B4', '0.0'],
      ],
    ]);
    // The first value that is not a cell value, in row and then column order, is reported.
    expectXlvError(() => writeTable([{ a: 1 }, { c: [1], b: {} }], { columns: ['a', 'b', 'c'] }), 'XLV_E_INVALID_ROWS', { row: 3, key: 'b' });
    expectXlvError(() => writeTable([[1], [1, [2], {}]]), 'XLV_E_INVALID_ROWS', { row: 2, key: '1' });
  });
});

describe('schema issues of wide tables (K5)', () => {
  // 16 384 columns, 100 data rows, and a schema that reports 1 000 issues per row, all on the
  // last column.
  const header: Record<string, string> = {};
  for (let i = 1; i <= 16_384; i++) header[`c${String(i)}`] = 'x';
  const bytes = writeTable([header, ...Array.from({ length: 100 }, () => ({ c1: 1 }))], { preset: 'plain' });
  const issues = Array.from({ length: 1000 }, () => ({ message: 'bad', path: ['c16384'] }));
  const schema: StandardSchemaV1 = { '~standard': { version: 1, vendor: 'test', validate: () => ({ issues }) } };

  it('maps the issues of readTable to columns in linear time', () => {
    const run = timed(() => readTable(bytes, { schema }));
    expect(run.value.issues).toHaveLength(101_000);
    expect(run.value.issues[0]).toMatchObject({ column: 'c16384', cell: 'XFD2' });
    expect(run.ms).toBeLessThan(5000);
  }, 120_000);

  it('maps the issues of readTableStream to columns in linear time', async () => {
    const t0 = performance.now();
    let count = 0;
    let first: unknown;
    for await (const item of readTableStream(bytes, { schema })) {
      count += item.issues.length;
      first ??= item.issues[0];
    }
    const ms = performance.now() - t0;
    expect(count).toBe(101_000);
    expect(first).toMatchObject({ column: 'c16384', cell: 'XFD2' });
    expect(ms).toBeLessThan(5000);
  }, 120_000);
});

describe('editing cells in new rows (M3)', () => {
  it('places many edits between many existing rows in linear time', () => {
    // Rows 1, 3, 5, ... exist; the edits go to rows 2, 4, 6, ..., which do not.
    const existing = 100_000;
    const data = Array.from({ length: existing }, (_, i) => `<row r="${String(2 * i + 1)}"><c r="A${String(2 * i + 1)}"><v>${String(i)}</v></c></row>`).join('');
    const bytes = buildXlsx({ sheets: [{ name: 'S', data }] });
    const edits = Array.from({ length: 20_000 }, (_, i) => ({ sheet: 'S', cell: `B${String(2 * i + 2)}`, value: i }));
    const run = timed(() => editCells(bytes, edits));
    const sheet = readWorkbook(run.value.bytes).getSheet('S');
    expect(sheet.cell('B2')?.value).toBe(0);
    expect(sheet.cell('B40000')?.value).toBe(19_999);
    expect(sheet.cell('A199999')?.value).toBe(99_999);
    const numbers = [...sheet.rows()].map((r) => r[0]?.row);
    expect(numbers.slice(0, 5)).toEqual([1, 2, 3, 4, 5]);
    expect(numbers).toHaveLength(existing + 20_000);
    expect(run.ms).toBeLessThan(6000);
  }, 120_000);
});

describe('diff with styles (M5)', () => {
  it('compares workbooks of 1 M styled cells quickly, equal and with every cell changed', () => {
    // One style with a font, a fill and a border, on every cell (written as XML: faster than the
    // writer for a million cells).
    const styles =
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.00"/></numFmts>' +
      '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/></font></fonts>' +
      '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/></patternFill></fill></fills>' +
      '<borders count="2"><border/><border><bottom style="thin"/></border></borders>' +
      '<cellStyleXfs count="1"><xf/></cellStyleXfs>' +
      '<cellXfs count="2"><xf/><xf numFmtId="164" fontId="1" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1"/></cellXfs>';
    const book = (offset: number): Uint8Array => {
      const rows: string[] = [];
      for (let r = 1; r <= 1000; r++) {
        let cells = '';
        for (let c = 1; c <= 1000; c++) cells += `<c r="${formatCellRef(r, c)}" s="1"><v>${String(r * c + offset)}</v></c>`;
        rows.push(`<row r="${String(r)}">${cells}</row>`);
      }
      return buildXlsx({ sheets: [{ name: 'S', data: rows.join('') }], styles });
    };
    const first = book(0);
    const a = readWorkbook(first);
    const b = readWorkbook(book(1));
    const c = readWorkbook(first);
    const same = timed(() => diff(a, c, { styles: true }));
    expect(same.value.equal).toBe(true);
    const run = timed(() => diff(a, b, { styles: true }));
    expect(run.value.summary.cellsChanged).toBe(1_000_000);
    expect(run.value.sheets[0]?.cells).toHaveLength(10_000);
    expect(run.value.sheets[0]?.cells[0]).toMatchObject({ ref: 'A1', kind: 'changed', before: { value: 1 }, after: { value: 2 } });
    expect(run.value.sheets[0]?.cells[0]?.before?.style).toMatchObject({ font: { bold: true } });
    expect(run.ms + same.ms).toBeLessThan(3000);
  }, 180_000);

  it('reports the same changes with and without styles on a small workbook', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 1, { font: { bold: true } }).set('A2', 2).set('A3', 3, { font: { italic: true } });
    const wb2 = createWorkbook();
    wb2.addSheet('S').set('A1', 1, { font: { italic: true } }).set('A2', 2).set('A4', 4);
    const result = diff(writeWorkbook(wb), writeWorkbook(wb2), { styles: true, maxCells: 2 });
    expect(result.summary).toMatchObject({ cellsAdded: 1, cellsRemoved: 1, cellsChanged: 1 });
    expect(result.sheets[0]?.truncated).toBe(true);
    expect(result.sheets[0]?.cells.map((c) => [c.ref, c.kind, c.before?.style?.font?.bold ?? null, c.after?.style?.font?.italic ?? null])).toEqual([
      ['A1', 'changed', true, true],
      ['A3', 'removed', null, null],
    ]);
    expect(diff(writeWorkbook(wb), writeWorkbook(wb2), { styles: false }).summary).toMatchObject({ cellsAdded: 1, cellsRemoved: 1, cellsChanged: 0 });
  });
});

describe('building workbooks with many sheets, merges and names (C3, C4, C6)', () => {
  it('adds 16 000 sheets with ids in linear time', () => {
    const wb = createWorkbook();
    const run = timed(() => {
      for (let i = 1; i <= 16_000; i++) wb.addSheet(`Sheet number ${String(i)}`, { sheetId: i + 100 });
    });
    expect(wb.sheets).toHaveLength(16_000);
    expectXlvError(() => wb.addSheet('SHEET NUMBER 7'), 'XLV_E_DUPLICATE_SHEET_NAME', { name: 'SHEET NUMBER 7', existing: 'Sheet number 7' });
    expectXlvError(() => wb.addSheet('Other', { sheetId: 107 }), 'XLV_E_INVALID_OPTION', { existing: 'Sheet number 7' });
    wb.addSheet('Other', { sheetId: 99 });
    expect(run.ms).toBeLessThan(3000);
  }, 60_000);

  it('merges 60 000 ranges in linear time', () => {
    const sheet = createWorkbook().addSheet('S');
    const run = timed(() => {
      for (let i = 1; i <= 60_000; i++) sheet.merge(`A${String(i)}:B${String(i)}`);
    });
    expect(run.ms).toBeLessThan(3000);
    expectXlvError(() => sheet.merge('B30000:C30001'), 'XLV_E_OVERLAPPING_MERGE', { other: 'A30000:B30000' });
  }, 60_000);

  it('merges many tall and wide ranges in linear time', () => {
    const sheet = createWorkbook().addSheet('S');
    const run = timed(() => {
      // Full-height columns, then rows of wide ranges beside small ones.
      for (let c = 1; c <= 8000; c++) sheet.merge(`${formatCellRef(1, c)}:${formatCellRef(1_048_576, c)}`);
      for (let r = 1; r <= 20_000; r++) sheet.merge(`${formatCellRef(r, 8001)}:${formatCellRef(r, 16_000)}`);
      for (let r = 1; r <= 20_000; r++) sheet.merge(`${formatCellRef(r, 16_001)}:${formatCellRef(r, 16_002)}`);
    });
    expect(run.ms).toBeLessThan(2500);
    expectXlvError(() => sheet.merge(`${formatCellRef(5, 16_002)}:${formatCellRef(6, 16_003)}`), 'XLV_E_OVERLAPPING_MERGE', { other: `${formatCellRef(5, 16_001)}:${formatCellRef(5, 16_002)}` });
    expectXlvError(() => sheet.merge('A5000:XFD5000'), 'XLV_E_OVERLAPPING_MERGE', { other: 'A1:A1048576' });
    sheet.merge('XFC1:XFD1048576');
  }, 60_000);

  it('reports the first merged range or table that a range overlaps, as before', () => {
    const wb = createWorkbook();
    const sheet = wb.addSheet('S');
    sheet.merge('C3:D4').merge('A1:B2').merge('E1:F10');
    sheet.addTable({ name: 'First', ref: 'H1:I3', columns: [{ name: 'a' }, { name: 'b' }] });
    sheet.addTable({ name: 'Second', ref: 'K1:L3', columns: [{ name: 'a' }, { name: 'b' }] });
    expectXlvError(() => sheet.merge('A1:F4'), 'XLV_E_OVERLAPPING_MERGE', { other: 'C3:D4' });
    expectXlvError(() => sheet.merge('G2:K2'), 'XLV_E_INVALID_TABLE', { name: 'First' });
    expectXlvError(() => sheet.addTable({ name: 'Third', ref: 'F1:H3', columns: [{ name: 'a' }, { name: 'b' }, { name: 'c' }] }), 'XLV_E_INVALID_TABLE', { name: 'Third' });
    expect(() => sheet.addTable({ name: 'Third', ref: 'F1:H3', columns: [{ name: 'a' }, { name: 'b' }, { name: 'c' }] })).toThrow(/overlaps table "First"/);
    expect(() => sheet.addTable({ name: 'Fourth', ref: 'A20:B30', columns: [{ name: 'a' }, { name: 'b' }] })).not.toThrow();
    expect(() => sheet.addTable({ name: 'Fifth', ref: 'D2:E3', columns: [{ name: 'a' }, { name: 'b' }] })).toThrow(/overlaps the merged range C3:D4/);
    expect(() => sheet.merge('A25:B25')).toThrow(/overlaps table "Fourth"/);
    sheet.merge('G1:G2');
  });

  it('defines 40 000 names in linear time', () => {
    const wb = createWorkbook();
    for (let i = 1; i <= 300; i++) wb.addSheet(`S${String(i)}`);
    wb.getSheet('S1').addTable({ name: 'Sales', ref: 'A1:B3', columns: [{ name: 'a' }, { name: 'b' }] });
    const run = timed(() => {
      for (let i = 1; i <= 40_000; i++) wb.defineName(`Name_${String(i)}`, '1', i % 2 === 0 ? { sheet: `S${String((i % 300) + 1)}` } : undefined);
    });
    expect(wb.definedNames).toHaveLength(40_000);
    expect(run.ms).toBeLessThan(3000);
    expectXlvError(() => wb.defineName('NAME_1', '2'), 'XLV_E_INVALID_DEFINED_NAME', { name: 'NAME_1' });
    expect(() => wb.defineName('NAME_1', '2')).toThrow('The workbook already has a defined name "Name_1".');
    expect(() => wb.defineName('NAME_2', '2', { sheet: 'S3' })).toThrow('The sheet "S3" already has a defined name "Name_2".');
    wb.defineName('NAME_2', '2', { sheet: 'S4' });
    wb.defineName('NAME_2', '2');
    expectXlvError(() => wb.defineName('X', '2', { sheet: 's1' }), 'XLV_E_SHEET_NOT_FOUND');
    expect(() => wb.defineName('SALES', '2')).toThrow('"SALES" is already the name of a table.');
    expect(() => wb.getSheet('S2').addTable({ name: 'name_5', ref: 'A1:B3', columns: [{ name: 'a' }, { name: 'b' }] })).toThrow('"Name_5" is already a defined name of the workbook.');
  }, 60_000);
});

describe('fromCSV output budget with the maxCells option (D-J01)', () => {
  it("uses the caller's maxCells, as parsing does", () => {
    const text = `${','.repeat(99)}\n${'1\n'.repeat(100)}`;
    expect(() => fromCSV(text, { header: false })).not.toThrow();
    expect(() => fromCSV(text, { header: false, maxCells: 1000 })).toThrow(expect.objectContaining({ code: 'XLV_E_LIMIT_EXCEEDED' }));
  });
});

describe('writeTable and fromJSON take maxCells (independent review of the patch, D-J01)', () => {
  const rows = Array.from({ length: 30_000 }, (_, i) => ({ [`c${String(i % 20)}`]: i }));
  const columns = Array.from({ length: 20 }, (_, i) => `c${String(i)}`);
  it('a lower maxCells refuses a sparse table, a higher one accepts it', async () => {
    const { writeTable, fromJSON } = await import('../../packages/xlvane/src/table');
    expect(() => writeTable(rows, { columns, maxCells: 100_000 })).toThrow(expect.objectContaining({ code: 'XLV_E_LIMIT_EXCEEDED' }));
    expect(() => fromJSON(rows, { maxCells: 100_000 })).toThrow(expect.objectContaining({ code: 'XLV_E_LIMIT_EXCEEDED' }));
    expect(writeTable(rows, { columns, maxCells: 1_000_000, preset: 'plain' }).length).toBeGreaterThan(0);
  });
});
