// Small hostile files must not make editing or reading slow (pre-release security audit,
// docs/decisions.md D-R03): drawings and VML parts are scanned without backtracking regular
// expressions, hyperlinks are resolved through a lookup table, edits are applied in one pass, and
// editing does not inflate parts it does not touch. Each input took a minute or more before the
// fix; the budgets are generous so that slow machines pass.
import { readFileSync } from 'node:fs';
import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { editCells, splice } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { deleteRows, insertRows } from '../../packages/xlvane/src/rows';
import { openWorkbookStream } from '../../packages/xlvane/src/stream-read';
import { buildXlsx } from '../helpers/xlsx-builder';

const NSR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const rels = (items: [string, string, string][]) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items
    .map(([id, t, tgt]) => `<Relationship Id="${id}" Type="${NSR}/${t}" Target="${tgt}"${tgt.startsWith('http') ? ' TargetMode="External"' : ''}/>`)
    .join('')}</Relationships>`;
const one = '<row r="1"><c r="A1"><v>1</v></c></row>';

function timed<T>(fn: () => T): { ms: number; value: T | Error } {
  const t0 = performance.now();
  let value: T | Error;
  try {
    value = fn();
  } catch (error) {
    value = error as Error;
  }
  return { ms: performance.now() - t0, value };
}

describe('linear time on hostile input', () => {
  it('reads drawing anchors with long white space in a marker quickly (was cubic)', () => {
    const drawing = `<xdr:wsDr xmlns:xdr="${XDR}"><xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>${' '.repeat(8000)}<xdr:x/></xdr:rowOff></xdr:from><xdr:ext cx="1" cy="1"/><xdr:pic></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`;
    const bytes = buildXlsx({
      sheets: [{ name: 'S', data: one, after: '<drawing r:id="rId1"/>' }],
      extraParts: { 'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', 'drawing', '../drawings/drawing1.xml']]), 'xl/drawings/drawing1.xml': drawing },
    });
    expect(timed(() => insertRows(bytes, { sheet: 'S', at: 1, count: 1 })).ms).toBeLessThan(2000);
  });

  it('moves the row of a picture marker and keeps the rest of the anchor', () => {
    const anchor = (row: number) =>
      `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>5</xdr:colOff><xdr:row>${String(row)}</xdr:row><xdr:rowOff>7</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${String(row + 4)}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>`;
    const bytes = buildXlsx({
      sheets: [{ name: 'S', data: one, after: '<drawing r:id="rId1"/>' }],
      extraParts: { 'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', 'drawing', '../drawings/drawing1.xml']]), 'xl/drawings/drawing1.xml': `<xdr:wsDr xmlns:xdr="${XDR}">${anchor(5)}</xdr:wsDr>` },
    });
    const out = new TextDecoder().decode(unzipSync(insertRows(bytes, { sheet: 'S', at: 2, count: 3 }).bytes)['xl/drawings/drawing1.xml']);
    expect(out).toBe(`<xdr:wsDr xmlns:xdr="${XDR}">${anchor(8)}</xdr:wsDr>`);
  });

  it('scans a VML part of unclosed shapes quickly (was quadratic)', () => {
    const vml = `<xml xmlns:v="urn:schemas-microsoft-com:vml">${'<v:shape/>'.repeat(40000)}</xml>`;
    const bytes = buildXlsx({
      sheets: [{ name: 'S', data: one, after: '<legacyDrawing r:id="rId1"/>' }],
      extraParts: { 'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', 'vmlDrawing', '../drawings/vmlDrawing1.vml']]), 'xl/drawings/vmlDrawing1.vml': vml },
    });
    expect(timed(() => insertRows(bytes, { sheet: 'S', at: 1, count: 1 })).ms).toBeLessThan(2000);
  });

  it('resolves many hyperlinks quickly (was quadratic)', () => {
    const n = 30000;
    const links = Array.from({ length: n }, (_, i) => `<hyperlink ref="A${String(i + 1)}" r:id="rId${String(i + 1)}"/>`).join('');
    const bytes = buildXlsx({
      sheets: [{ name: 'S', data: one, after: `<hyperlinks>${links}</hyperlinks>` }],
      extraParts: { 'xl/worksheets/_rels/sheet1.xml.rels': rels(Array.from({ length: n }, (_, i) => [`rId${String(i + 1)}`, 'hyperlink', `http://e/${String(i)}`])) },
    });
    const read = timed(() => readWorkbook(bytes));
    expect(read.ms).toBeLessThan(3000);
    expect((read.value as ReturnType<typeof readWorkbook>).getSheet('S').hyperlinks).toHaveLength(n);
  });

  it('moves many notes and defined names quickly (one-pass edits; was quadratic)', () => {
    const n = 30000;
    const comments = `<?xml version="1.0"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>a</author></authors><commentList>${Array.from({ length: n }, (_, i) => `<comment ref="A${String(i + 1)}" authorId="0"><text><t>x</t></text></comment>`).join('')}</commentList></comments>`;
    const names = Array.from({ length: n }, (_, i) => `<definedName name="n${String(i)}">S!$A$${String(i + 1)}</definedName>`).join('');
    const bytes = buildXlsx({
      sheets: [{ name: 'S', data: one }],
      workbookAfterSheets: `<definedNames>${names}</definedNames>`,
      extraParts: { 'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', 'comments', '../comments1.xml']]), 'xl/comments1.xml': comments },
    });
    const ins = timed(() => insertRows(bytes, { sheet: 'S', at: 1, count: 1 }));
    expect(ins.ms).toBeLessThan(5000);
    const book = readWorkbook((ins.value as ReturnType<typeof insertRows>).bytes);
    expect(book.getSheet('S').notes[0]?.ref).toBe('A2');
    expect(book.definedNames.find((d) => d.name === 'n9')?.refersTo).toBe('S!$A$11');
    expect(timed(() => deleteRows(bytes, { sheet: 'S', at: 5, count: 1 })).ms).toBeLessThan(5000);
  });
});

describe('editing does not inflate parts it does not touch', () => {
  it('copies unreferenced entries in compressed form, unchanged', { timeout: 30000 }, () => {
    const big = new Uint8Array(1024 * 1024);
    const extra: Record<string, Uint8Array> = {};
    for (let i = 0; i < 40; i++) extra[`junk/j${String(i)}.bin`] = big;
    const bytes = buildXlsx({ sheets: [{ name: 'S', data: one }], extraParts: extra });
    const edit = timed(() => editCells(bytes, [{ sheet: 'S', cell: 'B1', value: 2 }]));
    expect(edit.ms).toBeLessThan(2000);
    const out = unzipSync((edit.value as ReturnType<typeof editCells>).bytes);
    expect(out['junk/j7.bin']).toEqual(big);
    expect(readWorkbook((edit.value as ReturnType<typeof editCells>).bytes).getSheet('S').cell('B1')?.value).toBe(2);
  });

  it('keeps the bytes of every part an edit leaves alone', () => {
    const bytes = new Uint8Array(readFileSync(new URL('../corpus/libreoffice/formulas-merged-hidden.xlsx', import.meta.url)));
    const before = unzipSync(bytes);
    const after = unzipSync(editCells(bytes, [{ sheet: readWorkbook(bytes).sheets[0]!.name, cell: 'Z1', value: 'x' }]).bytes);
    for (const [name, data] of Object.entries(before)) {
      if (name.startsWith('xl/worksheets/') && !name.includes('_rels')) continue;
      // The edited sheet, the strings it adds and the workbook (recalculation flag) change.
      if (name === 'xl/sharedStrings.xml' || name === 'xl/workbook.xml') continue;
      expect(after[name], name).toEqual(data);
    }
  });
});

describe('the streaming reader bounds parts it holds in memory', () => {
  it('refuses a shared string table larger than the batch limit (256 MiB by default)', async () => {
    // A small maxEntryBytes stands in for the 256 MiB default; inflating 256 MiB in a unit test
    // would be too slow. Before the fix the shared strings were read under the 4 GiB streaming bound.
    const sst = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${'<si><t>abcdefgh</t></si>'.repeat(200)}</sst>`;
    const bytes = buildXlsx({ sheets: [{ name: 'S', data: '' }], sharedStrings: sst });
    await expect(openWorkbookStream(bytes, { limits: { maxEntryBytes: 1024 } })).rejects.toMatchObject({ code: 'XLV_E_LIMIT_EXCEEDED' });
    await expect(openWorkbookStream(bytes).then((s) => s.sheets.length)).resolves.toBe(1);
  });
});

describe('splice', () => {
  // The previous algorithm, applied from the end of the text to its start (quadratic).
  const slow = (source: string, edits: [number, number, string][]): string => {
    let out = source;
    for (const [s, e, t] of [...edits].sort((a, b) => b[0] - a[0] || b[1] - a[1])) out = out.slice(0, s) + t + out.slice(e);
    return out;
  };

  it('gives the same text as end-to-start splicing for edits that do not overlap', () => {
    let seed = 7;
    const rnd = (n: number): number => (seed = (seed * 1103515245 + 12345) % 2147483648) % n;
    for (let round = 0; round < 500; round++) {
      const source = 'abcdefghijklmnopqrstuvwxyz'.slice(0, 1 + rnd(26));
      const edits: [number, number, string][] = [];
      let pos = 0;
      while (pos <= source.length && edits.length < 6) {
        const start = pos + rnd(3);
        if (start > source.length) break;
        const end = Math.min(source.length, start + rnd(3));
        edits.push([start, end, `<${String(edits.length)}>`]);
        // Insertions (start = end) may share a position with the next edit.
        pos = end;
      }
      edits.sort(() => rnd(3) - 1);
      expect(splice(source, edits), JSON.stringify([source, edits])).toBe(slow(source, edits));
    }
  });

  it('applies 100 000 edits to a long text in linear time', () => {
    const source = 'x'.repeat(1_000_000);
    const edits = Array.from({ length: 100_000 }, (_, i): [number, number, string] => [i * 10, i * 10 + 1, 'y']);
    const run = timed(() => splice(source, edits));
    expect(run.ms).toBeLessThan(2000);
    expect((run.value as string).slice(0, 12)).toBe('yxxxxxxxxxyx');
  });
});
