// Pictures on worksheets: Drawings part (ECMA-376 Part 1 §12.3.8, §20.5), Image parts
// (§15.2.14), sheet background (§18.3.1.67); docs/facts/tables-drawings.md TD-30 … TD-39;
// docs/decisions.md D-D03. Images come from the Apache POI corpus (tests/corpus/LICENSES.md) or
// are built here.
import { readFileSync } from 'node:fs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { editCells } from '../../packages/xlvane/src/edit';
import { deleteRows, insertRows } from '../../packages/xlvane/src/rows';
import { imageFormat } from '../../packages/xlvane/src/images';
import { inspect } from '../../packages/xlvane/src/inspect';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';

const corpus = (name: string): Uint8Array => readFileSync(new URL(`../corpus/poi/${name}`, import.meta.url));
const mediaOf = (file: string, ext: string): Uint8Array => Object.entries(unzipSync(corpus(file))).find(([n]) => n.startsWith('xl/media/') && n.endsWith(ext))![1];
const PNG = mediaOf('WithDrawing.xlsx', '.png');
const JPEG = mediaOf('WithDrawing.xlsx', '.jpeg');
const GIF = new Uint8Array([...strToU8('GIF89a'), 1, 0, 1, 0, 0, 0, 0, 0x3b]);

function parts(bytes: Uint8Array): Record<string, string> {
  return Object.fromEntries(Object.entries(unzipSync(bytes)).map(([k, v]) => [k, strFromU8(v)]));
}

async function bytesOf(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** Replaces part texts of a package. */
function patch(bytes: Uint8Array, edit: (name: string, text: string) => string | null): Uint8Array {
  const files = unzipSync(bytes);
  const out: Record<string, Uint8Array> = {};
  for (const [name, data] of Object.entries(files)) {
    if (/\.(png|jpeg|gif)$/.test(name)) {
      if (edit(name, '') !== null) out[name] = data;
    } else {
      const t = edit(name, strFromU8(data));
      if (t !== null) out[name] = strToU8(t);
    }
  }
  return zipSync(out);
}

function sample(): Uint8Array {
  const wb = createWorkbook();
  const logo = wb.addImage(PNG);
  const photo = wb.addImage(JPEG);
  const sheet = wb.addSheet('Report').set('A1', 'Title').column(2, { width: 10 });
  sheet.addImage(logo, 'B2:D6', { description: 'Logo', hyperlink: 'https://example.com' });
  sheet.addImage(photo, { from: { col: 1.5, row: 8 }, to: 'F12', editAs: 'oneCell' });
  sheet.addImage(logo, { from: 'H2', size: { width: 100, height: 50 } });
  return writeWorkbook(wb);
}

describe('image formats', () => {
  it('are recognised from their first bytes (TD-31)', () => {
    expect(imageFormat(PNG)).toBe('png');
    expect(imageFormat(JPEG)).toBe('jpeg');
    expect(imageFormat(GIF)).toBe('gif');
    expect(imageFormat(strToU8('GIF87a...'))).toBe('gif');
    expect(imageFormat(mediaOf('WithDrawing.xlsx', '.emf'))).toBeNull();
    expect(imageFormat(new Uint8Array([0x89, 0x50]))).toBeNull();
  });

  it('addImage accepts PNG, JPEG and GIF only', () => {
    const wb = createWorkbook();
    expect([wb.addImage(PNG), wb.addImage(JPEG), wb.addImage(GIF)]).toEqual([0, 1, 2]);
    expectXlvError(() => wb.addImage(strToU8('<svg/>')), 'XLV_E_INVALID_IMAGE');
    expectXlvError(() => wb.addImage('logo.png' as unknown as Uint8Array), 'XLV_E_INVALID_IMAGE');
  });
});

describe('writing pictures', () => {
  it('writes a drawing part, its relationships, the media once and the content types', () => {
    const p = parts(sample());
    expect(Object.keys(p).filter((n) => n.startsWith('xl/media/'))).toEqual(['xl/media/image1.png', 'xl/media/image2.jpeg']);
    expect(p['[Content_Types].xml']).toContain('<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>');
    expect(p['[Content_Types].xml']).toContain('<Override PartName="/xl/media/image2.jpeg" ContentType="image/jpeg"/>');
    expect(p['xl/worksheets/_rels/sheet1.xml.rels']).toContain('relationships/drawing" Target="../drawings/drawing1.xml"');
    const rels = p['xl/drawings/_rels/drawing1.xml.rels']!;
    // One relationship per image, also when an image is placed twice (exceljs repeats them).
    expect(rels.match(/relationships\/image"/g)).toHaveLength(2);
    expect(rels).toContain('Target="https://example.com" TargetMode="External"');
    const d = p['xl/drawings/drawing1.xml']!;
    expect(d).toContain('<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row>');
    // Column B is 10 wide = 70 px; half of it is 35 px = 333 375 EMU (TD-33).
    expect(d).toContain('<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>333375</xdr:colOff>');
    expect(d).toContain('<xdr:ext cx="952500" cy="476250"/>');
    expect(d.match(/<xdr:cNvPr id="(\d+)"/g)).toEqual(['<xdr:cNvPr id="1"', '<xdr:cNvPr id="2"', '<xdr:cNvPr id="3"']);
    expect(d).toContain('descr="Logo"><a:hlinkClick r:id="rId2"/></xdr:cNvPr>');
    expect(d).not.toContain('creationId');
    expect(p['xl/worksheets/sheet1.xml']).toMatch(/<\/sheetData><drawing r:id="rId1"\/><\/worksheet>$/);
  });

  it('keeps the CT_Worksheet order with notes, a background and the drawing (TD-34)', () => {
    const wb = createWorkbook();
    const id = wb.addImage(GIF);
    wb.addSheet('S').set('A1', 1).note('A1', 'n').addImage(id, 'B2:C3').background(id);
    const p = parts(writeWorkbook(wb));
    expect(p['xl/worksheets/sheet1.xml']).toMatch(/<drawing r:id="rId\d"\/><legacyDrawing r:id="rId\d"\/><picture r:id="rId\d"\/><\/worksheet>$/);
    expect(Object.keys(p).filter((n) => n.startsWith('xl/media/'))).toEqual(['xl/media/image1.gif']);
  });

  it('stores identical images once across sheets', () => {
    const wb = createWorkbook();
    const a = wb.addImage(PNG);
    const b = wb.addImage(new Uint8Array(PNG));
    wb.addSheet('One').addImage(a, 'A1:B2');
    wb.addSheet('Two').addImage(b, 'A1:B2');
    const p = parts(writeWorkbook(wb));
    expect(Object.keys(p).filter((n) => n.startsWith('xl/media/'))).toEqual(['xl/media/image1.png']);
    expect(Object.keys(p).filter((n) => /drawings\/drawing\d\.xml$/.test(n))).toEqual(['xl/drawings/drawing1.xml', 'xl/drawings/drawing2.xml']);
  });

  it('validates the image, the position and the options', () => {
    const wb = createWorkbook();
    const id = wb.addImage(PNG);
    const s = wb.addSheet('S');
    expectXlvError(() => s.addImage(5, 'A1:B2'), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.addImage(id, 'A1:'), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.addImage(id, { from: 'C3', to: 'A1' }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.addImage(id, { from: 'C3', size: { width: 0, height: 10 } }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.addImage(id, { from: 'C3', size: { width: 10, height: 10 }, editAs: 'oneCell' } as never), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.addImage(id, { from: { col: -1, row: 0 }, to: 'B2' }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.addImage(id, { from: { col: 16_384, row: 0 }, to: 'B2' }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.addImage(id, 'A1:B2', { __proto__: { x: 1 }, alt: 'x' } as never), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.addImage(id, 'A1:B2', { hyperlink: 'a\u0001b' }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.background(7), 'XLV_E_INVALID_OPTION');
  });

  it('writes the same pictures with the streaming writer', async () => {
    const stream = writeWorkbookStream({
      images: [PNG, JPEG],
      sheets: [
        {
          name: 'Report',
          columns: [undefined, { width: 10 }],
          images: [
            { image: 0, anchor: 'B2:D6', description: 'Logo', hyperlink: 'https://example.com' },
            { image: 1, anchor: { from: { col: 1.5, row: 8 }, to: 'F12', editAs: 'oneCell' } },
            { image: 0, anchor: { from: 'H2', size: { width: 100, height: 50 } } },
          ],
          rows: [['Title']],
        },
      ],
    });
    const streamed = readWorkbook(await bytesOf(stream)).getSheet('Report').images;
    const batch = readWorkbook(sample()).getSheet('Report').images;
    const shape = (i: (typeof batch)[number]) => ({ ...i, image: i.image?.format });
    expect(streamed.map(shape)).toEqual(batch.map(shape));
  });
});

describe('reading pictures', () => {
  it('reads anchors, descriptions, links and images; images are shared objects', () => {
    const wb = readWorkbook(sample());
    const [a, b, c] = wb.getSheet('Report').images;
    expect(a).toMatchObject({ name: 'Picture 1', description: 'Logo', hyperlink: 'https://example.com', anchor: { type: 'twoCell', editAs: 'twoCell', from: { cell: 'B2', col: 1, row: 1 }, to: { cell: 'E7' } } });
    expect(b?.anchor).toEqual({ type: 'twoCell', editAs: 'oneCell', from: { cell: 'B9', col: 1, row: 8, colOffset: 333_375, rowOffset: 0 }, to: { cell: 'F12', col: 5, row: 11, colOffset: 0, rowOffset: 0 } });
    expect(c?.anchor).toEqual({ type: 'oneCell', from: { cell: 'H2', col: 7, row: 1, colOffset: 0, rowOffset: 0 }, size: { width: 952_500, height: 476_250 } });
    expect(a?.image).toBe(c?.image);
    expect(a?.image).toMatchObject({ part: '/xl/media/image1.png', format: 'png', contentType: 'image/png' });
    expect(a?.image?.bytes).toEqual(PNG);
    expect(wb.images.map((i) => i.format)).toEqual(['png', 'jpeg']);
    expect(wb.warnings).toEqual([]);
  });

  it('reads Excel files of the POI corpus and reports other drawing objects', () => {
    const wb = readWorkbook(corpus('WithDrawing.xlsx'));
    const s = wb.getSheet('Sheet1');
    expect(s.images.map((i) => i.image?.format ?? null)).toEqual(['jpeg', null, 'png', null, null]);
    expect(s.images.every((i) => i.anchor.type === 'twoCell')).toBe(true);
    expect(wb.warnings.map((w) => w.code)).toContain('XLV_W_DRAWING_CONTENT_NOT_READ');
    const absolute = readWorkbook(corpus('absolute-anchor-over-empty-sheet.xlsx'));
    expect(absolute.sheets.flatMap((sh) => sh.images.map((i) => i.anchor.type))).toContain('absolute');
  });

  it('accepts any namespace prefix and universal measures', () => {
    const file = patch(sample(), (name, text) =>
      name === 'xl/drawings/drawing1.xml' ? text.replace(/xdr:/g, 'wd:').replace('xmlns:xdr=', 'xmlns:wd=').replace('<wd:colOff>333375</wd:colOff>', '<wd:colOff>0.5in</wd:colOff>') : text,
    );
    const images = readWorkbook(file).getSheet('Report').images;
    expect(images).toHaveLength(3);
    expect(images[1]?.anchor.type === 'twoCell' && images[1].anchor.from.colOffset).toBe(457_200);
  });

  it('keeps reading when the drawing or an image is broken', () => {
    const broken = patch(sample(), (name, text) => (name === 'xl/drawings/drawing1.xml' ? '<xdr:wsDr' : text));
    const wb = readWorkbook(broken);
    expect(wb.getSheet('Report').images).toEqual([]);
    expect(wb.warnings.map((w) => w.code)).toEqual(['XLV_W_INVALID_VALUE']);
    const noMedia = patch(sample(), (name, text) => (name === 'xl/media/image2.jpeg' ? null : text));
    expect(readWorkbook(noMedia).getSheet('Report').images.map((i) => i.image?.format ?? null)).toEqual(['png', null, 'png']);
    const badMarker = patch(sample(), (name, text) => (name === 'xl/drawings/drawing1.xml' ? text.replace('<xdr:row>1</xdr:row>', '<xdr:row>-1</xdr:row>') : text));
    expect(readWorkbook(badMarker).getSheet('Report').images).toHaveLength(2);
  });

  it('bounds untrusted drawing XML by the read limits and never opens link targets', () => {
    const big = patch(sample(), (name, text) => (name === 'xl/drawings/drawing1.xml' ? text.replace('</xdr:wsDr>', `${'<xdr:sp/>'.repeat(50)}</xdr:wsDr>`) : text));
    expectXlvError(() => readWorkbook(big, { limits: { maxXmlTreeElements: 40 } }), 'XLV_E_LIMIT_EXCEEDED');
    const js = patch(sample(), (name, text) => (name === 'xl/drawings/_rels/drawing1.xml.rels' ? text.replace('https://example.com', 'javascript:alert(1)') : text));
    expect(readWorkbook(js).getSheet('Report').images[0]?.hyperlink).toBe('javascript:alert(1)');
  });

  it('reads the background image', () => {
    const wb = createWorkbook();
    wb.addSheet('Cover').background(wb.addImage(JPEG));
    const read = readWorkbook(writeWorkbook(wb));
    expect(read.getSheet('Cover').backgroundImage).toMatchObject({ format: 'jpeg', part: '/xl/media/image1.jpeg' });
    expect(read.images).toHaveLength(1);
  });

  it('inspect reports pictures', () => {
    const s = inspect(readWorkbook(sample())).sheets[0]!;
    expect(s.imageCount).toBe(3);
    expect(s.images).toEqual([
      { cell: 'B2', format: 'png' },
      { cell: 'B9', format: 'jpeg' },
      { cell: 'H2', format: 'png' },
    ]);
  });

  it('editCells keeps drawings and images byte for byte', () => {
    const file = sample();
    const out = editCells(file, [{ sheet: 'Report', cell: 'A2', value: 'x' }]).bytes;
    const before = unzipSync(file);
    const after = unzipSync(out);
    for (const name of ['xl/drawings/drawing1.xml', 'xl/drawings/_rels/drawing1.xml.rels', 'xl/media/image1.png', 'xl/media/image2.jpeg']) expect(after[name]).toEqual(before[name]);
  });
});

describe('inserting and deleting rows moves pictures (TD-38)', () => {
  const rowsOf = (bytes: Uint8Array) =>
    readWorkbook(bytes)
      .getSheet('Report')
      .images.map((i) => (i.anchor.type === 'twoCell' ? `${i.anchor.from.cell}:${i.anchor.to.cell}` : i.anchor.type === 'oneCell' ? i.anchor.from.cell : 'abs'));

  it('moves markers at or below the change and grows pictures that contain it', () => {
    expect(rowsOf(sample())).toEqual(['B2:E7', 'B9:F12', 'H2']);
    expect(rowsOf(insertRows(sample(), { sheet: 'Report', at: 4, count: 2 }).bytes)).toEqual(['B2:E9', 'B11:F14', 'H2']);
    expect(rowsOf(insertRows(sample(), { sheet: 'Report', at: 1 }).bytes)).toEqual(['B3:E8', 'B10:F13', 'H3']);
    expect(rowsOf(deleteRows(sample(), { sheet: 'Report', at: 7, count: 1 }).bytes)).toEqual(['B2:E7', 'B8:F11', 'H2']);
  });

  it('refuses a deletion that would cut a picture', () => {
    expectXlvError(() => deleteRows(sample(), { sheet: 'Report', at: 3, count: 1 }), 'XLV_E_NOT_SUPPORTED');
    expectXlvError(() => deleteRows(sample(), { sheet: 'Report', at: 9, count: 1 }), 'XLV_E_NOT_SUPPORTED');
  });

  it('refuses drawings with objects other than pictures', () => {
    const chart = patch(sample(), (name, text) => (name === 'xl/drawings/drawing1.xml' ? text.replace('</xdr:wsDr>', '<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="1" cy="1"/><xdr:graphicFrame/><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>') : text));
    expectXlvError(() => insertRows(chart, { sheet: 'Report', at: 1 }), 'XLV_E_NOT_SUPPORTED');
  });
});
