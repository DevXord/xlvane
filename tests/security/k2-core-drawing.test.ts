// A Drawings part shared by many sheets was decompressed and parsed once per sheet (1.3.0
// penetration test: 50 sheets sharing one drawing with 10 000 chart frames took 12 s). As for
// sheet, comments, VML and table parts (D-H01), the drawing belongs to the first worksheet that
// relates to it; the others get no pictures or charts from it and XLV_W_SHEET_PART_SHARED. The
// root element named in the warning about a foreign drawing part is shortened (util/clip.ts).
import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0]);

const rels = (list: readonly (readonly [string, string, string])[]): string =>
  `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${PR}">${list.map(([id, type, target]) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`).join('')}</Relationships>`;

const marker = (tag: string, col: number, row: number): string => `<xdr:${tag}><xdr:col>${String(col)}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${String(row)}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${tag}>`;
const chartFrame = (i: number): string =>
  `<xdr:twoCellAnchor>${marker('from', 1, i)}${marker('to', 5, i + 8)}<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${String(i + 2)}" name="Chart ${String(i + 1)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" r:id="rC"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
const picture = `<xdr:oneCellAnchor>${marker('from', 0, 0)}<xdr:ext cx="100" cy="100"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rI"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;

/** `sheets` worksheets that all relate to one drawing part with `frames` chart frames and a picture. */
function sharedDrawing(sheets: number, frames: number, drawingXml?: string): Uint8Array {
  const files: Record<string, Uint8Array> = {};
  const overrides: [string, string][] = [
    ['/xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'],
    ['/xl/drawings/drawing1.xml', 'application/vnd.openxmlformats-officedocument.drawing+xml'],
    ['/xl/charts/chart1.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'],
  ];
  for (let i = 1; i <= sheets; i++) {
    overrides.push([`/xl/worksheets/sheet${String(i)}.xml`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml']);
    files[`xl/worksheets/sheet${String(i)}.xml`] = strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheetData><row r="1"><c r="A1"><v>${String(i)}</v></c></row></sheetData><drawing r:id="rD"/></worksheet>`,
    );
    // The last sheet names the drawing with another case: part names compare without regard to case.
    files[`xl/worksheets/_rels/sheet${String(i)}.xml.rels`] = strToU8(rels([['rD', 'drawing', i === sheets && i > 1 ? '../Drawings/DRAWING1.xml' : '../drawings/drawing1.xml']]));
  }
  files['[Content_Types].xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>${overrides.map(([p, t]) => `<Override PartName="${p}" ContentType="${t}"/>`).join('')}</Types>`,
  );
  files['_rels/.rels'] = strToU8(rels([['rId1', 'officeDocument', 'xl/workbook.xml']]));
  files['xl/workbook.xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheets>${Array.from({ length: sheets }, (_, i) => `<sheet name="S${String(i + 1)}" sheetId="${String(i + 1)}" r:id="rS${String(i + 1)}"/>`).join('')}</sheets></workbook>`,
  );
  files['xl/_rels/workbook.xml.rels'] = strToU8(rels(Array.from({ length: sheets }, (_, i) => [`rS${String(i + 1)}`, 'worksheet', `worksheets/sheet${String(i + 1)}.xml`] as const)));
  files['xl/drawings/drawing1.xml'] = strToU8(drawingXml ?? `<?xml version="1.0" encoding="UTF-8"?><xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${R}">${picture}${Array.from({ length: frames }, (_, i) => chartFrame(i)).join('')}</xdr:wsDr>`);
  files['xl/drawings/_rels/drawing1.xml.rels'] = strToU8(rels([['rC', 'chart', '../charts/chart1.xml'], ['rI', 'image', '../media/image1.png']]));
  files['xl/charts/chart1.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8"?><c:chartSpace xmlns:c="${C}"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>S1!$A$1</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`);
  files['xl/media/image1.png'] = PNG;
  return zipSync(files, { level: 9 });
}

describe('a drawing part shared by several sheets', () => {
  it('belongs to the first sheet that relates to it; the others get XLV_W_SHEET_PART_SHARED', () => {
    const wb = readWorkbook(sharedDrawing(3, 2));
    expect(wb.getSheet('S1').charts).toHaveLength(2);
    expect(wb.getSheet('S1').images).toHaveLength(1);
    for (const name of ['S2', 'S3']) {
      expect(wb.getSheet(name).charts).toEqual([]);
      expect(wb.getSheet(name).images).toEqual([]);
      // The cells of the sheet are read as before.
      expect(wb.getSheet(name).cell('A1')?.value).toBe(Number(name.slice(1)));
    }
    expect(wb.warnings.map((w) => [w.code, w.location.sheet])).toEqual([
      ['XLV_W_SHEET_PART_SHARED', 'S2'],
      ['XLV_W_SHEET_PART_SHARED', 'S3'],
    ]);
    expect(wb.warnings[0]!.message).toContain('belongs to sheet "S1"');
    expect(wb.images).toHaveLength(1);
  });

  it('belongs to the first sheet that shows it, not to a sheet whose relationships only list it (review of D-K03)', async () => {
    const { unzipSync } = await import('fflate');
    const files = unzipSync(sharedDrawing(2, 1));
    files['xl/worksheets/sheet1.xml'] = strToU8(new TextDecoder().decode(files['xl/worksheets/sheet1.xml']).replace('<drawing r:id="rD"/>', ''));
    const wb = readWorkbook(zipSync(files));
    expect(wb.getSheet('S1').charts).toEqual([]);
    expect(wb.getSheet('S2').charts).toHaveLength(1);
    expect(wb.getSheet('S2').images).toHaveLength(1);
    expect(wb.warnings.map((w) => w.code)).not.toContain('XLV_W_SHEET_PART_SHARED');
  });

  it('parses the drawing once: 50 sheets and 10 000 chart frames in linear time', () => {
    const bytes = sharedDrawing(50, 10_000);
    const t0 = performance.now();
    const wb = readWorkbook(bytes, { limits: { maxCompressionRatio: Number.POSITIVE_INFINITY } });
    const ms = performance.now() - t0;
    expect(wb.getSheet('S1').charts).toHaveLength(10_000);
    expect(wb.sheets.reduce((n, s) => n + s.charts.length, 0)).toBe(10_000);
    expect(wb.warnings.filter((w) => w.code === 'XLV_W_SHEET_PART_SHARED')).toHaveLength(49);
    // Before: 12 s for the whole read (each sheet parsed the drawing again).
    expect(ms).toBeLessThan(4000);
  }, 60_000);
});

describe('a foreign drawing part', () => {
  it('shortens the root element named in the warning', () => {
    // 1000 characters with the default maxXmlNameLength (1024), 100 000 with a raised one.
    for (const [length, limits] of [[1000, {}], [100_000, { maxXmlNameLength: 1_000_000 }]] as const) {
      const local = 'x'.repeat(length);
      const wb = readWorkbook(sharedDrawing(1, 0, `<?xml version="1.0" encoding="UTF-8"?><${local}/>`), { limits });
      const warning = wb.warnings.find((w) => w.code === 'XLV_W_INVALID_VALUE');
      expect(warning?.message).toContain(`<${'x'.repeat(200)}…>`);
      expect(warning!.message.length).toBeLessThan(400);
    }
    // The cut does not split a surrogate pair (U+1F600 at code units 199 and 200).
    const wb = readWorkbook(sharedDrawing(1, 0, `<?xml version="1.0" encoding="UTF-8"?><${'x'.repeat(199)}\u{1F600}x/>`));
    expect(wb.warnings.find((w) => w.code === 'XLV_W_INVALID_VALUE')?.message).toContain(`<${'x'.repeat(199)}…>`);
  });
});
