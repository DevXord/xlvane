// insertRows and deleteRows after the K3 review (B6): charts, pivot caches and sparklines on other
// sheets that may refer to the changed sheet were found by searching the part's XML for the sheet
// name as written in the workbook. A name that needs quoting or escaping is written differently in
// those parts ("'O''Brien'!", "'R&amp;D'!", "A&lt;B", character references), so the rows moved and
// the references stayed. The name is now searched for after the XML references are decoded, also
// in its quoted form. Pivot cache definitions are also found by their content type and by the
// relationship type of the workbook (both as written by Microsoft Excel in the corpus files
// poi/ExcelPivotTableSample.xlsx, poi/54436.xlsx and poi/DataTableCities.xlsx), not only by name.
import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { deleteRows, insertRows } from '../../packages/xlvane/src/rows';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';

const M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PIVOT_CACHE_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml';

function raw(sheets: { name: string; xml: string }[], extra: Record<string, string> = {}, ct = '', wbRels = '', wbExtra = ''): Uint8Array {
  const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${String(i + 1)}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}${ct}</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="${M}" xmlns:r="${R}"><sheets>${sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${String(i + 1)}" r:id="rId${String(i + 1)}"/>`).join('')}</sheets>${wbExtra}</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${String(i + 1)}" Type="${R}/worksheet" Target="worksheets/sheet${String(i + 1)}.xml"/>`).join('')}${wbRels}</Relationships>`),
  };
  sheets.forEach((s, i) => (files[`xl/worksheets/sheet${String(i + 1)}.xml`] = strToU8(`<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${M}">${s.xml}</worksheet>`)));
  for (const [k, v] of Object.entries(extra)) files[k] = strToU8(v);
  return zipSync(files);
}

const data = `<sheetData>${[1, 2, 3].map((i) => `<row r="${String(i)}"><c r="A${String(i)}"><v>${String(i)}</v></c></row>`).join('')}</sheetData>`;

function sparklineSheet(formula: string): string {
  return `<sheetData/><extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>${formula}</xm:f><xm:sqref>B1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>`;
}

function pivotCache(sheet: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><pivotCacheDefinition xmlns="${M}"><cacheSource type="worksheet"><worksheetSource ref="A1:A3" sheet="${sheet}"/></cacheSource><cacheFields count="1"><cacheField name="1"><sharedItems/></cacheField></cacheFields></pivotCacheDefinition>`;
}

describe('B6: references to a sheet whose name is quoted or escaped in XML', () => {
  for (const name of ["O'Brien", 'R&D', 'A<B', "It's R&D"]) {
    it(`a chart on another sheet referring to "${name}"`, () => {
      const wb = createWorkbook();
      const d = wb.addSheet(name);
      for (let i = 1; i <= 6; i++) d.set(`A${String(i)}`, `m${String(i)}`).set(`B${String(i)}`, i);
      const q = `'${name.replace(/'/g, "''")}'`;
      wb.addSheet('Charts').addChart({ type: 'column', anchor: 'A1:F10', series: [{ categories: `${q}!A1:A6`, values: `${q}!B1:B6` }] });
      const bytes = writeWorkbook(wb);
      const error = expectXlvError(() => insertRows(bytes, { sheet: name, at: 1, count: 3 }), 'XLV_E_NOT_SUPPORTED');
      expect(error.message).toMatch(/may refer to sheet/);
      expectXlvError(() => deleteRows(bytes, { sheet: name, at: 1 }), 'XLV_E_NOT_SUPPORTED');
    });
  }

  for (const [label, name, formula] of [
    ['an entity', 'R&D', "'R&amp;D'!A1:A3"],
    ['a doubled quote', "O'Brien", "'O''Brien'!A1:A3"],
    ['an escaped "<"', 'A<B', "'A&lt;B'!A1:A3"],
    ['character references', 'R&D', "'&#82;&#x26;&#x44;'!A1:A3"],
    ['other letter case', 'R&D', "'r&amp;d'!A1:A3"],
  ] as const) {
    it(`a sparkline on another sheet referring to "${name}" with ${label}`, () => {
      expectXlvError(() => insertRows(raw([{ name, xml: data }, { name: 'S', xml: sparklineSheet(formula) }]), { sheet: name, at: 1 }), 'XLV_E_NOT_SUPPORTED');
    });
  }

  it('a sparkline that refers to another sheet only does not stop the rows', () => {
    const bytes = raw([{ name: 'R&D', xml: data }, { name: 'S', xml: sparklineSheet("'Q&amp;A'!A1:A3") }, { name: 'Q&A', xml: data }]);
    expect(() => insertRows(bytes, { sheet: 'R&D', at: 1 })).not.toThrow();
  });

  it('a pivot cache definition that names the sheet with an entity', () => {
    const bytes = raw([{ name: 'R&D', xml: data }], { 'xl/pivotCache/pivotCacheDefinition1.xml': pivotCache('R&amp;D') }, `<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="${PIVOT_CACHE_TYPE}"/>`);
    expectXlvError(() => insertRows(bytes, { sheet: 'R&D', at: 1 }), 'XLV_E_NOT_SUPPORTED');
  });
});

describe('B6: pivot cache definitions with other part names', () => {
  it('found by its content type', () => {
    const bytes = raw([{ name: 'Data', xml: data }], { 'xl/pc/def1.xml': pivotCache('Data') }, `<Override PartName="/xl/pc/def1.xml" ContentType="${PIVOT_CACHE_TYPE}"/>`);
    expect(() => insertRows(bytes, { sheet: 'Data', at: 1 })).toThrow(/may refer to sheet/);
  });

  it('found by the relationship of the workbook', () => {
    const bytes = raw([{ name: 'Data', xml: data }], { 'xl/pc/def1.xml': pivotCache('Data') }, '', `<Relationship Id="rIdP" Type="${R}/pivotCacheDefinition" Target="pc/def1.xml"/>`, `<pivotCaches><pivotCache cacheId="1" r:id="rIdP"/></pivotCaches>`);
    expect(() => insertRows(bytes, { sheet: 'Data', at: 1 })).toThrow(/may refer to sheet/);
  });

  it('found by both (the reproduction of the review)', () => {
    const bytes = raw([{ name: 'Data', xml: data }], { 'xl/pc/def1.xml': pivotCache('Data') }, `<Override PartName="/xl/pc/def1.xml" ContentType="${PIVOT_CACHE_TYPE}"/>`, `<Relationship Id="rIdP" Type="${R}/pivotCacheDefinition" Target="pc/def1.xml"/>`, `<pivotCaches><pivotCache cacheId="1" r:id="rIdP"/></pivotCaches>`);
    expect(() => deleteRows(bytes, { sheet: 'Data', at: 1 })).toThrow(/may refer to sheet/);
  });
});
