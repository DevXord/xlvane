// Tables: ECMA-376 Part 1 §18.5, §12.3.21, §18.3.1.94/.95; [MS-OI29500] §2.1.679 … §2.1.684,
// §2.1.1094; docs/facts/tables-drawings.md TD-01 … TD-12; docs/decisions.md D-D02.
import { readFileSync } from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { editCells } from '../../packages/xlvane/src/edit';
import { inspect } from '../../packages/xlvane/src/inspect';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { readTable } from '../../packages/xlvane/src/table';
import { columnReference } from '../../packages/xlvane/src/tables';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx, NS } from '../helpers/xlsx-builder';

const parts = (bytes: Uint8Array): Record<string, string> => Object.fromEntries(Object.entries(unzipSync(bytes)).map(([k, v]) => [k, strFromU8(v)]));
const rels = (items: string): string => `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
const corpus = (name: string): Uint8Array => readFileSync(new URL(`../corpus/poi/${name}`, import.meta.url));

function salesBook() {
  const wb = createWorkbook();
  wb.addSheet('Data').addTable({
    name: 'Sales',
    ref: 'B2',
    totalsRow: true,
    style: { theme: 'TableStyleMedium9', showRowStripes: true },
    columns: [
      { name: 'Region', totalsRowLabel: 'Total' },
      { name: 'Revenue', totalsRowFunction: 'sum' },
      { name: 'Units', totalsRowFunction: 'average', filterButton: false },
      { name: 'Double', calculatedColumnFormula: 'Sales[[#This Row],[Revenue]]*2' },
    ],
    rows: [
      ['North', 1200, 3],
      ['South', 980, 4],
    ],
  });
  return wb;
}

describe('writing tables', () => {
  it('writes the cells, the table part, its relationship, tableParts and the content type', () => {
    const p = parts(writeWorkbook(salesBook()));
    const table = p['xl/tables/table1.xml'] ?? '';
    expect(table).toContain('<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="B2:E5" totalsRowCount="1">');
    expect(table).toContain('<autoFilter ref="B2:E4"><filterColumn colId="2" hiddenButton="1"/></autoFilter>');
    expect(table).toContain('<tableColumn id="1" name="Region" totalsRowLabel="Total"/>');
    expect(table).toContain('<tableColumn id="2" name="Revenue" totalsRowFunction="sum"/>');
    expect(table).toContain('<tableColumn id="4" name="Double"><calculatedColumnFormula>Sales[[#This Row],[Revenue]]*2</calculatedColumnFormula></tableColumn>');
    expect(table).toContain('<tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>');
    expect(p['xl/worksheets/_rels/sheet1.xml.rels']).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"');
    const sheet = p['xl/worksheets/sheet1.xml'] ?? '';
    expect(sheet).toMatch(/<tableParts count="1"><tablePart r:id="rId1"\/><\/tableParts><\/worksheet>$/);
    expect(p['[Content_Types].xml']).toContain('<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>');
    const read = readWorkbook(writeWorkbook(salesBook())).getSheet('Data');
    expect(read.cell('B2')?.value).toBe('Region');
    expect(read.cell('C5')?.formula?.text).toBe('SUBTOTAL(109,Sales[Revenue])');
    expect(read.cell('D5')?.formula?.text).toBe('SUBTOTAL(101,Sales[Units])');
    expect(read.cell('B5')?.value).toBe('Total');
    expect(read.cell('E3')?.formula?.text).toBe('Sales[[#This Row],[Revenue]]*2');
  });

  it('numbers tables across the workbook and writes no totals row when not asked', () => {
    const wb = createWorkbook();
    wb.addSheet('A').addTable({ name: 'One', ref: 'A1', columns: [{ name: 'x' }], rows: [[1]] });
    wb.addSheet('B').addTable({ name: 'Two', ref: 'A1:B3', headerRow: true, columns: [{ name: 'x' }, { name: 'y' }] }).addTable({ name: 'Three', ref: 'D1', headerRow: false, columns: [{ name: 'z' }] });
    const p = parts(writeWorkbook(wb));
    expect(p['xl/tables/table2.xml']).toContain('id="2" name="Two" displayName="Two" ref="A1:B3" totalsRowShown="0">');
    expect(p['xl/tables/table3.xml']).toContain('ref="D1:D1" headerRowCount="0" totalsRowShown="0">');
    expect(readWorkbook(writeWorkbook(wb)).getSheet('B').tables.map((t) => t.ref)).toEqual(['A1:B3', 'D1']);
    expect(p['xl/tables/table3.xml']).not.toContain('<autoFilter');
    expect(p['xl/worksheets/_rels/sheet2.xml.rels']).toContain('Target="../tables/table3.xml"');
  });

  it('escapes column names in totals formulas as Excel does', () => {
    expect(columnReference('T', 'Units')).toBe('T[Units]');
    expect(columnReference('T', 'Field 4 ')).toBe('T[[Field 4 ]]');
    expect(columnReference('T', "Q1 [est]#'")).toBe("T[[Q1 '[est']'#'']]");
  });

  it('round-trips through the reader', () => {
    const [t] = readWorkbook(writeWorkbook(salesBook())).getSheet('Data').tables;
    expect(t).toMatchObject({
      name: 'Sales',
      displayName: 'Sales',
      ref: 'B2:E5',
      headerRowCount: 1,
      totalsRowCount: 1,
      autoFilter: 'B2:E4',
      style: { theme: 'TableStyleMedium9', showRowStripes: true, showFirstColumn: false },
    });
    expect(t?.columns.map((c) => [c.name, c.totalsRowFunction, c.totalsRowLabel, c.filterButton])).toEqual([
      ['Region', 'none', 'Total', true],
      ['Revenue', 'sum', null, true],
      ['Units', 'average', null, false],
      ['Double', 'none', null, true],
    ]);
    expect(t?.columns[3]?.calculatedColumnFormula).toBe('Sales[[#This Row],[Revenue]]*2');
  });

  it('refuses invalid tables with XLV_E_INVALID_TABLE', () => {
    const sheet = () => createWorkbook().addSheet('S');
    expectXlvError(() => sheet().addTable({ name: 'My Table', ref: 'A1', columns: [{ name: 'a' }] }), 'XLV_E_INVALID_TABLE');
    expectXlvError(() => sheet().addTable({ name: 'A1', ref: 'A1', columns: [{ name: 'a' }] }), 'XLV_E_INVALID_TABLE');
    expectXlvError(() => sheet().addTable({ name: 'T', ref: 'A1', columns: [{ name: 'a' }, { name: 'A' }] }), 'XLV_E_INVALID_TABLE');
    expectXlvError(() => sheet().addTable({ name: 'T', ref: 'A1', columns: [] }), 'XLV_E_INVALID_TABLE');
    expectXlvError(() => sheet().addTable({ name: 'T', ref: 'A1:C3', columns: [{ name: 'a' }] }), 'XLV_E_INVALID_TABLE');
    expectXlvError(() => sheet().addTable({ name: 'T', ref: 'A1:A2', totalsRow: true, columns: [{ name: 'a' }] }), 'XLV_E_INVALID_TABLE');
    expectXlvError(() => sheet().addTable({ name: 'T', ref: 'A1', columns: [{ name: 'a', totalsRowFunction: 'sum', totalsRowLabel: 'x' }] }), 'XLV_E_INVALID_TABLE');
    expectXlvError(() => sheet().addTable({ name: 'T', ref: 'A1', columns: [{ name: 'a', totalsRowFunction: 'custom' }] }), 'XLV_E_INVALID_TABLE');
    expectXlvError(() => sheet().addTable({ name: 'T', ref: 'A1', style: { theme: 'Fancy' }, columns: [{ name: 'a' }] }), 'XLV_E_INVALID_TABLE');
    expectXlvError(() => sheet().addTable({ name: 'T', ref: 'XFD1', columns: [{ name: 'a' }, { name: 'b' }] }), 'XLV_E_INVALID_TABLE');
    expectXlvError(() => sheet().addTable(JSON.parse('{"name":"T","ref":"A1","columns":[{"name":"a"}],"__proto__":{"x":1}}')), 'XLV_E_INVALID_TABLE');
    const wb = createWorkbook();
    const s = wb.addSheet('S').addTable({ name: 'T', ref: 'A1', columns: [{ name: 'a' }] });
    expectXlvError(() => wb.addSheet('U').addTable({ name: 't', ref: 'A1', columns: [{ name: 'a' }] }), 'XLV_E_INVALID_TABLE');
    expectXlvError(() => s.addTable({ name: 'T2', ref: 'A2', columns: [{ name: 'a' }] }), 'XLV_E_INVALID_TABLE');
    expectXlvError(() => s.merge('A1:B1'), 'XLV_E_INVALID_TABLE');
    expectXlvError(() => wb.defineName('t', 'S!$A$1'), 'XLV_E_INVALID_DEFINED_NAME');
    wb.defineName('Rate', 'S!$C$1');
    expectXlvError(() => s.addTable({ name: 'rate', ref: 'D1', columns: [{ name: 'a' }] }), 'XLV_E_INVALID_TABLE');
    const merged = createWorkbook().addSheet('M').merge('B2:C2');
    expectXlvError(() => merged.addTable({ name: 'T', ref: 'A1:C3', columns: [{ name: 'a' }, { name: 'b' }, { name: 'c' }] }), 'XLV_E_INVALID_TABLE');
  });

  it('streams tables declared with their whole range', async () => {
    const stream = writeWorkbookStream({
      sheets: [
        {
          name: 'Data',
          tables: [{ name: 'Sales', ref: 'A1:B3', totalsRow: true, columns: [{ name: 'Region', totalsRowLabel: 'Total' }, { name: 'Revenue', totalsRowFunction: 'sum' }] }],
          rows: [{ values: ['Region', 'Revenue'] }, { values: ['North', 1] }, { values: ['Total', { formula: 'SUBTOTAL(109,Sales[Revenue])' }] }],
        },
      ],
    });
    const bytes = new Uint8Array(await new Response(stream).arrayBuffer());
    const p = parts(bytes);
    expect(p['xl/tables/table1.xml']).toContain('ref="A1:B3" totalsRowCount="1"');
    expect(p['[Content_Types].xml']).toContain('/xl/tables/table1.xml');
    expect(readWorkbook(bytes).getSheet('Data').tables[0]?.name).toBe('Sales');
    expect(() => writeWorkbookStream({ sheets: [{ name: 'S', tables: [{ name: 'T', ref: 'A1', columns: [{ name: 'a' }] }], rows: [] }] })).toThrow(/whole range/);
  });
});

describe('reading tables', () => {
  it('reads tables saved by Excel (POI corpus)', () => {
    const wb = readWorkbook(corpus('DataTableCities.xlsx'));
    const tables = wb.sheets.flatMap((s) => s.tables).sort((a, b) => a.ref.localeCompare(b.ref));
    expect(tables.map((t) => [t.name, t.ref, t.autoFilter, t.totalsRowCount])).toEqual([
      ['BigCity', 'C2:F15', 'C2:F14', 1],
      ['SmallCity', 'I2:L15', 'I2:L14', 1],
    ]);
    expect(tables[0]?.columns.map((c) => c.totalsRowFunction)).toEqual(['none', 'average', 'average', 'sum']);
    const sample = readWorkbook(corpus('table-sample.xlsx')).sheets.flatMap((s) => s.tables)[0];
    expect(sample?.columns[3]?.calculatedColumnFormula).toBe('SUM(Tabelle1[[#This Row],[Field 2]:[Field 3]])');
  });

  it('reads hand-made parts, tolerates broken ones and Strict files', () => {
    const table = (ns: string, body: string): string => `<?xml version="1.0"?><x:table xmlns:x="${ns}" id="1" displayName="T" ref="a1:b3" headerRowCount="1">${body}</x:table>`;
    const build = (opts: { strict?: boolean; part: string; target?: string }) => {
      const ns = opts.strict ? NS.strict : NS.transitional;
      return buildXlsx({
        strict: opts.strict === true,
        sheets: [{ name: 'S', data: '<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c></row>' }],
        extraParts: {
          'xl/worksheets/_rels/sheet1.xml.rels': rels(`<Relationship Id="rId1" Type="${ns.rel}table" Target="${opts.target ?? '../tables/t.xml'}"/>`),
          'xl/tables/t.xml': opts.part,
        },
      });
    };
    for (const strict of [false, true]) {
      const ns = strict ? NS.strict : NS.transitional;
      const t = readWorkbook(build({ strict, part: table(ns.main, `<x:tableColumns count="2"><x:tableColumn id="1" name="a"/><x:tableColumn id="2" name="b_x0020_"/></x:tableColumns>`) })).getSheet('S').tables[0];
      expect(t).toMatchObject({ name: 'T', displayName: 'T', ref: 'A1:B3', style: null, autoFilter: null });
      expect(t?.columns.map((c) => c.name)).toEqual(['a', 'b ']);
    }
    const abs = readWorkbook(build({ part: table(NS.transitional.main, '<x:tableColumns><x:tableColumn id="1" name="a"/></x:tableColumns>'), target: '/xl/tables/t.xml' }));
    expect(abs.getSheet('S').tables).toHaveLength(1);
    const broken = readWorkbook(build({ part: '<table' }));
    expect(broken.getSheet('S').tables).toEqual([]);
    expect(broken.warnings.map((w) => w.code)).toContain('XLV_W_INVALID_VALUE');
    const missing = readWorkbook(build({ part: '<x/>', target: '../tables/none.xml' }));
    expect(missing.warnings.map((w) => w.code)).toContain('XLV_W_PART_MISSING');
    const badRef = readWorkbook(build({ part: `<table xmlns="${NS.transitional.main}" id="1" displayName="T" ref="nope"><tableColumns><tableColumn id="1" name="a"/></tableColumns></table>` }));
    expect(badRef.getSheet('S').tables).toEqual([]);
  });
});

describe('readTable by table name, inspect and editing', () => {
  it('reads a table by name (data rows only, header from the table)', () => {
    const bytes = writeWorkbook(salesBook());
    const result = readTable(bytes, { table: 'sales' });
    expect(result.sheet).toBe('Data');
    expect(result.rows).toEqual([
      { Region: 'North', Revenue: 1200, Units: 3, Double: null },
      { Region: 'South', Revenue: 980, Units: 4, Double: null },
    ]);
    expectXlvError(() => readTable(bytes, { table: 'Nope' }), 'XLV_E_TABLE_NOT_FOUND');
    expectXlvError(() => readTable(bytes, { table: 'Sales', sheet: 'Data' }), 'XLV_E_INVALID_OPTION');
  });

  it('inspect lists the tables of a sheet', () => {
    const sheet = inspect(readWorkbook(writeWorkbook(salesBook()))).sheets[0];
    expect(sheet?.tables).toEqual([{ name: 'Sales', ref: 'B2:E5', columns: ['Region', 'Revenue', 'Units', 'Double'] }]);
  });

  it('editCells keeps the table parts byte for byte', () => {
    const bytes = writeWorkbook(salesBook());
    const edited = editCells(bytes, [{ sheet: 'Data', cell: 'H1', value: 'note' }]).bytes;
    expect(parts(edited)['xl/tables/table1.xml']).toBe(parts(bytes)['xl/tables/table1.xml']);
  });
});

describe('review fixes (D-D05)', () => {
  const table = { name: 'T', ref: 'A1', columns: [{ name: 'A' }, { name: 'B' }], rows: [[1, 2]] };

  it('refuses a table that overlaps the sheet filter, in either order and in the streaming writer', () => {
    const one = createWorkbook().addSheet('S');
    one.addTable(table);
    expectXlvError(() => one.autoFilter('A1:A5'), 'XLV_E_INVALID_OPTION');
    one.autoFilter('D1:E3');
    const two = createWorkbook().addSheet('S');
    two.autoFilter('A1:A5');
    expectXlvError(() => two.addTable(table), 'XLV_E_INVALID_TABLE');
    expectXlvError(() => writeWorkbookStream({ sheets: [{ name: 'S', autoFilter: true, tables: [{ name: 'T', ref: 'A1:B2', columns: [{ name: 'A' }, { name: 'B' }] }], rows: [['A', 'B'], [1, 2]] }] }), 'XLV_E_INVALID_TABLE');
  });

  it('refuses table names equal to built-in names', () => {
    for (const name of ['Print_Area', 'print_titles', '_xlnm.Print_Area', 'Criteria', 'Extract', 'Database']) {
      expectXlvError(() => createWorkbook().addSheet('S').addTable({ ...table, name }), 'XLV_E_INVALID_TABLE');
    }
  });
});
