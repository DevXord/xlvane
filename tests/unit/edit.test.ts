import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { appendRows, editCells } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { readTable } from '../../packages/xlvane/src/table';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx, c, CT, NS, row } from '../helpers/xlsx-builder';

const dec = new TextDecoder();
const parts = (bytes: Uint8Array): Record<string, string> =>
  Object.fromEntries(Object.entries(unzipSync(bytes)).map(([k, v]) => [k, dec.decode(v)]));

function base(): Uint8Array {
  const wb = createWorkbook();
  const s = wb.addSheet('Data');
  s.set('A1', 'name', { font: { bold: true } }).set('B1', 'total', { font: { bold: true } });
  s.set('A2', 'Alice').set('B2', 10, { numFmt: '#,##0.00' }).set('A3', 'Bob').set('B3', 20);
  s.set('C2', { formula: 'B2*2', result: 20 });
  s.merge('D1:E1').freeze({ rows: 1 }).column('A', { width: 30 });
  wb.addSheet('Other').set('A1', 'keep me');
  return writeWorkbook(wb);
}

/** A package with parts xlvane does not model (comments, a drawing), a calc chain and an unknown attribute. */
function rich(): Uint8Array {
  const sheet =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<worksheet xmlns="${NS.transitional.main}" xmlns:r="${NS.transitional.r}" xmlns:x14ac="urn:x14ac" xmlns:mc="urn:mc">` +
    `<dimension ref="A1:B3"/><sheetData>` +
    `<row r="1" spans="1:2" x14ac:dyDescent="0.25">${c('A1', 0, 't="s"')}${c('B1', 5)}</row>` +
    `<row r="3">${c('A3', 1, 't="s"')}<c r="B3"><f>B1*2</f><v>10</v></c></row>` +
    `</sheetData><conditionalFormatting sqref="B1:B3"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>` +
    `<legacyDrawing r:id="rId1"/></worksheet>`;
  return buildXlsx({
    sheets: [{ name: 'S', data: sheet }],
    sharedStrings: ['first', 'third'],
    styles: '<fonts count="1"><font><sz val="10"/><name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>',
    extraParts: {
      'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${NS.transitional.rel}vmlDrawing" Target="../drawings/vmlDrawing1.vml"/></Relationships>`,
      'xl/drawings/vmlDrawing1.vml': '<xml><v:shape id="unknown"/></xml>',
      'xl/comments1.xml': '<comments/>',
      'xl/calcChain.xml': `<?xml version="1.0"?><calcChain xmlns="${NS.transitional.main}"><c r="B3" i="1"/></calcChain>`,
      // The builder's own relationships and content types, plus the calculation chain.
      'xl/_rels/workbook.xml.rels':
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="${NS.transitional.rel}worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId2" Type="${NS.transitional.rel}sharedStrings" Target="sharedStrings.xml"/>` +
        `<Relationship Id="rId3" Type="${NS.transitional.rel}styles" Target="styles.xml"/>` +
        `<Relationship Id="rId4" Type="${NS.transitional.rel}calcChain" Target="calcChain.xml"/></Relationships>`,
      '[Content_Types].xml':
        `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="${CT.relationships}"/><Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>` +
        `<Override PartName="/xl/workbook.xml" ContentType="${CT.workbook}"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="${CT.worksheet}"/>` +
        `<Override PartName="/xl/sharedStrings.xml" ContentType="${CT.sharedStrings}"/><Override PartName="/xl/styles.xml" ContentType="${CT.styles}"/>` +
        `<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>`,
    },
  });
}

describe('editCells', () => {
  it('changes values and keeps everything else', () => {
    const { bytes, warnings } = editCells(base(), [
      { sheet: 'Data', cell: 'B2', value: 15 },
      { sheet: 'Data', cell: 'A4', value: 'Carol' },
      { sheet: 'Data', cell: 'B4', value: 30.5 },
      { sheet: 'Data', cell: 'D2', value: true },
    ]);
    const wb = readWorkbook(bytes);
    const s = wb.getSheet('Data');
    expect(readTable(wb, { sheet: 'Data', range: 'A1:B4' }).rows).toEqual([
      { name: 'Alice', total: 15 },
      { name: 'Bob', total: 20 },
      { name: 'Carol', total: 30.5 },
    ]);
    expect(s.cell('B2')?.numFmt).toBe('#,##0.00'); // existing style kept
    expect(s.cell('A1')?.style.font?.bold).toBe(true);
    expect(s.cell('D2')?.value).toBe(true);
    expect(s.merges).toEqual(['D1:E1']);
    expect(s.freeze).toEqual({ rows: 1, cols: 0 });
    expect(s.columns[0]?.width).toBe(30);
    expect(s.dimension).toBe('A1:D4');
    expect(wb.getSheet('Other').cell('A1')?.value).toBe('keep me');
    expect(warnings.map((w) => w.code)).toEqual(['XLV_W_FORMULAS_NOT_RECALCULATED']);
  });

  it('writes formulas, errors, dates and removes values', () => {
    const { bytes } = editCells(base(), [
      { sheet: 'Data', cell: 'C3', value: { formula: '=B3*2', result: 40 } },
      { sheet: 'Data', cell: 'C2', value: null },
      { sheet: 'Data', cell: 'E5', value: { error: '#N/A' } },
      { sheet: 'Data', cell: 'F2', value: new Date(Date.UTC(2024, 0, 31)) },
    ]);
    const s = readWorkbook(bytes).getSheet('Data');
    expect(s.cell('C3')?.formula?.text).toBe('B3*2');
    expect(s.cell('C3')?.value).toBe(40);
    expect(s.cell('C2')).toBeUndefined();
    expect(s.cell('E5')?.value).toBe('#N/A');
    expect(s.cell('F2')?.value).toEqual(new Date(Date.UTC(2024, 0, 31)));
    expect(s.cell('F2')?.numFmt).toBe('yyyy-mm-dd');
  });

  it('applies styles on top of the existing cell style', () => {
    const { bytes } = editCells(base(), [
      { sheet: 'Data', cell: 'B2', value: 11, style: { fill: { color: 'FF0000' } } },
      { sheet: 'Data', cell: 'A1', value: 'Name', style: { font: { italic: true } } },
    ]);
    const s = readWorkbook(bytes).getSheet('Data');
    expect(s.cell('B2')?.style.fill).toEqual({ pattern: 'solid', fgColor: 'FFFF0000' });
    expect(s.cell('B2')?.numFmt).toBe('#,##0.00');
    expect(s.cell('A1')?.style.font).toMatchObject({ bold: true, italic: true });
    expect(s.cell('A3')?.style).toEqual({});
  });

  it('keeps parts, attributes and elements it does not model', () => {
    const before = parts(rich());
    const { bytes, warnings } = editCells(rich(), [
      { sheet: 'S', cell: 'B1', value: 7 },
      { sheet: 'S', cell: 'A2', value: 'second' },
      { sheet: 'S', cell: 'A3', value: 'third' },
    ]);
    const after = parts(bytes);
    for (const name of ['xl/drawings/vmlDrawing1.vml', 'xl/comments1.xml', 'xl/worksheets/_rels/sheet1.xml.rels', 'xl/calcChain.xml', '_rels/.rels', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml']) {
      expect(after[name], name).toBe(before[name]);
    }
    const sheet = after['xl/worksheets/sheet1.xml']!;
    expect(sheet).toContain('<row r="1" spans="1:2" x14ac:dyDescent="0.25">');
    expect(sheet).toContain('<conditionalFormatting sqref="B1:B3">');
    expect(sheet).toContain('<legacyDrawing r:id="rId1"/>');
    expect(sheet).toContain('<c r="B3"><f>B1*2</f><v>10</v></c>');
    // The existing string "third" is reused; "second" is appended to the shared string table.
    expect(after['xl/sharedStrings.xml']).toContain('uniqueCount="3"');
    const s = readWorkbook(bytes).getSheet('S');
    expect([s.cell('A1')?.value, s.cell('B1')?.value, s.cell('A2')?.value, s.cell('A3')?.value]).toEqual(['first', 7, 'second', 'third']);
    expect(warnings.map((w) => w.code)).toEqual(['XLV_W_FORMULAS_NOT_RECALCULATED']);
  });

  it('removes the calculation chain only when formulas change, and asks for recalculation', () => {
    const plain = parts(editCells(rich(), [{ sheet: 'S', cell: 'B1', value: 1 }]).bytes);
    expect(plain['xl/calcChain.xml']).toBeDefined();
    // Values that formulas depend on changed, so recalculation is requested even here.
    expect(plain['xl/workbook.xml']).toContain('<calcPr fullCalcOnLoad="1"/>');
    const changed = parts(editCells(rich(), [{ sheet: 'S', cell: 'B3', value: { formula: 'B1*3' } }]).bytes);
    expect(changed['xl/calcChain.xml']).toBeUndefined();
    expect(changed['[Content_Types].xml']).not.toContain('calcChain');
    expect(changed['xl/_rels/workbook.xml.rels']).not.toContain('calcChain');
    expect(changed['xl/workbook.xml']).toContain('<calcPr fullCalcOnLoad="1"/>');
    readWorkbook(editCells(rich(), [{ sheet: 'S', cell: 'B3', value: { formula: 'B1*3' } }]).bytes);
  });

  it('creates a shared string table when the package has none', () => {
    const noSst = buildXlsx({ sheets: [{ name: 'S', data: row(1, c('A1', 1)) }] });
    const { bytes } = editCells(noSst, [{ sheet: 'S', cell: 'A2', value: 'text' }]);
    expect(readWorkbook(bytes).getSheet('S').cell('A2')?.value).toBe('text');
    expect(parts(bytes)['[Content_Types].xml']).toContain('sharedStrings+xml');
  });

  it('adds explicit references when rows or cells have none', () => {
    const implicit = buildXlsx({ sheets: [{ name: 'S', data: '<row><c><v>1</v></c><c><v>2</v></c></row><row><c><v>3</v></c></row>' }] });
    const { bytes } = editCells(implicit, [{ sheet: 'S', cell: 'B1', value: 9 }, { sheet: 'S', cell: 'A3', value: 4 }]);
    const s = readWorkbook(bytes).getSheet('S');
    expect([...s.cells()].map((cell) => [cell.ref, cell.value])).toEqual([
      ['A1', 1],
      ['B1', 9],
      ['A2', 3],
      ['A3', 4],
    ]);
  });

  it('validates edits', () => {
    expectXlvError(() => editCells(base(), [{ sheet: 'Nope', cell: 'A1', value: 1 }]), 'XLV_E_SHEET_NOT_FOUND');
    expectXlvError(() => editCells(base(), [{ sheet: 'Data', cell: 'A0', value: 1 }]), 'XLV_E_INVALID_CELL_REF');
    const bad = expectXlvError(() => editCells(base(), [{ sheet: 'Data', cell: 'A1', value: Number.NaN }]), 'XLV_E_INVALID_VALUE');
    expect(bad.location).toMatchObject({ sheet: 'Data', cell: 'A1' });
    expectXlvError(() => editCells(base(), [{ sheet: 'Data', cell: 'A1', value: 1, style: { bold: true } as never }]), 'XLV_E_INVALID_STYLE');
    expectXlvError(() => editCells(base(), 'x' as never), 'XLV_E_INVALID_INPUT');
    expectXlvError(() => editCells(readWorkbook(base()) as never, []), 'XLV_E_INVALID_INPUT');
  });
});

describe('appendRows', () => {
  it('appends objects under the matching header columns', () => {
    const { bytes } = appendRows(base(), 'Data', [
      { name: 'Carol', total: 30 },
      { total: 40, name: 'Dan' },
    ]);
    const t = readTable(bytes, { sheet: 'Data', range: 'A1:B5' });
    expect(t.rows.map((r) => [r['name'], r['total']])).toEqual([
      ['Alice', 10],
      ['Bob', 20],
      ['Carol', 30],
      ['Dan', 40],
    ]);
    expect(t.rowNumbers).toEqual([2, 3, 4, 5]);
  });

  it('appends arrays starting at the first used column', () => {
    const { bytes } = appendRows(base(), 'Other', [['x', 1]]);
    const s = readWorkbook(bytes).getSheet('Other');
    expect([s.cell('A2')?.value, s.cell('B2')?.value]).toEqual(['x', 1]);
  });

  it('rejects unknown keys with a suggestion', () => {
    const error = expectXlvError(() => appendRows(base(), 'Data', [{ nmae: 'x' }]), 'XLV_E_TABLE_INVALID');
    expect(error.hint).toContain('name');
  });
});

describe('sheets with invalid references (D-903)', () => {
  it('are read with a warning but not edited', () => {
    const xlsx = buildXlsx({ sheets: [{ name: 'S', data: '<row r="1"><c r="A1"><v>1</v></c><c r="B"><v>2</v></c></row>' }] });
    const error = expectXlvError(() => editCells(xlsx, [{ sheet: 'S', cell: 'C1', value: 3 }]), 'XLV_E_INVALID_CELL_REF');
    expect(error.details).toMatchObject({ ref: 'B' });
    expectXlvError(() => appendRows(xlsx, 'S', [[1]]), 'XLV_E_INVALID_CELL_REF');
  });
});
