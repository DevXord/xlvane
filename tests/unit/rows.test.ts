// Surgical insertRows / deleteRows (docs/decisions.md D-B04; reference rules D-B01).
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { XlvError } from '../../packages/xlvane/src/errors';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { deleteRows, insertRows } from '../../packages/xlvane/src/rows';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { buildXlsx, c, row } from '../helpers/xlsx-builder';

const part = (bytes: Uint8Array, name: string): string => strFromU8(unzipSync(bytes)[name] ?? new Uint8Array());
const codeOf = (fn: () => unknown): string => {
  try {
    fn();
  } catch (error) {
    return error instanceof XlvError ? error.code : 'CRASH';
  }
  return 'none';
};

function sample(): Uint8Array {
  return buildXlsx({
    sheets: [
      {
        name: 'Main',
        before: '<dimension ref="A1:C10"/>',
        data:
          row(1, c('A1', 1), c('B1', 10)) +
          row(4, c('A4', 4)) +
          row(5, c('A5', 5)) +
          row(6, c('A6', 6)) +
          row(10, c('A10', 10), '<c r="C10"><f>SUM(A1:A6)+$A$10</f><v>0</v></c>'),
        after:
          '<autoFilter ref="A1:C10"/><mergeCells count="2"><mergeCell ref="A5:B6"/><mergeCell ref="B8:C9"/></mergeCells>' +
          '<conditionalFormatting sqref="A4:A6"><cfRule type="expression" priority="1"><formula>A4&gt;$A$10</formula></cfRule></conditionalFormatting>' +
          '<dataValidations count="1"><dataValidation type="list" sqref="B4:B10"><formula1>$A$1:$A$6</formula1></dataValidation></dataValidations>' +
          '<hyperlinks><hyperlink ref="A6" location="Main!A10"/></hyperlinks>' +
          '<rowBreaks count="1" manualBreakCount="1"><brk id="6" max="16383" man="1"/></rowBreaks>',
      },
      { name: 'Other', data: row(1, '<c r="A1"><f>Main!A5*2+SUM(Main!A1:A10)</f><v>0</v></c>', '<c r="B1"><f>A1+1</f><v>0</v></c>') },
    ],
    workbookAfterSheets: '<definedNames><definedName name="Data">Main!$A$4:$A$6</definedName><definedName name="Total">Main!$C$10</definedName></definedNames>',
  });
}

describe('insertRows', () => {
  it('moves rows, cells and every reference on all sheets', () => {
    const { bytes, warnings } = insertRows(sample(), { sheet: 'Main', at: 5, count: 2 });
    const main = part(bytes, 'xl/worksheets/sheet1.xml');
    expect(main).toContain('<dimension ref="A1:C12"/>');
    expect(main).toContain('<row r="4"><c r="A4" >');
    expect(main).toContain('<row r="7"><c r="A7" ><v>5</v>');
    expect(main).toContain('<row r="12"><c r="A12" >');
    expect(main).toContain('<c r="C12"><f>SUM(A1:A8)+$A$12</f>');
    expect(main).toContain('<autoFilter ref="A1:C12"/>');
    expect(main).toContain('<mergeCell ref="A7:B8"/><mergeCell ref="B10:C11"/>');
    expect(main).toContain('<conditionalFormatting sqref="A4:A8"><cfRule type="expression" priority="1"><formula>A4&gt;$A$12</formula>');
    expect(main).toContain('<dataValidation type="list" sqref="B4:B12"><formula1>$A$1:$A$8</formula1>');
    expect(main).toContain('<hyperlink ref="A8" location="Main!A12"/>');
    expect(main).toContain('<brk id="8" max="16383" man="1"/>');
    const other = part(bytes, 'xl/worksheets/sheet2.xml');
    expect(other).toContain('<f>Main!A7*2+SUM(Main!A1:A12)</f>');
    expect(other).toContain('<f>A1+1</f>');
    const wb = part(bytes, 'xl/workbook.xml');
    expect(wb).toContain('<definedName name="Data">Main!$A$4:$A$8</definedName>');
    expect(wb).toContain('<definedName name="Total">Main!$C$12</definedName>');
    expect(wb).toContain('fullCalcOnLoad="1"');
    expect(warnings.map((w) => w.code)).toEqual(['XLV_W_FORMULAS_NOT_RECALCULATED']);
    const read = readWorkbook(bytes);
    expect(read.getSheet('Main').cell('A7')?.value).toBe(5);
    expect(read.getSheet('Main').cell('A5')).toBeUndefined();
  });

  it('keeps every part that holds no row positions byte for byte', () => {
    const source = sample();
    const { bytes } = insertRows(source, { sheet: 'Main', at: 20, count: 1 });
    const before = unzipSync(source);
    const after = unzipSync(bytes);
    for (const name of ['[Content_Types].xml', '_rels/.rels', 'xl/_rels/workbook.xml.rels']) expect(after[name]).toEqual(before[name]);
  });

  it('copies the formatting of the row above or below', () => {
    const file = buildXlsx({
      sheets: [{ name: 'S', data: '<row r="1" s="2" customFormat="1" ht="30" customHeight="1"><c r="A1" s="3"><v>1</v></c><c r="B1"/></row><row r="2"><c r="A2" s="4"><v>2</v></c></row>' }],
    });
    const above = part(insertRows(file, { sheet: 'S', at: 2, count: 2, style: 'above' }).bytes, 'xl/worksheets/sheet1.xml');
    expect(above).toContain('<row r="2" s="2" customFormat="1" ht="30" customHeight="1"><c r="A2" s="3"/></row><row r="3" s="2" customFormat="1" ht="30" customHeight="1"><c r="A3" s="3"/></row><row r="4"><c r="A4" s="4"><v>2</v>');
    const below = part(insertRows(file, { sheet: 'S', at: 2, count: 1, style: 'below' }).bytes, 'xl/worksheets/sheet1.xml');
    expect(below).toContain('<row r="2"><c r="A2" s="4"/></row><row r="3"><c r="A3" s="4"><v>2</v>');
  });

  it('expands shared formulas of the changed sheet into plain formulas', () => {
    const file = buildXlsx({
      sheets: [
        {
          name: 'S',
          data:
            row(1, c('A1', 1), '<c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c>') +
            row(2, c('A2', 2), '<c r="B2"><f t="shared" si="0"/><v>4</v></c>') +
            row(3, c('A3', 3), '<c r="B3"><f t="shared" si="0"/><v>6</v></c>'),
        },
      ],
    });
    const out = insertRows(file, { sheet: 'S', at: 2, count: 1 }).bytes;
    const s = readWorkbook(out).getSheet('S');
    expect([s.cell('B1')?.formula?.text, s.cell('B3')?.formula?.text, s.cell('B4')?.formula?.text]).toEqual(['A1*2', 'A3*2', 'A4*2']);
    expect(part(out, 'xl/worksheets/sheet1.xml')).not.toContain('t="shared"');
  });

  it('moves notes in the Comments part and their VML shapes', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('Notes');
    s.set('A1', 'x').set('A3', 'y').note('A3', 'hello').note('A1', 'top');
    const out = insertRows(writeWorkbook(wb), { sheet: 'Notes', at: 2, count: 2 }).bytes;
    const read = readWorkbook(out).getSheet('Notes');
    expect(read.notes.map((n) => [n.ref, n.text])).toEqual([
      ['A1', 'top'],
      ['A5', 'hello'],
    ]);
    const vml = part(out, 'xl/drawings/vmlDrawing1.vml');
    expect(vml).toContain('<x:Row>4</x:Row>');
    expect(vml).toContain('<x:Row>0</x:Row>');
  });

  it('refuses tables, drawings and formulas it cannot adjust', () => {
    const drawing = buildXlsx({ sheets: [{ name: 'S', data: row(1, c('A1', 1)), after: '<drawing r:id="rId1"/>' }] });
    expect(codeOf(() => insertRows(drawing, { sheet: 'S', at: 1 }))).toBe('XLV_E_NOT_SUPPORTED');
    const spill = buildXlsx({ sheets: [{ name: 'S', data: row(1, '<c r="A1"><f>@B1:B3</f></c>') }] });
    expect(codeOf(() => insertRows(spill, { sheet: 'S', at: 1 }))).toBe('XLV_E_NOT_SUPPORTED');
    const full = buildXlsx({ sheets: [{ name: 'S', data: row(1_048_576, c('A1048576', 1)) }] });
    expect(codeOf(() => insertRows(full, { sheet: 'S', at: 1 }))).toBe('XLV_E_LIMIT_EXCEEDED');
  });

  it('validates its options', () => {
    const file = sample();
    expect(codeOf(() => insertRows(file, { sheet: 'Nope', at: 1 }))).toBe('XLV_E_SHEET_NOT_FOUND');
    expect(codeOf(() => insertRows(file, { sheet: 'Main', at: 0 }))).toBe('XLV_E_INVALID_OPTION');
    expect(codeOf(() => insertRows(file, { sheet: 'Main', at: 1, count: 0 }))).toBe('XLV_E_INVALID_OPTION');
    expect(codeOf(() => insertRows(file, { sheet: 'Main', at: 1, style: 'x' as 'above' }))).toBe('XLV_E_INVALID_OPTION');
    expect(codeOf(() => insertRows(file, { sheet: 'Main', at: 1, extra: 1 } as never))).toBe('XLV_E_INVALID_OPTION');
  });
});

describe('deleteRows', () => {
  it('removes rows, shrinks ranges and turns deleted references into #REF!', () => {
    const { bytes } = deleteRows(sample(), { sheet: 'Main', at: 5, count: 2 });
    const main = part(bytes, 'xl/worksheets/sheet1.xml');
    expect(main).toContain('<dimension ref="A1:C8"/>');
    expect(main).not.toContain('<v>5</v>');
    expect(main).toContain('<row r="8"><c r="A8" ><v>10</v></c><c r="C8"><f>SUM(A1:A4)+$A$8</f>');
    // R-14: the merge whose top-left cell is deleted is dropped.
    expect(main).toContain('<mergeCells count="1"><mergeCell ref="B6:C7"/></mergeCells>');
    expect(main).toContain('<conditionalFormatting sqref="A4:A4">');
    expect(main).not.toContain('<hyperlinks');
    expect(main).not.toContain('<rowBreaks');
    expect(part(bytes, 'xl/worksheets/sheet2.xml')).toContain('<f>#REF!*2+SUM(Main!A1:A8)</f>');
    expect(part(bytes, 'xl/workbook.xml')).toContain('<definedName name="Data">Main!$A$4:$A$4</definedName>');
  });

  it('removes notes of deleted rows', () => {
    const wb = createWorkbook();
    wb.addSheet('N').set('A1', 1).set('A2', 2).note('A2', 'gone').note('A1', 'kept');
    const out = deleteRows(writeWorkbook(wb), { sheet: 'N', at: 2 }).bytes;
    expect(readWorkbook(out).getSheet('N').notes.map((n) => n.ref)).toEqual(['A1']);
    expect(part(out, 'xl/drawings/vmlDrawing1.vml').match(/<v:shape /g)?.length).toBe(1);
  });

  it('insert then delete of the same rows restores the cells', () => {
    const source = sample();
    const back = deleteRows(insertRows(source, { sheet: 'Main', at: 3, count: 4 }).bytes, { sheet: 'Main', at: 3, count: 4 }).bytes;
    const a = readWorkbook(source);
    const b = readWorkbook(back);
    for (const name of ['Main', 'Other']) {
      expect([...b.getSheet(name).cells()].map((x) => [x.ref, x.value, x.formula?.text ?? null])).toEqual([...a.getSheet(name).cells()].map((x) => [x.ref, x.value, x.formula?.text ?? null]));
    }
    expect(b.definedNames).toEqual(a.definedNames);
  });
});
