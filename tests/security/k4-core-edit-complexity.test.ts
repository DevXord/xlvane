// editCells and appendRows after the K3 review (A3): every new cell was compared with every merged
// range of its sheet (the merged-cell warning), every edit with every array formula of the sheet,
// and every new cell with every column definition, so 200 000 merged ranges and 20 000 edits took
// 44 s. They are now looked up in a range index (util/range-index.ts) and a column table. Counted
// instead of timed: the number of callbacks of Array.prototype.find and Array.prototype.some per
// edit must not grow with the number of ranges in the file.
import { strToU8, zipSync } from 'fflate';
import { describe, expect, it, vi } from 'vitest';
import { appendRows, editCells, type CellEdit } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';

const M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function sheetFile(sheetXml: string, styles = ''): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>${styles === '' ? '' : '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'}</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="${M}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/>${styles === '' ? '' : `<Relationship Id="rId2" Type="${R}/styles" Target="styles.xml"/>`}</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${M}">${sheetXml}</worksheet>`),
    ...(styles === '' ? {} : { 'xl/styles.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="${M}">${styles}</styleSheet>`) }),
  });
}

/** Callbacks run by Array.prototype.find and Array.prototype.some while `fn` runs. */
function callbacks(fn: () => unknown): number {
  let count = 0;
  const find = Array.prototype.find;
  const some = Array.prototype.some;
  const findSpy = vi.spyOn(Array.prototype, 'find').mockImplementation(function (this: unknown[], predicate: (value: unknown, index: number, array: unknown[]) => unknown, thisArg?: unknown) {
    return find.call(this, (v: unknown, i: number, a: unknown[]) => (count++, predicate.call(thisArg, v, i, a)));
  });
  const someSpy = vi.spyOn(Array.prototype, 'some').mockImplementation(function (this: unknown[], predicate: (value: unknown, index: number, array: unknown[]) => unknown, thisArg?: unknown) {
    return some.call(this, (v: unknown, i: number, a: unknown[]) => (count++, predicate.call(thisArg, v, i, a)));
  });
  try {
    fn();
  } finally {
    findSpy.mockRestore();
    someSpy.mockRestore();
  }
  return count;
}

const EDITS = 2000;
const edits = (n: number): CellEdit[] => Array.from({ length: n }, (_, i) => ({ sheet: 'S', cell: `A${String(i + 1)}`, value: i }));

describe('A3: editCells cost per edit does not grow with the ranges of the file', () => {
  it('merged ranges (merged-cell warning)', () => {
    const n = 20_000;
    const bytes = sheetFile(`<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><mergeCells count="${String(n)}">${Array.from({ length: n }, (_, k) => `<mergeCell ref="C${String(k + 1)}:D${String(k + 1)}"/>`).join('')}</mergeCells>`);
    const one = callbacks(() => editCells(bytes, edits(1)));
    const many = callbacks(() => editCells(bytes, edits(EDITS)));
    expect(many - one).toBeLessThan(EDITS * 50);
  }, 60_000);

  it('still warns about a value written inside a merged range', () => {
    const bytes = sheetFile(`<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><mergeCells count="2"><mergeCell ref="C1:D1"/><mergeCell ref="B5:F9"/></mergeCells>`);
    const { warnings } = editCells(bytes, [
      { sheet: 'S', cell: 'C1', value: 1 },
      { sheet: 'S', cell: 'E7', value: 2 },
      { sheet: 'S', cell: 'G7', value: 3 },
    ]);
    expect(warnings.filter((w) => w.code === 'XLV_W_MERGED_CELL_WRITE').map((w) => w.cells)).toEqual([['E7']]);
  });

  it('array formulas (cells of an array formula cannot be edited)', () => {
    const n = 20_000;
    const bytes = sheetFile(`<sheetData>${Array.from({ length: n }, (_, k) => `<row r="${String(k + 1)}"><c r="C${String(k + 1)}"><f t="array" ref="C${String(k + 1)}">1</f><v>1</v></c></row>`).join('')}</sheetData>`);
    const one = callbacks(() => editCells(bytes, edits(1)));
    const many = callbacks(() => editCells(bytes, edits(EDITS)));
    expect(many - one).toBeLessThan(EDITS * 50);
  }, 60_000);

  it('still refuses a cell inside an array formula', () => {
    const bytes = sheetFile(`<sheetData><row r="2"><c r="B2"><f t="array" ref="B2:C4">1</f><v>1</v></c></row></sheetData>`);
    expect(() => editCells(bytes, [{ sheet: 'S', cell: 'C3', value: 1 }])).toThrow(/array formula B2:C4/);
    expect(() => editCells(bytes, [{ sheet: 'S', cell: 'D3', value: 1 }])).not.toThrow();
  });

  it('column definitions (column style of new cells)', () => {
    const n = 16_000;
    const styles = '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>';
    const bytes = sheetFile(`<cols>${Array.from({ length: n }, (_, k) => `<col min="${String(k + 1)}" max="${String(k + 1)}" style="1"/>`).join('')}</cols><sheetData/>`, styles);
    const at = (count: number): CellEdit[] => Array.from({ length: count }, (_, i) => ({ sheet: 'S', cell: `XFD${String(i + 1)}`, value: i }));
    const one = callbacks(() => editCells(bytes, at(1)));
    const many = callbacks(() => editCells(bytes, at(EDITS)));
    expect(many - one).toBeLessThan(EDITS * 50);
  }, 60_000);

  it('still gives new cells the style of their column, the first col element that covers it', () => {
    const styles = '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/><xf numFmtId="10" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs>';
    const bytes = sheetFile('<cols><col min="2" max="3" style="1"/><col min="3" max="5" style="2"/></cols><sheetData/>', styles);
    const { bytes: out } = editCells(bytes, ['A1', 'B1', 'C1', 'E1', 'F1'].map((cell) => ({ sheet: 'S', cell, value: 1 })));
    const sheet = readWorkbook(out).getSheet('S');
    expect(['A1', 'B1', 'C1', 'E1', 'F1'].map((ref) => sheet.cell(ref)?.styleIndex)).toEqual([0, 1, 1, 2, 0]);
  });

  it('appendRows below many merged ranges', () => {
    const n = 20_000;
    const bytes = sheetFile(`<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><mergeCells count="${String(n)}">${Array.from({ length: n }, (_, k) => `<mergeCell ref="C${String(k + 1)}:D${String(k + 1)}"/>`).join('')}</mergeCells>`);
    const rows = (count: number) => Array.from({ length: count }, (_, i) => [i]);
    const one = callbacks(() => appendRows(bytes, 'S', rows(1)));
    const many = callbacks(() => appendRows(bytes, 'S', rows(EDITS)));
    expect(many - one).toBeLessThan(EDITS * 50);
  }, 60_000);
});
