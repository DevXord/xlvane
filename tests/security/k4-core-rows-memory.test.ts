// insertRows and deleteRows after the K3 review: like editCells before A2, they kept an object per
// row, cell and formula of every worksheet (the element scanner of edit.ts) and an edit per moved
// cell, so a 4.4 MB file with 4.9 million number cells ran out of memory ("xlvane edit
// --insert-rows 1", exit 134). The worksheets are now scanned into typed arrays and the new text is
// written in pieces as it is produced. Inserted rows that copy the cell formats of a row
// (`style: 'above'` or `'below'`) are counted against `limits.maxCells` before they are built.
import { strToU8, zipSync } from 'fflate';
import { describe, expect, it, vi } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { deleteRows, insertRows } from '../../packages/xlvane/src/rows';
import { formatCellRef } from '../../packages/xlvane/src/sml/cell-ref';
import { expectXlvError } from '../helpers/expect-error';

// Counts the elements the element scanner of edit.ts keeps, for whatever part it scans.
const kept = vi.hoisted(() => ({ elements: 0 }));
vi.mock('../../packages/xlvane/src/edit', async (importOriginal) => {
  const mod = await importOriginal<typeof import('../../packages/xlvane/src/edit')>();
  return {
    ...mod,
    scan: (...args: Parameters<typeof mod.scan>) => {
      const result = mod.scan(...args);
      kept.elements += result.elements.length;
      return result;
    },
  };
});

const M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const STYLES = '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs>';

function workbook(sheets: { name: string; xml: string }[]): Uint8Array {
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${String(i + 1)}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="${M}" xmlns:r="${R}"><sheets>${sheets.map((s, i) => `<sheet name="${s.name}" sheetId="${String(i + 1)}" r:id="rId${String(i + 1)}"/>`).join('')}</sheets></workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${String(i + 1)}" Type="${R}/worksheet" Target="worksheets/sheet${String(i + 1)}.xml"/>`).join('')}<Relationship Id="rIdS" Type="${R}/styles" Target="styles.xml"/></Relationships>`),
    'xl/styles.xml': strToU8(`<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="${M}">${STYLES}</styleSheet>`),
  };
  sheets.forEach((s, i) => (files[`xl/worksheets/sheet${String(i + 1)}.xml`] = strToU8(`<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${M}">${s.xml}</worksheet>`)));
  return zipSync(files);
}

/** `rows` × `cols` numbers, with or without cell references. */
function grid(rows: number, cols: number, refs: boolean): string {
  const parts: string[] = ['<sheetData>'];
  for (let r = 1; r <= rows; r++) {
    parts.push(refs ? `<row r="${String(r)}">` : '<row>');
    for (let c = 1; c <= cols; c++) parts.push(refs ? `<c r="${formatCellRef(r, c)}"><v>${String(r * 100 + c)}</v></c>` : `<c><v>${String(r * 100 + c)}</v></c>`);
    parts.push('</row>');
  }
  parts.push('</sheetData>');
  return parts.join('');
}

/** A sheet of `rows` formulas that refer to sheet Data, one per row. */
function formulas(rows: number): string {
  return `<sheetData>${Array.from({ length: rows }, (_, i) => `<row r="${String(i + 1)}"><c r="A${String(i + 1)}"><f>Data!B${String(i + 1)}*2</f><v>0</v></c></row>`).join('')}</sheetData>`;
}

function values(bytes: Uint8Array, sheet: string): unknown[][] {
  return [...readWorkbook(bytes).getSheet(sheet).rows()].map((r) => r.map((c) => c.formula?.text ?? c.value));
}

describe('insertRows and deleteRows keep no object per cell', () => {
  for (const refs of [true, false]) {
    it(`scans sheets of cells ${refs ? 'with' : 'without'} references without an element per cell`, () => {
      const bytes = workbook([
        { name: 'Data', xml: grid(200, 50, refs) },
        { name: 'Calc', xml: formulas(2000) },
      ]);
      kept.elements = 0;
      const inserted = insertRows(bytes, { sheet: 'Data', at: 3, count: 2 });
      // 10 000 cells and 2 000 formulas; only the elements around them may be kept.
      expect(kept.elements).toBeLessThan(500);
      kept.elements = 0;
      const deleted = deleteRows(bytes, { sheet: 'Data', at: 3, count: 2 });
      expect(kept.elements).toBeLessThan(500);

      const before = values(bytes, 'Data');
      expect(values(inserted.bytes, 'Data')).toEqual(before);
      const sheet = readWorkbook(inserted.bytes).getSheet('Data');
      expect([sheet.cell('A2')?.value, sheet.cell('A3')?.value, sheet.cell('A5')?.value, sheet.cell('AX202')?.value]).toEqual([201, undefined, 301, 20050]);
      expect(readWorkbook(inserted.bytes).getSheet('Calc').cell('A3')?.formula?.text).toBe('Data!B5*2');
      expect(readWorkbook(deleted.bytes).getSheet('Data').cell('A3')?.value).toBe(501);
      expect(readWorkbook(deleted.bytes).getSheet('Calc').cell('A3')?.formula?.text).toBe('#REF!*2');
      expect(readWorkbook(deleted.bytes).getSheet('Calc').cell('A5')?.formula?.text).toBe('Data!B3*2');
    });
  }

  it('moves shared and array formulas of the changed sheet as before', () => {
    const xml = `<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row><row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c><c r="C3"><f t="array" ref="C3:C4">A3:A4</f><v>3</v></c></row></sheetData>`;
    const out = insertRows(workbook([{ name: 'Data', xml }]), { sheet: 'Data', at: 2 });
    const sheet = readWorkbook(out.bytes).getSheet('Data');
    expect(['B1', 'B3', 'B4', 'C4'].map((ref) => sheet.cell(ref)?.formula?.text)).toEqual(['A1*2', 'A3*2', 'A4*2', 'A4:A5']);
    expect(sheet.cell('C4')?.formula?.ref).toBe('C4:C5');
  });
});

describe('inserted rows that copy cell formats count against maxCells', () => {
  const styled = `<sheetData><row r="1">${Array.from({ length: 10 }, (_, i) => `<c r="${formatCellRef(1, i + 1)}" s="1"><v>1</v></c>`).join('')}</row></sheetData>`;

  it('refuses before building the rows', () => {
    const bytes = workbook([{ name: 'Data', xml: styled }]);
    expectXlvError(() => insertRows(bytes, { sheet: 'Data', at: 2, count: 100, style: 'above', limits: { maxCells: 500 } }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxCells', value: 500 });
    const { bytes: out } = insertRows(bytes, { sheet: 'Data', at: 2, count: 100, style: 'above', limits: { maxCells: 1010 } });
    // The result reads with the same limit.
    expect([...readWorkbook(out, { limits: { maxCells: 1010 } }).getSheet('Data').styledCells()].length).toBe(1000);
  });

  it('refuses millions of formatted cells with the default limits', () => {
    const bytes = workbook([{ name: 'Data', xml: styled }]);
    expectXlvError(() => insertRows(bytes, { sheet: 'Data', at: 2, count: 1_000_000, style: 'above' }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxCells' });
  });
});
