// Rich text: runs with their own fonts in shared strings and inline strings (ECMA-376 Part 1
// §18.4; facts F-SST-08, F-SST-09 in docs/facts/sst-styles.md; decision D-A01).
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { editCells } from '../../packages/xlvane/src/edit';
import { DEFAULT_READ_LIMITS } from '../../packages/xlvane/src/limits';
import type { RichTextRun } from '../../packages/xlvane/src/model';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { SML_NS } from '../../packages/xlvane/src/sml/namespaces';
import { parseSharedStrings } from '../../packages/xlvane/src/sml/shared-strings';
import { openWorkbookStream } from '../../packages/xlvane/src/stream-read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { WarningCollector } from '../../packages/xlvane/src/util/warnings';
import { createWorkbook, writeWorkbook, type CellInput } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';
import { expectPrototypesUntouched } from '../helpers/prototypes';
import { buildXlsx } from '../helpers/xlsx-builder';

const T = SML_NS.transitional;
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

async function collect(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** A workbook with one sheet "S" and `value` in A1. */
function oneCell(value: CellInput): Uint8Array {
  const workbook = createWorkbook();
  workbook.addSheet('S').set('A1', value);
  return writeWorkbook(workbook);
}

function part(bytes: Uint8Array, name: string): string {
  const file = unzipSync(bytes)[name];
  if (!file) throw new Error(`no part ${name}`);
  return strFromU8(file);
}

const RUNS: RichTextRun[] = [
  { text: 'Total: ', font: { bold: true, color: 'FFFF0000', size: 12, name: 'Arial' } },
  { text: '1,250', font: { italic: true, underline: 'double', strike: true } },
  { text: ' EUR' },
];

describe('reading rich text', () => {
  it('keeps the plain value and returns the runs of shared strings with their fonts', () => {
    const sst = `${DECL}<sst xmlns="${T}"><si><t>plain</t></si><si><r><rPr><b/><sz val="11"/><color rgb="FF0000FF"/><rFont val="Calibri"/><family val="2"/></rPr><t xml:space="preserve">Bold </t></r><r><t>normal</t></r><rPh sb="0" eb="1"><t>ph</t></rPh></si><si><t>lead</t><r><rPr><i val="0"/><u val="double"/></rPr><t>x</t></r></si></sst>`;
    const wb = readWorkbook(
      buildXlsx({ sharedStrings: sst, sheets: [{ name: 'S', data: '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>' }] }),
    );
    const s = wb.getSheet('S');
    expect(s.cell('A1')?.value).toBe('plain');
    expect(s.cell('A1')?.richText).toBeUndefined();
    expect(s.cell('B1')?.value).toBe('Bold normal');
    expect(s.cell('B1')?.richText).toEqual([{ text: 'Bold ', font: { bold: true, size: 11, color: 'FF0000FF', name: 'Calibri', family: 2 } }, { text: 'normal' }]);
    // A leading t is a run without properties; i val="0" is not italic.
    expect(s.cell('C1')?.richText).toEqual([{ text: 'lead' }, { text: 'x', font: { underline: 'double' } }]);
    expect(Object.isFrozen(s.cell('B1')?.richText)).toBe(true);
  });

  it('reads inline rich text, including in Strict files', () => {
    for (const strict of [false, true]) {
      const wb = readWorkbook(
        buildXlsx({
          strict,
          sheets: [
            {
              name: 'S',
              data: '<row r="1"><c r="A1" t="inlineStr"><is><r><rPr><strike/><rFont val="Arial"/></rPr><t>a</t></r><r><t>b</t></r></is></c><c r="B1" t="inlineStr"><is><t>only</t></is></c></row>',
            },
          ],
        }),
      );
      const s = wb.getSheet('S');
      expect(s.cell('A1')?.value).toBe('ab');
      expect(s.cell('A1')?.richText).toEqual([{ text: 'a', font: { strike: true, name: 'Arial' } }, { text: 'b' }]);
      expect(s.cell('B1')?.richText).toBeUndefined();
    }
  });

  it('streaming gives the same runs as the batch reader', async () => {
    const bytes = oneCell({ richText: RUNS });
    const stream = await openWorkbookStream(bytes);
    const cells = [];
    for await (const row of stream.rows('S')) cells.push(...row.cells);
    expect(cells[0]?.richText).toEqual(readWorkbook(bytes).getSheet('S').cell('A1')?.richText);
  });

  it('reports invalid run properties as warnings and leaves them out', () => {
    const warnings = new WarningCollector();
    const rich = new Map<number, readonly RichTextRun[]>();
    const xml = `<sst xmlns="${T}"><si><r><rPr><sz val="big"/><u val="wavy"/><b/></rPr><t>x</t></r></si></sst>`;
    expect(parseSharedStrings(xml, { part: '/xl/sharedStrings.xml', limits: DEFAULT_READ_LIMITS, warnings }, rich)).toEqual(['x']);
    expect(rich.get(0)).toEqual([{ text: 'x', font: { bold: true } }]);
    expect(new Set(warnings.list().map((w) => w.code))).toEqual(new Set(['XLV_W_INVALID_VALUE']));
  });

  it('does not pollute prototypes through run properties', () => {
    {
      const rich = new Map<number, readonly RichTextRun[]>();
      const xml = `<sst xmlns="${T}"><si><r><rPr><__proto__ val="1"/><constructor/><rFont val="__proto__"/></rPr><t>__proto__</t></r></si></sst>`;
      parseSharedStrings(xml, { part: '/xl/sharedStrings.xml', limits: DEFAULT_READ_LIMITS, warnings: new WarningCollector() }, rich);
      expect(rich.get(0)).toEqual([{ text: '__proto__', font: { name: '__proto__' } }]);
    }
    expectXlvError(() => oneCell(JSON.parse('{"richText":[{"text":"x","font":{"__proto__":{"bold":true}}}]}') as CellInput), 'XLV_E_INVALID_STYLE');
    expectPrototypesUntouched();
  });
});

describe('writing rich text', () => {
  it('writes runs into the shared string table with rPr in the order Office requires', () => {
    const bytes = oneCell({ richText: RUNS });
    const sst = part(bytes, 'xl/sharedStrings.xml');
    expect(sst).toContain(
      '<si><r><rPr><b/><sz val="12"/><color rgb="FFFF0000"/><rFont val="Arial"/></rPr><t xml:space="preserve">Total: </t></r>' +
        '<r><rPr><i/><strike/><u val="double"/></rPr><t>1,250</t></r><r><t xml:space="preserve"> EUR</t></r></si>',
    );
    const cell = readWorkbook(bytes).getSheet('S').cell('A1');
    expect(cell?.value).toBe('Total: 1,250 EUR');
    expect(cell?.richText).toEqual(RUNS);
  });

  it('shares equal items and keeps plain and rich text with the same characters apart', () => {
    const workbook = createWorkbook();
    const sheet = workbook.addSheet('S');
    sheet.set('A1', { richText: [{ text: 'x', font: { bold: true } }] });
    sheet.set('A2', { richText: [{ text: 'x', font: { bold: true } }] });
    sheet.set('A3', 'x');
    sheet.set('A4', { richText: [{ text: 'x' }] });
    const bytes = writeWorkbook(workbook);
    const sst = part(bytes, 'xl/sharedStrings.xml');
    expect(sst).toContain('count="4" uniqueCount="3"');
    const s = readWorkbook(bytes).getSheet('S');
    expect(s.cell('A2')?.richText).toEqual([{ text: 'x', font: { bold: true } }]);
    expect(s.cell('A3')?.richText).toBeUndefined();
    expect(s.cell('A4')?.richText).toEqual([{ text: 'x' }]);
  });

  it('writes inline rich text when the streaming writer has no shared string table', async () => {
    const bytes = await collect(writeWorkbookStream({ sheets: [{ name: 'S', rows: [[{ richText: RUNS }, 'plain']] }] }));
    expect(part(bytes, 'xl/worksheets/sheet1.xml')).toContain('t="inlineStr"><is><r><rPr><b/>');
    expect(readWorkbook(bytes).getSheet('S').cell('A1')?.richText).toEqual(RUNS);
    const shared = await collect(writeWorkbookStream({ sharedStrings: true, sheets: [{ name: 'S', rows: [[{ richText: RUNS }]] }] }));
    expect(part(shared, 'xl/sharedStrings.xml')).toContain('<rFont val="Arial"/>');
    expect(readWorkbook(shared).getSheet('S').cell('A1')?.richText).toEqual(RUNS);
  });

  it('editCells adds rich text and does not reuse a rich item of the file for plain text', () => {
    const sst = `${DECL}<sst xmlns="${T}" count="1" uniqueCount="1"><si><r><rPr><b/></rPr><t>x</t></r></si></sst>`;
    const source = buildXlsx({ sharedStrings: sst, sheets: [{ name: 'S', data: '<row r="1"><c r="A1" t="s"><v>0</v></c></row>' }] });
    const edited = editCells(source, [
      { sheet: 'S', cell: 'B1', value: 'x' },
      { sheet: 'S', cell: 'C1', value: { richText: RUNS } },
    ]).bytes;
    const s = readWorkbook(edited).getSheet('S');
    expect(s.cell('A1')?.richText).toEqual([{ text: 'x', font: { bold: true } }]);
    expect(s.cell('B1')?.value).toBe('x');
    expect(s.cell('B1')?.richText).toBeUndefined();
    expect(s.cell('C1')?.richText).toEqual(RUNS);
  });

  it('rejects invalid rich text with XLV_E_INVALID_VALUE or XLV_E_INVALID_STYLE', () => {
    const set = (value: unknown) => () => createWorkbook().addSheet('S').set('A1', value as CellInput);
    expectXlvError(set({ richText: [] }), 'XLV_E_INVALID_VALUE');
    expectXlvError(set({ richText: 'x' }), 'XLV_E_INVALID_VALUE');
    expectXlvError(set({ richText: [{ text: 'x' }], extra: 1 }), 'XLV_E_INVALID_VALUE');
    expectXlvError(set({ richText: [{ text: 1 }] }), 'XLV_E_INVALID_VALUE');
    expectXlvError(set({ richText: [{ text: 'x', size: 3 }] }), 'XLV_E_INVALID_VALUE');
    expectXlvError(set({ richText: [{ text: '\ud800' }] }), 'XLV_E_INVALID_VALUE');
    expectXlvError(set({ richText: [{ text: 'x', font: { size: 500 } }] }), 'XLV_E_INVALID_STYLE', { property: 'font.size', run: 0 });
    expectXlvError(set({ richText: Array.from({ length: 32_768 }, () => ({ text: '' })) }), 'XLV_E_INVALID_VALUE');
    expectXlvError(set({ richText: [{ text: 'a'.repeat(20_000) }, { text: 'b'.repeat(20_000) }] }), 'XLV_E_INVALID_VALUE');
    // Rich text is a cell value, not a formula result.
    expectXlvError(set({ formula: 'A2', result: { richText: [{ text: 'x' }] } }), 'XLV_E_INVALID_VALUE');
  });

  it('round-trips the limits: 32,767 characters spread over runs', () => {
    const runs = [{ text: 'a'.repeat(32_000) }, { text: 'b'.repeat(767), font: { bold: true } }];
    const bytes = oneCell({ richText: runs });
    expect((readWorkbook(bytes).getSheet('S').cell('A1')?.value as string).length).toBe(32_767);
  });
});
