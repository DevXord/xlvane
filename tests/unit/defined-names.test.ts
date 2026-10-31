// Defined names (ECMA-376 Part 1 §18.2.5, §18.17.2.5; [MS-OI29500] §2.1.568; docs/facts
// sml-write.md W2-23 … W2-25; docs/decisions.md D-A03).
import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { editCells } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { openWorkbookStream } from '../../packages/xlvane/src/stream-read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError, expectXlvErrorAsync } from '../helpers/expect-error';
import { buildXlsx } from '../helpers/xlsx-builder';

const dec = new TextDecoder();

async function bytesOf(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

describe('reading defined names', () => {
  it('reads names with scope, hidden flag and comment, and decodes ST_Xstring escapes', () => {
    const bytes = buildXlsx({
      sheets: [
        { name: 'Data', data: '' },
        { name: 'Other', data: '' },
      ],
      workbookAfterSheets:
        '<definedNames>' +
        '<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$B$3</definedName>' +
        '<definedName name="Rate" comment="VAT_x000A_rate">Data!$B$1</definedName>' +
        '<definedName name="Local" localSheetId="1">Other!$A$1&amp;"x"</definedName>' +
        '<definedName name="Bad" localSheetId="7">1</definedName>' +
        '<definedName name="Odd" localSheetId="x" hidden="maybe">2</definedName>' +
        '</definedNames>',
    });
    const wb = readWorkbook(bytes);
    expect(wb.definedNames).toEqual([
      { name: '_xlnm._FilterDatabase', refersTo: 'Data!$A$1:$B$3', sheet: 'Data', hidden: true, comment: null },
      { name: 'Rate', refersTo: 'Data!$B$1', sheet: null, hidden: false, comment: 'VAT\nrate' },
      { name: 'Local', refersTo: 'Other!$A$1&"x"', sheet: 'Other', hidden: false, comment: null },
      { name: 'Bad', refersTo: '1', sheet: null, hidden: false, comment: null },
      { name: 'Odd', refersTo: '2', sheet: null, hidden: false, comment: null },
    ]);
    // Both problems are in the workbook part, so they are reported as one warning with a count.
    expect(wb.warnings.map((w) => [w.code, w.count])).toEqual([['XLV_W_INVALID_VALUE', 2]]);
  });

  it('gives an empty list without definedNames', () => {
    expect(readWorkbook(buildXlsx()).definedNames).toEqual([]);
  });

  it('keeps names that look like Object prototype keys as plain data', () => {
    const wb = readWorkbook(buildXlsx({ workbookAfterSheets: '<definedNames><definedName name="__proto__">1</definedName><definedName name="constructor">2</definedName></definedNames>' }));
    expect(wb.definedNames.map((n) => n.name)).toEqual(['__proto__', 'constructor']);
    expect(({} as Record<string, unknown>)['refersTo']).toBeUndefined();
  });
});

describe('writing defined names', () => {
  it('writes names after the autofilter names and reads them back', () => {
    const wb = createWorkbook();
    const data = wb.addSheet('Data');
    data.set('A1', 'x').autoFilter('A1:A3');
    wb.addSheet("Bob's");
    wb.defineName('Rate', '=Data!$B$1', { comment: 'VAT' });
    wb.defineName('Local', "'Bob''s'!$A$1", { sheet: "Bob's", hidden: true });
    wb.defineName('Ünïcode_名前', '0.19');
    const bytes = writeWorkbook(wb);
    const xml = dec.decode(unzipSync(bytes)['xl/workbook.xml']);
    expect(xml).toContain(
      '<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\'Data\'!$A$1:$A$3</definedName>' +
        '<definedName name="Rate" comment="VAT">Data!$B$1</definedName>' +
        `<definedName name="Local" localSheetId="1" hidden="1">'Bob''s'!$A$1</definedName>`,
    );
    expect(readWorkbook(bytes).definedNames.slice(1)).toEqual([
      { name: 'Rate', refersTo: 'Data!$B$1', sheet: null, hidden: false, comment: 'VAT' },
      { name: 'Local', refersTo: "'Bob''s'!$A$1", sheet: "Bob's", hidden: true, comment: null },
      { name: 'Ünïcode_名前', refersTo: '0.19', sheet: null, hidden: false, comment: null },
    ]);
    expect(wb.definedNames.map((n) => n.name)).toEqual(['Rate', 'Local', 'Ünïcode_名前']);
  });

  it('accepts the documented name characters', () => {
    const wb = createWorkbook();
    wb.addSheet('S');
    for (const name of ['a', '_x', '\\ab', 'A.b?c', 'Name1', 'ABCD1', 'R1C1x', 'TRUEx', '__proto__']) expect(() => wb.defineName(name, '1')).not.toThrow();
  });

  it('rejects invalid names with XLV_E_INVALID_DEFINED_NAME', () => {
    const wb = createWorkbook();
    wb.addSheet('S');
    for (const name of ['', 'a'.repeat(256), '1abc', 'a b', 'a-b', 'a!b', '$A', '.a', '?a', 'A1', 'xfd1048576', 'XFE1', 'R', 'c', 'R1C1', 'r2', 'C3R4', 'TRUE', 'false', '_xlnm.Print_Area', '_XLfoo', '\\a', 'a\uD800']) {
      expectXlvError(() => wb.defineName(name, '1'), 'XLV_E_INVALID_DEFINED_NAME');
    }
    expectXlvError(() => wb.defineName(1 as never, '1'), 'XLV_E_INVALID_DEFINED_NAME');
    for (const refersTo of ['', '=', '  ', 'x'.repeat(8193), 3]) expectXlvError(() => wb.defineName('Ok', refersTo as never), 'XLV_E_INVALID_DEFINED_NAME');
    expectXlvError(() => wb.defineName('Ok', '1', { comment: 'c'.repeat(256) }), 'XLV_E_INVALID_DEFINED_NAME');
    expectXlvError(() => wb.defineName('Ok', '1', { hiden: true } as never), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => wb.defineName('Ok', '1', { hidden: 'yes' } as never), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => wb.defineName('Ok', '1', { sheet: 'Nope' }), 'XLV_E_SHEET_NOT_FOUND');
  });

  it('requires unique names per scope, ignoring case', () => {
    const wb = createWorkbook();
    wb.addSheet('A');
    wb.addSheet('B');
    wb.defineName('Total', '1');
    expectXlvError(() => wb.defineName('TOTAL', '2'), 'XLV_E_INVALID_DEFINED_NAME');
    wb.defineName('total', '3', { sheet: 'A' });
    wb.defineName('Total', '4', { sheet: 'B' });
    expectXlvError(() => wb.defineName('TOTAL', '5', { sheet: 'A' }), 'XLV_E_INVALID_DEFINED_NAME');
  });

  it('streaming writer writes names, checking them up front or after an iterable of sheets', async () => {
    const bytes = await bytesOf(
      writeWorkbookStream({
        sheets: [{ name: 'Data', rows: [[1], [2]] }],
        definedNames: [{ name: 'Items', refersTo: 'Data!$A$1:$A$2' }, { name: 'One', refersTo: '1', sheet: 'Data', hidden: true }],
      }),
    );
    expect(readWorkbook(bytes).definedNames).toEqual([
      { name: 'Items', refersTo: 'Data!$A$1:$A$2', sheet: null, hidden: false, comment: null },
      { name: 'One', refersTo: '1', sheet: 'Data', hidden: true, comment: null },
    ]);
    expect((await openWorkbookStream(bytes)).definedNames).toHaveLength(2);
    expectXlvError(() => writeWorkbookStream({ sheets: [{ name: 'S', rows: [] }], definedNames: [{ name: 'A1', refersTo: '1' }] }), 'XLV_E_INVALID_DEFINED_NAME');
    expectXlvError(() => writeWorkbookStream({ sheets: [{ name: 'S', rows: [] }], definedNames: [{ name: 'x', refersTo: '1' }, { name: 'X', refersTo: '2' }] }), 'XLV_E_INVALID_DEFINED_NAME');
    expectXlvError(() => writeWorkbookStream({ sheets: [{ name: 'S', rows: [] }], definedNames: [{ name: 'x', refersTo: '1', sheet: 'T' }] }), 'XLV_E_SHEET_NOT_FOUND');
    expectXlvError(() => writeWorkbookStream({ sheets: [{ name: 'S', rows: [] }], definedNames: {} as never }), 'XLV_E_INVALID_OPTION');
    const sheets = (function* () {
      yield { name: 'Late', rows: [[1]] };
    })();
    const late = await bytesOf(writeWorkbookStream({ sheets, definedNames: [{ name: 'L', refersTo: 'Late!$A$1', sheet: 'Late' }] }));
    expect(readWorkbook(late).definedNames[0]?.sheet).toBe('Late');
    const missing = (function* () {
      yield { name: 'Only', rows: [[1]] };
    })();
    await expectXlvErrorAsync(() => bytesOf(writeWorkbookStream({ sheets: missing, definedNames: [{ name: 'L', refersTo: '1', sheet: 'Nope' }] })), 'XLV_E_SHEET_NOT_FOUND');
  });

  it('editCells keeps the defined names', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 1);
    wb.defineName('Keep', 'S!$A$1');
    const edited = editCells(writeWorkbook(wb), [{ sheet: 'S', cell: 'A2', value: 2 }]).bytes;
    expect(readWorkbook(edited).definedNames.map((n) => n.name)).toEqual(['Keep']);
  });
});
