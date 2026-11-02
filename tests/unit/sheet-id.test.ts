// Sheet ids (`sheet@sheetId`, ECMA-376 E-09: xsd:unsignedInt, required, unique; [MS-OI29500]
// §2.1.578: Office uses 1 to 65534, W2-20): given to addSheet or chosen when written, read back
// as `sheet.sheetId`, and kept by in-place edits (D-F03).
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { editCells } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { openWorkbookStream } from '../../packages/xlvane/src/stream-read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx } from '../helpers/xlsx-builder';

const sheetsXml = (bytes: Uint8Array): string => strFromU8(unzipSync(bytes)['xl/workbook.xml']!).replace(/^[\s\S]*<sheets>|<\/sheets>[\s\S]*$/g, '');

describe('writing', () => {
  it('numbers sheets 1, 2, 3 when no id is given, as before', () => {
    const wb = createWorkbook();
    wb.addSheet('A');
    wb.addSheet('B');
    expect(sheetsXml(writeWorkbook(wb))).toBe('<sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2"/>');
    expect(wb.sheets.map((s) => s.sheetId)).toEqual([null, null]);
  });

  it('writes given ids and gives the other sheets the lowest free ones', () => {
    const wb = createWorkbook();
    wb.addSheet('A');
    wb.addSheet('B', { sheetId: 1 });
    wb.addSheet('C', { sheetId: 7 });
    wb.addSheet('D');
    expect(wb.sheets.map((s) => s.sheetId)).toEqual([null, 1, 7, null]);
    const bytes = writeWorkbook(wb);
    // r:id and the part names stay positional; only sheetId changes.
    expect(sheetsXml(bytes)).toBe('<sheet name="A" sheetId="2" r:id="rId1"/><sheet name="B" sheetId="1" r:id="rId2"/><sheet name="C" sheetId="7" r:id="rId3"/><sheet name="D" sheetId="3" r:id="rId4"/>');
    expect(readWorkbook(bytes).sheets.map((s) => [s.name, s.sheetId])).toEqual([['A', 2], ['B', 1], ['C', 7], ['D', 3]]);
  });

  it('accepts whole numbers from 1 to 65534, each once', () => {
    for (const bad of [0, 65535, 1.5, -1, Number.NaN, '3', null]) {
      expectXlvError(() => createWorkbook().addSheet('S', { sheetId: bad as number }), 'XLV_E_INVALID_OPTION');
    }
    createWorkbook().addSheet('S', { sheetId: 65534 });
    const wb = createWorkbook();
    wb.addSheet('A', { sheetId: 4 });
    const e = expectXlvError(() => wb.addSheet('B', { sheetId: 4 }), 'XLV_E_INVALID_OPTION');
    expect(e.message).toContain('"A"');
    expect(wb.sheets).toHaveLength(1);
  });

  it('writes ids in the streaming writer', async () => {
    const stream = writeWorkbookStream({ sheets: [{ name: 'A', sheetId: 9, rows: [[1]] }, { name: 'B', rows: [[2]] }] });
    const bytes = new Uint8Array(await new Response(stream).arrayBuffer());
    expect(sheetsXml(bytes)).toBe('<sheet name="A" sheetId="9" r:id="rId1"/><sheet name="B" sheetId="1" r:id="rId2"/>');
    expect((await openWorkbookStream(bytes)).sheets.map((s) => s.sheetId)).toEqual([9, 1]);
  });
});

describe('reading', () => {
  it('gives the id of each sheet', async () => {
    const bytes = buildXlsx({ sheets: [{ name: 'A', sheetId: '5' }, { name: 'B', sheetId: ' 2 ' }] });
    const wb = readWorkbook(bytes);
    expect(wb.sheets.map((s) => s.sheetId)).toEqual([5, 2]);
    expect(wb.warnings).toEqual([]);
    expect((await openWorkbookStream(bytes)).sheets.map((s) => s.sheetId)).toEqual([5, 2]);
  });

  it('gives null with a warning for a missing, invalid or repeated id', () => {
    const wb = readWorkbook(buildXlsx({ sheets: [{ name: 'A', sheetId: null }, { name: 'B', sheetId: 'x' }, { name: 'C', sheetId: '4294967296' }, { name: 'D', sheetId: '3' }, { name: 'E', sheetId: '3' }] }));
    expect(wb.sheets.map((s) => s.sheetId)).toEqual([null, null, null, 3, null]);
    const invalid = wb.warnings.filter((w) => w.code === 'XLV_W_INVALID_VALUE');
    expect(invalid.map((w) => w.location?.sheet)).toEqual(['A', 'B', 'C', 'E']);
  });

  it('keeps the ids of a file in editCells', () => {
    const bytes = buildXlsx({ sheets: [{ name: 'A', sheetId: '7', data: '<row r="1"><c r="A1"><v>1</v></c></row>' }, { name: 'B', sheetId: '3' }] });
    const edited = editCells(bytes, [{ sheet: 'A', cell: 'B1', value: 2 }]).bytes;
    expect(readWorkbook(edited).sheets.map((s) => s.sheetId)).toEqual([7, 3]);
  });
});
