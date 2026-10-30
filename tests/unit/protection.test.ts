// Sheet and workbook protection, without passwords (ECMA-376 Part 1 §18.3.1.85, §18.2.29;
// [MS-OI29500] §2.1.586, §2.1.654; docs/facts/sheet-rules.md SR-20 … SR-22; D-C03). Password
// hashes are read and kept as data, never interpreted; writing them is phase C2.
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { editCells } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx } from '../helpers/xlsx-builder';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const sheetXml = (bytes: Uint8Array): string => strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']!);

describe('writing', () => {
  it('writes sheet="1" and only the flags that differ from the schema defaults, between sheetData and autoFilter', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 'x').autoFilter('A1:B5').protect({ sort: false, autoFilter: false, selectLockedCells: true });
    const xml = sheetXml(writeWorkbook(wb));
    expect(xml).toContain('</sheetData><sheetProtection sheet="1" selectLockedCells="1" sort="0" autoFilter="0"/><autoFilter ref="A1:B5"/>');
  });

  it('protect() with no options uses the schema defaults; null removes it', () => {
    const wb = createWorkbook();
    wb.addSheet('S').protect();
    expect(sheetXml(writeWorkbook(wb))).toContain('<sheetProtection sheet="1"/>');
    const off = createWorkbook();
    off.addSheet('S').protect().protect(null);
    expect(sheetXml(writeWorkbook(off))).not.toContain('sheetProtection');
  });

  it('refuses a password until phase C2 and unknown flags', () => {
    const e = expectXlvError(() => createWorkbook().addSheet('S').protect({ password: 'x' } as never), 'XLV_E_NOT_SUPPORTED');
    expect(e.hint).toContain('without a password');
    expectXlvError(() => createWorkbook().addSheet('S').protect({ sorting: false } as never), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => createWorkbook().addSheet('S').protect({ sort: 'no' } as never), 'XLV_E_INVALID_OPTION');
  });

  it('round-trips the flags', () => {
    const wb = createWorkbook();
    wb.addSheet('S').protect({ formatColumns: false, objects: true });
    const p = readWorkbook(writeWorkbook(wb)).getSheet('S').protection;
    expect(p).toMatchObject({ sheet: true, formatColumns: false, objects: true, formatCells: true, hashValue: null, password: null });
  });
});

describe('reading', () => {
  const read = (protection: string, workbook = '') =>
    readWorkbook(buildXlsx({ sheets: [{ name: 'S', data: `<worksheet xmlns="${NS}"><sheetData/>${protection}</worksheet>` }], workbookAfterSheets: workbook }));

  it('is null without the element and applies the schema defaults', () => {
    expect(read('').sheets[0]!.protection).toBeNull();
    expect(read('<sheetProtection/>').sheets[0]!.protection).toMatchObject({ sheet: false, objects: false, formatCells: true, selectLockedCells: false, sort: true });
  });

  it('keeps the password verifier as data (corpus form)', () => {
    const p = read(
      '<sheetProtection algorithmName="SHA-512" hashValue="7oosxvc+NBQ=" saltValue="Md8b5c0CY+y/ULfm8qFdig==" spinCount="100000" sheet="1" formatCells="0" objects="1" scenarios="1"/>',
    ).sheets[0]!.protection;
    expect(p).toMatchObject({ sheet: true, algorithmName: 'SHA-512', hashValue: '7oosxvc+NBQ=', saltValue: 'Md8b5c0CY+y/ULfm8qFdig==', spinCount: 100000, formatCells: false, objects: true });
    expect(read('<sheetProtection password="CBEB" sheet="true"/>').sheets[0]!.protection?.password).toBe('CBEB');
  });

  it('warns about invalid booleans', () => {
    const wb = read('<sheetProtection sheet="maybe"/>');
    expect(wb.sheets[0]!.protection?.sheet).toBe(false);
    expect(wb.warnings.map((w) => w.code)).toContain('XLV_W_INVALID_VALUE');
  });

  it('reads workbook protection', () => {
    expect(read('').protection).toBeNull();
    const p = read('', '<workbookProtection workbookAlgorithmName="SHA-512" workbookHashValue="64Ct" workbookSaltValue="lUZr" workbookSpinCount="100000" lockStructure="1"/>').protection;
    expect(p).toMatchObject({ lockStructure: true, lockWindows: false, workbookAlgorithmName: 'SHA-512', workbookSpinCount: 100000, workbookPassword: null });
  });

  it('editCells keeps the protection of the file', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 1).protect({ sort: false });
    const edited = editCells(writeWorkbook(wb), [{ sheet: 'S', cell: 'B1', value: 2 }]);
    expect(readWorkbook(edited.bytes).getSheet('S').protection).toMatchObject({ sheet: true, sort: false });
  });
});
