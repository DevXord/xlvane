import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { serialToDate } from '../../packages/xlvane/src/sml/dates';
import type { Cell, Workbook } from '../../packages/xlvane/src/model';
import { expectXlvError } from '../helpers/expect-error';
import { expectPrototypesUntouched } from '../helpers/prototypes';
import { buildXlsx, c, CT, inline, row } from '../helpers/xlsx-builder';

const STYLES = `
  <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/></numFmts>
  <cellXfs count="4">
    <xf numFmtId="0"/>
    <xf numFmtId="14" applyNumberFormat="1"/>
    <xf numFmtId="164" applyNumberFormat="1"/>
    <xf numFmtId="4" applyNumberFormat="1"/>
  </cellXfs>`;

function pick(cell: Cell | undefined): Pick<Cell, 'type' | 'value'> | undefined {
  return cell && { type: cell.type, value: cell.value };
}

function warningCodes(wb: Workbook): string[] {
  return wb.warnings.map((w) => w.code);
}

describe('readWorkbook', () => {
  const wb = readWorkbook(
    buildXlsx({
      sharedStrings: ['Name', 'Amount', 'Alice'],
      styles: STYLES,
      sheets: [
        {
          name: 'Sales',
          before: '<dimension ref="A1:H3"/>',
          data:
            row(1, c('A1', 0, 't="s"'), c('B1', 1, 't="s"')) +
            row(
              2,
              c('A2', 2, 't="s"'),
              c('B2', 1234.5),
              c('C2', 1, 't="b"'),
              c('D2', '#DIV/0!', 't="e"'),
              c('E2', 45000, 's="1"'),
              c('F2', 45000.5, 's="2"'),
              c('G2', 1234.5, 's="3"'),
              inline('H2', 'inline & text'),
            ) +
            row(3, '<c r="A3" t="str"><f>A2&amp;"!"</f><v>Alice!</v></c>', '<c r="B3"><f>B2*2</f><v>2469</v></c>'),
          after: '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
        },
        { name: 'Hidden', state: 'hidden', data: row(1, c('A1', 1)) },
      ],
    }),
  );

  it('lists sheets in workbook order with their properties', () => {
    expect(wb.sheets.map((s) => [s.name, s.index, s.kind, s.state, s.part])).toEqual([
      ['Sales', 0, 'worksheet', 'visible', '/xl/worksheets/sheet1.xml'],
      ['Hidden', 1, 'worksheet', 'hidden', '/xl/worksheets/sheet2.xml'],
    ]);
    expect(wb.dateSystem).toBe('1900');
    expect(wb.conformance).toBe('transitional');
    expect(wb.format).toBe('xlsx');
    expect(wb.hasMacros).toBe(false);
    expect(wb.externalLinkCount).toBe(0);
  });

  it('reads every cell type', () => {
    const s = wb.getSheet('Sales');
    expect(pick(s.cell('A1'))).toEqual({ type: 'string', value: 'Name' });
    expect(pick(s.cell('A2'))).toEqual({ type: 'string', value: 'Alice' });
    expect(pick(s.cell('B2'))).toEqual({ type: 'number', value: 1234.5 });
    expect(pick(s.cell('C2'))).toEqual({ type: 'boolean', value: true });
    expect(pick(s.cell('D2'))).toEqual({ type: 'error', value: '#DIV/0!' });
    expect(pick(s.cell('H2'))).toEqual({ type: 'string', value: 'inline & text' });
  });

  it('turns date-formatted numbers into dates, built-in and custom formats alike', () => {
    const s = wb.getSheet('Sales');
    expect(pick(s.cell('E2'))).toEqual({ type: 'date', value: serialToDate(45000, '1900') });
    expect(pick(s.cell('F2'))).toEqual({ type: 'date', value: serialToDate(45000.5, '1900') });
    expect(s.cell('F2')?.numFmt).toBe('yyyy\\-mm\\-dd');
    expect(pick(s.cell('G2'))).toEqual({ type: 'number', value: 1234.5 });
    expect(s.cell('G2')?.numFmt).toBe('#,##0.00');
    expect(s.cell('B2')?.numFmt).toBeNull();
  });

  it('reads formulas with their cached values', () => {
    const s = wb.getSheet('Sales');
    expect(s.cell('A3')).toMatchObject({ type: 'string', value: 'Alice!', formula: { kind: 'normal', text: 'A2&"!"' } });
    expect(s.cell('B3')).toMatchObject({ type: 'number', value: 2469, formula: { kind: 'normal', text: 'B2*2', ref: null } });
    expect(s.cell('B2')?.formula).toBeNull();
  });

  it('reports dimension, merges, cell counts and row order', () => {
    const s = wb.getSheet('Sales');
    expect(s.dimension).toBe('A1:H3');
    expect(s.merges).toEqual(['A1:B1']);
    expect(s.cellCount).toBe(12);
    expect([...s.rows()].map((r) => r.map((cell) => cell.ref))).toEqual([
      ['A1', 'B1'],
      ['A2', 'B2', 'C2', 'D2', 'E2', 'F2', 'G2', 'H2'],
      ['A3', 'B3'],
    ]);
    expect([...s.cells()]).toHaveLength(12);
  });

  it('throws XLV_E_SHEET_NOT_FOUND with the available sheets and a suggestion', () => {
    const error = expectXlvError(() => wb.getSheet('sales'), 'XLV_E_SHEET_NOT_FOUND');
    expect(error.hint).toContain('Sales');
    expect(error.details['available']).toEqual(['Sales', 'Hidden']);
    expect(error.location).toEqual({ sheet: 'sales' });
  });

  it('accepts an ArrayBuffer', () => {
    const bytes = buildXlsx();
    const copy = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    expect(readWorkbook(copy).sheets).toHaveLength(1);
  });
});

describe('cell layout', () => {
  it('places cells without r attributes by position and sorts out-of-order rows and cells', () => {
    const wb = readWorkbook(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            data: '<row r="3"><c r="C3"><v>3</v></c><c r="A3"><v>1</v></c></row><row><c><v>10</v></c><c><v>11</v></c></row><row r="1"><c><v>0</v></c></row>',
          },
        ],
      }),
    );
    const s = wb.getSheet('S');
    expect([...s.rows()].map((r) => r.map((cell) => `${cell.ref}=${String(cell.value)}`))).toEqual([
      ['A1=0'],
      ['A3=1', 'C3=3'],
      ['A4=10', 'B4=11'],
    ]);
  });

  it('does not store cells that have neither a value nor a formula', () => {
    const s = readWorkbook(buildXlsx({ sheets: [{ name: 'S', data: row(1, '<c r="A1" s="1"/>', c('B1', 5)) }] })).getSheet('S');
    expect(s.cell('A1')).toBeUndefined();
    expect(s.cellCount).toBe(1);
  });

  it('does not carry inline-string text over to the next cell on malformed input', () => {
    const s = readWorkbook(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            data: row(1, '<c r="A1" t="inlineStr"><is><t>AAA<f/>BBB</t></is></c>', '<c r="B1" t="inlineStr"><is><t>X</t></is></c>'),
          },
        ],
      }),
    ).getSheet('S');
    expect(s.cell('A1')?.value).toBe('AAABBB');
    expect(s.cell('B1')?.value).toBe('X');
  });

  it('ignores rows and cells outside sheetData', () => {
    const xml =
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>' +
      '<extLst><row r="1"><c r="A1"><v>1</v></c></row></extLst></worksheet>';
    expect(readWorkbook(buildXlsx({ sheets: [{ name: 'S', data: xml }] })).getSheet('S').cellCount).toBe(0);
  });

  it('warns when the styles or shared strings part is missing', () => {
    const wb = readWorkbook(buildXlsx({ styles: '', sharedStrings: [], omit: ['xl/styles.xml', 'xl/sharedStrings.xml'] }));
    expect(wb.warnings.filter((w) => w.code === 'XLV_W_PART_MISSING')).toHaveLength(2);
  });

  it('warns about duplicate cells and keeps the last one', () => {
    const wb = readWorkbook(buildXlsx({ sheets: [{ name: 'S', data: row(1, c('A1', 1), c('A1', 2)) }] }));
    expect(wb.getSheet('S').cell('A1')?.value).toBe(2);
    expect(wb.warnings).toContainEqual(expect.objectContaining({ code: 'XLV_W_DUPLICATE_CELL', count: 1, cells: ['A1'] }));
  });

  it.each(['A0', 'XFE1', 'A1048577', '1A', 'A-1', ''])('rejects the invalid cell reference "%s"', (ref) => {
    const xlsx = buildXlsx({ sheets: [{ name: 'S', data: `<row r="1"><c r="${ref}"><v>1</v></c></row>` }] });
    const error = expectXlvError(() => readWorkbook(xlsx), 'XLV_E_INVALID_CELL_REF');
    expect(error.location).toMatchObject({ sheet: 'S', part: '/xl/worksheets/sheet1.xml' });
  });

  it('enforces maxCells across sheets', () => {
    const xlsx = buildXlsx({
      sheets: [
        { name: 'A', data: row(1, c('A1', 1), c('B1', 2)) },
        { name: 'B', data: row(1, c('A1', 1), c('B1', 2)) },
      ],
    });
    expect(readWorkbook(xlsx, { limits: { maxCells: 4 } }).sheets).toHaveLength(2);
    expectXlvError(() => readWorkbook(xlsx, { limits: { maxCells: 3 } }), 'XLV_E_LIMIT_EXCEEDED', { limit: 'maxCells' });
  });
});

describe('formulas', () => {
  const wb = readWorkbook(
    buildXlsx({
      sheets: [
        {
          name: 'F',
          data: row(
            1,
            '<c r="A1"><f t="shared" ref="A1:A3" si="0">B1*2</f><v>2</v></c>',
            '<c r="B1"><f t="array" ref="B1:B2">SUM(C1:C2*2)</f><v>4</v></c>',
            '<c r="C1"><f>1/0</f></c>',
          ) + row(2, '<c r="A2"><f t="shared" si="0"/><v>4</v></c>'),
        },
      ],
    }),
  );
  const s = wb.getSheet('F');

  it('reads shared master, shared child and array formulas', () => {
    expect(s.cell('A1')?.formula).toEqual({ kind: 'shared', text: 'B1*2', ref: 'A1:A3', sharedIndex: 0 });
    expect(s.cell('A2')?.formula).toEqual({ kind: 'shared', text: null, ref: null, sharedIndex: 0 });
    expect(s.cell('A2')?.value).toBe(4);
    expect(s.cell('B1')?.formula).toEqual({ kind: 'array', text: 'SUM(C1:C2*2)', ref: 'B1:B2', sharedIndex: null });
  });

  it('keeps formula cells without a cached value and warns', () => {
    expect(s.cell('C1')).toMatchObject({ type: 'empty', value: null, formula: { text: '1/0' } });
    expect(warningCodes(wb)).toEqual(
      expect.arrayContaining(['XLV_W_SHARED_FORMULA_NOT_EXPANDED', 'XLV_W_FORMULA_NO_CACHED_VALUE']),
    );
  });
});

describe('workbook properties', () => {
  it('uses the 1904 date system when workbookPr says so', () => {
    const wb = readWorkbook(buildXlsx({ workbookPr: 'date1904="1"', styles: STYLES, sheets: [{ name: 'S', data: row(1, c('A1', 0, 's="1"')) }] }));
    expect(wb.dateSystem).toBe('1904');
    expect(wb.getSheet('S').cell('A1')?.value).toEqual(serialToDate(0, '1904'));
  });

  it('reads ISO 8601 date cells (t="d")', () => {
    const s = readWorkbook(buildXlsx({ sheets: [{ name: 'S', data: row(1, c('A1', '2024-02-29T13:45:00', 't="d"')) }] })).getSheet('S');
    expect(s.cell('A1')).toMatchObject({ type: 'date', value: new Date(Date.UTC(2024, 1, 29, 13, 45, 0)) });
  });

  it('reads Strict conformance packages', () => {
    const wb = readWorkbook(
      buildXlsx({ strict: true, sharedStrings: ['strict'], sheets: [{ name: 'S', data: row(1, c('A1', 0, 't="s"')) }] }),
    );
    expect(wb.conformance).toBe('strict');
    expect(wb.getSheet('S').cell('A1')?.value).toBe('strict');
  });

  it.each([
    [CT.workbook, 'xlsx'],
    [CT.template, 'xltx'],
    [CT.xlsm, 'xlsm'],
    ['application/vnd.ms-excel.template.macroEnabled.main+xml', 'xltm'],
    ['application/xml', 'unknown'],
  ])('derives the format from the content type %s', (type, format) => {
    expect(readWorkbook(buildXlsx({ workbookContentType: type })).format).toBe(format);
  });

  it('reports macros and external links without touching them', () => {
    const wb = readWorkbook(buildXlsx({ vba: true, externalLinks: 2 }));
    expect(wb.hasMacros).toBe(true);
    expect(wb.externalLinkCount).toBe(2);
    expect(warningCodes(wb)).toEqual(expect.arrayContaining(['XLV_W_MACROS_PRESENT', 'XLV_W_EXTERNAL_LINKS']));
  });

  it('warns about very hidden sheets', () => {
    const wb = readWorkbook(buildXlsx({ sheets: [{ name: 'V', state: 'veryHidden' }] }));
    expect(wb.sheets[0]?.state).toBe('veryHidden');
    expect(warningCodes(wb)).toContain('XLV_W_VERY_HIDDEN_SHEET');
  });

  it('keeps chartsheets as sheets without cells', () => {
    const wb = readWorkbook(buildXlsx({ sheets: [{ name: 'Data', data: row(1, c('A1', 1)) }, { name: 'Chart', kind: 'chartsheet' }] }));
    expect(wb.getSheet('Chart')).toMatchObject({ kind: 'chartsheet', cellCount: 0 });
  });

  it('warns when a sheet part is missing', () => {
    const wb = readWorkbook(buildXlsx({ sheets: [{ name: 'Gone', missingPart: true }] }));
    expect(wb.getSheet('Gone')).toMatchObject({ part: null, cellCount: 0 });
    expect(warningCodes(wb)).toContain('XLV_W_SHEET_PART_MISSING');
  });

  it('warns about shared string indexes that do not exist', () => {
    const wb = readWorkbook(buildXlsx({ sharedStrings: ['only'], sheets: [{ name: 'S', data: row(1, c('A1', 5, 't="s"'), c('B1', 0, 't="s"')) }] }));
    expect(wb.getSheet('S').cell('B1')?.value).toBe('only');
    expect(wb.getSheet('S').cell('A1')).toBeUndefined();
    expect(wb.warnings).toContainEqual(expect.objectContaining({ code: 'XLV_W_INVALID_VALUE', cells: ['A1'] }));
  });

  it('warns about values that do not match the cell type', () => {
    const wb = readWorkbook(buildXlsx({ sheets: [{ name: 'S', data: row(1, c('A1', 'abc'), c('B1', 'maybe', 't="b"'), c('C1', '0x10')) }] }));
    const s = wb.getSheet('S');
    expect(s.cell('A1')).toMatchObject({ type: 'string', value: 'abc' });
    expect(s.cell('C1')).toMatchObject({ type: 'string', value: '0x10' });
    expect(wb.warnings).toContainEqual(expect.objectContaining({ code: 'XLV_W_INVALID_VALUE', count: 3 }));
  });
});

describe('invalid input', () => {
  it('rejects values that are not bytes', () => {
    expectXlvError(() => readWorkbook('report.xlsx' as unknown as Uint8Array), 'XLV_E_INVALID_INPUT');
    expectXlvError(() => readWorkbook(null as unknown as Uint8Array), 'XLV_E_INVALID_INPUT');
  });

  it('rejects unknown options', () => {
    const error = expectXlvError(() => readWorkbook(buildXlsx(), { limit: {} } as never), 'XLV_E_INVALID_OPTION');
    expect(error.hint).toContain('limits');
  });

  it('reports a package without a workbook part', () => {
    expectXlvError(() => readWorkbook(buildXlsx({ omit: ['xl/workbook.xml'] })), 'XLV_E_PART_MISSING');
    expectXlvError(() => readWorkbook(buildXlsx({ omit: ['_rels/.rels'] })), 'XLV_E_PART_MISSING');
  });

  it('rejects a workbook part with the wrong root element', () => {
    expectXlvError(
      () => readWorkbook(buildXlsx({ extraParts: { 'xl/workbook.xml': '<document xmlns="urn:x"/>' } })),
      'XLV_E_PART_INVALID',
    );
  });
});

describe('prototype pollution', () => {
  it('treats sheet names and strings like __proto__ as plain data', () => {
    const wb = readWorkbook(
      buildXlsx({
        sharedStrings: ['__proto__', 'constructor'],
        sheets: [
          { name: '__proto__', data: row(1, c('A1', 0, 't="s"'), c('B1', 1, 't="s"')) },
          { name: 'constructor', data: '' },
          { name: 'prototype', data: '' },
        ],
      }),
    );
    expect(wb.getSheet('__proto__').cell('A1')?.value).toBe('__proto__');
    expect(wb.getSheet('constructor').name).toBe('constructor');
    expectXlvError(() => wb.getSheet('toString'), 'XLV_E_SHEET_NOT_FOUND');
    expectPrototypesUntouched();
  });
});
