// Conditional formatting and differential formats (ECMA-376 Part 1 §18.3.1.10, §18.3.1.11,
// §18.3.1.15, §18.3.1.16, §18.3.1.18, §18.3.1.28, §18.3.1.43, §18.3.1.49, §18.8.14, §18.8.15;
// [MS-OI29500] §2.1.595, §2.1.596, §2.1.601, §2.1.602, §2.1.611, §2.1.1085, §2.1.1092;
// docs/facts/sheet-rules.md SR-07 … SR-16; docs/decisions.md D-C02).
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { editCells } from '../../packages/xlvane/src/edit';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { deleteRows, insertRows } from '../../packages/xlvane/src/rows';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import type { XlvErrorCode } from '../../packages/xlvane/src/errors';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx } from '../helpers/xlsx-builder';

/** Expects an XlvError with `code` whose message, details or hint match `pattern`. */
function expectRefused(fn: () => unknown, code: XlvErrorCode, pattern?: RegExp): void {
  const error = expectXlvError(fn, code);
  if (pattern) expect(`${error.message} ${JSON.stringify(error.details)} ${error.hint}`).toMatch(pattern);
}

const part = (bytes: Uint8Array, name: string): string => strFromU8(unzipSync(bytes)[name] as Uint8Array);

describe('reading conditional formatting', () => {
  const file = buildXlsx({
    styles:
      '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border/></borders><cellXfs count="1"><xf/></cellXfs>' +
      '<dxfs count="2">' +
      '<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>' +
      '<dxf><numFmt numFmtId="164" formatCode="0.0%"/><border><left style="thin"/></border></dxf>' +
      '</dxfs>',
    sheets: [
      {
        name: 'S',
        data: '<row r="1"><c r="A1"><v>1</v></c></row>',
        after:
          '<conditionalFormatting sqref="A1:A10">' +
          '<cfRule type="cellIs" dxfId="0" priority="2" operator="greaterThan"><formula>100</formula></cfRule>' +
          '<cfRule type="top10" dxfId="1" priority="3" percent="1" rank="5"/>' +
          '</conditionalFormatting>' +
          '<conditionalFormatting sqref="B1:B10 D1">' +
          '<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color theme="5"/></colorScale></cfRule>' +
          '<cfRule type="dataBar" priority="4"><dataBar minLength="0" showValue="0"><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">{1}</x14:id></ext></extLst></cfRule>' +
          '<cfRule type="iconSet" priority="5"><iconSet iconSet="3Arrows" reverse="1"><cfvo type="percent" val="0"/><cfvo type="percent" val="33" gte="0"/><cfvo type="percent" val="67"/></iconSet></cfRule>' +
          '<cfRule type="expression" dxfId="9" priority="6" stopIfTrue="1"><formula>MOD(ROW(),2)=0</formula></cfRule>' +
          '<cfRule type="nonsense" priority="7"/>' +
          '</conditionalFormatting>',
      },
    ],
  });

  it('reads the rules with their differential formats', () => {
    const wb = readWorkbook(file);
    const cf = wb.getSheet('S').conditionalFormats;
    expect(cf.map((c) => c.sqref)).toEqual(['A1:A10', 'B1:B10 D1']);
    const [cellIs, top] = cf[0]?.rules ?? [];
    expect(cellIs).toMatchObject({
      type: 'cellIs',
      priority: 2,
      operator: 'greaterThan',
      formulas: ['100'],
      stopIfTrue: false,
      style: { font: { bold: true, color: 'FF9C0006' }, fill: { pattern: 'solid', fgColor: 'FFFFC7CE' } },
    });
    expect(top).toMatchObject({ type: 'top10', rank: 5, percent: true, bottom: false, style: { numFmt: '0.0%', border: { left: { style: 'thin' } } } });
    const rules = cf[1]?.rules ?? [];
    expect(rules.map((r) => r.type)).toEqual(['colorScale', 'dataBar', 'iconSet', 'expression']);
    expect(rules[0]?.colorScale).toEqual({
      cfvo: [
        { type: 'min', value: null, gte: true },
        { type: 'percentile', value: '50', gte: true },
        { type: 'max', value: null, gte: true },
      ],
      colors: ['FFF8696B', 'FFFFEB84', { theme: 5 }],
    });
    expect(rules[1]?.dataBar).toMatchObject({ minLength: 0, maxLength: 90, showValue: false, color: 'FF638EC6' });
    expect(rules[2]?.iconSet).toMatchObject({ iconSet: '3Arrows', reverse: true, showValue: true, percent: true });
    expect(rules[2]?.iconSet?.cfvo[1]).toEqual({ type: 'percent', value: '33', gte: false });
    expect(rules[3]).toMatchObject({ stopIfTrue: true, style: null, formulas: ['MOD(ROW(),2)=0'] });
    expect(wb.warnings.filter((w) => w.code === 'XLV_W_INVALID_VALUE').length).toBeGreaterThan(0);
  });
});

describe('writing conditional formatting', () => {
  it('writes rules after the merges and before the validations, with deduplicated dxfs', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('S').set('A1', 1).merge('C1:D1');
    const red = { font: { bold: true, color: '9C0006' }, fill: { pattern: 'solid', fgColor: 'FFC7CE' } } as const;
    s.conditionalFormat('A1:A10', [
      { type: 'cellIs', operator: 'greaterThan', formulas: [100], style: red },
      { type: 'expression', formulas: ['=MOD(ROW(),2)=0'], style: red, stopIfTrue: true },
    ]);
    s.conditionalFormat('B1:B10', [{ type: 'top10', rank: 3, bottom: true, style: { numFmt: '0.0%' } }]);
    s.validate('E1', { type: 'list', formula1: '"a"' });
    const bytes = writeWorkbook(wb);
    const xml = part(bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain(
      '</mergeCells><conditionalFormatting sqref="A1:A10"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100</formula></cfRule>' +
        '<cfRule type="expression" dxfId="0" priority="2" stopIfTrue="1"><formula>MOD(ROW(),2)=0</formula></cfRule></conditionalFormatting>' +
        '<conditionalFormatting sqref="B1:B10"><cfRule type="top10" dxfId="1" priority="3" bottom="1" rank="3"/></conditionalFormatting><dataValidations',
    );
    const styles = part(bytes, 'xl/styles.xml');
    expect(styles).toContain(
      '</cellStyles><dxfs count="2"><dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>',
    );
    expect(styles).toMatch(/<dxf><numFmt numFmtId="(\d+)" formatCode="0\.0%"\/><\/dxf><\/dxfs>/);
    const back = readWorkbook(bytes).getSheet('S').conditionalFormats;
    expect(back[0]?.rules[0]).toMatchObject({ type: 'cellIs', priority: 1, formulas: ['100'], style: { font: { bold: true, color: 'FF9C0006' }, fill: { pattern: 'solid', fgColor: 'FFFFC7CE' } } });
    expect(back[1]?.rules[0]).toMatchObject({ type: 'top10', rank: 3, bottom: true, style: { numFmt: '0.0%' } });
  });

  it('writes colour scales, data bars and icon sets without dxfs', () => {
    const wb = createWorkbook();
    wb.addSheet('S').conditionalFormat('A1:A9', [
      { type: 'colorScale', colorScale: { cfvo: [{ type: 'min' }, { type: 'max' }], colors: ['F8696B', { theme: 5, tint: 0.4 }] } },
      { type: 'dataBar', dataBar: { cfvo: [{ type: 'num', value: 0 }, { type: 'formula', value: '=MAX($A$1:$A$9)' }], color: '638EC6', minLength: 0, maxLength: 100 } },
      { type: 'iconSet', iconSet: { iconSet: '4Arrows', cfvo: [{ type: 'percent', value: 0 }, { type: 'percent', value: 25 }, { type: 'percent', value: 50, gte: false }, { type: 'percent', value: 75 }], showValue: false } },
    ]);
    const bytes = writeWorkbook(wb);
    const xml = part(bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color theme="5" tint="0.4"/></colorScale></cfRule>');
    expect(xml).toContain('<cfRule type="dataBar" priority="2"><dataBar minLength="0" maxLength="100"><cfvo type="num" val="0"/><cfvo type="formula" val="MAX($A$1:$A$9)"/><color rgb="FF638EC6"/></dataBar></cfRule>');
    expect(xml).toContain('<iconSet iconSet="4Arrows" showValue="0"><cfvo type="percent" val="0"/><cfvo type="percent" val="25"/><cfvo type="percent" val="50" gte="0"/>');
    expect(part(bytes, 'xl/styles.xml')).not.toContain('<dxfs');
    const rules = readWorkbook(bytes).getSheet('S').conditionalFormats[0]?.rules ?? [];
    expect(rules[1]?.dataBar).toMatchObject({ minLength: 0, maxLength: 100, cfvo: [{ type: 'num', value: '0' }, { type: 'formula', value: 'MAX($A$1:$A$9)' }] });
    expect(rules[2]?.iconSet).toMatchObject({ iconSet: '4Arrows', showValue: false });
  });

  it('assigns increasing priorities and refuses duplicates', () => {
    const s = createWorkbook().addSheet('S');
    s.conditionalFormat('A1', [{ type: 'expression', formulas: ['TRUE'], priority: 5 }]);
    s.conditionalFormat('B1', [{ type: 'expression', formulas: ['TRUE'] }]);
    expectRefused(() => s.conditionalFormat('C1', [{ type: 'expression', formulas: ['TRUE'], priority: 6 }]), 'XLV_E_INVALID_OPTION', /priority/);
    expectRefused(() => s.conditionalFormat('C1', [{ type: 'expression', formulas: ['TRUE'], priority: 0 }]), 'XLV_E_INVALID_OPTION', /priority/);
  });

  it('enforces the per-type rules and the [MS-OI29500] limits', () => {
    const s = createWorkbook().addSheet('S');
    expectRefused(() => s.conditionalFormat('A1', [{ type: 'bogus' as never }]), 'XLV_E_INVALID_OPTION', /type/);
    expectRefused(() => s.conditionalFormat('A1', [{ type: 'cellIs', formulas: [1] }]), 'XLV_E_INVALID_OPTION', /operator/);
    expectRefused(() => s.conditionalFormat('A1', [{ type: 'expression' }]), 'XLV_E_INVALID_OPTION', /formulas/);
    expectRefused(() => s.conditionalFormat('A1', [{ type: 'containsText', text: 'x' }]), 'XLV_E_INVALID_OPTION', /formulas/);
    expectRefused(() => s.conditionalFormat('A1', [{ type: 'timePeriod', formulas: ['TRUE'] }]), 'XLV_E_INVALID_OPTION', /timePeriod/);
    expectRefused(() => s.conditionalFormat('A1', [{ type: 'expression', formulas: ['1', '2', '3', '4'] }]), 'XLV_E_INVALID_OPTION', /formulas/);
    expectRefused(() => s.conditionalFormat('A1', [{ type: 'expression', formulas: ['x'.repeat(8193)] }]), 'XLV_E_INVALID_OPTION', /8192/);
    expectRefused(() => s.conditionalFormat('A1', [{ type: 'expression', formulas: ['TRUE'], rank: 3 }]), 'XLV_E_INVALID_OPTION', /rank/);
    expectRefused(
      () => s.conditionalFormat('A1', [{ type: 'colorScale', colorScale: { cfvo: [{ type: 'min' }, { type: 'num', value: 1 }, { type: 'num', value: 2 }, { type: 'max' }], colors: ['FF0000', '00FF00', '0000FF', '000000'] } }]),
      'XLV_E_INVALID_OPTION',
      /3/,
    );
    expectRefused(() => s.conditionalFormat('A1', [{ type: 'colorScale', colorScale: { cfvo: [{ type: 'min' }, { type: 'max' }], colors: ['FF0000'] } }]), 'XLV_E_INVALID_OPTION', /colors/);
    expectRefused(() => s.conditionalFormat('A1', [{ type: 'dataBar', dataBar: { cfvo: [{ type: 'min' }, { type: 'max' }], color: '638EC6', maxLength: 101 } }]), 'XLV_E_INVALID_OPTION', /100/);
    expectRefused(() => s.conditionalFormat('A1', [{ type: 'iconSet', iconSet: { cfvo: [{ type: 'min' }, { type: 'percent', value: 50 }] } }]), 'XLV_E_INVALID_OPTION', /min/);
    expectRefused(() => s.conditionalFormat('A1', [{ type: 'colorScale', style: { font: { bold: true } }, colorScale: { cfvo: [{ type: 'min' }, { type: 'max' }], colors: ['FF0000', '00FF00'] } }]), 'XLV_E_INVALID_OPTION', /style/);
    expectRefused(() => s.conditionalFormat('A1', [{ type: 'expression', formulas: ['TRUE'], style: { alignment: { wrapText: true } } }]), 'XLV_E_INVALID_OPTION', /alignment/);
    expectRefused(() => s.conditionalFormat('A1', []), 'XLV_E_INVALID_OPTION', /rule/);
    expectRefused(() => s.conditionalFormat(Array.from({ length: 8193 }, (_, i) => `A${String(i + 1)}`).join(' '), [{ type: 'expression', formulas: ['TRUE'] }]), 'XLV_E_INVALID_OPTION', /8192/);
  });

  it('is written by the streaming writer, with dxfs in the styles part', async () => {
    const stream = writeWorkbookStream({
      sheets: [{ name: 'S', rows: [[1], [2]], conditionalFormats: [{ sqref: 'A1:A2', rules: [{ type: 'cellIs', operator: 'equal', formulas: [2], style: { font: { italic: true } } }] }] }],
    });
    const bytes = new Uint8Array(await new Response(stream).arrayBuffer());
    expect(readWorkbook(bytes).getSheet('S').conditionalFormats[0]?.rules[0]).toMatchObject({ type: 'cellIs', operator: 'equal', style: { font: { italic: true } } });
  });

  it('is kept by editCells and follows insertRows / deleteRows', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', 1).conditionalFormat('A2:A6', [{ type: 'expression', formulas: ['$A2>1'], style: { font: { bold: true } } }]);
    const bytes = writeWorkbook(wb);
    const edited = editCells(bytes, [{ sheet: 'S', cell: 'B1', value: 'x' }]);
    expect(readWorkbook(edited.bytes).getSheet('S').conditionalFormats).toEqual(readWorkbook(bytes).getSheet('S').conditionalFormats);
    const inserted = insertRows(bytes, { sheet: 'S', at: 3, count: 1 });
    const cf = readWorkbook(inserted.bytes).getSheet('S').conditionalFormats;
    expect(cf[0]).toMatchObject({ sqref: 'A2:A7', rules: [{ formulas: ['$A2>1'], style: { font: { bold: true } } }] });
    const deleted = deleteRows(inserted.bytes, { sheet: 'S', at: 2, count: 2 });
    expect(readWorkbook(deleted.bytes).getSheet('S').conditionalFormats[0]?.sqref).toBe('A2:A5');
  });
});
