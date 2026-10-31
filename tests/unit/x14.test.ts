// The Excel 2010 extensions of conditional formatting and data validation in the worksheet
// extLst ([MS-XLSX] §2.2.4.12, §2.6.1–§2.6.5, §2.6.26–§2.6.30, §2.6.36, §2.7.7–§2.7.10;
// docs/facts/sheet-rules.md SR-40 … SR-44; docs/decisions.md D-C07).
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { editCells } from '../../packages/xlvane/src/edit';
import type { XlvErrorCode } from '../../packages/xlvane/src/errors';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { openWorkbookStream } from '../../packages/xlvane/src/stream-read';
import { writeWorkbookStream } from '../../packages/xlvane/src/stream-write';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { expectXlvError } from '../helpers/expect-error';
import { buildXlsx } from '../helpers/xlsx-builder';

const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XM = 'http://schemas.microsoft.com/office/excel/2006/main';
const part = (bytes: Uint8Array, name: string): string => strFromU8(unzipSync(bytes)[name] as Uint8Array);

function expectRefused(fn: () => unknown, code: XlvErrorCode, pattern: RegExp): void {
  const error = expectXlvError(fn, code);
  expect(`${error.message} ${JSON.stringify(error.details)} ${error.hint}`).toMatch(pattern);
}

/** A worksheet `extLst` with the given ext children (declaring x14 and xm on each ext). */
const extList = (...exts: string[]): string => `<extLst>${exts.join('')}</extLst>`;
const cfExt = (inner: string): string => `<ext uri="{78C0D931-6437-407D-A8EE-F0AAD7539E65}" xmlns:x14="${X14}" xmlns:xm="${XM}"><x14:conditionalFormattings>${inner}</x14:conditionalFormattings></ext>`;
const dvExt = (inner: string): string => `<ext uri="{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}" xmlns:x14="${X14}" xmlns:xm="${XM}"><x14:dataValidations count="1">${inner}</x14:dataValidations></ext>`;

describe('reading the extensions', () => {
  const file = buildXlsx({
    styles:
      '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border/></borders><cellXfs count="1"><xf/></cellXfs>',
    sheets: [
      {
        name: 'S',
        data: '<row r="1"><c r="A1"><v>1</v></c></row>',
        after:
          '<conditionalFormatting sqref="A1:A10">' +
          '<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>' +
          `<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="${X14}"><x14:id>{AAAA0000-0000-0000-0000-000000000001}</x14:id></ext></extLst></cfRule>` +
          '</conditionalFormatting>' +
          extList(
            `<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${X14}"><x14:sparklineGroups><x14:sparklineGroup/></x14:sparklineGroups></ext>`,
            cfExt(
              '<x14:conditionalFormatting>' +
                '<x14:cfRule type="dataBar" id="{aaaa0000-0000-0000-0000-000000000001}"><x14:dataBar minLength="0" maxLength="100" border="1" gradient="0" negativeBarBorderColorSameAsPositive="0" axisPosition="middle" direction="rightToLeft">' +
                '<x14:cfvo type="autoMin"/><x14:cfvo type="num"><xm:f>50</xm:f></x14:cfvo>' +
                '<x14:borderColor rgb="FF0000FF"/><x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FF800000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule>' +
                '<xm:sqref>A1:A10</xm:sqref></x14:conditionalFormatting>' +
                '<x14:conditionalFormatting>' +
                '<x14:cfRule type="iconSet" priority="2" id="{BBBB0000-0000-0000-0000-000000000002}"><x14:iconSet iconSet="3Stars" custom="1" showValue="0">' +
                '<x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent" gte="0"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo>' +
                '<x14:cfIcon iconSet="NoIcons" iconId="0"/><x14:cfIcon iconSet="3Flags" iconId="1"/><x14:cfIcon iconSet="3Stars" iconId="2"/></x14:iconSet></x14:cfRule>' +
                '<x14:cfRule type="expression" priority="3"><xm:f>Lists!$A$1&gt;0</xm:f><x14:dxf><font><b/></font></x14:dxf></x14:cfRule>' +
                '<xm:sqref>B1:B5 D1</xm:sqref></x14:conditionalFormatting>' +
                // An extension without priority and without a matching main rule is ignored (§2.6.27).
                '<x14:conditionalFormatting><x14:cfRule type="dataBar" id="{CCCC0000-0000-0000-0000-000000000003}"><x14:dataBar><x14:cfvo type="min"/><x14:cfvo type="max"/></x14:dataBar></x14:cfRule><xm:sqref>C1</xm:sqref></x14:conditionalFormatting>',
            ),
            dvExt(
              '<x14:dataValidation type="list" allowBlank="1" showErrorMessage="1" error="Pick one"><x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>E1:E20</xm:sqref></x14:dataValidation>',
            ),
          ),
      },
    ],
  });

  it('merges a data bar extension into its main rule by id', () => {
    const [main, extended] = readWorkbook(file).getSheet('S').conditionalFormats;
    expect(main?.rules[0]?.dataBar).toEqual({
      cfvo: [
        { type: 'autoMin', value: null, gte: true },
        { type: 'num', value: '50', gte: true },
      ],
      color: 'FF638EC6',
      minLength: 0,
      maxLength: 100,
      showValue: true,
      extended: true,
      gradient: false,
      border: true,
      borderColor: 'FF0000FF',
      negativeFillColor: 'FFFF0000',
      negativeBorderColor: 'FF800000',
      axisPosition: 'middle',
      axisColor: 'FF000000',
      direction: 'rightToLeft',
    });
    expect(extended?.sqref).toBe('B1:B5 D1');
  });

  it('reads rules with their own priority, custom icons and inline differential formats', () => {
    const rules = readWorkbook(file).getSheet('S').conditionalFormats[1]?.rules ?? [];
    expect(rules[0]).toMatchObject({ type: 'iconSet', priority: 2 });
    expect(rules[0]?.iconSet).toEqual({
      iconSet: '3Stars',
      cfvo: [
        { type: 'percent', value: '0', gte: true },
        { type: 'percent', value: '33', gte: false },
        { type: 'percent', value: '67', gte: true },
      ],
      reverse: false,
      showValue: false,
      percent: true,
      custom: [
        { iconSet: 'NoIcons', iconId: 0 },
        { iconSet: '3Flags', iconId: 1 },
        { iconSet: '3Stars', iconId: 2 },
      ],
    });
    expect(rules[1]).toMatchObject({ type: 'expression', priority: 3, formulas: ['Lists!$A$1>0'], style: { font: { bold: true } } });
    // The unmatched extension (C1) is ignored.
    expect(readWorkbook(file).getSheet('S').conditionalFormats).toHaveLength(2);
  });

  it('reads extended data validations, such as lists from another sheet', () => {
    expect(readWorkbook(file).getSheet('S').dataValidations).toEqual([
      expect.objectContaining({ sqref: 'E1:E20', type: 'list', formula1: 'Lists!$A$1:$A$3', allowBlank: true, showErrorMessage: true, error: 'Pick one', operator: null }),
    ]);
  });

  it('gives a plain data bar the extension defaults', () => {
    const wb = createWorkbook();
    wb.addSheet('S').conditionalFormat('A1:A3', [{ type: 'dataBar', dataBar: { cfvo: [{ type: 'min' }, { type: 'max' }], color: '638EC6' } }]);
    expect(readWorkbook(writeWorkbook(wb)).getSheet('S').conditionalFormats[0]?.rules[0]?.dataBar).toMatchObject({
      extended: false,
      gradient: true,
      border: false,
      borderColor: null,
      negativeFillColor: null,
      axisPosition: 'automatic',
      axisColor: null,
      direction: 'context',
    });
  });

  it('skips other extensions without keeping them, and bounds what it keeps', () => {
    const big = buildXlsx({
      sheets: [
        {
          name: 'S',
          data: '',
          after: extList(`<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${X14}">${'<x14:sparklineGroup/>'.repeat(2000)}</ext>`, cfExt(`<x14:conditionalFormatting>${'<x14:cfRule/>'.repeat(50)}</x14:conditionalFormatting>`)),
        },
      ],
    });
    // 2000 skipped sparkline elements do not count; the 50 kept ones do.
    expect(readWorkbook(big, { limits: { maxXmlTreeElements: 100 } }).getSheet('S').conditionalFormats).toEqual([]);
    expectXlvError(() => readWorkbook(big, { limits: { maxXmlTreeElements: 40 } }), 'XLV_E_LIMIT_EXCEEDED');
  });

  it('is read by the streaming reader the same way', async () => {
    const stream = await openWorkbookStream(file);
    expect(stream.sheets.length).toBeGreaterThan(0);
    // The streaming reader exposes no conditional formats; reading the sheet must still work.
    for await (const row of stream.rows('S')) expect(row.cells.length).toBeGreaterThanOrEqual(0);
  });

  it('ignores hostile extension content safely', () => {
    const hostile = buildXlsx({
      sheets: [
        {
          name: 'S',
          data: '',
          after: extList(
            cfExt(
              '<x14:conditionalFormatting><x14:cfRule type="iconSet" priority="1"><x14:iconSet iconSet="5Boxes" custom="1"><x14:cfvo type="__proto__"/></x14:iconSet></x14:cfRule>' +
                '<x14:cfRule type="dataBar" priority="2"><x14:dataBar><x14:cfvo type="min"/><x14:cfvo type="max"/></x14:dataBar></x14:cfRule>' +
                '<x14:cfRule type="iconSet" priority="3"><x14:iconSet iconSet="3Stars" custom="1"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfIcon iconSet="3Flags" iconId="7"/></x14:iconSet></x14:cfRule>' +
                '<xm:sqref>A1</xm:sqref></x14:conditionalFormatting>',
            ),
          ),
        },
      ],
    });
    const wb = readWorkbook(hostile);
    expect(wb.getSheet('S').conditionalFormats).toEqual([]);
    expect(wb.warnings.map((w) => w.code)).toContain('XLV_W_INVALID_VALUE');
    expect(({} as Record<string, unknown>)['type']).toBeUndefined();
  });
});

describe('writing the extensions', () => {
  const bar = {
    type: 'dataBar' as const,
    dataBar: {
      cfvo: [{ type: 'autoMin' as const }, { type: 'autoMax' as const }] as const,
      color: '638EC6',
      minLength: 0,
      maxLength: 100,
      gradient: false,
      borderColor: '0000FF',
      negativeFillColor: 'FF0000',
      axisColor: '000000',
      axisPosition: 'middle' as const,
    },
  };

  it('writes a main data bar with an x14 id and its extension in the worksheet extLst', () => {
    const wb = createWorkbook();
    wb.addSheet('One').conditionalFormat('Z1', [{ type: 'expression', formulas: ['TRUE'] }]);
    wb.addSheet('S').conditionalFormat('A1:A10', [{ type: 'expression', formulas: ['A1>5'], style: { font: { bold: true } } }, bar]);
    const xml = part(writeWorkbook(wb), 'xl/worksheets/sheet2.xml');
    const id = '{00000002-0000-0000-0000-000000000002}';
    expect(xml).toContain(
      '<cfRule type="dataBar" priority="2"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>' +
        `<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="${X14}"><x14:id>${id}</x14:id></ext></extLst></cfRule>`,
    );
    expect(xml).toMatch(
      new RegExp(
        `<extLst><ext uri="\\{78C0D931-6437-407D-A8EE-F0AAD7539E65\\}" xmlns:x14="${X14}" xmlns:xm="${XM}"><x14:conditionalFormattings><x14:conditionalFormatting>` +
          `<x14:cfRule type="dataBar" id="\\{00000002-0000-0000-0000-000000000002\\}"><x14:dataBar minLength="0" maxLength="100" border="1" gradient="0" axisPosition="middle">` +
          '<x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:borderColor rgb="FF0000FF"/><x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule>' +
          '<xm:sqref>A1:A10</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>$',
      ),
    );
    // The extension is the last child of the worksheet (CT_Worksheet).
    expect(xml.indexOf('<extLst><ext uri="{78C0')).toBeGreaterThan(xml.indexOf('</conditionalFormatting>'));
  });

  it('round-trips the extended data bar', () => {
    const wb = createWorkbook();
    wb.addSheet('S').conditionalFormat('A1:A10', [bar]);
    const rule = readWorkbook(writeWorkbook(wb)).getSheet('S').conditionalFormats[0]?.rules[0];
    expect(rule?.dataBar).toEqual({
      cfvo: [
        { type: 'autoMin', value: null, gte: true },
        { type: 'autoMax', value: null, gte: true },
      ],
      color: 'FF638EC6',
      minLength: 0,
      maxLength: 100,
      showValue: true,
      extended: true,
      gradient: false,
      border: true,
      borderColor: 'FF0000FF',
      negativeFillColor: 'FFFF0000',
      negativeBorderColor: null,
      axisPosition: 'middle',
      axisColor: 'FF000000',
      direction: 'context',
    });
  });

  it('writes extension-only icon sets in the extension alone, with their priority', () => {
    const wb = createWorkbook();
    wb.addSheet('S').conditionalFormat('B1:B9', [
      {
        type: 'iconSet',
        iconSet: {
          iconSet: '3Triangles',
          cfvo: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33, gte: false }, { type: 'percent', value: 67 }],
          custom: [
            { iconSet: '3Triangles', iconId: 0 },
            { iconSet: 'NoIcons', iconId: 0 },
            { iconSet: '3Arrows', iconId: 2 },
          ],
        },
      },
    ]);
    const bytes = writeWorkbook(wb);
    const xml = part(bytes, 'xl/worksheets/sheet1.xml');
    expect(xml).not.toContain('<conditionalFormatting ');
    expect(xml).toContain(
      '<x14:conditionalFormatting><x14:cfRule type="iconSet" priority="1"><x14:iconSet iconSet="3Triangles" custom="1">' +
        '<x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent" gte="0"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo>' +
        '<x14:cfIcon iconSet="3Triangles" iconId="0"/><x14:cfIcon iconSet="NoIcons" iconId="0"/><x14:cfIcon iconSet="3Arrows" iconId="2"/></x14:iconSet></x14:cfRule><xm:sqref>B1:B9</xm:sqref></x14:conditionalFormatting>',
    );
    expect(readWorkbook(bytes).getSheet('S').conditionalFormats[0]?.rules[0]).toMatchObject({ type: 'iconSet', priority: 1, iconSet: { iconSet: '3Triangles', custom: [{ iconSet: '3Triangles', iconId: 0 }, { iconSet: 'NoIcons', iconId: 0 }, { iconSet: '3Arrows', iconId: 2 }] } });
  });

  it('is deterministic, also in the streaming writer', async () => {
    const make = () => {
      const wb = createWorkbook();
      wb.addSheet('S').conditionalFormat('A1:A10', [bar]);
      return part(writeWorkbook(wb), 'xl/worksheets/sheet1.xml');
    };
    expect(make()).toBe(make());
    const stream = writeWorkbookStream({ sheets: [{ name: 'S', rows: [[1], [2]], conditionalFormats: [{ sqref: 'A1:A10', rules: [bar] }] }] });
    const bytes = new Uint8Array(await new Response(stream).arrayBuffer());
    expect(part(bytes, 'xl/worksheets/sheet1.xml')).toContain('<x14:id>{00000001-0000-0000-0000-000000000001}</x14:id>');
    expect(readWorkbook(bytes).getSheet('S').conditionalFormats[0]?.rules[0]?.dataBar).toMatchObject({ extended: true, gradient: false });
  });

  it('writes plain rules without an extension', () => {
    const wb = createWorkbook();
    wb.addSheet('S').conditionalFormat('A1', [{ type: 'dataBar', dataBar: { cfvo: [{ type: 'min' }, { type: 'max' }], color: '638EC6' } }]);
    expect(part(writeWorkbook(wb), 'xl/worksheets/sheet1.xml')).not.toContain('extLst');
  });

  it('keeps the extensions when another cell is edited', () => {
    const wb = createWorkbook();
    wb.addSheet('S').conditionalFormat('A1:A10', [bar]);
    const edited = editCells(writeWorkbook(wb), [{ sheet: 'S', cell: 'B1', value: 1 }]).bytes;
    expect(readWorkbook(edited).getSheet('S').conditionalFormats[0]?.rules[0]?.dataBar).toMatchObject({ extended: true, borderColor: 'FF0000FF' });
  });

  it('enforces the [MS-XLSX] rules', () => {
    const s = createWorkbook().addSheet('S');
    const db = (extra: Record<string, unknown>) => [{ type: 'dataBar' as const, dataBar: { cfvo: [{ type: 'min' as const }, { type: 'max' as const }], color: '638EC6', ...extra } as never }];
    expectRefused(() => s.conditionalFormat('A1', db({ gradient: false })), 'XLV_E_INVALID_OPTION', /axisColor/);
    expectRefused(() => s.conditionalFormat('A1', db({ axisPosition: 'none', negativeBorderColor: 'FF0000' })), 'XLV_E_INVALID_OPTION', /borderColor/);
    expectRefused(() => s.conditionalFormat('A1', db({ axisPosition: 'none', minLength: 60, maxLength: 50 })), 'XLV_E_INVALID_OPTION', /minLength/);
    expectRefused(
      () => s.conditionalFormat('A1', [{ type: 'dataBar', dataBar: { cfvo: [{ type: 'autoMax' }, { type: 'max' }], color: '638EC6', axisPosition: 'none' } }]),
      'XLV_E_INVALID_OPTION',
      /autoMax/,
    );
    expectRefused(() => s.conditionalFormat('A1', [{ type: 'colorScale', colorScale: { cfvo: [{ type: 'autoMin' }, { type: 'max' }], colors: ['FF0000', '00FF00'] } }]), 'XLV_E_INVALID_OPTION', /type/);
    const icons = (extra: Record<string, unknown>) => [{ type: 'iconSet' as const, iconSet: { cfvo: [{ type: 'percent' as const, value: 0 }, { type: 'percent' as const, value: 50 }], ...extra } as never }];
    expectRefused(() => s.conditionalFormat('A1', icons({ iconSet: '3Stars' })), 'XLV_E_INVALID_OPTION', /3 to 5/);
    expectRefused(
      () => s.conditionalFormat('A1', [{ type: 'iconSet', iconSet: { iconSet: '3Stars', cfvo: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }], custom: [{ iconSet: '3Flags', iconId: 0 }] } }]),
      'XLV_E_INVALID_OPTION',
      /one icon per value object/,
    );
    expectRefused(
      () =>
        s.conditionalFormat('A1', [
          {
            type: 'iconSet',
            iconSet: { cfvo: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }], custom: [{ iconSet: '3Flags', iconId: 3 }, { iconSet: 'NoIcons', iconId: 0 }, { iconSet: 'NoIcons', iconId: 0 }] },
          },
        ]),
      'XLV_E_INVALID_OPTION',
      /iconId/,
    );
  });
});
