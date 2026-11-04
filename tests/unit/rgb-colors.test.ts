// CT_Color@rgb is ST_UnsignedIntHex, 4 bytes = 8 hex digits in ARGB order (W2-16). Some writers
// put 6 digits (RRGGBB) there, e.g. Apache POI test file 52348.xlsx (`<color rgb="000000"/>`,
// `<fgColor rgb="808080"/>`). The reader takes exactly 6 hex digits as an opaque RGB colour
// (alpha FF, as the writers do for 6-digit input) and still warns with XLV_W_INVALID_VALUE.
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { DEFAULT_READ_LIMITS } from '../../packages/xlvane/src/limits';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { SML_NS } from '../../packages/xlvane/src/sml/namespaces';
import { parseStyles } from '../../packages/xlvane/src/sml/styles';
import { WarningCollector } from '../../packages/xlvane/src/util/warnings';
import { buildXlsx } from '../helpers/xlsx-builder';

function parse(content: string) {
  const warnings = new WarningCollector();
  const styles = parseStyles(`<styleSheet xmlns="${SML_NS.transitional}">${content}</styleSheet>`, { part: '/xl/styles.xml', limits: DEFAULT_READ_LIMITS, warnings });
  return { styles, warnings: warnings.list() };
}

const PARTS =
  '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="10"/><color rgb="000000"/><name val="Arial"/></font></fonts>' +
  '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="808080"/><bgColor rgb="ff00aa"/></patternFill></fill></fills>' +
  '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF0000"/></left><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1"/></cellXfs>' +
  '<dxfs count="1"><dxf><font><color rgb="00FF00"/></font></dxf></dxfs>';

describe('six-digit rgb colours', () => {
  it('are read as opaque RGB in fonts, fills, borders and dxfs, with a warning', () => {
    const { styles, warnings } = parse(PARTS);
    expect(styles.cellStyle(1, null)).toEqual({
      font: { name: 'Arial', size: 10, color: 'FF000000' },
      fill: { pattern: 'solid', fgColor: 'FF808080', bgColor: 'FFFF00AA' },
      border: { left: { style: 'thin', color: 'FFFF0000' } },
    });
    expect(styles.dxf(0)).toEqual({ font: { color: 'FF00FF00' } });
    expect(warnings).toMatchObject([{ code: 'XLV_W_INVALID_VALUE', count: 5 }]);
    expect(warnings[0]?.message).toBe('The font colour "000000" is not an ARGB value; it was read as the opaque RGB colour "FF000000".');
  });

  it('other lengths and non-hex values are still ignored with a warning', () => {
    for (const bad of ['0000000', 'FFFFF', 'GG0000', '', '0x0000', ' 000000', 'FF0000FF00']) {
      const { styles, warnings } = parse(`<fonts count="1"><font><color rgb="${bad}"/></font></fonts><cellXfs count="1"><xf numFmtId="0" fontId="0"/></cellXfs>`);
      expect(styles.defaultFont(), bad).toEqual({});
      expect(warnings, bad).toMatchObject([{ code: 'XLV_W_INVALID_VALUE', count: 1 }]);
      expect(warnings[0]?.message, bad).toContain('it is ignored');
    }
  });

  it('eight digits are read without a warning', () => {
    const { styles, warnings } = parse('<fonts count="1"><font><color rgb="80123abc"/></font></fonts>');
    expect(styles.defaultFont()).toEqual({ color: '80123ABC' });
    expect(warnings).toEqual([]);
  });

  it('are read in rich text runs and the sheet tab colour', () => {
    const sst =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      `<sst xmlns="${SML_NS.transitional}" count="1" uniqueCount="1"><si><r><rPr><color rgb="0000FF"/></rPr><t>blue</t></r><r><t>plain</t></r></si></sst>`;
    const wb = readWorkbook(
      buildXlsx({
        sharedStrings: sst,
        sheets: [{ name: 'S', before: '<sheetPr><tabColor rgb="00B050"/></sheetPr>', data: '<row r="1"><c r="A1" t="s"><v>0</v></c></row>' }],
      }),
    );
    const s = wb.getSheet('S');
    expect(s.cell('A1')?.richText).toEqual([{ text: 'blue', font: { color: 'FF0000FF' } }, { text: 'plain' }]);
    expect(s.tabColor).toBe('FF00B050');
    expect(wb.warnings.map((w) => w.code)).toEqual(['XLV_W_INVALID_VALUE', 'XLV_W_INVALID_VALUE']);
  });

  it('reads the colours of Apache POI 52348.xlsx', () => {
    const wb = readWorkbook(new Uint8Array(readFileSync(new URL('../corpus/poi/52348.xlsx', import.meta.url))));
    const styles = [...wb.getSheet('balance').cells(), ...wb.getSheet('balance').styledCells()].map((c) => c.style);
    const fills = new Set(styles.map((s) => (s.fill && 'fgColor' in s.fill ? s.fill.fgColor : undefined)).filter((c) => c !== undefined));
    const fontColors = new Set(styles.map((s) => s.font?.color).filter((c) => c !== undefined));
    expect(fills.has('FF808080')).toBe(true);
    expect(fontColors.has('FF000000')).toBe(true);
    expect(wb.warnings.filter((w) => w.code === 'XLV_W_INVALID_VALUE').every((w) => w.message.includes('opaque RGB') || !w.message.includes('ARGB'))).toBe(true);
  });
});
