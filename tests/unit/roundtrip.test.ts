import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import type { CellStyle } from '../../packages/xlvane/src/style';

const DEFAULT_FONT = { name: 'Calibri', size: 11 };

describe('round trip: values', () => {
  const wb = createWorkbook();
  const s = wb.addSheet('Values');
  s.set('A1', 'text')
    .set('A2', '  spaces  ')
    .set('A3', 'line\nbreak\r\nand CR\r')
    .set('A4', 'zażółć 😀')
    .set('A5', '_x0041_ \u0001 ￿ <&>"\'')
    .set('A6', '')
    .set('B1', 0)
    .set('B2', -0.001)
    .set('B3', 1e20)
    .set('B4', 0.1 + 0.2)
    .set('B5', Number.MAX_VALUE)
    .set('B6', -5e-324)
    .set('C1', true)
    .set('C2', false)
    .set('D1', { error: '#DIV/0!' })
    .set('D2', { error: '#N/A' })
    .set('E1', new Date(Date.UTC(2024, 0, 31)))
    .set('E2', new Date(Date.UTC(2024, 1, 29, 13, 45, 30)))
    .set('E3', new Date(Date.UTC(1900, 0, 1)))
    .set('E4', new Date(Date.UTC(1900, 2, 1)))
    .set('E5', new Date(Date.UTC(9999, 11, 31)))
    .set('F1', { formula: 'B1+1', result: 1 })
    .set('F2', { formula: '=A1&"!"', result: 'text!' })
    .set('F3', { formula: 'B1>0', result: false })
    .set('F4', { formula: '1/0', result: { error: '#DIV/0!' } })
    .set('F5', { formula: 'TODAY()' })
    .set('F6', { formula: 'E1+1', result: new Date(Date.UTC(2024, 1, 1)) });
  const read = readWorkbook(writeWorkbook(wb)).getSheet('Values');
  const v = (ref: string) => read.cell(ref)?.value;
  const t = (ref: string) => read.cell(ref)?.type;

  it('keeps strings exactly', () => {
    expect(v('A1')).toBe('text');
    expect(v('A2')).toBe('  spaces  ');
    expect(v('A3')).toBe('line\nbreak\r\nand CR\r');
    expect(v('A4')).toBe('zażółć 😀');
    expect(v('A5')).toBe('_x0041_ \u0001 ￿ <&>"\'');
    expect(v('A6')).toBe('');
  });

  it('keeps numbers exactly', () => {
    expect(v('B1')).toBe(0);
    expect(v('B2')).toBe(-0.001);
    expect(v('B3')).toBe(1e20);
    expect(v('B4')).toBe(0.1 + 0.2);
    expect(v('B5')).toBe(Number.MAX_VALUE);
    expect(v('B6')).toBe(-5e-324);
  });

  it('keeps booleans and errors', () => {
    expect([v('C1'), t('C1'), v('C2')]).toEqual([true, 'boolean', false]);
    expect([v('D1'), t('D1'), v('D2')]).toEqual(['#DIV/0!', 'error', '#N/A']);
  });

  it('keeps dates, with a date number format', () => {
    expect(v('E1')).toEqual(new Date(Date.UTC(2024, 0, 31)));
    expect(v('E2')).toEqual(new Date(Date.UTC(2024, 1, 29, 13, 45, 30)));
    expect(v('E3')).toEqual(new Date(Date.UTC(1900, 0, 1)));
    expect(v('E4')).toEqual(new Date(Date.UTC(1900, 2, 1)));
    expect(v('E5')).toEqual(new Date(Date.UTC(9999, 11, 31)));
    expect(read.cell('E1')?.numFmt).toBe('yyyy-mm-dd');
    expect(read.cell('E2')?.numFmt).toBe('yyyy-mm-dd hh:mm:ss');
  });

  it('keeps formulas and cached results', () => {
    expect(read.cell('F1')?.formula?.text).toBe('B1+1');
    expect(v('F1')).toBe(1);
    expect(read.cell('F2')?.formula?.text).toBe('A1&"!"');
    expect(v('F2')).toBe('text!');
    expect(v('F3')).toBe(false);
    expect([v('F4'), t('F4')]).toEqual(['#DIV/0!', 'error']);
    expect(read.cell('F5')?.formula?.text).toBe('TODAY()');
    expect(v('F5')).toBeNull();
    expect(v('F6')).toEqual(new Date(Date.UTC(2024, 1, 1)));
  });
});

describe('round trip: 1904 date system', () => {
  it('keeps dates and the date system', () => {
    const wb = createWorkbook({ dateSystem: '1904' });
    wb.addSheet('S').set('A1', new Date(Date.UTC(1904, 0, 1))).set('A2', new Date(Date.UTC(2030, 5, 15, 6)));
    const read = readWorkbook(writeWorkbook(wb));
    expect(read.dateSystem).toBe('1904');
    expect(read.getSheet('S').cell('A1')?.value).toEqual(new Date(Date.UTC(1904, 0, 1)));
    expect(read.getSheet('S').cell('A2')?.value).toEqual(new Date(Date.UTC(2030, 5, 15, 6)));
  });
});

describe('round trip: sheets', () => {
  it('keeps names, order and visibility', () => {
    const wb = createWorkbook();
    wb.addSheet('First & <second>');
    wb.addSheet('Hidden', { state: 'hidden' });
    wb.addSheet('Very', { state: 'veryHidden' });
    wb.addSheet("It's 😀");
    const read = readWorkbook(writeWorkbook(wb));
    expect(read.sheets.map((s) => [s.name, s.state])).toEqual([
      ['First & <second>', 'visible'],
      ['Hidden', 'hidden'],
      ['Very', 'veryHidden'],
      ["It's 😀", 'visible'],
    ]);
  });
});

describe('round trip: styles', () => {
  const styles: Record<string, CellStyle> = {
    A1: { font: { bold: true, italic: true, underline: 'double', strike: true, color: 'FF0000', size: 14, name: 'Arial' } },
    A2: { font: { underline: 'single' } },
    A3: { fill: { color: '00FF00' } },
    A4: { fill: { pattern: 'darkGrid', fgColor: '80112233', bgColor: 'FFFFFF' } },
    A5: {
      border: {
        top: { style: 'thin' },
        right: { style: 'medium', color: '0000FF' },
        bottom: { style: 'double' },
        left: { style: 'dashDot', color: 'FF123456' },
      },
    },
    A6: { alignment: { horizontal: 'center', vertical: 'top', wrapText: true, indent: 2, textRotation: 90 } },
    A7: { numFmt: '#,##0.00' },
    A8: { numFmt: '0.00%' },
    A9: { numFmt: '[Red]"neg "0;"pos "0' },
  };
  const wb = createWorkbook();
  const s = wb.addSheet('S');
  for (const [ref, style] of Object.entries(styles)) s.set(ref, 1.5, style);
  s.set('B1', 2);
  const read = readWorkbook(writeWorkbook(wb)).getSheet('S');

  it('reads back fonts, completed with the default font', () => {
    expect(read.cell('A1')?.style.font).toEqual({
      name: 'Arial',
      size: 14,
      bold: true,
      italic: true,
      underline: 'double',
      strike: true,
      color: 'FFFF0000',
    });
    expect(read.cell('A2')?.style.font).toEqual({ ...DEFAULT_FONT, underline: 'single' });
  });

  it('reads back fills with ARGB colours', () => {
    expect(read.cell('A3')?.style.fill).toEqual({ pattern: 'solid', fgColor: 'FF00FF00' });
    expect(read.cell('A4')?.style.fill).toEqual({ pattern: 'darkGrid', fgColor: '80112233', bgColor: 'FFFFFFFF' });
  });

  it('reads back borders and alignment', () => {
    expect(read.cell('A5')?.style.border).toEqual({
      top: { style: 'thin' },
      right: { style: 'medium', color: 'FF0000FF' },
      bottom: { style: 'double' },
      left: { style: 'dashDot', color: 'FF123456' },
    });
    expect(read.cell('A6')?.style.alignment).toEqual(styles.A6!.alignment);
  });

  it('reads back number formats, reusing built-in ids', () => {
    expect(read.cell('A7')?.numFmt).toBe('#,##0.00');
    expect(read.cell('A7')?.style.numFmt).toBe('#,##0.00');
    expect(read.cell('A8')?.numFmt).toBe('0.00%');
    expect(read.cell('A9')?.numFmt).toBe('[Red]"neg "0;"pos "0');
  });

  it('gives unstyled cells an empty style', () => {
    expect(read.cell('B1')?.style).toEqual({});
    expect(read.cell('B1')?.numFmt).toBeNull();
  });
});

describe('round trip: layout', () => {
  const wb = createWorkbook();
  const s = wb.addSheet('S');
  s.set('A1', 'h').set('C5', 1);
  s.column('A', { width: 25.5 }).column(3, { hidden: true }).column('D', { width: 8, hidden: true });
  s.row(1, { height: 30 }).row(4, { hidden: true });
  s.merge('A10:C12').merge('E1:F1');
  s.freeze({ rows: 1, cols: 2 });
  s.autoFilter('A1:C5');
  const second = wb.addSheet('T');
  second.set('A1', 1).freeze({ rows: 3 });
  const read = readWorkbook(writeWorkbook(wb));
  const rs = read.getSheet('S');

  it('keeps column widths and hidden columns', () => {
    expect(rs.columns).toEqual([
      { min: 1, max: 1, width: 25.5, hidden: false },
      { min: 3, max: 3, width: null, hidden: true },
      { min: 4, max: 4, width: 8, hidden: true },
    ]);
  });

  it('keeps row heights and hidden rows, even for rows without cells', () => {
    expect(rs.rowInfo(1)).toEqual({ height: 30, hidden: false });
    expect(rs.rowInfo(4)).toEqual({ height: null, hidden: true });
    expect(rs.rowInfo(2)).toBeUndefined();
  });

  it('keeps merges, frozen panes and the autofilter', () => {
    expect(rs.merges).toEqual(['A10:C12', 'E1:F1']);
    expect(rs.freeze).toEqual({ rows: 1, cols: 2 });
    expect(rs.autoFilter).toBe('A1:C5');
    expect(read.getSheet('T').freeze).toEqual({ rows: 3, cols: 0 });
    expect(read.getSheet('T').autoFilter).toBeNull();
  });
});
