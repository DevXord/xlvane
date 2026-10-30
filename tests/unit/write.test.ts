import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { createWorkbook, writeWorkbook } from '../../packages/xlvane/src/write';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { expectXlvError } from '../helpers/expect-error';

const dec = new TextDecoder();

function parts(bytes: Uint8Array): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, data] of Object.entries(unzipSync(bytes))) out[name] = dec.decode(data);
  return out;
}

describe('createWorkbook / addSheet', () => {
  it('adds sheets in order and finds them by name', () => {
    const wb = createWorkbook();
    wb.addSheet('One');
    wb.addSheet('Two', { state: 'hidden' });
    expect(wb.sheets.map((s) => s.name)).toEqual(['One', 'Two']);
    expect(wb.getSheet('Two').state).toBe('hidden');
    expect(wb.dateSystem).toBe('1900');
  });

  it('rejects invalid sheet names', () => {
    const wb = createWorkbook();
    for (const name of ['', 'a'.repeat(32), 'a/b', 'a\\b', 'a?', 'a*', 'a:b', '[x]', "'quoted", "quoted'"]) {
      expectXlvError(() => wb.addSheet(name), 'XLV_E_INVALID_SHEET_NAME');
    }
    expect(() => wb.addSheet('a'.repeat(31))).not.toThrow();
  });

  it('rejects duplicate sheet names ignoring case', () => {
    const wb = createWorkbook();
    wb.addSheet('Sales');
    expectXlvError(() => wb.addSheet('SALES'), 'XLV_E_DUPLICATE_SHEET_NAME');
  });

  it('suggests close names when a sheet is not found', () => {
    const wb = createWorkbook();
    wb.addSheet('Sales');
    const error = expectXlvError(() => wb.getSheet('Sale'), 'XLV_E_SHEET_NOT_FOUND');
    expect(error.hint).toContain('Sales');
  });

  it('rejects unknown options', () => {
    expectXlvError(() => createWorkbook({ dateSystem: '1901' as '1900' }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => createWorkbook({ datesystem: '1904' } as never), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => createWorkbook().addSheet('A', { state: 'gone' as 'hidden' }), 'XLV_E_INVALID_OPTION');
  });

  it('refuses to write a workbook without sheets', () => {
    expectXlvError(() => writeWorkbook(createWorkbook()), 'XLV_E_INVALID_INPUT');
  });
});

describe('NewSheet.set', () => {
  it('validates references', () => {
    const s = createWorkbook().addSheet('S');
    expectXlvError(() => s.set('A0', 1), 'XLV_E_INVALID_CELL_REF');
    expectXlvError(() => s.set('XFE1', 1), 'XLV_E_INVALID_CELL_REF');
    expect(() => s.set('xfd1048576', 1)).not.toThrow();
  });

  it('rejects values that a spreadsheet cannot hold', () => {
    const s = createWorkbook().addSheet('S');
    expectXlvError(() => s.set('A1', Number.NaN), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => s.set('A1', Infinity), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => s.set('A1', 'x'.repeat(32_768)), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => s.set('A1', 'bad \uD800 surrogate'), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => s.set('A1', new Date(Number.NaN)), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => s.set('A1', new Date(Date.UTC(1899, 0, 1))), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => s.set('A1', { error: '#OOPS' } as never), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => s.set('A1', { formula: '' }), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => s.set('A1', { formula: 'A2', result: Number.NaN }), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => s.set('A1', Symbol('x') as never), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => s.set('A1', { other: 1 } as never), 'XLV_E_INVALID_VALUE');
    expect(() => s.set('A1', 'x'.repeat(32_767))).not.toThrow();
  });

  it('rejects dates before 1904 in the 1904 date system', () => {
    const s = createWorkbook({ dateSystem: '1904' }).addSheet('S');
    expectXlvError(() => s.set('A1', new Date(Date.UTC(1903, 11, 31))), 'XLV_E_INVALID_VALUE');
  });

  it('keeps the current style when set without one, as spreadsheet applications do', () => {
    const wb = createWorkbook();
    wb.addSheet('S').set('A1', new Date(Date.UTC(2024, 0, 31))).set('A1', 45322);
    expect(readWorkbook(writeWorkbook(wb)).getSheet('S').cell('A1')?.value).toEqual(new Date(Date.UTC(2024, 0, 31)));
  });

  it('removes a cell when set to null without a style', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('S').set('A1', 1).set('A1', null);
    expect(s.cellCount).toBe(0);
    expect(readWorkbook(writeWorkbook(wb)).getSheet('S').cellCount).toBe(0);
  });
});

describe('NewSheet styles', () => {
  it('validates style objects', () => {
    const s = createWorkbook().addSheet('S');
    const bad: unknown[] = [
      { font: { size: 0 } },
      { font: { size: 410 } },
      { font: { name: '' } },
      { font: { color: 'red' } },
      { font: { underline: 'wavy' } },
      { fill: { color: '12345' } },
      { fill: { pattern: 'stripes' } },
      { border: { top: { style: 'fat' } } },
      { border: { diagonal: { style: 'fat' } } },
      { fill: { type: 'gradient', gradient: 'linear', stops: [] } },
      { font: { vertAlign: 'up' } },
      { font: { family: 6 } },
      { font: { charset: 256 } },
      { alignment: { readingOrder: 2 } },
      { protection: { locked: 'no' } },
      { alignment: { horizontal: 'middle' } },
      { alignment: { indent: -1 } },
      { alignment: { textRotation: 181 } },
      { numFmt: '' },
      { bold: true },
    ];
    for (const style of bad) expectXlvError(() => s.set('A1', 1, style as never), 'XLV_E_INVALID_STYLE');
  });

  it('applies a style to a range, merging with existing cell styles', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('S');
    s.set('A1', 'x', { font: { bold: true } });
    s.style('A1:B2', { fill: { color: 'FFFF00' } });
    const read = readWorkbook(writeWorkbook(wb)).getSheet('S');
    expect(read.cell('A1')?.style.font?.bold).toBe(true);
    expect(read.cell('A1')?.style.fill).toEqual({ pattern: 'solid', fgColor: 'FFFFFF00' });
  });

  it('refuses to style huge ranges', () => {
    const s = createWorkbook().addSheet('S');
    expectXlvError(() => s.style('A1:XFD1048576', { font: { bold: true } }), 'XLV_E_LIMIT_EXCEEDED');
  });

  it('writes each distinct style once', () => {
    const wb = createWorkbook();
    const s = wb.addSheet('S');
    for (let r = 1; r <= 50; r++) s.set(`A${String(r)}`, r, { font: { bold: true }, numFmt: '0.00' });
    const styles = parts(writeWorkbook(wb))['xl/styles.xml']!;
    expect(styles.match(/<xf /g)?.length).toBe(1 + 2); // cellStyleXfs default + default + bold
  });
});

describe('NewSheet layout', () => {
  it('validates columns, rows, merges, panes and filters', () => {
    const s = createWorkbook().addSheet('S');
    expectXlvError(() => s.column('XFE', { width: 10 }), 'XLV_E_INVALID_CELL_REF');
    expectXlvError(() => s.column(0, { width: 10 }), 'XLV_E_INVALID_CELL_REF');
    expectXlvError(() => s.column('A', { width: 256 }), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => s.row(0, { height: 10 }), 'XLV_E_INVALID_CELL_REF');
    expectXlvError(() => s.row(1, { height: 410 }), 'XLV_E_INVALID_VALUE');
    expectXlvError(() => s.merge('A1'), 'XLV_E_INVALID_CELL_REF');
    expectXlvError(() => s.freeze({ rows: -1 }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.freeze({ rows: 1_048_576 }), 'XLV_E_INVALID_OPTION');
    expectXlvError(() => s.autoFilter('nope'), 'XLV_E_INVALID_CELL_REF');
  });

  it('rejects overlapping merges', () => {
    const s = createWorkbook().addSheet('S');
    s.merge('A1:C3');
    expectXlvError(() => s.merge('C3:D4'), 'XLV_E_OVERLAPPING_MERGE');
    expect(() => s.merge('D4:E5')).not.toThrow();
  });
});

describe('writeWorkbook output', () => {
  const wb = createWorkbook();
  const s = wb.addSheet('Data');
  s.set('A1', 'Name').set('B1', 'Amount').set('A2', 'Alice').set('B2', 10).set('A3', 'Alice');
  wb.addSheet('Empty');
  const files = parts(writeWorkbook(wb));

  it('contains the expected parts', () => {
    expect(Object.keys(files).sort()).toEqual(
      [
        '[Content_Types].xml',
        '_rels/.rels',
        'xl/workbook.xml',
        'xl/_rels/workbook.xml.rels',
        'xl/styles.xml',
        'xl/sharedStrings.xml',
        'xl/worksheets/sheet1.xml',
        'xl/worksheets/sheet2.xml',
      ].sort(),
    );
  });

  it('stores repeated strings once in the shared string table', () => {
    expect(files['xl/sharedStrings.xml']).toContain('count="4" uniqueCount="3"');
  });

  it('is deterministic', () => {
    expect(writeWorkbook(wb)).toEqual(writeWorkbook(wb));
  });

  it('marks the first sheet as selected', () => {
    expect(files['xl/worksheets/sheet1.xml']).toContain('tabSelected="1"');
    expect(files['xl/worksheets/sheet2.xml']).not.toContain('tabSelected');
  });
});
