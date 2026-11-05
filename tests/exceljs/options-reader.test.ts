// xlvane tests for the xlsx read/write options, the WorkbookReader options and events, and the
// module enums of @xlvane/exceljs (docs/decisions.md D-E03; docs/facts/exceljs.md X4-25, X4-26).
// The exceljs behaviour they follow is cited from exceljs 4.4.0 (bench/node_modules/exceljs).
import { describe, expect, it } from 'vitest';
import { readWorkbook } from '../../packages/xlvane/src/read';
import { ExcelJS } from './helpers';

const sample = async (): Promise<Uint8Array> => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  ws.addRow(['a', 'b', 'c']);
  ws.addRow([1, 2, 3]);
  ws.getCell('A1').font = { bold: true };
  ws.getCell('B2').value = { text: 'site', hyperlink: 'https://example.com' };
  ws.getCell('C2').value = new Date(Date.UTC(2024, 0, 31));
  ws.mergeCells('A3:B3');
  ws.getCell('A3').value = 'merged';
  ws.getCell('A1').dataValidation = { type: 'list', formulae: ['"a,b"'] };
  ws.autoFilter = 'A1:C2';
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  return wb.xlsx.writeBuffer();
};

describe('xlsx read options', () => {
  it('ignoreNodes skips the named worksheet elements, as exceljs worksheet-xform.js L372', async () => {
    const bytes = await sample();
    const wb = await new ExcelJS.Workbook().xlsx.load(bytes, { ignoreNodes: ['mergeCells', 'hyperlinks', 'dataValidations', 'autoFilter', 'sheetViews'] });
    const ws = wb.getWorksheet('Data');
    expect(ws.getCell('B3').isMerged).toBe(false);
    expect(ws.getCell('B2').value).toBe('site');
    expect(ws.getCell('A1').dataValidation).toBeUndefined();
    expect(ws.autoFilter).toBeNull();
    expect(ws.views).toEqual([]);
    const all = await new ExcelJS.Workbook().xlsx.load(bytes, { ignoreNodes: ['sheetData', 'mergeCells'] });
    expect(all.getWorksheet('Data').rowCount).toBe(0);
  });

  it('maxRows and maxCols refuse bigger worksheets with XLV_E_LIMIT_EXCEEDED', async () => {
    const bytes = await sample();
    await expect(new ExcelJS.Workbook().xlsx.load(bytes, { maxRows: 2 })).rejects.toMatchObject({ code: 'XLV_E_LIMIT_EXCEEDED', details: { limit: 'maxRows' } });
    await expect(new ExcelJS.Workbook().xlsx.load(bytes, { maxCols: 2 })).rejects.toMatchObject({ code: 'XLV_E_LIMIT_EXCEEDED', details: { limit: 'maxCols' } });
    await expect(new ExcelJS.Workbook().xlsx.load(bytes, { maxRows: 3, maxCols: 3 })).resolves.toBeDefined();
    await expect(new ExcelJS.Workbook().xlsx.load(bytes, { maxRows: -1 })).rejects.toMatchObject({ code: 'XLV_E_INVALID_OPTION' });
  });

  it('base64 reads the file from base64 text, as exceljs xlsx.js L259', async () => {
    const bytes = await sample();
    const base64 = btoa(String.fromCharCode(...bytes));
    const wb = await new ExcelJS.Workbook().xlsx.load(base64, { base64: true });
    expect(wb.getWorksheet('Data').getCell('A2').value).toBe(1);
  });

  it('createInputStream throws as in exceljs 4', () => {
    expect(() => new ExcelJS.Workbook().xlsx.createInputStream()).toThrow(/deprecated/);
  });
});

describe('xlsx write options', () => {
  it('useStyles: false writes no styles but keeps the date format, as exceljs (styles-xform.js Mock)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'bold';
    ws.getCell('A1').font = { bold: true };
    ws.getCell('B1').value = new Date(Date.UTC(2024, 0, 31));
    ws.getColumn(3).font = { italic: true };
    const plain = readWorkbook(await wb.xlsx.writeBuffer({ useStyles: false })).getSheet('S');
    expect(plain.cell('A1')?.style.font?.bold).toBeUndefined();
    expect(plain.cell('B1')?.numFmt).toBe('mm-dd-yy');
    expect(plain.columns.find((c) => c.min <= 3 && c.max >= 3)?.style).toBeUndefined();
    const styled = readWorkbook(await wb.xlsx.writeBuffer()).getSheet('S');
    expect(styled.cell('A1')?.style.font?.bold).toBe(true);
  });

  it('accepts useSharedStrings and zip options without changing the content', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S').addRow(['x', 1]);
    const a = readWorkbook(await wb.xlsx.writeBuffer({ useSharedStrings: false, zip: { compression: 'DEFLATE', compressionOptions: { level: 9 } } }));
    expect(a.getSheet('S').cell('A1')?.value).toBe('x');
  });
});

describe('WorkbookReader options and events', () => {
  it('validates the options against WorkbookReader.Options and refuses what it cannot do', () => {
    expect(() => new ExcelJS.stream.xlsx.WorkbookReader(new Uint8Array(), { styles: 'emit' })).toThrow(expect.objectContaining({ code: 'XLV_E_INVALID_OPTION' }));
    expect(() => new ExcelJS.stream.xlsx.WorkbookReader(new Uint8Array(), { sharedStrings: 'emit' })).toThrow(expect.objectContaining({ code: 'XLV_E_NOT_SUPPORTED' }));
    expect(() => new ExcelJS.stream.xlsx.WorkbookReader(new Uint8Array(), { hyperlinks: 'emit' })).toThrow(expect.objectContaining({ code: 'XLV_E_NOT_SUPPORTED' }));
    expect(() => new ExcelJS.stream.xlsx.WorkbookReader(new Uint8Array(), { hyperlinks: 'cache', entries: 'emit', styles: 'cache' })).not.toThrow();
  });

  it('styles: "cache" gives the cells their styles; the default does not', async () => {
    const bytes = await sample();
    const fonts = async (options: Record<string, unknown>) => {
      const out: unknown[] = [];
      for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(bytes, options)) for await (const row of ws) if (row.number === 1) out.push(row.getCell(1).font?.bold ?? null);
      return out;
    };
    expect(await fonts({})).toEqual([null]);
    expect(await fonts({ styles: 'cache' })).toEqual([true]);
  });

  it('read() emits worksheet, row, end and finished events, and entry events with entries: "emit"', async () => {
    const bytes = await sample();
    const events: string[] = [];
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(bytes, { entries: 'emit' });
    reader.on('entry', (e: { type: string }) => events.push(`entry:${e.type}`));
    reader.on('worksheet', (ws: any) => {
      events.push(`worksheet:${String(ws.name)}:${String(ws.state)}`);
      ws.on('row', (row: any) => events.push(`row:${String(row.number)}`));
      ws.on('finished', () => events.push(`finished:${String(ws.dimensions.range)}`));
    });
    reader.on('end', () => events.push('end'));
    reader.on('finished', () => events.push('finished'));
    await reader.read();
    expect(events).toEqual(['entry:workbook', 'entry:worksheet', 'worksheet:Data:visible', 'row:1', 'row:2', 'row:3', 'finished:A1:C3', 'end', 'finished']);
  });

  it('emits error events, and throws an unhandled error like Node.js', async () => {
    const errors: unknown[] = [];
    const reader = new ExcelJS.stream.xlsx.WorkbookReader(new Uint8Array([1, 2, 3]), {});
    reader.on('error', (e: unknown) => errors.push(e));
    await reader.read();
    expect(errors).toHaveLength(1);
    await expect(new ExcelJS.stream.xlsx.WorkbookReader(new Uint8Array([1, 2, 3]), {}).read()).rejects.toBeDefined();
  });

  it('gives the worksheet reader exceljs members (columns, getColumn, column keys, destroy)', async () => {
    const bytes = await sample();
    for await (const ws of new ExcelJS.stream.xlsx.WorkbookReader(bytes, {})) {
      expect(ws.columns).toBeNull();
      expect(ws.getColumn(2).number).toBe(2);
      expect(ws.columns).toHaveLength(2);
      expect(ws.workbook).toBeInstanceOf(ExcelJS.stream.xlsx.WorkbookReader);
      expect(() => ws.destroy()).toThrow(/Invalid Operation/);
    }
  });
});

describe('module enums', () => {
  it('has RelationshipType, DocumentType and ReadingOrder with exceljs values (lib/doc/enums.js)', () => {
    expect(ExcelJS.RelationshipType).toEqual({ None: 0, OfficeDocument: 1, Worksheet: 2, CalcChain: 3, SharedStrings: 4, Styles: 5, Theme: 6, Hyperlink: 7 });
    expect(ExcelJS.DocumentType).toEqual({ Xlsx: 1 });
    expect(ExcelJS.ReadingOrder).toEqual({ LeftToRight: 1, RightToLeft: 2 });
  });
});
