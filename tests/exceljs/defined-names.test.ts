// Ported from exceljs v4.4.0 spec/unit/doc/defined-names.spec.js (MIT, Copyright (c) 2014-2019
// Guyon Roche), changes: Mocha/Chai → Vitest; `verquire('doc/defined-names')` → the
// `DefinedNames` class of @xlvane/exceljs (exceljs exposes the same object as
// `workbook.definedNames`). Two tests added for the read and write path of workbook names.
import { describe, expect, it } from 'vitest';
import { DefinedNames } from '../../packages/exceljs/src/defined-names';
import { ExcelJS, roundTrip } from './helpers';

describe('DefinedNames', () => {
  it('adds names for cells', () => {
    const dn = new DefinedNames();

    dn.add('blort!A1', 'foo');
    expect(dn.getNames('blort!A1')).toEqual(['foo']);
    expect(dn.getNames('blort!$A$1')).toEqual(['foo']);

    dn.add('blort!$B$4', 'bar');
    expect(dn.getNames('blort!B4')).toEqual(['bar']);
    expect(dn.getNames('blort!$B$4')).toEqual(['bar']);

    dn.add("'blo rt'!$B$4", 'bar');
    expect(dn.getNames("'blo rt'!$B$4")).toEqual(['bar']);
    dn.add("'blo ,!rt'!$B$4", 'bar');
    expect(dn.getNames("'blo ,!rt'!$B$4")).toEqual(['bar']);
  });

  it('removes names for cells', () => {
    const dn = new DefinedNames();

    dn.add('blort!A1', 'foo');
    dn.add('blort!A1', 'bar');
    dn.remove('blort!A1', 'foo');

    expect(dn.getNames('blort!A1')).toEqual(['bar']);
  });

  it('gets the right ranges for a name', () => {
    const dn = new DefinedNames();

    dn.add('blort!A1', 'vertical');
    dn.add('blort!A2', 'vertical');
    dn.add('blort!A3', 'vertical');

    dn.add('blort!C1', 'horizontal');
    dn.add('blort!D1', 'horizontal');
    dn.add('blort!E1', 'horizontal');

    dn.add('blort!C3', 'square');
    dn.add('blort!D3', 'square');
    dn.add('blort!C4', 'square');
    dn.add('blort!D4', 'square');

    dn.add('other!A1', 'single');

    expect(dn.getRanges('vertical')).toEqual({ name: 'vertical', ranges: ['blort!$A$1:$A$3'] });
    expect(dn.getRanges('horizontal')).toEqual({ name: 'horizontal', ranges: ['blort!$C$1:$E$1'] });
    expect(dn.getRanges('square')).toEqual({ name: 'square', ranges: ['blort!$C$3:$D$4'] });
    expect(dn.getRanges('single')).toEqual({ name: 'single', ranges: ['other!$A$1'] });
  });

  it('splices', () => {
    const dn = new DefinedNames();
    dn.add('vertical!A1', 'vertical');
    dn.add('vertical!A2', 'vertical');
    dn.add('vertical!A3', 'vertical');
    dn.add('vertical!A4', 'vertical');

    dn.add('horizontal!A1', 'horizontal');
    dn.add('horizontal!B1', 'horizontal');
    dn.add('horizontal!C1', 'horizontal');
    dn.add('horizontal!D1', 'horizontal');

    for (const col of ['A', 'B', 'C', 'D']) {
      for (const row of [1, 2, 3, 4]) dn.add(`square!${col}${String(row)}`, 'square');
    }

    dn.add('single!A1', 'singleA1');
    dn.add('single!D1', 'singleD1');
    dn.add('single!A4', 'singleA4');
    dn.add('single!D4', 'singleD4');

    dn.spliceRows('vertical', 2, 2, 1);
    dn.spliceColumns('horizontal', 2, 2, 1);
    dn.spliceRows('square', 2, 2, 1);
    dn.spliceColumns('square', 2, 2, 1);
    dn.spliceRows('single', 2, 2, 1);
    dn.spliceColumns('single', 2, 2, 1);

    expect(dn.getRanges('vertical')).toEqual({ name: 'vertical', ranges: ['vertical!$A$1', 'vertical!$A$3'] });
    expect(dn.getRanges('horizontal')).toEqual({ name: 'horizontal', ranges: ['horizontal!$A$1', 'horizontal!$C$1'] });
    expect(dn.getRanges('square')).toEqual({ name: 'square', ranges: ['square!$A$1', 'square!$C$1', 'square!$A$3', 'square!$C$3'] });
    expect(dn.getRanges('singleA1')).toEqual({ name: 'singleA1', ranges: ['single!$A$1'] });
    expect(dn.getRanges('singleD1')).toEqual({ name: 'singleD1', ranges: ['single!$C$1'] });
    expect(dn.getRanges('singleA4')).toEqual({ name: 'singleA4', ranges: ['single!$A$3'] });
    expect(dn.getRanges('singleD4')).toEqual({ name: 'singleD4', ranges: ['single!$C$3'] });
  });

  it('creates matrix from model', () => {
    const dn = new DefinedNames();

    dn.model = [];
    dn.add('blort!A1', 'bar');
    dn.remove('blort!A1', 'foo');

    expect(dn.getNames('blort!A1')).toEqual(['bar']);
  });

  it('skips values with invalid range', () => {
    const dn = new DefinedNames();
    dn.model = [
      { name: 'eq', ranges: ['"="'] },
      { name: 'ref', ranges: ['#REF!'] },
      { name: 'single', ranges: ['Sheet3!$A$1'] },
      { name: 'range', ranges: ['Sheet3!$A$2:$F$2228'] },
    ];

    expect(dn.model).toEqual([
      { name: 'single', ranges: ['Sheet3!$A$1'] },
      { name: 'range', ranges: ['Sheet3!$A$2:$F$2228'] },
    ]);
  });

  // Added for xlvane: the written names and what is read back.
  it('writes each name as one definedName with its ranges, quoting sheet names as exceljs', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('Data');
    wb.addWorksheet('My Sheet');
    wb.definedNames.add('Data!$A$1:$B$2', 'block');
    wb.definedNames.add("'My Sheet'!$C$3", 'block');
    const read = (await import('xlvane')).readWorkbook(await wb.xlsx.writeBuffer());
    expect(read.definedNames).toEqual([{ name: 'block', refersTo: "Data!$A$1:$B$2,'My Sheet'!$C$3", sheet: null, hidden: false, comment: null }]);
    const wb2 = await roundTrip(wb);
    expect(wb2.definedNames.model).toEqual([{ name: 'block', ranges: ['Data!$A$1:$B$2', "'My Sheet'!$C$3"] }]);
  });

  it('drops built-in names, scopes and non-range names on read, as exceljs', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Data');
    ws.getCell('A1').value = 1;
    ws.autoFilter = 'A1:B1';
    const { createWorkbook, writeWorkbook } = await import('xlvane');
    const book = createWorkbook();
    const sheet = book.addSheet('Data');
    sheet.set('A1', 1);
    sheet.autoFilter('A1:B1');
    book.defineName('local', 'Data!$A$1', { sheet: 'Data' });
    book.defineName('rate', '0.25');
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(writeWorkbook(book));
    expect(wb2.definedNames.model).toEqual([{ name: 'local', ranges: ['Data!$A$1'] }]);
    expect(wb2.getWorksheet('Data').getCell('A1').name).toBe('local');
  });
});
