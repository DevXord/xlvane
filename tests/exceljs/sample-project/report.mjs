// Shared by esm.mjs and cjs.cjs through dynamic import: builds the same workbook with whatever
// exceljs module object it is given, so the scripts differ only in how they load exceljs.
// Plain exceljs 4.4.0 usage, following the exceljs README.

/** A 283×168 PNG (exceljs v4.4.0 spec/integration/data/image.png, MIT). */
const PNG_BASE64 = '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';

/** Writes report.xlsx in `dir` with the in-memory API and stream.xlsx.WorkbookWriter, then reads both back. */
export async function run(ExcelJS, Workbook, dir) {
  const path = `${dir}/report.xlsx`;
  const streamPath = `${dir}/streamed.xlsx`;

  // --- write with the document API ---
  const workbook = new Workbook();
  workbook.creator = 'Sample';
  workbook.lastModifiedBy = 'Sample editor';
  workbook.title = 'Sales report';
  workbook.company = 'Sample Ltd';
  workbook.manager = 'Sample manager';
  workbook.created = new Date(Date.UTC(2024, 0, 1, 9, 30));
  workbook.modified = new Date(Date.UTC(2024, 1, 2, 17, 0));
  const sheet = workbook.addWorksheet('Sales', { views: [{ state: 'frozen', xSplit: 0, ySplit: 1, zoomScale: 120 }], properties: { tabColor: { argb: 'FF00B050' } } });
  sheet.columns = [
    { header: 'Id', key: 'id', width: 8 },
    { header: 'Product', key: 'product', width: 24 },
    { header: 'Amount', key: 'amount', width: 12 },
    { header: 'Date', key: 'date', width: 14 },
  ];
  sheet.addRow({ id: 1, product: 'Tea', amount: 12.5, date: new Date(Date.UTC(2024, 0, 31)) });
  sheet.addRow({ id: 2, product: 'Coffee', amount: 20, date: new Date(Date.UTC(2024, 1, 1)) });
  sheet.addRow([3, 'Cocoa', 7.25, new Date(Date.UTC(2024, 1, 2))]);
  sheet.getCell('C5').value = { formula: 'SUM(C2:C4)', result: 39.75 };
  sheet.getCell('A5').value = 'Total';

  const header = sheet.getRow(1);
  header.font = { bold: true };
  header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
  header.border = { bottom: { style: 'thin' } };
  sheet.getColumn('amount').numFmt = '#,##0.00';
  sheet.getCell('B2').alignment = { horizontal: 'center', wrapText: true };
  sheet.mergeCells('A7:D7');
  sheet.getCell('A7').value = 'Merged note';
  sheet.autoFilter = 'A1:D4';
  sheet.getCell('A9').value = {
    richText: [{ font: { bold: true }, text: 'Note: ' }, { font: { italic: true, color: { argb: 'FFFF0000' } }, text: 'prices' }, { text: ' exclude tax' }],
  };
  sheet.getColumn(6).font = { name: 'Arial', size: 10, italic: true };
  sheet.getCell('F2').value = 'later';
  sheet.getCell('C2').name = 'firstAmount';
  workbook.definedNames.add('Sales!$C$2:$C$4', 'amounts');
  sheet.getCell('A10').value = { text: 'Price list', hyperlink: 'https://example.com/prices?year=2024&q=1' };
  sheet.getCell('A11').value = { text: 'Mail us', hyperlink: 'mailto:sales@example.com' };
  sheet.getCell('A12').value = { text: 'Go to total', hyperlink: '#Sales!C5' };
  sheet.getCell('A5').note = 'Checked by finance';
  sheet.getCell('B4').note = { texts: [{ font: { bold: true }, text: 'Cocoa' }, { text: ' is new this year' }] };
  // Data validation, conditional formatting, protection and printing (no password: xlvane
  // does not hash passwords yet).
  for (const a of ['B2', 'B3', 'B4']) sheet.getCell(a).dataValidation = { type: 'list', allowBlank: true, formulae: ['"Tea,Coffee,Cocoa"'] };
  sheet.getCell('C2').dataValidation = { type: 'decimal', operator: 'between', showErrorMessage: true, error: 'Use 0 to 1000', formulae: [0, 1000] };
  sheet.addConditionalFormatting({
    ref: 'C2:C4',
    rules: [
      { type: 'cellIs', operator: 'greaterThan', priority: 1, formulae: [15], style: { font: { bold: true } } },
      { type: 'expression', priority: 2, formulae: ['$C2<10'], style: { font: { italic: true } } },
    ],
  });
  // An x14 data bar (solid fill, axis in the middle) and an x14 icon set.
  sheet.addConditionalFormatting({
    ref: 'C2:C4',
    rules: [
      { type: 'dataBar', priority: 3, gradient: false, axisPosition: 'middle', axisColor: { argb: 'FF000000' }, direction: 'leftToRight', cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: 'FF638EC6' } },
      { type: 'iconSet', priority: 4, iconSet: '3Stars', cfvo: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }] },
    ],
  });
  await sheet.protect('', { sort: true, autoFilter: true, selectLockedCells: true, selectUnlockedCells: true });
  sheet.pageSetup.orientation = 'landscape';
  sheet.pageSetup.paperSize = 9;
  sheet.pageSetup.margins = { left: 0.5, right: 0.5, top: 1, bottom: 1, header: 0.4, footer: 0.4 };
  sheet.pageSetup.printArea = 'A1:D7';
  sheet.pageSetup.printTitlesRow = '1:1';
  sheet.headerFooter.oddHeader = '&CSales report';
  sheet.headerFooter.oddFooter = '&RPage &P of &N';
  // Row and column edits on plain values (no formulas, merges, notes or names: exceljs 4.4.0 does
  // not move those, xlvane does, docs/exceljs-compat.md).
  const edits = workbook.addWorksheet('Edits');
  for (let n = 1; n <= 6; n++) edits.addRow([`r${n}`, n]);
  edits.getRow(2).font = { bold: true };
  edits.insertRow(3, ['inserted', 0], 'i');
  edits.spliceRows(5, 1);
  edits.duplicateRow(1, 2, true);
  edits.spliceColumns(2, 0, ['c1', 'c2', 'c3']);
  edits.insertRows(2, [['a', 'b'], ['c', 'd']]);
  edits.getRow(1).splice(1, 1);
  // A table with a totals row and pictures (exceljs README "Tables" and "Images").
  const media = workbook.addWorksheet('Media');
  media.addTable({
    name: 'Fruit',
    ref: 'B2',
    totalsRow: true,
    style: { theme: 'TableStyleLight9', showRowStripes: true },
    columns: [{ name: 'Name', totalsRowLabel: 'Total' }, { name: 'Count', totalsRowFunction: 'sum' }],
    rows: [['Apple', 3], ['Pear', 5]],
  });
  const logo = workbook.addImage({ base64: `data:image/png;base64,${PNG_BASE64}`, extension: 'png' });
  media.addImage(logo, 'E2:G8');
  media.addImage(logo, { tl: { col: 8, row: 1 }, ext: { width: 120, height: 60 } });
  media.addBackgroundImage(logo);
  const extra = workbook.addWorksheet('Extra', { properties: { outlineLevelRow: 2 } });
  extra.getRow(1).values = ['angle', 'path', 'diagonal', 'rotated'];
  extra.getCell('A1').fill = { type: 'gradient', gradient: 'angle', degree: 90, stops: [{ position: 0, color: { argb: 'FFFFFFFF' } }, { position: 1, color: { argb: 'FF4472C4' } }] };
  extra.getCell('B1').fill = { type: 'gradient', gradient: 'path', center: { left: 0.5, top: 0.5 }, stops: [{ position: 0, color: { argb: 'FFFF0000' } }, { position: 1, color: { argb: 'FF00FF00' } }] };
  extra.getCell('C1').border = { diagonal: { up: true, down: false, style: 'thin', color: { argb: 'FF000000' } } };
  extra.getCell('D1').alignment = { textRotation: -45, shrinkToFit: true, readingOrder: 'rtl' };
  extra.getCell('D1').protection = { locked: false };
  extra.fillFormula('A3:B4', 'ROW()*COLUMN()', [[3, 6], [4, 8]]);
  extra.getRow(3).outlineLevel = 1;
  extra.getRow(4).outlineLevel = 2;
  extra.getColumn(2).outlineLevel = 1;
  workbook.views = [{ x: 10, y: 20, width: 15000, height: 9000, firstSheet: 0, activeTab: 0, visibility: 'visible' }];
  await workbook.xlsx.writeFile(path);

  // --- read it back ---
  const read = new Workbook();
  await read.xlsx.readFile(path);
  const ws = read.getWorksheet('Sales');
  const rows = [];
  ws.eachRow((row, rowNumber) => {
    const cells = [];
    row.eachCell((cell, colNumber) => {
      const v = cell.value;
      cells.push([colNumber, v instanceof Date ? v.toISOString() : v]);
    });
    rows.push([rowNumber, cells]);
  });
  const view = ws.views[0] ?? {};

  // --- streaming writer ---
  const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: streamPath, useStyles: true, useSharedStrings: false });
  const streamed = writer.addWorksheet('Log');
  streamed.addRow(['n', 'square']).commit();
  for (let n = 1; n <= 1000; n++) streamed.addRow([n, n * n]).commit();
  streamed.commit();
  await writer.commit();
  const readStreamed = new Workbook();
  await readStreamed.xlsx.readFile(streamPath);
  const log = readStreamed.getWorksheet('Log');

  const readEdits = read.getWorksheet('Edits');
  const editRows = [];
  const editBold = [];
  readEdits.eachRow((row, rowNumber) => {
    editRows.push([rowNumber, row.values.slice(1)]);
    row.eachCell((cell) => {
      if (cell.font?.bold) editBold.push(cell.address);
    });
  });
  // --- CSV (workbook.csv) ---
  const csvBook = new Workbook();
  const csvSheet = csvBook.addWorksheet('Csv');
  csvSheet.addRow(['Name', 'Amount', 'Paid', 'When', 'Note']);
  csvSheet.addRow(['Tea, green', 12.5, true, new Date(Date.UTC(2024, 0, 31, 8, 15)), 'said "hi"']);
  csvSheet.addRow(['Coffee', -3, false, null, { error: '#N/A' }]);
  csvSheet.getCell('A5').value = { text: 'site', hyperlink: 'https://example.com/a' };
  csvSheet.getCell('B5').value = { formula: 'B2*2', result: 25 };
  const csvDefault = (await csvBook.csv.writeBuffer()).toString();
  const csvOptions = { dateFormat: 'DD/MM/YYYY HH:mm', dateUTC: true, includeEmptyRows: false, formatterOptions: { delimiter: ';', rowDelimiter: '\r\n', includeEndRowDelimiter: true } };
  const csvCustom = (await csvBook.csv.writeBuffer(csvOptions)).toString();
  const csvPath = `${dir}/data.csv`;
  await csvBook.csv.writeFile(csvPath);
  const csvRead = new Workbook();
  const csvWs = await csvRead.csv.readFile(csvPath, { sheetName: 'Imported' });
  const csvRows = [];
  csvWs.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    csvRows.push([rowNumber, row.values.slice(1).map((v) => (v instanceof Date ? v.toISOString() : v))]);
  });

  const readMedia = read.getWorksheet('Media');
  const table = readMedia.getTables()[0];
  const tableModel = table.model ?? table.table;
  const readExtra = read.getWorksheet('Extra');
  const extraOut = {
    a1fill: readExtra.getCell('A1').fill,
    b1fill: readExtra.getCell('B1').fill,
    c1border: readExtra.getCell('C1').border,
    d1alignment: readExtra.getCell('D1').alignment,
    d1protection: readExtra.getCell('D1').protection ?? null,
    formulas: ['A3', 'B3', 'A4', 'B4'].map((a) => [a, readExtra.getCell(a).value, readExtra.getCell(a).formula ?? null, readExtra.getCell(a).formulaType]),
    outline: { row3: readExtra.getRow(3).outlineLevel, row4: readExtra.getRow(4).outlineLevel, row4collapsed: readExtra.getRow(4).collapsed, colB: readExtra.getColumn(2).outlineLevel, outlineLevelRow: readExtra.properties.outlineLevelRow ?? null },
    bookViews: read.views,
  };
  return {
    extra: extraOut,
    csv: { default: csvDefault, custom: csvCustom, sheet: csvWs.name, rows: csvRows },
    media: {
      table: { name: tableModel.name, tableRef: tableModel.tableRef, headerRow: tableModel.headerRow, totalsRow: tableModel.totalsRow, columns: tableModel.columns.map((c) => c.name), b4: readMedia.getCell('B4').value, c5: readMedia.getCell('C5').value, b5: readMedia.getCell('B5').value },
      images: readMedia.getImages().map((i) => ({ tl: [i.range.tl.nativeCol, i.range.tl.nativeRow], br: i.range.br ? [i.range.br.nativeCol, i.range.br.nativeRow] : null, ext: i.range.ext ?? null, bytes: read.getImage(i.imageId).buffer.length })),
      background: read.getImage(readMedia.getBackgroundImageId())?.buffer.length ?? null,
    },
    sheetNames: read.worksheets.map((w) => w.name),
    rowCount: ws.rowCount,
    rows,
    b2: ws.getCell('B2').value,
    c5: ws.getCell('C5').value,
    c2numFmt: ws.getCell('C2').numFmt,
    a1font: { bold: ws.getCell('A1').font?.bold ?? false },
    a1fill: ws.getCell('A1').fill,
    a1borderBottom: ws.getCell('A1').border?.bottom?.style ?? null,
    b2alignment: { horizontal: ws.getCell('B2').alignment?.horizontal ?? null, wrapText: ws.getCell('B2').alignment?.wrapText ?? null },
    merged: { a7: ws.getCell('A7').value, d7isMerged: ws.getCell('D7').isMerged, d7master: ws.getCell('D7').master.address },
    frozen: { state: view.state ?? null, ySplit: view.ySplit ?? null },
    autoFilter: ws.autoFilter ?? null,
    widthB: ws.getColumn(2).width,
    valueTypes: { number: ExcelJS.ValueType.Number, formula: ExcelJS.ValueType.Formula, c5type: ws.getCell('C5').type },
    streamed: { rowCount: log.rowCount, last: log.getCell('B1001').value },
    properties: { creator: read.creator, lastModifiedBy: read.lastModifiedBy, title: read.title, created: read.created?.toISOString(), modified: read.modified?.toISOString() },
    a9: ws.getCell('A9').value,
    a9text: ws.getCell('A9').text,
    a9type: ws.getCell('A9').type,
    d2numFmt: ws.getCell('D2').numFmt ?? null,
    columnF: { font: ws.getColumn(6).font ?? null, width: ws.getColumn(6).width ?? null, f2font: ws.getCell('F2').font ?? null },
    names: { model: read.definedNames.model, c2: ws.getCell('C2').names, c3: ws.getCell('C3').name ?? null },
    company: { company: read.company ?? null, manager: read.manager ?? null },
    tabColor: ws.properties.tabColor ?? null,
    view: { state: view.state ?? null, zoomScale: view.zoomScale ?? null, showGridLines: view.showGridLines ?? null, topLeftCell: view.topLeftCell ?? null },
    notes: { a5: ws.getCell('A5').note ?? null, b4texts: ws.getCell('B4').note?.texts ?? null, a1: ws.getCell('A1').note ?? null },
    edits: { rows: editRows, bold: editBold },
    rules: {
      b3validation: ws.getCell('B3').dataValidation ?? null,
      c2validation: ws.getCell('C2').dataValidation ?? null,
      conditional: ws.conditionalFormattings.map((cf) => ({
        ref: cf.ref,
        rules: cf.rules.map((r) => ({
          type: r.type,
          operator: r.operator ?? null,
          priority: r.priority,
          formulae: r.formulae ?? null,
          bold: r.style?.font?.bold ?? null,
          italic: r.style?.font?.italic ?? null,
          bar: r.type === 'dataBar' ? { gradient: r.gradient, minLength: r.minLength, maxLength: r.maxLength, axisPosition: r.axisPosition, axisColor: r.axisColor, direction: r.direction, color: r.color, cfvo: r.cfvo } : null,
          icons: r.type === 'iconSet' ? { iconSet: r.iconSet, cfvo: r.cfvo } : null,
        })),
      })),
      protection: ws.sheetProtection ?? null,
      page: {
        orientation: ws.pageSetup.orientation ?? null,
        paperSize: ws.pageSetup.paperSize ?? null,
        margins: ws.pageSetup.margins ?? null,
        printArea: ws.pageSetup.printArea ?? null,
        printTitlesRow: ws.pageSetup.printTitlesRow ?? null,
      },
      header: ws.headerFooter?.oddHeader ?? null,
      footer: ws.headerFooter?.oddFooter ?? null,
      names: read.definedNames.model.map((d) => d.name),
    },
    links: {
      a10: ws.getCell('A10').value,
      a10type: ws.getCell('A10').type,
      a10text: ws.getCell('A10').text,
      a11hyperlink: ws.getCell('A11').hyperlink ?? null,
      a12: ws.getCell('A12').value,
    },
  };
}
