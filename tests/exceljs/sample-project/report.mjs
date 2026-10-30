// Shared by esm.mjs and cjs.cjs through dynamic import: builds the same workbook with whatever
// exceljs module object it is given, so the scripts differ only in how they load exceljs.
// Plain exceljs 4.4.0 usage, following the exceljs README.

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
  return {
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
