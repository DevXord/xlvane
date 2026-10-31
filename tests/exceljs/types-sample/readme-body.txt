// Code from the exceljs 4.4.0 README (sections "Interface" … "Value Types"), with the type
// annotations a TypeScript user writes. tests/exceljs/types-compat.test.ts prepends an import of
// `ExcelJS` and type-checks this file under several module settings. It is not run.

export async function readmeSample(input: ExcelJS.Buffer): Promise<ExcelJS.Buffer> {
  const workbook: ExcelJS.Workbook = new ExcelJS.Workbook();
  workbook.creator = 'Me';
  workbook.lastModifiedBy = 'Her';
  workbook.created = new Date(1985, 8, 30);
  workbook.modified = new Date();
  workbook.lastPrinted = new Date(2016, 9, 27);
  workbook.properties.date1904 = true;
  workbook.calcProperties.fullCalcOnLoad = true;
  workbook.views = [{ x: 0, y: 0, width: 10000, height: 20000, firstSheet: 0, activeTab: 1, visibility: 'visible' }];

  const sheet: ExcelJS.Worksheet = workbook.addWorksheet('My Sheet', { properties: { tabColor: { argb: 'FFC0000' } } });
  workbook.addWorksheet('Grid', { views: [{ showGridLines: false }] });
  workbook.addWorksheet('Frozen', { views: [{ state: 'frozen', xSplit: 1, ySplit: 1 }] });
  workbook.addWorksheet('Header', { headerFooter: { firstHeader: 'Hello Exceljs', firstFooter: 'Hello World' } });
  const worksheet = workbook.addWorksheet('A4', { pageSetup: { paperSize: 9, orientation: 'landscape' } });
  const options: Partial<ExcelJS.AddWorksheetOptions> = { state: 'hidden' };
  workbook.addWorksheet('Hidden', options);
  const removed = workbook.addWorksheet('Removed');
  workbook.removeWorksheet(removed.id);

  workbook.eachSheet((ws: ExcelJS.Worksheet, sheetId: number) => {
    console.log(ws.name, sheetId);
  });
  const byName: ExcelJS.Worksheet | undefined = workbook.getWorksheet('My Sheet');
  const byId = workbook.getWorksheet(1);
  console.log(byName?.name, byId?.name, workbook.worksheets.length);
  sheet.state = 'veryHidden';
  const state: ExcelJS.WorksheetState = sheet.state;
  console.log(state);
  sheet.properties.outlineLevelCol = 2;
  const props: ExcelJS.WorksheetProperties = sheet.properties;
  console.log(props.defaultRowHeight, sheet.rowCount, sheet.actualRowCount, sheet.columnCount, sheet.actualColumnCount);

  // Page setup, headers and footers, views, auto filter
  const margins: ExcelJS.Margins = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };
  worksheet.pageSetup.margins = margins;
  worksheet.pageSetup.printArea = 'A1:G20';
  worksheet.pageSetup.printTitlesRow = '1:3';
  worksheet.pageSetup.printTitlesColumn = 'A:C';
  const paper: ExcelJS.PaperSize = ExcelJS.PaperSize.A4;
  worksheet.pageSetup.paperSize = paper;
  worksheet.headerFooter.differentFirst = true;
  worksheet.headerFooter.oddFooter = 'Page &P of &N';
  const hf: Partial<ExcelJS.HeaderFooter> = worksheet.headerFooter;
  console.log(hf);
  worksheet.views = [{ state: 'frozen', xSplit: 2, ySplit: 3, topLeftCell: 'G10', activeCell: 'A1' }];
  worksheet.views = [{ state: 'split', xSplit: 2000, ySplit: 3000, topLeftCell: 'G10', activeCell: 'A1' }];
  const view: Partial<ExcelJS.WorksheetView> = { state: 'normal', zoomScale: 120 };
  worksheet.views = [view];
  worksheet.autoFilter = 'A1:C1';
  worksheet.autoFilter = { from: 'A1', to: 'C1' };
  worksheet.autoFilter = { from: { row: 3, column: 1 }, to: { row: 5, column: 12 } };
  const filter: ExcelJS.AutoFilter = { from: 'D3', to: { row: 7, column: 5 } };
  worksheet.autoFilter = filter;

  // Columns
  sheet.columns = [
    { header: 'Id', key: 'id', width: 10 },
    { header: 'Name', key: 'name', width: 32 },
    { header: 'D.O.B.', key: 'DOB', width: 10, outlineLevel: 1 },
  ];
  const idCol: ExcelJS.Column = sheet.getColumn('id');
  const nameCol = sheet.getColumn('B');
  const dobCol = sheet.getColumn(3);
  dobCol.header = 'Date of Birth';
  dobCol.header = ['Date of Birth', 'A.K.A. D.O.B.'];
  dobCol.key = 'dob';
  dobCol.width = 15;
  dobCol.hidden = true;
  sheet.getColumn(4).outlineLevel = 0;
  console.log(sheet.getColumn(4).collapsed, idCol.number, nameCol.letter);
  dobCol.eachCell((cell: ExcelJS.Cell, rowNumber: number) => {
    console.log(cell.address, rowNumber);
  });
  dobCol.eachCell({ includeEmpty: true }, (cell, rowNumber) => console.log(cell.value, rowNumber));
  idCol.values = [1, 2, 3, 5, 8];
  nameCol.values = [, 'a', 'b', 'c'];
  sheet.spliceColumns(3, 2);
  sheet.spliceColumns(3, 1, ['new column values']);

  // Rows
  const row: ExcelJS.Row = sheet.getRow(5);
  const rows: ExcelJS.Row[] | undefined = sheet.getRows(5, 2);
  const lastRow: ExcelJS.Row | undefined = sheet.lastRow;
  console.log(rows?.length, lastRow?.number);
  row.height = 42.5;
  row.hidden = true;
  sheet.getRow(4).outlineLevel = 0;
  row.getCell(1).value = 5;
  row.getCell('name').value = 'Zeb';
  row.getCell('C').value = new Date();
  row.values = [1, 2, 3];
  row.values = [, 10, , 20];
  row.values = { id: 13, name: 'Thing 1', dob: new Date() };
  row.splice(3, 2);
  row.splice(4, 1, 'new value 1', 'new value 2');
  const cellCount: number = row.cellCount;
  const numValues: number = row.actualCellCount;
  console.log(cellCount, numValues);
  sheet.eachRow((r: ExcelJS.Row, rowNumber: number) => {
    console.log(`Row ${rowNumber} = ${JSON.stringify(r.values)}`);
  });
  sheet.eachRow({ includeEmpty: true }, (r, rowNumber) => console.log(r.number, rowNumber));
  row.eachCell((cell, colNumber) => console.log(`Cell ${colNumber} = ${String(cell.value)}`));
  row.eachCell({ includeEmpty: true }, (cell, colNumber) => console.log(cell.address, colNumber));
  row.addPageBreak();
  row.commit();
  sheet.addRow({ id: 1, name: 'John Doe', dob: new Date(1970, 1, 1) });
  sheet.addRow([3, 'Sam', new Date()]);
  const rowValues: ExcelJS.RowValues = [];
  rowValues[1] = 4;
  rowValues[5] = 'Kyle';
  sheet.addRow(rowValues);
  sheet.addRow([3, 'Sam', new Date()], 'i');
  const newRows: ExcelJS.Row[] = sheet.addRows([
    [5, 'Bob', new Date()],
    { id: 6, name: 'Barbara', dob: new Date() },
  ]);
  console.log(newRows.length);
  sheet.insertRow(1, { id: 1, name: 'John Doe', dob: new Date(1970, 1, 1) });
  sheet.insertRows(1, [[5, 'Bob', new Date()]], 'o');
  sheet.spliceRows(4, 3);
  sheet.spliceRows(4, 1, [1, 2, 3]);
  sheet.duplicateRow(1, 2, true);

  // Cells
  const cell: ExcelJS.Cell = sheet.getCell('C3');
  cell.value = new Date(1968, 5, 1);
  const text: string = cell.text;
  const type: ExcelJS.ValueType = cell.type;
  console.log(text, type === ExcelJS.ValueType.Date, cell.effectiveType);
  sheet.mergeCells('A4:B5');
  sheet.mergeCells('K10', 'M12');
  sheet.mergeCells(10, 11, 12, 13);
  sheet.unMergeCells('A4');
  sheet.getCell('A1').name = 'PI';
  sheet.getCell('A1').names = ['thing1', 'thing2'];
  sheet.getCell('A1').removeName('thing1');

  // Data validations
  const dv: ExcelJS.DataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: ['"One,Two,Three,Four"'],
  };
  sheet.getCell('A1').dataValidation = dv;
  sheet.getCell('A2').dataValidation = {
    type: 'whole',
    operator: 'notEqual',
    showErrorMessage: true,
    formulae: [5],
    errorStyle: 'error',
    errorTitle: 'Five',
    error: 'The value must not be Five',
  };
  sheet.getCell('A3').dataValidation = {
    type: 'decimal',
    operator: 'between',
    allowBlank: true,
    showInputMessage: true,
    formulae: [1.5, 7],
    promptTitle: 'Decimal',
    prompt: 'The value must between 1.5 and 7',
  };
  sheet.getCell('A4').dataValidation = { type: 'date', operator: 'lessThan', showErrorMessage: true, allowBlank: true, formulae: [new Date(2016, 0, 1)] };
  const op: ExcelJS.DataValidationOperator = 'greaterThan';
  sheet.getCell('A5').dataValidation = { type: 'textLength', operator: op, formulae: [15] };

  // Comments
  sheet.getCell('A1').note = 'Hello, ExcelJS!';
  const note: ExcelJS.Comment = {
    texts: [
      { font: { size: 12, color: { theme: 0 }, name: 'Calibri', family: 2, scheme: 'minor' }, text: 'This is ' },
      { font: { italic: true, size: 12, color: { theme: 0 }, name: 'Calibri', scheme: 'minor' }, text: 'a' },
    ],
    margins: { insetmode: 'custom', inset: [0.25, 0.25, 0.35, 0.35] },
    protection: { locked: 'True', lockText: 'False' },
    editAs: 'twoCells',
  };
  sheet.getCell('B1').note = note;

  // Tables
  const table: ExcelJS.Table = sheet.addTable({
    name: 'MyTable',
    ref: 'A1',
    headerRow: true,
    totalsRow: true,
    style: { theme: 'TableStyleDark3', showRowStripes: true },
    columns: [
      { name: 'Date', totalsRowLabel: 'Totals:', filterButton: true },
      { name: 'Amount', totalsRowFunction: 'sum', filterButton: false },
    ],
    rows: [
      [new Date('2019-07-20'), 70.1],
      [new Date('2019-07-21'), 70.6],
    ],
  });
  table.addRow([new Date('2019-08-01'), 42]);
  table.removeRows(0, 1);
  const column: ExcelJS.TableColumnProperties = table.getColumn(1) as unknown as ExcelJS.TableColumnProperties;
  console.log(column.name);
  table.commit();

  // Styles
  sheet.getCell('A1').numFmt = '0.00%';
  sheet.getColumn(3).numFmt = '"£"#,##0.00;[Red]\\-"£"#,##0.00';
  sheet.getRow(2).font = { name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', bold: true };
  const font: Partial<ExcelJS.Font> = { name: 'Arial Black', color: { argb: 'FF00FF00' }, family: 2, size: 14, italic: true };
  sheet.getCell('A2').font = font;
  sheet.getCell('A3').font = { vertAlign: 'superscript' };
  const alignment: Partial<ExcelJS.Alignment> = { vertical: 'middle', horizontal: 'center', wrapText: true };
  sheet.getCell('A4').alignment = alignment;
  sheet.getCell('A5').alignment = { textRotation: 'vertical', readingOrder: 'rtl', indent: 1 };
  const borders: Partial<ExcelJS.Borders> = {
    top: { style: 'thin' },
    left: { style: 'double', color: { argb: 'FF00FF00' } },
    bottom: { style: 'thick' },
    right: { style: 'thin' },
    diagonal: { up: true, down: true, style: 'thick', color: { argb: 'FFFF0000' } },
  };
  sheet.getCell('A6').border = borders;
  const solid: ExcelJS.Fill = { type: 'pattern', pattern: 'darkVertical', fgColor: { argb: 'FFFF0000' } };
  sheet.getCell('A7').fill = solid;
  const gradient: ExcelJS.FillGradientAngle = {
    type: 'gradient',
    gradient: 'angle',
    degree: 0,
    stops: [
      { position: 0, color: { argb: 'FF0000FF' } },
      { position: 0.5, color: { argb: 'FFFFFFFF' } },
      { position: 1, color: { argb: 'FF0000FF' } },
    ],
  };
  sheet.getCell('A8').fill = gradient;
  const path: ExcelJS.FillGradientPath = {
    type: 'gradient',
    gradient: 'path',
    center: { left: 0.5, top: 0.5 },
    stops: [
      { position: 0, color: { argb: 'FFFF0000' } },
      { position: 1, color: { argb: 'FF00FF00' } },
    ],
  };
  sheet.getCell('A9').fill = path;
  const style: Partial<ExcelJS.Style> = { font: { bold: true }, protection: { locked: false, hidden: true } };
  sheet.getCell('A10').style = style;
  sheet.getCell('A11').protection = { locked: false };
  const richText: ExcelJS.CellRichTextValue = {
    richText: [
      { text: 'This is ' },
      { font: { italic: true }, text: 'italic' },
    ],
  };
  sheet.getCell('A12').value = richText;
  const rt: ExcelJS.RichText = { text: 'plain' };
  console.log(rt.text);

  // Conditional formatting
  sheet.addConditionalFormatting({
    ref: 'A1:E7',
    rules: [
      {
        type: 'expression',
        priority: 1,
        formulae: ['MOD(ROW()+COLUMN(),2)=0'],
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF00FF00' } } },
      },
    ],
  });
  const cellIs: ExcelJS.CellIsRuleType = { type: 'cellIs', operator: 'greaterThan', formulae: [5], priority: 2, style: {} };
  const colorScale: ExcelJS.ColorScaleRuleType = {
    type: 'colorScale',
    priority: 3,
    cfvo: [{ type: 'min' }, { type: 'max' }],
    color: [{ argb: 'FFF8696B' }, { argb: 'FF63BE7B' }],
  };
  const rule: ExcelJS.ConditionalFormattingRule = cellIs;
  const cf: ExcelJS.ConditionalFormattingOptions = { ref: 'B1:B10', rules: [rule, colorScale] };
  sheet.addConditionalFormatting(cf);
  sheet.removeConditionalFormatting('');

  // Outline levels
  sheet.getColumn(3).outlineLevel = 1;
  sheet.getRow(3).outlineLevel = 1;
  sheet.properties.outlineProperties = { summaryBelow: false, summaryRight: false };

  // Images
  const imageId: number = workbook.addImage({ buffer: input, extension: 'png' });
  const imageId2 = workbook.addImage({ base64: 'iVBORw0KGgo=', extension: 'jpeg' });
  sheet.addBackgroundImage(imageId);
  sheet.addImage(imageId2, 'B2:D6');
  sheet.addImage(imageId, { tl: { col: 0, row: 0 }, ext: { width: 500, height: 200 } });
  const range: ExcelJS.ImagePosition = { tl: { col: 0, row: 0 }, ext: { width: 50, height: 20 } };
  sheet.addImage(imageId, range);
  const images = sheet.getImages();
  console.log(images.length);
  // Typed without `undefined` in exceljs (index.d.ts L1327, L1342-1346, L1365, L1780); found by
  // the outside-project tests before 1.0.0 (docs/decisions.md D-R09).
  console.log(sheet.getImages()[0].range.tl.nativeCol);
  const image: ExcelJS.Image = workbook.getImage(imageId);
  console.log(image.extension);
  sheet.getTable('MyTable').addRow([new Date('2019-07-22'), 70.9]);
  sheet.fillFormula('E1:E3', 'A1*2', (r, c) => r + c);
  sheet.fillFormula('F1:F3', 'A1*3', [3, 6, 9]);

  // Sheet protection
  await sheet.protect('the-password', { selectLockedCells: true, formatCells: true });
  const protection: Partial<ExcelJS.WorksheetProtection> = { sort: true };
  await sheet.protect('', protection);
  sheet.unprotect();

  // Value types
  sheet.getCell('B1').value = null;
  sheet.getCell('B2').value = 5;
  sheet.getCell('B3').value = 'Hello, World!';
  sheet.getCell('B4').value = new Date(2017, 2, 15);
  const hyperlink: ExcelJS.CellHyperlinkValue = { text: 'www.mylink.com', hyperlink: 'http://www.mylink.com', tooltip: 'www.mylink.com' };
  sheet.getCell('B5').value = hyperlink;
  sheet.getCell('B6').value = { text: 'Sheet2', hyperlink: '#\'Sheet2\'!A1' };
  const formula: ExcelJS.CellFormulaValue = { formula: 'A1+A2', result: 7, date1904: false };
  sheet.getCell('B7').value = formula;
  sheet.getCell('B8').value = { formula: 'A1+A2', result: undefined, date1904: false };
  const shared: ExcelJS.CellSharedFormulaValue = { sharedFormula: 'B7', result: 10, date1904: false };
  sheet.getCell('B9').value = shared;
  sheet.fillFormula('C1:C10', 'A1+B1', (r: number, c: number) => r * c);
  const formulaType: ExcelJS.FormulaType = sheet.getCell('C1').formulaType;
  console.log(formulaType === ExcelJS.FormulaType.Shared);
  sheet.getCell('B10').value = true;
  const err: ExcelJS.CellErrorValue = { error: '#N/A' };
  sheet.getCell('B11').value = err;
  const value: ExcelJS.CellValue = sheet.getCell('B11').value;
  console.log(value, ExcelJS.ErrorValue.NotApplicable);

  // Defined names
  workbook.definedNames.add('Sheet1!A1', 'name');
  console.log(workbook.definedNames.getNames('Sheet1!A1'));

  // File I/O
  await workbook.xlsx.load(input);
  const buffer: ExcelJS.Buffer = await workbook.xlsx.writeBuffer();
  const csv: ExcelJS.Buffer = await workbook.csv.writeBuffer({ sheetName: 'My Sheet' });
  console.log(csv.byteLength);
  const read: ExcelJS.Worksheet = await workbook.csv.readFile('in.csv');
  console.log(read.name);

  // Streaming I/O
  const writerOptions: Partial<ExcelJS.stream.xlsx.WorkbookWriterOptions> = { useStyles: true, useSharedStrings: true };
  const writer: ExcelJS.stream.xlsx.WorkbookWriter = new ExcelJS.stream.xlsx.WorkbookWriter(writerOptions);
  const streamSheet = writer.addWorksheet('Stream', { pageSetup: { fitToPage: true, fitToHeight: 5, fitToWidth: 7 } });
  streamSheet.addRow([1, 2, 3]).commit();
  await streamSheet.commit();
  await writer.commit();
  const readerOptions: Partial<ExcelJS.stream.xlsx.WorkbookStreamReaderOptions> = { worksheets: 'emit', sharedStrings: 'cache' };
  const reader: ExcelJS.stream.xlsx.WorkbookReader = new ExcelJS.stream.xlsx.WorkbookReader('in.xlsx', readerOptions);
  for await (const worksheetReader of reader) {
    for await (const r of worksheetReader) console.log(r.values);
  }

  // Enums of the module
  console.log(ExcelJS.RelationshipType.Hyperlink, ExcelJS.DocumentType.Xlsx, ExcelJS.ReadingOrder.RightToLeft);
  return buffer;
}
