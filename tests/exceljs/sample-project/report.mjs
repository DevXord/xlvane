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
  const sheet = workbook.addWorksheet('Sales', { views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }] });
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
  };
}
