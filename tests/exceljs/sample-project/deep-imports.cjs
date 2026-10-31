// Deep imports that exceljs 4.4.0 documents (README: "exceljs/dist/es5" for older Node.js, and
// the browser bundles in dist/). @xlvane/compat maps them to its main entry (docs/exceljs-compat.md).
const paths = ['exceljs/dist/es5', 'exceljs/dist/es5/index.js', 'exceljs/dist/exceljs.js', 'exceljs/dist/exceljs.min.js'];
const out = {};
for (const p of paths) {
  const ExcelJS = require(p);
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('S').getCell('A1').value = 1;
  out[p] = { workbook: typeof ExcelJS.Workbook, value: wb.getWorksheet('S').getCell('A1').value, valueTypeNumber: ExcelJS.ValueType.Number };
}
console.log(JSON.stringify(out));
