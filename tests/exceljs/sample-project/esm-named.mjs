// Named imports. exceljs 4.4.0 rejects them ("Named export 'Workbook' not found ... CommonJS
// module"); @xlvane/exceljs is an ES module and offers them in addition. Not part of the
// cross-check against exceljs.
import ExcelJS, { Workbook, ValueType, stream } from 'exceljs';
import { run } from './report.mjs';

const dir = process.argv[2] ?? '.';
const summary = await run(ExcelJS, Workbook, dir);
summary.namedImportsAreSame = Workbook === ExcelJS.Workbook && ValueType === ExcelJS.ValueType && stream === ExcelJS.stream;
console.log(JSON.stringify(summary));
