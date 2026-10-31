// ESM code written for exceljs 4.4.0: the default import (exceljs is a CommonJS package, so
// this is the import form that works with it).
import ExcelJS from 'exceljs';
import { run } from './report.mjs';

const dir = process.argv[2] ?? '.';
const summary = await run(ExcelJS, ExcelJS.Workbook, dir);
console.log(JSON.stringify(summary));
