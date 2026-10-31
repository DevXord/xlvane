// CommonJS code written for exceljs: require() of the package.
const ExcelJS = require('exceljs');

const dir = process.argv[2] ?? '.';
import('./report.mjs')
  .then(({ run }) => run(ExcelJS, ExcelJS.Workbook, dir))
  .then((summary) => {
    summary.requireType = typeof ExcelJS;
    console.log(JSON.stringify(summary));
  })
  .catch((error) => {
    console.error(error);
    process.exit(1);
  });
