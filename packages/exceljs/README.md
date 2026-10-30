# @xlvane/exceljs

The commonly used part of the [exceljs](https://github.com/exceljs/exceljs) 4.4.0 API, built on [xlvane](../xlvane). Code written for exceljs keeps working after one change in `package.json`:

```json
"dependencies": {
  "exceljs": "npm:@xlvane/exceljs@^1"
}
```

The package is not published to npm yet.

```js
import ExcelJS from 'exceljs'; // or: const ExcelJS = require('exceljs');

const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet('Sales');
sheet.columns = [
  { header: 'Region', key: 'region', width: 15 },
  { header: 'Revenue', key: 'revenue', width: 12 },
];
sheet.addRow({ region: 'North', revenue: 1200 });
await workbook.xlsx.writeFile('sales.xlsx');
```

- Requires Node.js 22.12 or later for `require()` (the package is an ES module; `require()` of ES modules works without a flag from 22.12). ESM `import` works the same way, and named imports such as `import { Workbook } from 'exceljs'` work too.
- Methods and options outside the supported subset throw an `XlvError` with code `XLV_E_NOT_SUPPORTED` and the name of the feature. Nothing is dropped silently.
- What is supported, partly supported and not supported, with the source of each exceljs behaviour: [docs/exceljs-compat.md](../../docs/exceljs-compat.md).

License: MIT
