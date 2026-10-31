# @xlvane/compat

The [exceljs](https://github.com/exceljs/exceljs) 4.4.0 API, built on [xlvane](https://www.npmjs.com/package/xlvane), with the differences listed in docs/exceljs-compat.md (link below). Code written for exceljs keeps working after one change in `package.json`:

```json
"dependencies": {
  "exceljs": "npm:@xlvane/compat@^1"
}
```

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
- The few exceljs features that are not offered (row and column operations on a streamed worksheet, which exceljs itself does not have, and deep imports of exceljs internals such as `exceljs/lib/...`) throw an `XlvError` with code `XLV_E_NOT_SUPPORTED` and the name of the feature, or fail to resolve. Nothing is dropped silently.
- What is supported, partly supported and not supported, with the source of each exceljs behaviour: [docs/exceljs-compat.md](https://github.com/DevXord/xlvane/blob/main/docs/exceljs-compat.md) (also in `node_modules/xlvane/docs/exceljs-compat.md` after installing). Errors are described in `node_modules/xlvane/docs/errors.md`.

@xlvane/compat is an independent project. It is not affiliated with or endorsed by the exceljs project or its authors, and it contains no exceljs code; its API and type declarations follow exceljs 4.4.0, whose MIT license notice is included in LICENSE.

License: MIT
