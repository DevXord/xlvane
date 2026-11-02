# @xlvane/compat

The [exceljs](https://github.com/exceljs/exceljs) 4.4.0 API, built on [xlvane](https://www.npmjs.com/package/xlvane), with the differences listed below and in [docs/exceljs-compat.md](https://github.com/DevXord/xlvane/blob/main/docs/exceljs-compat.md). Code written for exceljs keeps working after one change in `package.json`:

```json
"dependencies": {
  "exceljs": "npm:@xlvane/compat@^1"
}
```

npm installs the package under the name `exceljs` (an npm alias), so `import ExcelJS from 'exceljs'` and `require('exceljs')` load `@xlvane/compat` and the code stays unchanged.

## Installing

In a project that uses exceljs, either install the alias:

```sh
npm install exceljs@npm:@xlvane/compat@^1
```

or let the `xlvane` command line prepare the project:

```sh
npx xlvane migrate              # report: the change to package.json and places in the code to check
npx xlvane migrate --force      # change package.json, then run npm install
```

`xlvane migrate` changes only `package.json`, and only with `--force`. It scans the files that mention exceljs for API outside the supported subset; the matching is textual, so a finding is a place to look at.

## Example

```js
import ExcelJS from 'exceljs';

const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet('Sales');
sheet.columns = [
  { header: 'Region', key: 'region', width: 15 },
  { header: 'Revenue', key: 'revenue', width: 12 },
];
sheet.addRow({ region: 'North', revenue: 1200 });
await workbook.xlsx.writeFile('sales.xlsx');
```

In CommonJS, `const ExcelJS = require('exceljs');` gives the same object; call `writeFile` inside an async function there.

## Requirements

- Node.js 22.12 or later (`engines.node`). The package is an ES module; `require()` of ES modules works without a flag from Node.js 22.12. Named imports such as `import { Workbook } from 'exceljs'` work too.
- The same test bundle runs in Node.js 22, 24 and 26, Bun, Deno, Chromium and the Cloudflare Workers runtime (workerd). Methods that read or write files (`xlsx.readFile`, `xlsx.writeFile`, `csv.readFile`, `csv.writeFile`, the `filename` forms of the streaming classes and of `addImage`) work in Node.js, Bun and Deno; in browsers and Workers they throw `XLV_E_NOT_SUPPORTED`, and a bundler leaves the file helpers out.
- Dependencies: `xlvane` only, which depends on `fflate` only.
- TypeScript: the type names of exceljs's `index.d.ts` are exported (`ExcelJS.Worksheet`, `ExcelJS.Fill`, …) and work with `moduleResolution` `bundler` and `nodenext` and with `module: commonjs`. In a CommonJS project, TypeScript refuses `import ExcelJS = require('exceljs')` of an ES module under `module: node16` (TS1471); use `nodenext` or `commonjs`.

## Differences from exceljs

The full list, with the exceljs source of each behaviour, is in [docs/exceljs-compat.md](https://github.com/DevXord/xlvane/blob/main/docs/exceljs-compat.md) (also in `node_modules/xlvane/docs/exceljs-compat.md` after installing). The ones most likely to matter:

- **Not offered:** row and column operations on a streamed worksheet (`insertRow(s)`, `spliceRows`, `duplicateRow`, `spliceColumns`; exceljs's streaming worksheet does not have them either) throw an `XlvError` with code `XLV_E_NOT_SUPPORTED` and the name of the method. Deep imports of exceljs internals (`exceljs/lib/...`) do not resolve; `exceljs/dist/es5` and the `exceljs/dist/exceljs*.js` files do. Nothing is dropped silently.
- **CSV:** text that starts with `=`, `+`, `-`, `@`, a tab, CR or LF is written with a leading `'`, so spreadsheet applications do not run it as a formula. `formulaInjection: "off"` writes it unchanged, as exceljs does.
- **Sheet ids:** `worksheet.id` is written as `sheetId` and kept on loading, as in exceljs; an id outside 1 to 65534 is written as the lowest unused one, and the streaming `WorkbookReader` numbers worksheets 1, 2, 3… in file order.
- **Views:** only the first view of a worksheet is written; a frozen view always gets the cell right of and below the panes as `topLeftCell`.
- **Notes:** `protection.locked` and `lockText` read back as `'True'`/`'False'`, the values Excel writes.
- **Module object:** it is frozen, so adding or replacing its properties (monkey-patching) has no effect.
- **Streaming writer speed:** `stream.xlsx.WorkbookWriter` compresses in JavaScript (fflate) and took about twice as long as exceljs, which uses Node.js zlib (300 000 rows: 6.1 s against 2.8 s).
- **Input files** are treated as untrusted, as in xlvane: ZIP and XML limits, no DTDs or external entities. A file over a limit is rejected with an `XlvError` (`XLV_E_*` code, location and hint; `node_modules/xlvane/docs/errors.md`).

## Links

- Changes in each version: [CHANGELOG.md](https://github.com/DevXord/xlvane/blob/main/CHANGELOG.md)
- Bugs and questions: [GitHub issues](https://github.com/DevXord/xlvane/issues)
- Reporting a vulnerability: [SECURITY.md](https://github.com/DevXord/xlvane/blob/main/SECURITY.md)

## License

MIT. @xlvane/compat is an independent project. It is not affiliated with or endorsed by the exceljs project or its authors, and it contains no exceljs code; its API and type declarations follow exceljs 4.4.0, whose MIT license notice is included in LICENSE.
