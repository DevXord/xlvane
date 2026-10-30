# xlvane

Read, inspect and write XLSX files; input files are treated as untrusted. ESM only, one runtime dependency ([fflate](https://github.com/101arrowz/fflate)). The core works with `Uint8Array` and has no Node.js dependencies; file helpers are in `xlvane/node`.

Not published to npm yet. Streaming and the task-oriented API are planned (see [docs/PLAN.md](../../docs/PLAN.md)).

## Reading a workbook

```ts
import { readWorkbook } from 'xlvane';

const workbook = readWorkbook(bytes); // Uint8Array or ArrayBuffer
const sheet = workbook.getSheet('Sales');
console.log(sheet.cell('B2')?.value);
for (const row of sheet.rows()) {
  console.log(row.map((cell) => cell.value));
}
```

In Node.js:

```ts
import { readWorkbookFile } from 'xlvane/node';

const workbook = await readWorkbookFile('report.xlsx');
```

Cell values are numbers, strings, booleans, `Date`s (for date-formatted numbers and ISO date cells; the UTC fields of the `Date` hold the spreadsheet's date and time), error codes such as `"#DIV/0!"` (with `type: "error"`) or `null`. Formula cells carry `formula.text` and the cached result stored in the file; xlvane does not calculate formulas. Cells that have neither a value nor a formula are not stored.

## Writing a workbook

```ts
import { createWorkbook, writeWorkbook } from 'xlvane';

const workbook = createWorkbook(); // { dateSystem: '1904' } for the 1904 date system
const sheet = workbook.addSheet('Sales');
const header = { font: { bold: true }, fill: { color: 'D9E1F2' }, border: { bottom: { style: 'thin' } } } as const;
sheet.set('A1', 'Region', header).set('B1', 'Revenue', header).set('C1', 'Date', header);
sheet.set('A2', 'North').set('B2', 1250.5, { numFmt: '#,##0.00' }).set('C2', new Date(Date.UTC(2024, 0, 31)));
sheet.set('B3', { formula: 'SUM(B2:B2)', result: 1250.5 });
sheet.column('A', { width: 20 }).row(1, { height: 24 });
sheet.merge('A5:C5').freeze({ rows: 1 }).autoFilter('A1:C2');
const bytes = writeWorkbook(workbook); // Uint8Array
```

In Node.js, `writeWorkbookFile(workbook, 'sales.xlsx')` from `xlvane/node` writes the file; it never replaces an existing file unless you pass `{ overwrite: true }`.

- Values: numbers, strings, booleans, `Date`s (stored as serial numbers with a `yyyy-mm-dd` or `yyyy-mm-dd hh:mm:ss` format unless the style sets one), `{ error: '#N/A' }`, and `{ formula, result }`. xlvane writes the formula text and the cached result you give; it does not calculate formulas.
- Styles: font (name, size, bold, italic, underline, strike, colour), fill (solid colour or pattern), borders (top, right, bottom, left), alignment (horizontal, vertical, wrap, indent, rotation) and number format. Identical styles are stored once. `sheet.style(range, style)` merges a style into every cell of a range.
- Layout: column widths and hidden columns, row heights and hidden rows, merged ranges, frozen rows and columns, autofilter, hidden and very hidden sheets.
- Invalid input is rejected when it is set, with an `XlvError` saying what and where, for example `XLV_E_INVALID_SHEET_NAME` or `XLV_E_INVALID_STYLE` with the allowed values.
- The output is deterministic: the same workbook always gives the same bytes. The test suite opens written files that use all of these features in LibreOffice and checks the values, styles and layout it sees.

Reading returns the same information: `cell.style`, `sheet.columns`, `sheet.rowInfo(row)`, `sheet.freeze` and `sheet.autoFilter`.

## Inspecting a workbook

`inspect()` returns plain JSON data about the structure of a workbook: sheets and their visibility, used ranges, a detected header row, column types with example values, merged cells, formulas and warnings. The schema is documented in [docs/inspect-json.md](../../docs/inspect-json.md).

```ts
import { inspect } from 'xlvane';
import { inspectFile } from 'xlvane/node';

const report = inspect(bytes);          // or: await inspectFile('report.xlsx')
console.log(JSON.stringify(report, null, 2));
```

## Errors and warnings

Problems with input files throw `XlvError` with a stable `code` (for example `XLV_E_SHEET_NOT_FOUND`), a `location` (part, sheet, cell), a `hint` saying what to do next and a `docsUrl`. `error.toJSON()` returns the same as data. Recoverable oddities are reported as warnings in `workbook.warnings`. All codes are described in [docs/errors.md](../../docs/errors.md).

## Untrusted input and limits

Every file is treated as untrusted:

- ZIP: sizes are measured on the bytes actually decompressed, not taken from headers; overlapping entries, duplicate names, encrypted entries, unsupported compression methods and multi-disk archives are rejected; CRC-32 is verified.
- XML: an own parser that rejects DTDs (so no external entities or entity expansion), recognises only the predefined entities, and does not recurse.
- Data from the file never becomes a key of a plain object, so names like `__proto__` are harmless.
- Macros are never run and external links are never followed; both are reported.

Limits can be changed per call, for example `readWorkbook(bytes, { limits: { maxCells: 50_000_000 } })`. Use `Infinity` to switch a limit off.

| Limit | Default | Protects against |
|---|---|---|
| `maxInputBytes` | 256 MiB | oversized input |
| `maxEntries` | 10 000 | archives with huge numbers of entries |
| `maxEntryBytes` | 256 MiB | one part decompressing to a huge size |
| `maxTotalBytes` | 512 MiB | all parts together decompressing to a huge size |
| `maxCompressionRatio` | 200 (checked after `compressionRatioMinBytes`, 1 MiB) | zip bombs |
| `maxXmlDepth` | 128 | deeply nested XML |
| `maxXmlAttributes` | 256 | elements with huge numbers of attributes |
| `maxXmlNameLength` | 1 024 | huge element or attribute names |
| `maxXmlTextLength` | 8 Mi characters | huge text nodes or attribute values |
| `maxXmlTreeElements` | 1 000 000 | huge workbook, styles or relationship parts (elements plus attributes) |
| `maxCells` | 5 000 000 | workbooks too large to hold in memory (measured: about 180 bytes per numeric cell, Node.js 26) |

## Package size

**Size budget:** 70 kB of gzip-compressed JavaScript for `dist/*.js` (level 9, unminified as published; measured 58.4 kB on 2026-10-30). `npm run check:pkg` fails when the budget is exceeded. The only dependency, fflate, adds about 22 kB gzip (its `esm/browser.js`, measured the same way).

Not supported: password-protected workbooks and legacy `.xls` files (both rejected with `XLV_E_UNSUPPORTED_CONTAINER`).

License: MIT
