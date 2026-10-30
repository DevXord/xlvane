# xlvane

Read and inspect XLSX files from untrusted sources. ESM only, one runtime dependency ([fflate](https://github.com/101arrowz/fflate)). The core works with `Uint8Array` and has no Node.js dependencies; file helpers are in `xlvane/node`.

Not published to npm yet. Writing workbooks is planned (see [docs/PLAN.md](../../docs/PLAN.md)).

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

Not supported: password-protected workbooks and legacy `.xls` files (both rejected with `XLV_E_UNSUPPORTED_CONTAINER`).

License: MIT
