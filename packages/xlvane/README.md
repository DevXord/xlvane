# xlvane

Read, inspect and write XLSX files; input files are treated as untrusted. ESM only, one runtime dependency ([fflate](https://github.com/101arrowz/fflate)). The core works with `Uint8Array` and has no Node.js dependencies; file helpers are in `xlvane/node`.

```sh
npm install xlvane
```

The package also contains `docs/` (error codes, JSON schemas, exceljs compatibility), the Agent Skill `skills/xlvane/SKILL.md` and `llms-full.txt` (all of it in one file), for reading without network access.

## Tested environments

The core API (reading, writing, inspect, tables, CSV, editing, diff, streaming with Web Streams) and `@xlvane/exceljs` are tested in Node.js 22, 24 and 26, Bun, Deno, Chromium (Playwright) and the Cloudflare Workers runtime (workerd), with the same test bundle (`npm run test:runtimes`). Reading is tested on 340 files of the Apache POI test corpus (written by Excel for Windows, Mac and the web, LibreOffice, ONLYOFFICE and others, and damaged files) and on files saved by LibreOffice; see `tests/corpus/`.

## Command line in 5 minutes

The package installs an `xlvane` command (Node.js 22 or later): `npm install xlvane` in a project and run `npx xlvane`, or `npm install -g xlvane`.

```sh
xlvane                                    # what the tool does and the common commands
xlvane inspect report.xlsx                # sheets, used ranges, header rows, column types, warnings
xlvane read report.xlsx --sheet Sales     # a sheet as a table (first 50 rows; --limit, --range)
xlvane convert report.xlsx sales.csv      # XLSX to CSV or JSON, or CSV/JSON to XLSX
xlvane create data.json report.xlsx       # a formatted report from JSON or CSV
xlvane edit report.xlsx --set "Sales!B2=1250" -o updated.xlsx
xlvane diff report.xlsx updated.xlsx      # what changed; exit code 1 when they differ
xlvane validate report.xlsx --columns "Region:string,Revenue:number"
xlvane help edit                          # usage and examples of one command
```

- Nothing is overwritten without `--force`. Commands that write accept `--dry-run` and end with the path they wrote.
- Errors say what happened, where, and what to do next, often with a ready-to-run `Try:` command; typos get "Did you mean …?".
- `--json` prints one JSON document per command with a documented schema ([docs/cli-json.md](../../docs/cli-json.md)). Exit codes: 0 success, 1 a check found something (`diff`, `validate`), 2 usage error, 3 input or file error.
- Colour is used only on a terminal and never carries information alone; `NO_COLOR` turns it off.
- `xlvane migrate` prepares a project for `@xlvane/exceljs`.

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

### Rich text, document properties, defined names, column styles

```ts
import { createWorkbook, readWorkbook, writeWorkbook } from 'xlvane';

const workbook = createWorkbook({ properties: { title: 'Q3 sales', creator: 'Finance', created: new Date(Date.UTC(2026, 9, 1)) } });
const sheet = workbook.addSheet('Sales');
sheet.set('A1', { richText: [{ text: 'Total ', font: { bold: true } }, { text: '(EUR)', font: { color: '808080' } }] });
sheet.set('B1', 1250.5).set('C1', new Date(Date.UTC(2026, 9, 1)), { numFmt: 14 });
sheet.column('B', { width: 14, style: { numFmt: '#,##0.00' } });
workbook.defineName('Total', 'Sales!$B$1');

const read = readWorkbook(writeWorkbook(workbook));
console.log(read.properties.title, read.definedNames, read.getSheet('Sales').cell('A1')?.richText);
```

- **Rich text:** `{ richText: [{ text, font? }] }` writes text whose parts have their own fonts. Reading keeps `cell.value` as the plain text and adds `cell.richText` with the runs. Run fonts take the same properties as cell fonts.
- **Document properties:** `properties` (`title`, `subject`, `creator`, `keywords`, `description`, `lastModifiedBy`, `category`, `created`, `modified`, `lastPrinted`, and a few more) are stored in `docProps/core.xml`, only when at least one is set. Read them from `workbook.properties`. `company`, `manager`, `application` and `appVersion` go to `docProps/app.xml`.
- **Defined names:** `workbook.defineName(name, refersTo, { sheet?, hidden?, comment? })`; `workbook.definedNames` lists the names of a file that was read, built-in names such as `_xlnm._FilterDatabase` included. Invalid names throw `XLV_E_INVALID_DEFINED_NAME`.
- **Column styles:** `column(c, { style })` stores a default style for the column. Applications use it for cells added later; xlvane does not apply it to the cells you set. Reading gives `sheet.columns[i].style`.
- **Built-in number formats:** `numFmt` may be a built-in format number such as `14` (a date shown in the reader's local date order). A format code string such as `'yyyy-mm-dd'` is shown the same everywhere.

### Hyperlinks, notes, tab colour and view

```ts
import { createWorkbook, readWorkbook, writeWorkbook } from 'xlvane';

const workbook = createWorkbook({ properties: { company: 'Example Ltd' } });
const sheet = workbook.addSheet('Links').tabColor('FF0070C0').view({ zoom: 120, showGridLines: false });
sheet.set('A1', 'Docs').link('A1', { url: 'https://example.com/docs', tooltip: 'Open the docs' });
sheet.set('A2', 'Back to top').link('A2', { location: 'Links!A1' });
sheet.set('B2', 1250).note('B2', { text: 'Estimate from March', author: 'Ann' });

const read = readWorkbook(writeWorkbook(workbook)).getSheet('Links');
console.log(read.hyperlinks, read.notes, read.cell('B2')?.note?.text, read.tabColor, read.view.zoom);
```

- **Hyperlinks:** `link(range, { url | location, tooltip?, display? })`. A `url` (web page, `mailto:`, file) is stored as an external relationship of the sheet; a `location` (a place in the workbook) is stored without one. The cell text is set with `set`. Reading gives `sheet.hyperlinks` and `cell.hyperlink`; xlvane never opens a target, and `XLV_W_UNUSUAL_HYPERLINK` reports targets other than http(s) or mailto.
- **Notes (comments):** `note(cell, text | { text | runs, author? })`. The cell needs no value. Notes are written hidden, as spreadsheet applications show them when the pointer rests on the cell. Reading gives `sheet.notes` and `cell.note` (`text`, `author`, and `runs` for rich text). Threaded comments of current Excel versions are not read; their placeholder notes are, with `XLV_W_THREADED_COMMENTS`.
- **Tab colour and view:** `tabColor(color)`, and `view({ zoom, showGridLines, showRowColHeaders, rightToLeft, layout, split })`. Reading gives `sheet.tabColor` and `sheet.view`.
- **Streaming:** `writeWorkbookStream` takes `hyperlinks` and `notes` per row (`{ values, hyperlinks: [...], notes: [...] }`, by column position) and `tabColor`/`view` per sheet. `openWorkbookStream` sets `cell.note` on the rows it returns and lists a sheet's notes with `notes(sheet)`; it does not set `cell.hyperlink`.
- `editCells` and `appendRows` keep existing links and notes but cannot add new ones yet.

### Sheet rules and printing

```ts
import { createWorkbook, readWorkbook, writeWorkbook } from 'xlvane';

const workbook = createWorkbook();
const sheet = workbook.addSheet('Report').set('A1', 'Quarterly report');
sheet.pageSetup({ orientation: 'landscape', paperSize: 9, margins: { left: 0.5, right: 0.5 }, fitToPage: true, fitToWidth: 1, fitToHeight: 0 });
sheet.headerFooter({ oddHeader: '&C&A', oddFooter: '&LPrinted &D&RPage &P of &N' });
sheet.printArea('A1:H40').printTitles({ rows: '1:2' }).pageBreaks({ rows: [21] });
sheet.protect({ sort: false, autoFilter: false });

const read = readWorkbook(writeWorkbook(workbook)).getSheet('Report');
console.log(read.pageSetup.orientation, read.printArea, read.printTitles, read.protection?.sheet);
```

- **Page setup:** `pageSetup({ margins, orientation, paperSize, scale, fitToPage, fitToWidth, fitToHeight, pageOrder, blackAndWhite, draft, cellComments, errors, firstPageNumber, horizontalDpi, verticalDpi, copies, printOptions })`. Margins are in inches. Values outside the limits Office documents (for example margins of 49 inches or more, a scale other than 10–400 or 0) throw `XLV_E_INVALID_OPTION` with the allowed range.
- **Headers and footers:** `headerFooter({ oddHeader, oddFooter, evenHeader, evenFooter, firstHeader, firstFooter, differentOddEven, differentFirst, … })`, with the codes `&L`, `&C`, `&R` (sections), `&P` (page), `&N` (pages), `&D` (date), `&A` (sheet), `&F` (file), `&&` (an ampersand) and others; at most 255 characters each.
- **Print area, titles and page breaks:** `printArea(range | ranges)`, `printTitles({ rows: '1:2', columns: 'A' })` and `pageBreaks({ rows: [21], columns: ['F'] })` (the rows and columns that start a new page). The print area and titles are stored as the built-in names `_xlnm.Print_Area` and `_xlnm.Print_Titles`; `insertRows` and `deleteRows` move them.
- **Protection:** `protect(flags)` turns on sheet protection. Each flag is `true` when the action is locked, as in the file: `protect({ sort: false })` keeps sorting allowed. Flags not given keep the schema defaults. A password is not supported yet (`XLV_E_NOT_SUPPORTED`). Reading gives `sheet.protection` (flags and the stored password verifier, never interpreted) and `workbook.protection`.
- **Reading** gives `sheet.pageSetup` (with defaults), `sheet.headerFooter`, `sheet.pageBreaks`, `sheet.printArea` and `sheet.printTitles`. `writeWorkbookStream` takes the same settings per sheet (`pageSetup`, `headerFooter`, `pageBreaks`, `printArea`, `printTitles`, `protect`).

## Tasks in one call

```ts
import { readTable, writeTable, editCells, appendRows, diff, toCSV, fromCSV, toJSON, fromJSON } from 'xlvane';
import { z } from 'zod'; // or any Standard Schema library: Valibot, ArkType, …

// Rows as objects, validated; problems come with cell addresses.
const { rows, issues } = readTable(bytes, { sheet: 'Orders', schema: z.object({ Customer: z.string(), Total: z.number() }) });

// A formatted report: bold header, frozen first row, filter, fitted widths.
const report = writeTable(rows, { sheet: 'Report', columns: ['Customer', { key: 'Total', numFmt: '#,##0.00' }] });

// Change cells or add rows in an existing file; everything else in the file stays as it was.
const { bytes: edited } = editCells(bytes, [{ sheet: 'Orders', cell: 'B7', value: 120.5 }]);
const { bytes: appended } = appendRows(bytes, 'Orders', [{ Customer: 'Dan', Total: 42 }]);

// What changed between two versions, as JSON.
console.log(diff(bytes, edited).sheets[0]?.cells);

// CSV and JSON in both directions. toCSV escapes text that spreadsheets would run as formulas.
const csv = toCSV(bytes, { sheet: 'Orders' });
const fromText = fromCSV('sku,stock\r\nA-1,12\r\n');
const json = toJSON(bytes).rows;
const fromData = fromJSON([{ sku: 'A-1', stock: 12 }]);
```

- `readTable` finds the header row like `inspect()`, or takes `header: 3` or `header: false`. With a schema, `rows` has the schema's output type. Asynchronous schemas need `readTableAsync`.
- `editCells` and `appendRows` change only the cells they write, plus new entries in the shared strings and styles. Charts, comments, conditional formatting, data validation, defined names and every other part are kept byte for byte. xlvane does not calculate formulas. When the workbook has formulas, the result carries `XLV_W_FORMULAS_NOT_RECALCULATED` and the file asks spreadsheet applications to recalculate on load.
- `fromCSV` never creates formulas. It converts plain numbers, booleans and ISO dates (`types: "infer"`); numbers with leading zeros or more than 15 digits stay text.
- Shared formulas are translated to each cell when reading. If a formula cannot be translated safely, its `formula.text` is `null` and the reader reports `XLV_W_SHARED_FORMULA_NOT_EXPANDED`.

### Inserting and deleting rows

```ts
import { insertRows, deleteRows } from 'xlvane';

declare const bytes: Uint8Array;
// Two rows before row 5 of "Orders", formatted like row 4.
const { bytes: wider } = insertRows(bytes, { sheet: 'Orders', at: 5, count: 2, style: 'above' });
// Rows 10 to 12 of "Orders".
const { bytes: shorter } = deleteRows(wider, { sheet: 'Orders', at: 10, count: 3 });
```

- Like `editCells`, they change an existing file in place and keep every other part as it was.
- References follow the rows as in a spreadsheet application: formulas on every sheet, merged cells, hyperlinks, notes, the autofilter, conditional formats, data validation, page breaks and defined names. A range grows when rows are inserted inside it and shrinks when some of its rows are deleted; a reference whose cells are all deleted becomes `#REF!`. The rules are in docs/facts/ref-adjust.md (observed in LibreOffice; not yet checked against Excel).
- Sheets with tables, pictures, charts, pivot tables or form controls, and formulas xlvane cannot rewrite safely (for example 3-D references such as `Jan:Mar!A1` that include the sheet), are refused with `XLV_E_NOT_SUPPORTED`. Data that would be pushed past row 1048576 gives `XLV_E_LIMIT_EXCEEDED`.
- Sheet views (selection, frozen panes) are not changed.
- `adjustFormula` and `adjustRange` apply the same rules to formula text you hold yourself.

## Large files: streaming

```ts
import { openWorkbookStream, writeWorkbookStream, readTableStream } from 'xlvane';
import { openWorkbookStreamFile, writeWorkbookStreamFile } from 'xlvane/node';

// Read row by row: worksheet cells are never all in memory. A Blob (File, fs.openAsBlob) also
// keeps the file itself out of memory.
const book = await openWorkbookStreamFile('large.xlsx');
for await (const row of book.rows('Data')) console.log(row.number, row.cells.map((c) => c.value));

// Rows as validated objects, one at a time.
for await (const { value, issues } of readTableStream(file, { schema })) { /* … */ }

// Write: rows are pulled from a (async) iterable only as fast as the output is consumed.
function* rows() {
  yield { values: ['id', 'value'], style: { font: { bold: true } } };
  for (let i = 1; i <= 1_000_000; i++) yield [i, Math.sqrt(i)];
}
await writeWorkbookStreamFile('out.xlsx', { sheets: [{ name: 'Data', freeze: { rows: 1 }, autoFilter: true, rows: rows() }] });
const stream = writeWorkbookStream({ sheets: [{ name: 'Data', rows: rows() }] }); // ReadableStream<Uint8Array>
```

- Memory stays constant: the test suite reads and writes 50 000 and 400 000 rows and checks that the retained heap does not grow. It measured under 1 MB in both cases.
- Streamed text is stored in the cells (inline strings) by default. Pass `sharedStrings: true` for a shared string table, which is smaller when values repeat but kept in memory until the end.
- Rows come in file order. A shared formula whose master cell comes later in the file cannot be translated while streaming.
- `maxCells` does not apply when streaming; single parts may decompress to 4 GiB (`STREAM_LIMIT_DEFAULTS`). All other limits apply.

## Benchmark

100 000 rows × 20 columns (numbers, texts, dates, booleans). Median of 5 runs, each in a fresh Node.js 22.22.0 process. Machine: Intel Xeon @ 2.10 GHz, 4 logical cores, 15.7 GiB RAM, Linux (a cloud VM). Measured on 2026-10-30. Peak RSS includes Node.js and the imported libraries; the baseline with all three libraries imported is 135 MiB. Method, raw data and caveats: [bench/perf/RESULTS.md](../../bench/perf/RESULTS.md). Reproduce with `npm run bench`.

| Write | Time, s | Peak RSS, MiB | | Read | Time, s | Peak RSS, MiB |
|---|---|---|---|---|---|---|
| xlvane `writeWorkbook` | 9.37 | 1005 | | xlvane `readWorkbook` | 4.52 | 690 |
| xlvane `writeWorkbookStream` | 4.91 | 170 | | xlvane `openWorkbookStream` | 5.32 | 309 |
| exceljs 4.4.0 `Workbook` | 18.72 | 2102 | | exceljs `xlsx.load` | 6.80 | 889 |
| exceljs `WorkbookWriter` (stream) | 6.31 | 194 | | exceljs `WorkbookReader` (stream) | 4.60 | 309 |
| SheetJS 0.18.5 `XLSX.write` | 10.96 | 1002 | | SheetJS `XLSX.read` | 9.04 | 838 |

The streaming reader of exceljs was faster than xlvane's in this run (4.60 s against 5.32 s). SheetJS is the npm release 0.18.5, which has CVE-2023-30533; the benchmark reads only files it generated itself. The results come from one machine and one data mix.

## Inspecting a workbook

`inspect()` returns plain JSON data about the structure of a workbook: sheets and their visibility, used ranges, a detected header row, column types with example values, merged cells, formulas, rich-text cells, column styles, document properties, defined names and warnings. The schema is documented in [docs/inspect-json.md](../../docs/inspect-json.md).

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
| `maxSharedStringChars` | 512 Mi characters | many cells pointing at one huge shared string (a few MB that become many GB when values are copied) |

## Package size

**Size budget:** 175 kB of gzip-compressed JavaScript for the library files in `dist/` (level 9, unminified as published; measured 122.9 kB on 2026-10-30, after phase A2 hyperlinks). **CLI size budget:** 40 kB for `dist/cli.js`, which is not loaded by `import 'xlvane'` (measured 20.3 kB). `npm run check:pkg` fails when the budget is exceeded. The only dependency, fflate, adds about 22 kB gzip (its `esm/browser.js`, measured the same way).

Not supported: password-protected workbooks and legacy `.xls` files (both rejected with `XLV_E_UNSUPPORTED_CONTAINER`).

License: MIT
