# xlvane

Read, inspect and write XLSX files; input files are treated as untrusted. ESM only, one runtime dependency ([fflate](https://github.com/101arrowz/fflate)). The core works with `Uint8Array` and has no Node.js dependencies; file helpers are in `xlvane/node`.

```sh
npm install xlvane
```

The package also contains `docs/` (error codes, JSON schemas, exceljs compatibility), the Agent Skill `skills/xlvane/SKILL.md` and `llms-full.txt` (all of it in one file), for reading without network access: after installing, they are in `node_modules/xlvane/`. Links below point to the same files on GitHub.

## Quick start

```ts
import { inspect, readTable, writeTable } from 'xlvane';
import { readWorkbookFile } from 'xlvane/node';

const workbook = await readWorkbookFile('orders.xlsx');
console.log(inspect(workbook).sheets.map((s) => [s.name, s.usedRange])); // what is in the file
const { rows } = readTable(workbook, { sheet: 'Orders' }); // [{ Customer: 'Ann', Total: 120.5 }, …]
const big = rows.filter((row) => typeof row['Total'] === 'number' && row['Total'] > 100);
const bytes = writeTable(big, { sheet: 'Big orders' }); // a formatted .xlsx file as a Uint8Array
```

From a shell: `npx xlvane inspect orders.xlsx`, then `npx xlvane read orders.xlsx --sheet Orders`.

## Contents

- [Command line in 5 minutes](#command-line-in-5-minutes)
- [Reading a workbook](#reading-a-workbook) · [Writing a workbook](#writing-a-workbook) (rich text, hyperlinks, notes, images, printing, [data validation and conditional formatting](#data-validation-and-conditional-formatting), [tables](#tables))
- [Tasks in one call](#tasks-in-one-call): tables of rows, CSV, JSON, editing files, diff · [inserting and deleting rows](#inserting-and-deleting-rows)
- [Large files: streaming](#large-files-streaming)
- [Inspecting a workbook](#inspecting-a-workbook) · [Errors and warnings](#errors-and-warnings) · [Untrusted input and limits](#untrusted-input-and-limits)
- [Limitations](#limitations) · [Package size](#package-size)
- exceljs users: [`@xlvane/compat`](https://www.npmjs.com/package/@xlvane/compat) and `xlvane migrate`

## Tested environments

The core API (reading, writing, inspect, tables, CSV, editing, diff, streaming with Web Streams) and `@xlvane/compat` are tested in Node.js 22, 24 and 26, Bun, Deno, Chromium (Playwright) and the Cloudflare Workers runtime (workerd), with the same test bundle (`npm run test:runtimes`). Reading is tested on 340 files of the Apache POI test corpus (written by Excel for Windows, Mac and the web, LibreOffice, ONLYOFFICE and others, and damaged files) and on files saved by LibreOffice; see `tests/corpus/`.

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
- `--json` prints one JSON document per command with a documented schema ([docs/cli-json.md](https://github.com/DevXord/xlvane/blob/main/docs/cli-json.md)). Exit codes: 0 success, 1 a check found something (`diff`, `validate`), 2 usage error, 3 input or file error.
- Colour (red errors, yellow warnings, green results, bold headings) is used only on a terminal and never carries information alone; `NO_COLOR` or `TERM=dumb` turn it off, `FORCE_COLOR=1` turns it on. Control characters from files, such as an escape character in a cell, are shown as `\x1b` and never reach the terminal.
- Tables and summaries go to separate streams: `xlvane read report.xlsx > out.txt` writes only the table; the row count and warnings notes go to stderr.
- `xlvane migrate` prepares a project for `@xlvane/compat`.

## Reading a workbook

```ts
import { readWorkbook } from 'xlvane';

declare const bytes: Uint8Array; // the file's contents
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

Cell values are numbers, strings, booleans, `Date`s (for date-formatted numbers and ISO date cells; the UTC fields of the `Date` hold the spreadsheet's date and time), error codes such as `"#DIV/0!"` (with `type: "error"`) or `null`. Formula cells carry `formula.text` and the cached result stored in the file; xlvane does not calculate formulas. Cells that have neither a value nor a formula are not part of `cells()`; those that hold only a style are listed by `sheet.styledCells()` (and `styledCells` of each streamed row), so tables, CSV and `inspect` see only data.

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

- Values: numbers, strings, booleans, `Date`s (stored as serial numbers with a `yyyy-mm-dd` or `yyyy-mm-dd hh:mm:ss` format unless the style sets one), `{ error: '#N/A' }`, and `{ formula, result }`, also as an array formula (`{ formula, array: 'A1:A3' }`) or a shared formula (`{ formula, shared: 'B1:B9' }` on the first cell, `{ sharedFormula: 'B1' }` on the others). xlvane writes the formula text and the cached result you give; it does not calculate formulas.
- Styles: font (name, size, bold, italic, underline, strike, colour, superscript and subscript, family, scheme, charset, outline, shadow), fill (solid colour, pattern, or linear and path gradient), borders (top, right, bottom, left, diagonal), alignment (horizontal, vertical, wrap, shrink to fit, indent, rotation, reading order), cell protection (`locked`, `hidden`) and number format. Identical styles are stored once. `sheet.style(range, style)` merges a style into every cell of a range.
- Layout: column widths and hidden columns, row heights and hidden rows, outline levels of rows and columns (`row(n, { outlineLevel: 1, collapsed })`, `column('B', { outlineLevel: 1 })`), merged ranges, frozen rows and columns, autofilter, hidden and very hidden sheets. `sheet.properties({ defaultRowHeight, outlineLevelRow, outlineLevelCol, summaryBelow, summaryRight })` sets the sheet's format properties.
- Workbook: `workbook.views([{ x, y, width, height, firstSheet, activeTab, visibility }])` sets the workbook windows and `workbook.calculation({ fullCalcOnLoad: true })` asks applications to recalculate when the file is opened.
- Writer options: `writeWorkbook(workbook, { sharedStrings: false })` writes text inline instead of in a shared string table; `{ compressionLevel: 0 }` (0–9, default 6) stores or compresses the parts. The same options work for `writeWorkbookFile` and, with `compressionLevel`, `writeWorkbookStream`.
- Invalid input is rejected when it is set, with an `XlvError` saying what and where, for example `XLV_E_INVALID_SHEET_NAME` or `XLV_E_INVALID_STYLE` with the allowed values.
- The output is deterministic: the same workbook always gives the same bytes. The exception is a protection password (`protect({ password })`), whose verifier gets a random salt; pass a fixed salt through `hashPassword` for identical bytes. The test suite opens written files that use all of these features in LibreOffice and checks the values, styles and layout it sees.

Reading returns the same information: `cell.style`, `cell.formula`, `sheet.columns`, `sheet.rowInfo(row)`, `sheet.properties`, `sheet.freeze` and `sheet.autoFilter`.

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
- **Notes (comments):** `note(cell, text | { text | runs, author?, box? })`; `box` sets the note box: `margins` in millimetres, `autoMargins`, `locked`, `lockText`, `anchor` (`'absolute'`, `'oneCell'`, `'twoCell'`); reading gives the same as `note.box`. The cell needs no value. Notes are written hidden, as spreadsheet applications show them when the pointer rests on the cell. Reading gives `sheet.notes` and `cell.note` (`text`, `author`, and `runs` for rich text). Threaded comments of current Excel versions are not read; their placeholder notes are, with `XLV_W_THREADED_COMMENTS`.
- **Tab colour and view:** `tabColor(color)`, and `view({ zoom, zoomScaleNormal, showGridLines, showRowColHeaders, showRuler, rightToLeft, layout, split, topLeftCell, activeCell })`. Reading gives `sheet.tabColor`, `sheet.view`, `sheet.properties`, `workbook.views` and `workbook.calculation`.
- **Streaming:** `writeWorkbookStream` takes `hyperlinks` and `notes` per row (`{ values, hyperlinks: [...], notes: [...] }`, by column position) and `tabColor`/`view` per sheet. `openWorkbookStream` sets `cell.note` on the rows it returns and lists a sheet's notes with `notes(sheet)`; it does not set `cell.hyperlink`.
- `editCells` and `appendRows` keep existing links and notes but cannot add new ones yet.

### Images

```ts
import { createWorkbook, readWorkbook, writeWorkbook } from 'xlvane';

declare const png: Uint8Array; // the bytes of a PNG, JPEG or GIF file
const workbook = createWorkbook();
const logo = workbook.addImage(png);
const sheet = workbook.addSheet('Report').set('A1', 'Quarterly report');
sheet.addImage(logo, 'E1:G4', { description: 'Company logo', hyperlink: 'https://example.com' });
sheet.addImage(logo, { from: 'A6', size: { width: 160, height: 90 } });

const read = readWorkbook(writeWorkbook(workbook));
for (const picture of read.getSheet('Report').images) console.log(picture.anchor, picture.image?.format);
```

- **Adding:** `workbook.addImage(bytes)` takes PNG, JPEG or GIF bytes (recognised from the first bytes; other formats throw `XLV_E_INVALID_IMAGE`) and returns a number; the same image placed many times is stored once. `sheet.addImage(image, anchor, { name?, description?, hyperlink? })` places it over a range (`"E1:G4"`), between two corners (`{ from, to, editAs }`, each a cell or a 0-based `{ col, row }` with fractions) or at a corner with a size in pixels (`{ from, size }`). `sheet.background(image)` sets the sheet background. Pixels are converted at 96 per inch (9525 EMU); fractional positions use the sheet's column widths and row heights.
- **Reading:** `sheet.images` lists the pictures (`anchor` in EMU with the cell of each corner, `image` with `format`, `contentType` and `bytes`, `description`, `hyperlink`), `sheet.backgroundImage` the background, and `workbook.images` every image once. Charts, shapes and other drawing objects are not read; `XLV_W_DRAWING_CONTENT_NOT_READ` reports them, and `editCells` keeps them unchanged.
- **Rows:** `insertRows` and `deleteRows` move pictures with their rows. A deletion that would cut a picture, and a drawing with charts or shapes, are refused with `XLV_E_NOT_SUPPORTED`.
- **Streaming:** `writeWorkbookStream({ images: [png], sheets: [{ images: [{ image: 0, anchor: 'B2:D6' }], background: 0, rows }] })`.

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
- **Protection:** `protect(flags)` turns on sheet protection. Each flag is `true` when the action is locked, as in the file: `protect({ sort: false })` keeps sorting allowed. Flags not given keep the schema defaults. `protect({ password: 'secret' })` stores a SHA-512 password verifier (100,000 iterations, random salt; `hashPassword` gives a fixed salt or another algorithm), and `workbook.protect({ password })` protects the workbook structure. A protection password is not encryption: the content stays readable, and spreadsheet applications only refuse changes without it. Reading gives `sheet.protection` and `workbook.protection`; `verifyPassword(protection, password)` checks a password against them, including legacy 4-digit verifiers.
- **Reading** gives `sheet.pageSetup` (with defaults), `sheet.headerFooter`, `sheet.pageBreaks`, `sheet.printArea` and `sheet.printTitles`. `writeWorkbookStream` takes the same settings per sheet (`pageSetup`, `headerFooter`, `pageBreaks`, `printArea`, `printTitles`, `protect`).

### Data validation and conditional formatting

```ts
import { createWorkbook, readWorkbook, writeWorkbook } from 'xlvane';

const workbook = createWorkbook();
const sheet = workbook.addSheet('Orders').set('A1', 'Status').set('B1', 'Quantity').set('C1', 'Total');
sheet.validate('A2:A500', { type: 'list', formula1: '"Open,Shipped,Paid"', showErrorMessage: true, error: 'Pick a status from the list.' });
sheet.validate('B2:B500', { type: 'whole', operator: 'greaterThan', formula1: 0, showInputMessage: true, prompt: 'Quantity, at least 1' });
sheet.conditionalFormat('C2:C500', [
  { type: 'cellIs', operator: 'greaterThan', formulas: [1000], style: { fill: { pattern: 'solid', fgColor: 'FFC7CE' } } },
  { type: 'colorScale', colorScale: { cfvo: [{ type: 'min' }, { type: 'max' }], colors: ['F8696B', '63BE7B'] } },
]);

const read = readWorkbook(writeWorkbook(workbook)).getSheet('Orders');
console.log(read.dataValidations.length, read.conditionalFormats.map((f) => f.sqref));
```

- **Data validation:** `validate(range, rule)` with `type` `list`, `whole`, `decimal`, `date`, `time`, `textLength` or `custom`, an `operator`, `formula1`/`formula2`, input and error messages. `validate(range, null)` removes the validations added for exactly that range. Rules that Office would reject (a missing `formula1`, a message longer than the Office limit) throw `XLV_E_INVALID_OPTION` naming the property.
- **Conditional formatting:** `conditionalFormat(range, rules)` with `cellIs`, `expression`, `colorScale`, `dataBar`, `iconSet`, `top10`, `aboveAverage`, `containsText`, `timePeriod`, `duplicateValues` and the other rule types of the file format; `style` is a differential style (font, fill, border, number format). Priorities are numbered in the order the rules are added unless given.
- **Reading** gives `sheet.dataValidations` and `sheet.conditionalFormats`; `insertRows` and `deleteRows` move them with their rows, and `editCells` keeps them unchanged.

### Tables

```ts
import { createWorkbook, readTable, readWorkbook, writeWorkbook } from 'xlvane';

const workbook = createWorkbook();
workbook.addSheet('Data').addTable({
  name: 'Sales',
  ref: 'B2',
  totalsRow: true,
  style: { theme: 'TableStyleMedium9', showRowStripes: true },
  columns: [{ name: 'Region', totalsRowLabel: 'Total' }, { name: 'Revenue', totalsRowFunction: 'sum' }],
  rows: [['North', 1200], ['South', 980]],
});
const bytes = writeWorkbook(workbook);
console.log(readWorkbook(bytes).getSheet('Data').tables.map((t) => [t.displayName, t.ref]));
console.log(readTable(bytes, { table: 'Sales' }).rows); // [{ Region: 'North', Revenue: 1200 }, …]
```

- **Writing:** `addTable({ name, displayName, ref, headerRow, totalsRow, style, columns, rows })` writes the header row (the column names), the data rows and the totals row into the cells, and the table part. `ref` is the top-left cell or the whole range. Totals use `SUBTOTAL` (`sum`, `average`, `count`, `countNums`, `min`, `max`, `stdDev`, `var`), a label (`totalsRowLabel`) or a custom formula; `calculatedColumnFormula` fills the empty data cells of a column. Styles are the 60 built-in table styles. Names follow the rules of defined names (no spaces) and are unique in the workbook together with the defined names; tables must not overlap other tables or merged cells; violations throw `XLV_E_INVALID_TABLE`.
- **Reading:** `sheet.tables` lists each table with its range, header and totals rows, columns (totals functions and formulas) and style. `readTable(bytes, { table: 'Sales' })` reads the data rows of a table by name; an unknown name throws `XLV_E_TABLE_NOT_FOUND` with the available names.
- **Streaming:** `writeWorkbookStream` takes `tables` per sheet with the whole range in `ref`; the header, data and totals values are streamed as rows.
- **Editing:** `editCells` keeps tables as they are; `insertRows` and `deleteRows` move, grow and shrink them with their rows and refuse to delete a table's header row, totals row or all of its data rows.

## Tasks in one call

```ts
import { readTable, writeTable, editCells, appendRows, diff, toCSV, fromCSV, toJSON, fromJSON } from 'xlvane';
import { z } from 'zod'; // or any Standard Schema library: Valibot, ArkType, …

declare const bytes: Uint8Array; // the file's contents

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
- Warnings about the file (for example a shared formula that could not be expanded) are in `workbook.warnings`; `readTable`, `readTableAsync`, `toJSON` and `toCSV` also pass each of them to `onWarning`: `readTable(bytes, { onWarning: (w) => console.warn(w.code, w.message) })`.
- Dates: `readTable` returns `Date`s (their UTC fields hold the sheet's date and time); `dates: "iso"`, `toJSON` and `toCSV` give ISO 8601 text without a time zone, `2024-01-31` or `2024-01-31T13:45:00`.
- `editCells` and `appendRows` change only the cells they write, plus new entries in the shared strings and styles. Charts, comments, conditional formatting, data validation, defined names and every other part are kept byte for byte. xlvane does not calculate formulas. When the workbook has formulas, the result carries `XLV_W_FORMULAS_NOT_RECALCULATED` and the file asks spreadsheet applications to recalculate on load.
- `fromCSV` never creates formulas. It converts plain numbers, booleans and ISO dates (`types: "infer"`); numbers with leading zeros, whole numbers with more than 15 digits (IDs) and numbers a double cannot hold (`1e400`, `1e-400`) stay text. Every number `toCSV` writes reads back as the same number, except whole numbers above 15 digits. At most `maxCells` fields are read (default 5 000 000).
- Round trips: XLSX → CSV → XLSX keeps numbers, booleans and dates. XLSX → JSON → XLSX keeps numbers and booleans, but dates come back as text, because JSON has no date type and `fromJSON` does not guess; convert such fields to `Date` before `fromJSON` or `writeTable`.
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
- Tables move, grow and shrink with their rows; deleting a table's header row, totals row or all of its data rows is refused.
- Pictures move with their rows. Drawings with charts, shapes or form controls, a deletion that would cut a picture, parts xlvane cannot rewrite (charts or pivot tables that may refer to the sheet), and formulas xlvane cannot rewrite safely (for example 3-D references such as `Jan:Mar!A1` that include the sheet) are refused with `XLV_E_NOT_SUPPORTED`. Data that would be pushed past row 1048576 gives `XLV_E_LIMIT_EXCEEDED`.
- They also refuse sheets whose conditional formats use the [MS-XLSX] extension, which xlvane writes for data bars with an extension setting (`gradient`, `borderColor`, `negativeFillColor`, `negativeBorderColor`, `axisPosition`, `axisColor`, `direction`, or an `autoMin`/`autoMax` threshold) and for the icon sets `3Stars`, `3Triangles`, `5Boxes` and `NoIcons`; insert the rows before adding such rules, or add the rules again after the change.
- Sheet views (selection, frozen panes) are not changed.
- `adjustFormula` and `adjustRange` apply the same rules to formula text you hold yourself.

## Large files: streaming

```ts
import { openWorkbookStream, writeWorkbookStream, readTableStream } from 'xlvane';
import { openWorkbookStreamFile, writeWorkbookStreamFile } from 'xlvane/node';
import { z } from 'zod';

declare const file: Blob; // a File from an <input>, or fs.openAsBlob(path) in Node.js
const schema = z.object({ id: z.number(), value: z.number() });

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
- `readTableStream` differs from `readTable` in what it can see in advance: it detects the header row among the first rows it reads and fixes the columns there (the columns of the header row, or of the first rows without one), so a column that first appears further down is not part of the objects. `readTable` takes the columns of the whole sheet. For files that fit in memory, `readTable` gives the complete table.

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

`inspect()` returns plain JSON data about the structure of a workbook: sheets and their visibility, used ranges, a detected header row, column types with example values, merged cells, formulas, rich-text cells, column styles, document properties, defined names and warnings. The schema is documented in [docs/inspect-json.md](https://github.com/DevXord/xlvane/blob/main/docs/inspect-json.md).

```ts
import { inspect } from 'xlvane';
import { inspectFile } from 'xlvane/node';

declare const bytes: Uint8Array;
const report = inspect(bytes);          // or: await inspectFile('report.xlsx')
console.log(JSON.stringify(report, null, 2));
```

## Errors and warnings

Problems with input files throw `XlvError` with a stable `code` (for example `XLV_E_SHEET_NOT_FOUND`), a `location` (part, sheet, cell), a `hint` saying what to do next and a `docsUrl`. `error.toJSON()` returns the same as data. Recoverable oddities are reported as warnings in `workbook.warnings`. All codes are described in [docs/errors.md](https://github.com/DevXord/xlvane/blob/main/docs/errors.md).

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

**Size budget:** 300 kB of gzip-compressed JavaScript for the library files in `dist/` (level 9, unminified as published; measured 208.6 kB on 2026-10-31). **CLI size budget:** 40 kB for `dist/cli.js`, which is not loaded by `import 'xlvane'` (measured 25.7 kB on 2026-10-31). `npm run check:pkg` fails when the budget is exceeded. The only dependency, fflate, adds about 22 kB gzip (its `esm/browser.js`, measured the same way).

## Limitations

- xlvane does not calculate formulas. It reads and writes formula text and the cached results stored in the file; edited files ask spreadsheet applications to recalculate when they are opened.
- Password-protected (encrypted) workbooks and legacy `.xls` files are rejected with `XLV_E_UNSUPPORTED_CONTAINER`. A sheet protection password is not encryption and is supported.
- Charts, shapes, pivot tables, slicers and form controls are not modelled: reading skips them (with a warning for drawings), `editCells` keeps them unchanged, and row operations refuse sheets where they would have to move.
- The command line reads files, not standard input.

License: MIT
