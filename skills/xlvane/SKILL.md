---
name: xlvane
description: Read, inspect, write, edit, convert, compare and validate Excel XLSX files with the xlvane library or its xlvane command line. Use when a task involves .xlsx files, sheets turned into JSON or CSV, spreadsheet reports, or checking changes to a workbook.
license: MIT
compatibility: Node.js 22 or later. The library core works on Uint8Array and has no Node.js dependencies; the CLI and the file helpers in xlvane/node need Node.js.
---

# xlvane: XLSX files

xlvane reads, writes and edits `.xlsx` files. It treats input files as untrusted (size and ZIP limits, no external entities, macros are never run, links are never followed). It does not calculate formulas: it reads the cached results stored in the file and writes the formula text you give.

Two ways to use it, with the same behaviour:

| Way | When | Start |
|---|---|---|
| Library `xlvane` | You write JavaScript or TypeScript code | `npm install xlvane` |
| Command line `xlvane` | You run shell commands | `npx xlvane help` |

## Recommended order of work

1. **Look before reading.** `inspect` lists the sheets (including hidden ones), used ranges, the detected header row, column types with sample values, merged cells, formulas and warnings. Use it to pick the sheet and header row.
2. **Read rows as objects** with `readTable`; pass a Standard Schema (Zod, Valibot, ArkType) to validate them. Problems come back as a list with the cell address.
3. **Write** a new report with `writeTable`, or **change** an existing file with `editCells` / `appendRows`, which keep everything else in the file as it was.
4. **Check the result** with `diff` before handing the file over.

## Library

```ts
import { readFile, writeFile } from 'node:fs/promises';
import { appendRows, diff, editCells, inspect, readTable, writeTable } from 'xlvane';

const bytes = await readFile('report.xlsx');

// 1. Structure: sheets, header rows, column types, warnings.
const report = inspect(bytes);
const sales = report.sheets.find((s) => s.name === 'Sales');
console.log(sales?.headerRow, sales?.columns.map((c) => `${c.header}: ${c.dominantType}`));

// 2. Rows as objects keyed by the header row; dates are Date objects.
const table = readTable(bytes, { sheet: 'Sales', header: 'auto' });
console.log(table.rows.length, table.issues);

// 3a. A new formatted report (bold header, frozen first row, autofilter, column widths).
await writeFile('summary.xlsx', writeTable([{ region: 'North', total: 1250 }], { sheet: 'Summary' }));

// 3b. Change cells and append rows; styles, formulas and other sheets stay as they were.
const edited = editCells(bytes, [{ sheet: 'Sales', cell: 'B2', value: 1300 }]);
const appended = appendRows(edited.bytes, 'Sales', [{ Region: 'West', Revenue: 700 }]);
await writeFile('updated.xlsx', appended.bytes);

// 4. What changed.
const changes = diff(bytes, appended.bytes);
console.log(changes.summary);
```

With a schema:

```ts
import { readTable } from 'xlvane';
import { z } from 'zod';

declare const bytes: Uint8Array;
const Row = z.object({ Region: z.string(), Revenue: z.number() });
const { rows, issues } = readTable(bytes, { sheet: 'Sales', schema: Row });
// rows: { Region: string; Revenue: number }[]; issues name the cell of every invalid value.
console.log(rows.length, issues.map((i) => `${i.cell ?? '?'}: ${i.message}`));
```

Rich text, document properties, defined names and column styles:

```ts
import { createWorkbook, readWorkbook } from 'xlvane';

declare const bytes: Uint8Array;
const workbook = readWorkbook(bytes);
console.log(workbook.properties.title, workbook.definedNames.map((n) => `${n.name}=${n.refersTo}`));
console.log(workbook.sheets[0]?.cell('A1')?.richText); // runs with fonts; cell.value is the plain text

const out = createWorkbook({ properties: { title: 'Report' } });
out.addSheet('Data').set('A1', { richText: [{ text: 'Net ', font: { bold: true } }, { text: 'sales' }] }).column('B', { style: { numFmt: '#,##0.00' } });
out.defineName('Net', 'Data!$B$2');
```

Hyperlinks, notes (comments), tab colour and view:

```ts
import { createWorkbook, readWorkbook } from 'xlvane';

declare const bytes: Uint8Array;
const sheet = readWorkbook(bytes).getSheet('Sales');
console.log(sheet.hyperlinks.map((l) => `${l.ref} → ${l.target ?? l.location ?? ''}`), sheet.notes.map((n) => `${n.ref}: ${n.text}`));

const out = createWorkbook({ properties: { company: 'Example Ltd' } });
out.addSheet('Data').tabColor('FF0070C0').set('A1', 'Docs').link('A1', { url: 'https://example.com' }).note('B2', { text: 'Check', author: 'Ann' });
```

Link targets are never opened. Notes are written hidden and need no cell value. `editCells` keeps existing links and notes but cannot add new ones.

Printing and protection:

```ts
import { createWorkbook, readWorkbook } from 'xlvane';

declare const bytes: Uint8Array;
const sheet = readWorkbook(bytes).getSheet('Sales');
console.log(sheet.pageSetup.orientation, sheet.printArea, sheet.protection?.sheet);

const out = createWorkbook();
out.addSheet('Report').pageSetup({ orientation: 'landscape', margins: { left: 0.5 } }).headerFooter({ oddFooter: '&RPage &P of &N' }).printArea('A1:H40').protect({ sort: false });
```

Protection flags are `true` when the action is locked. Passwords are not supported yet.

### Inserting and deleting rows

```ts
import { deleteRows, insertRows } from 'xlvane';

declare const bytes: Uint8Array;
const { bytes: wider } = insertRows(bytes, { sheet: 'Sales', at: 5, count: 2, style: 'above' });
const { bytes: shorter } = deleteRows(wider, { sheet: 'Sales', at: 5, count: 2 });
```

Formulas on every sheet, merged cells, links, notes, conditional formats, data validation and defined names follow the rows; references to deleted cells become `#REF!`. Sheets with tables, pictures or charts are refused with `XLV_E_NOT_SUPPORTED`.

`inspect()` reports the same items (`properties`, `definedNames`, and per sheet `richTextCellCount`, `styledColumns`, `hyperlinkCount` and `noteCount`). A column style is a default for cells added later; xlvane does not apply it to cells you set.

In Node.js, `xlvane/node` adds `readWorkbookFile`, `inspectFile` and `writeWorkbookFile` (which never replaces an existing file unless `{ overwrite: true }` is passed). `toCSV`, `fromCSV`, `toJSON` and `fromJSON` convert tables. For very large files use `openWorkbookStream` and `writeWorkbookStream`.

Code written for the `exceljs` API can run on xlvane through `@xlvane/exceljs` (`"exceljs": "npm:@xlvane/exceljs@^1"` in package.json); unsupported exceljs features throw `XLV_E_NOT_SUPPORTED`.

## Command line

```sh
xlvane inspect report.xlsx                        # structure and warnings
xlvane read report.xlsx --sheet Sales --limit 20  # a sheet as a table
xlvane convert report.xlsx sales.csv              # XLSX to CSV/JSON, or CSV/JSON to XLSX
xlvane create data.json report.xlsx               # a formatted report from JSON or CSV
xlvane edit report.xlsx --set "Sales!B2=1300" -o updated.xlsx
xlvane edit report.xlsx --sheet Sales --insert-rows 5 --count 2 -o wider.xlsx
xlvane diff report.xlsx updated.xlsx              # exit code 1 when they differ
xlvane validate report.xlsx --columns "Region:string,Revenue:number"
```

Add `--json` to any command for one JSON document with a documented schema (`docs/cli-json.md` in the package). Exit codes: 0 success, 1 a check found something, 2 usage error, 3 input or file error. Nothing is overwritten without `--force`; commands that write accept `--dry-run`.

## Errors

Every error is an `XlvError` with a stable `code` (`XLV_E_*`), a `location` (sheet, cell, part), a `hint` that says what to do, and `details` (for example the list of available sheets). Read the hint and correct the call; the codes are described in `docs/errors.md` in the package. Warnings (`XLV_W_*`) never stop reading; check `workbook.warnings` or `inspect(...).warnings`.

Common cases:

- `XLV_E_SHEET_NOT_FOUND`: `details.available` lists the sheets, and the hint names a close match when there is one.
- `XLV_E_HEADER_NOT_FOUND`: pass the header row number, or `header: false` to use column letters.
- `XLV_E_FILE_EXISTS`: choose another output path, or overwrite explicitly.
- `XLV_E_LIMIT_EXCEEDED`: the file is larger than a limit; raise the limit only for trusted files.

## Things to know

- Formulas are not calculated. After `editCells`, cached results that depend on edited cells are out of date until a spreadsheet program recalculates them (the file asks it to on opening).
- Dates in spreadsheets have no time zone; xlvane returns them as `Date` objects whose UTC fields hold the spreadsheet's date and time.
- `.xls` (the old binary format) and encrypted files are not supported and are reported with an error code.
