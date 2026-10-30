---
name: xlvane
description: Read, inspect, write, edit, convert, compare and validate Excel XLSX files with the xlvane library, its xlvane command line or its MCP server. Use when a task involves .xlsx files, sheets turned into JSON or CSV, spreadsheet reports, or checking changes to a workbook.
license: MIT
compatibility: Node.js 22 or later. The library core works on Uint8Array and has no Node.js dependencies; the CLI, the file helpers in xlvane/node and the MCP server need Node.js.
---

# xlvane: XLSX files

xlvane reads, writes and edits `.xlsx` files. It treats input files as untrusted (size and ZIP limits, no external entities, macros are never run, links are never followed). It does not calculate formulas: it reads the cached results stored in the file and writes the formula text you give.

Status: the packages are not published to npm yet; until they are, build them from https://github.com/DevXord/xlvane.

Three ways to use it, each with the same behaviour:

| Way | When | Start |
|---|---|---|
| Library `xlvane` | You write JavaScript or TypeScript code | `npm install xlvane` |
| Command line `xlvane` | You run shell commands | `npx xlvane help` |
| MCP server `@xlvane/mcp` | Your client calls MCP tools | `npx @xlvane/mcp --root <dir>` |

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

In Node.js, `xlvane/node` adds `readWorkbookFile`, `inspectFile` and `writeWorkbookFile` (which never replaces an existing file unless `{ overwrite: true }` is passed). `toCSV`, `fromCSV`, `toJSON` and `fromJSON` convert tables. For very large files use `openWorkbookStream` and `writeWorkbookStream`.

Code written for the `exceljs` API can run on xlvane through `@xlvane/exceljs` (`"exceljs": "npm:@xlvane/exceljs@^1"` in package.json); unsupported exceljs features throw `XLV_E_NOT_SUPPORTED`.

## Command line

```sh
xlvane inspect report.xlsx                        # structure and warnings
xlvane read report.xlsx --sheet Sales --limit 20  # a sheet as a table
xlvane convert report.xlsx sales.csv              # XLSX to CSV/JSON, or CSV/JSON to XLSX
xlvane create data.json report.xlsx               # a formatted report from JSON or CSV
xlvane edit report.xlsx --set "Sales!B2=1300" -o updated.xlsx
xlvane diff report.xlsx updated.xlsx              # exit code 1 when they differ
xlvane validate report.xlsx --columns "Region:string,Revenue:number"
```

Add `--json` to any command for one JSON document with a documented schema (`docs/cli-json.md` in the package). Exit codes: 0 success, 1 a check found something, 2 usage error, 3 input or file error. Nothing is overwritten without `--force`; commands that write accept `--dry-run`.

## MCP server

`npx @xlvane/mcp --root ./data` serves the tools `inspect_workbook`, `read_table`, `diff` and `validate` for files inside `./data` only. With `--allow-write` it also serves `write_table`, `edit_cells` and `convert`; they create new files and replace an existing file only when the call passes `"overwrite": true`. Paths outside the roots, also through `..` or symbolic links, are refused with `XLV_E_PATH_NOT_ALLOWED`.

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
