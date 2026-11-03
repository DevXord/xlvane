# Changelog

The date after each version is the day (UTC) it was released: for 1.0.0 to 1.2.2 the day of publication in the `time` field of the npm registry, from 1.2.3 the day of the release tag (docs/RELEASING.md).

## 1.3.0 (unreleased)

Charts, and security fixes from a penetration test of the release candidate (docs/decisions.md D-K03). Please update all three packages.

### Security fixes

- **CSV formula injection (all versions before 1.3.0):** a byte order mark (U+FEFF/U+FFFE) before `=`, `+`, `-` or `@` hid the formula sign from the escape, and spreadsheet applications drop the mark; such text is now escaped (`toCSV`, `xlvane convert`, `@xlvane/compat` `csv.write*`, also as bytes in `latin1` and `utf16le` output). In `@xlvane/compat`, values returned by `map` that are not strings (for example `new String('=1+1')`) are checked as the text written. Fields holding `,`, `;`, a tab or `|` are quoted whatever the delimiter. `utf16le` output: LibreOffice reads it as 8-bit text and drops the zero bytes, so characters such as `㴀㄀⬀㄀` (bytes `00 3d 00 31 …`) became `=1+1`; while formulas are escaped, utf16le output starts with a byte order mark, fields whose bytes start a formula get the leading `'`, and fields whose bytes would put a separator before `=` are refused (`XLV_E_INVALID_OPTION`, option `encoding`). An empty delimiter, and with `latin1` or `utf16le` a delimiter other than ASCII, is refused.
- **@xlvane/compat denial of service:** `xlsx.load` and `stream.xlsx.WorkbookReader` no longer build the moved text of every shared formula (a 4.6 KB file ran out of memory); merged cells are limited to `max(1 048 576, cells in the file)` on load and written in blocks (a 2.4 KB file made `writeBuffer` run out of memory); `csv.read` adds rows while it parses and stops at 1 048 576 rows, 16 384 fields in a row or 5 000 000 fields in all. Rows from `stream.xlsx.WorkbookReader` copied into `stream.xlsx.WorkbookWriter` keep their shared formulas.
- **Charts:** series and chart groups count against `maxCells` when reading (a 4.9 MB file ran out of memory); `writeWorkbook` limits the cached text of charts (a `RangeError` before); a drawing part shared by several sheets is read once (`XLV_W_SHEET_PART_SHARED`); `insertRows`/`deleteRows` find chart parts by content type, so a renamed chart part no longer lets them leave its references wrong.
- **@xlvane/mcp:** `..` in the text of a symbolic link is resolved as the system resolves it, so it no longer reveals what exists outside the roots; `write_table` and `convert` from JSON respect `--max-cells`.
- **Text round trip (all versions):** text holding `_x` and four hex digits followed by a character that is written as an escape (a control character, U+FFFE or U+FFFF), for example `"_xAF0F\uFFFF"`, was read back changed; the writer now escapes its underscore. Found by fuzzing.
- **Command line:** a dangling symbolic link as output is replaced with `--force` (it was reported as existing with a suggestion to add `--force` again); special files are reported as "not a regular file"; files with long names can be replaced.

### xlvane

- **Charts:** `sheet.addChart({ type, anchor, series, title, legend, stacked, axes })` writes column, bar, line, pie, scatter and area charts over ranges of the workbook (also in `writeWorkbookStream`, per sheet `charts`); `sheet.charts` reads the charts of a worksheet's drawing (types, stacking, title, legend, axes and series references; references are never evaluated). New error code `XLV_E_INVALID_CHART`. `XLV_W_DRAWING_CONTENT_NOT_READ` no longer counts charts. `inspect` (and `xlvane inspect`, MCP `inspect_workbook`) reports `chartCount` and `charts` (at most 10 charts per sheet, at most 10 chart types each, with `typesTruncated`, and titles cut to 100 characters). A chart that cannot be read (a part over a per-part XML limit, not a chart part, not UTF-8 or UTF-16, not well-formed) is skipped with a warning instead of failing the workbook; cached chart values do not count against `maxXmlTreeElements`, so workbooks with charts over hundreds of thousands of cells read as before. `writeWorkbook` writes chart caches of at most 4 000 000 cells per workbook. Chart references quote sheet names of letters followed by digits (such as `'Sheet1'!$A$1`).

## 1.2.3 (unreleased)

A security release: three ways around the CSV formula-injection escape, an information leak in the MCP server and denial of service by crafted files. Found by penetration tests of the published versions (docs/decisions.md D-J01). Please update all three packages.

### xlvane

- **CSV:** while formulas are escaped (the default), `toCSV` refuses a delimiter that starts a formula (`=`, `+`, `-`, `@` and their full-width and small forms): a line with an empty first field started with it (affects 1.0.0 to 1.2.2). `xlvane convert --delimiter` likewise.
- **Linear time:** notes with many VML parts, pictures with many relationships, header detection (`readTable`, `inspect`), schema issues of `readTable`/`readTableStream`, `editCells` adding many rows, `diff` with styles, and writing many sheets, merges, tables and defined names.
- **Budgets:** `fromCSV`, `fromJSON` and `writeTable` write at most `max(maxCells, 16 × cells written)` positions (very sparse tables beyond 5 000 000 positions are refused; the new `maxCells` option of `writeTable` and `fromJSON` raises the bound); `readTable` and `toJSON` count header names repeated in every row against `maxSharedStringChars`.
- **Streaming:** one `<row>` element holds at most 16 384 cells (before: millions, until the process ran out of memory); unreferenced `.rels` entries are no longer decompressed; part names with percent-encoded characters are found, as in `readWorkbook`.
- Messages that quote text from a file (declared encoding, `date1904`, sheet state) cut it to 200 characters.
- **Command line:** `--force` uses a random temporary file name, removes only a file it created, and does not carry setuid/setgid bits to the new file; `migrate` refuses a `package.json` that is a symbolic link; "Try:" commands put options before `--`; `convert` to JSON reports a text too long for a string as `XLV_E_LIMIT_EXCEEDED`.

### @xlvane/compat

- **CSV:** with `encoding: "latin1"`, characters above U+00FF are written as `?` before fields are checked (U+013D became `=` and U+010A a new row); while formulas are escaped, a `delimiter` holding a formula sign or control character (a lone tab is allowed), a `rowDelimiter` other than a line break, a `quote` other than `"` or `false`, an `escape` other than the quote, and with `quote: false` a field starting with `"` are refused (affects 1.0.0 to 1.2.2). `formulaInjection: "off"` keeps exceljs's behaviour.
- **Load:** cells of defined names count against the load budget (a name over the whole sheet ran out of memory); data validations expand at most 100 000 cells per workbook; 16 000 sheets load in about a second (14 s before); bad sheet names and overlapping merges throw `XlvError` codes instead of plain `Error`s on load.
- **Streams:** `xlsx.read`, `csv.read` and `stream.xlsx.WorkbookReader` stop collecting input after 256 MiB.
- **Write:** shared and array formulas and many images of the same size are handled in near-linear time.
- `csv.read` date formats with `MMM`, `MMMM`, `A`, `a` or `Do` no longer take quadratic time on long digit runs.

### @xlvane/mcp

- **Paths:** a path through a symbolic link inside a root to an outside directory (or a dangling link to an outside file) gave different errors for existing and missing files, revealing what exists outside the roots (affects 1.2.0 to 1.2.2). It is now always `XLV_E_PATH_NOT_ALLOWED`.
- New option `--max-cells` (default 5 000 000) for every workbook read, edited or created from CSV or JSON; `convert` to JSON stops at `--max-file-size` while building the text; a replaced file does not keep setuid/setgid bits; messages on stderr show control characters as `\uXXXX`.

### All packages

- The release workflow fails when the npm registry cannot be asked, instead of passing with nothing staged.

## 1.2.2 (2026-11-02)

### @xlvane/mcp

- The name for the official MCP Registry is `io.github.DevXord/xlvane` (`mcpName` and `server.json`): the registry compares the namespace with case, and the lower-case name of 1.2.1 was refused. No change in behaviour.

### All packages

- The release workflow stages each package unless its version is already on npm, and goes on after a failed package, so a repeated run finishes the release. No change in the packages themselves.

## 1.2.1 (2026-11-02)

### @xlvane/mcp

- `package.json` has `mcpName` (`io.github.devxord/xlvane`), and `packages/mcp/server.json` describes the server for the official MCP Registry, so the server can be listed there. No change in behaviour.

### xlvane, @xlvane/compat

- No changes besides the version.

## 1.2.0 (2026-11-02)

### @xlvane/mcp (new)

- An MCP (Model Context Protocol) server for XLSX files, on the official MCP TypeScript SDK (`@modelcontextprotocol/server` 2.4.0, protocol revision 2026-07-28), over stdio:
  - start it with `npx @xlvane/mcp --root ./data` or `xlvane mcp --root ./data`;
  - `inspect_workbook`, `read_table`, `diff` and `validate` are always available;
  - `write_table`, `edit_cells` and `convert` only with `--allow-write`.
- **Directories:** files are read and written only inside the `--root` directories, after symbolic links are resolved. A file is checked again after it is opened. Writes never replace a file without `"overwrite": true`.
- **Limits:**
  - files above `--max-file-size` (default 64 MiB) are refused;
  - a tool result is at most `--max-response-chars` (default 1 Mi characters of JSON): `read_table` returns the rows that fit, with `nextOffset`;
  - `validate` lists at most 1 000 problems.
- **Errors:** tool errors carry the `XlvError` code, location and hint.
- **API:** `createServer()` and `runCli()` for other transports and tests.

### xlvane

- The command `xlvane mcp` starts the server from `@xlvane/mcp` (a separate package), and prints how to install it when it is missing (`XLV_E_NOT_INSTALLED`). The error code `XLV_E_PATH_NOT_ALLOWED` is used by the server.
- The Agent Skill describes the MCP server.
- **Fix:** the streaming reader (`openWorkbookStream`, `readTableStream`) could refuse a valid file with "Comment is longer than the limit" (1.1.1). This happened when the `--` that ends an XML comment was the last thing in one piece of decompressed data and more than `maxXmlTextLength` (8 Mi characters) of the part followed without another `--`. It was found by fuzzing.

### @xlvane/compat

- No changes besides the version.

## 1.1.1 (2026-11-02)

A security release: small, crafted files could make reading, converting or the command line take minutes or run out of memory. Upgrading is recommended for everyone who opens files from untrusted sources. No API changes besides two new warning codes.

### Security

- **Output budgets:** `toCSV`, `readTable`, `toJSON`, `readTableStream`, `xlvane read` and `xlvane convert` write at most `max(maxCells, 16 × cells read)` positions; beyond that they throw `XLV_E_LIMIT_EXCEEDED` (`details.limit` `maxCells`). Before, two cells A1 and XFD1048576 made `toCSV` write billions of empty fields, and a header 16 384 columns wide made `readTable` build 16 384 properties per cell. Pass a smaller `range` to read part of a large sheet.
- **Shared formulas:** the text of cells sharing a formula counts against `maxSharedStringChars` (the length of each cell's moved formula, apart from shared strings). Before, a 13 KB file sharing a 1 MiB formula with 2 000 cells took about 2 GB once the formulas were read.
- **Sheets sharing one part:** a worksheet part used by several sheets is read once; the other sheets are returned without cells and with the new warning `XLV_W_SHEET_PART_SHARED`. The same holds for comments, VML and table parts (the other sheets are read without those notes or tables). Before, such a part was decompressed and parsed once per sheet, outside the limit on decompressed bytes.
- **Hyperlinks:** linking hyperlinks to cells is limited; beyond the limit `cell.hyperlink` is not set, with the new warning `XLV_W_LIMIT_REACHED` (`sheet.hyperlinks` lists every link).
- **Linear time:** the streaming XML parser on long comments, attributes or text arriving in many pieces; opening a ZIP archive from a `Blob` with many entries; `insertRows`/`deleteRows` on sheets with comments holding `<drawing `; rows of style-only cells written right to left. Repeated header names in `readTable`, `toJSON`, `readTableStream` and `xlvane read` (16 384 columns named `x`: 12 s before). XML comments and processing instructions longer than `maxXmlTextLength` are refused, as text is; the streaming reader also stops early on an unfinished CDATA section, attribute value, reference or XML declaration that is already too long.
- **`insertRows`/`deleteRows`** no longer take a `drawing` element inside an XML comment for the sheet's drawing.
- **Command line:** text from files is escaped also when colours are on: before, an escape sequence that looked like one of the command line's colours passed through (for example in a sheet name shown by `inspect`, `diff`, `read` or `validate`); `--json` writes control and bidirectional characters as `\uXXXX` escapes (the same JSON value). `xlvane migrate` does not follow symbolic links, and `--force` changes the `exceljs` entry of the dependency section instead of the first `"exceljs"` in the file (which could be in `overrides`).
- **CSV formula injection:** text that starts with a control character (NUL included) or with the small forms `﹦`, `﹢`, `﹣`, `﹫` also gets the leading `'`, as a precaution.
- `readTableStream` without a header no longer throws a `RangeError` on rows with very many cells, and its output budget follows `limits.maxCells` when given.

### @xlvane/compat

- `xlsx.load` limits the cells it creates for merged ranges, tables and column definitions (`XLV_E_LIMIT_EXCEEDED`); before, one merged range covering the whole sheet ran out of memory. Checking a new merge against earlier ones takes linear time. Many sheets declaring a million rows in their dimension are refused instead of being scanned row by row.
- `workbook.csv` escapes the same additional formula-like starts as `toCSV`.

## 1.1.0 (2026-11-02)

### xlvane

- **Sheet ids:** `addSheet(name, { sheetId })` and the `sheetId` of a streamed sheet set the sheet's id in the file (1 to 65534, unique); other sheets get the lowest unused ids, so workbooks without given ids are written as before. Reading gives `sheet.sheetId` and `StreamSheetInfo.sheetId` (`null` with `XLV_W_INVALID_VALUE` for a missing, invalid or repeated id), and `inspect()` / `xlvane inspect --json` report it per sheet.
- **Split windows:** `view({ split: { x, y, topLeftCell } })` writes the top-left cell of the bottom-right pane (`pane@topLeftCell`), and reading gives it as `view.split.topLeftCell`.
- **`insertRows`/`deleteRows`** move the conditional formats and data validations of the [MS-XLSX] extension (extended data bars, the icon sets `3Stars`, `3Triangles`, `5Boxes`, `NoIcons`, rules written by Excel) instead of refusing the sheet, with the same results as LibreOffice. Sparklines, slicers and other extended content that refers to cells are still refused.
- **Fixes:**
  - `insertRows`/`deleteRows` now move conditional format thresholds that refer to cells (`{ type: 'formula', value: '$C$12' }`); before, they kept the old reference.
  - `insertRows`/`deleteRows` wrote a moved value containing `$` followed by a digit wrongly, for example a hyperlink location `Main!$A$12` became `Main!$A location=4`.
  - A sheet named like a word in the extension namespaces (for example `Main`) no longer makes `insertRows`/`deleteRows` refuse because of conditional formats on other sheets.
- **Command line:** tables, cut cells and wrapped text count terminal columns, so CJK text and emoji stay aligned; the `xlvane` command no longer makes Node.js print its own warning when both `NO_COLOR` and `FORCE_COLOR` are set (started through `npx`, npm itself may still print it).

### @xlvane/compat

- `worksheet.id` is written as `sheetId` and kept when a file is loaded and by the `model` setter, as in exceljs, so `getWorksheet(id)` finds the same sheet after removing one, writing and loading.
- A split view's `topLeftCell` is written and read back, as in exceljs.

## 1.0.1 (2026-11-02)

### xlvane

- `xlvane migrate` no longer lists `addImage({ filename })` as not supported: @xlvane/compat reads the file when the workbook is written, as exceljs does. It still lists `ModelContainer` and deep imports of `exceljs/lib/...`.
- README: package size measured for 1.0.0, benchmark link that works outside the repository, Agent Skill install command, Deno and bundler imports, links to the changelog, issues and security policy.
- Published from GitHub Actions with npm provenance (trusted publishing); 1.0.0 was published without it.

### @xlvane/compat

- README: installing (npm alias or `xlvane migrate`), requirements and runtimes, TypeScript settings, the differences from exceljs most likely to matter, links.
- docs/exceljs-compat.md lists `ModelContainer` as not supported (constructing it throws `XLV_E_NOT_SUPPORTED`, as in 1.0.0).

## 1.0.0 (2026-11-02)

First release of `xlvane` and `@xlvane/compat`.

### xlvane

- **Reading** XLSX, XLSM, XLTX and XLTM files from untrusted sources:
  - ZIP and XML limits; no DTDs or external entities; macros are never run and links are never followed;
  - every error has an `XLV_E_*` code, location and hint (docs/errors.md);
  - damaged rows and cell references are skipped with a warning, as LibreOffice does.
- **`inspect()`**: sheets, used ranges, header rows, column types and samples, merges, formulas, rich-text cells, column styles, document properties, defined names, warnings (docs/inspect-json.md).
- **Rich text** (`{ richText: [...] }`, `cell.richText`), **document properties** (`docProps/core.xml`, `workbook.properties`), **defined names** (`defineName`, `workbook.definedNames`), **column default styles** and **built-in number format ids** (such as `14`), read and written.
- **Hyperlinks** (external and in-workbook, written per ECMA-376), **notes** (comments, with rich text and author), **tab colour**, **sheet view** (zoom, grid lines, split) and **extended document properties** (`docProps/app.xml`: company, manager), read and written, also by the streaming writer.
- **Pictures** (PNG, JPEG, GIF): `addImage` with range, corner or size anchors and a sheet background; reading gives `sheet.images` with anchors and image bytes; `insertRows`/`deleteRows` move them.
- **Tables** (`addTable`, `sheet.tables`, `readTable(bytes, { table })`): header, totals row with `SUBTOTAL` functions, calculated columns and the 60 built-in table styles, read and written, also by the streaming writer; `insertRows`/`deleteRows` move them.
- **Writing** with styles (fonts with every ECMA-376 property, pattern and gradient fills, borders including diagonals, alignment, number formats, cell protection), column widths, row heights, outline levels, merges, frozen and split panes, autofilter, hidden sheets, workbook views and full recalculation on load; deterministic output. Style values and counts are checked against the limits Office documents in [MS-OI29500]. Cells that have only a style are kept on read (`sheet.styledCells()`).
- **Formulas**: plain, array and shared formulas read and written.
- **Sheet rules**: data validation and conditional formatting (including Excel 2010 data bars and icon sets), sheet and workbook protection with passwords ([MS-OFFCRYPTO] §2.4.2.4; `hashPassword`, `verifyPassword`), page setup, headers and footers, page breaks, print area and print titles.
- **Task API**:
  - `readTable` with Standard Schema validation, `writeTable`;
  - `toCSV`/`fromCSV` with CSV-injection escaping, `toJSON`/`fromJSON`;
  - `editCells`/`appendRows`, which keep everything else in the file unchanged;
  - `insertRows`/`deleteRows`, which move formulas, merged cells, links, notes, conditional formats, data validation and defined names with the rows (`xlvane edit --insert-rows` / `--delete-rows`); `adjustFormula`/`adjustRange` for formula text;
  - `diff`.
- **Streaming**: `openWorkbookStream`, `writeWorkbookStream`, `readTableStream`, `writeTableStream`.
- **Node.js helpers** in `xlvane/node`.
- **The `xlvane` command line**:
  - `inspect`, `read`, `convert`, `create`, `edit`, `diff`, `validate`, `migrate`;
  - `--json` with documented schemas (docs/cli-json.md).
- **Agent Skill** (`skills/xlvane/SKILL.md`), the documentation and `llms-full.txt` are included in the package.
- **Tested** in Node.js 22, 24 and 26, Bun, Deno, Chromium and Cloudflare Workers (workerd), and on 340 files of the Apache POI test corpus.

### @xlvane/compat

- **The exceljs 4.4.0 API** on xlvane, for switching with `"exceljs": "npm:@xlvane/compat@^1"`: workbooks, worksheets, rows, columns, cells, styles, rich text, hyperlinks, notes, formulas (incl. shared and array formulas and `fillFormula`), merges, defined names, data validation, conditional formatting, protection, page setup, tables, images, row and column operations, `workbook.csv`, streaming `WorkbookWriter`/`WorkbookReader`, and the documented deep imports.
- Where exceljs writes data the specification or Excel rejects, or loses data, xlvane writes valid data and keeps it; every difference is listed in docs/exceljs-compat.md. Row and column operations move formulas, merges, notes and names as a spreadsheet application does. CSV output escapes formula-like text by default (`formulaInjection: 'off'` gives exceljs's output).

### Limitations

- Formulas are not calculated: xlvane reads and writes formula text and the cached results stored in the file. Edited files ask spreadsheet applications to recalculate when they are opened.
- Legacy `.xls` files and password-encrypted workbooks are not read (`XLV_E_UNSUPPORTED_CONTAINER`). Sheet and workbook protection passwords, which are not encryption, are supported.
- Charts, shapes, pivot tables, slicers and form controls are not modelled: reading skips them (drawings with a warning), `editCells` keeps them unchanged, and `insertRows`/`deleteRows` refuse sheets where they would have to move.
- `insertRows`/`deleteRows` also refuse sheets whose conditional formats use the [MS-XLSX] extension, which xlvane writes for data bars with an extension setting (`gradient`, `borderColor`, `negativeFillColor`, `negativeBorderColor`, `axisPosition`, `axisColor`, `direction`, or an `autoMin`/`autoMax` threshold) and for the icon sets `3Stars`, `3Triangles`, `5Boxes` and `NoIcons`; insert the rows before adding such rules, or add the rules again after the change.
- `readTableStream` fixes its columns from the first rows it reads; `readTable` uses the whole sheet.
- JSON has no date type: `fromJSON` keeps ISO date strings as text (`fromCSV` recognises them).
- The command line reads files only, not standard input.
