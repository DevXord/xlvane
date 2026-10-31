# Changelog

## 1.0.0

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
