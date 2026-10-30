# Changelog

## 1.0.0

First release of `xlvane` and `@xlvane/exceljs`.

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
- **Writing** with styles (fonts, fills, borders, alignment, number formats), column widths, row heights, merges, frozen panes, autofilter and hidden sheets; deterministic output. Style values and counts are checked against the limits Office documents in [MS-OI29500].
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

### @xlvane/exceljs

- **The commonly used exceljs 4.4.0 API** on xlvane, for switching with `"exceljs": "npm:@xlvane/exceljs@^1"`. Unsupported features throw `XLV_E_NOT_SUPPORTED` (docs/exceljs-compat.md).
- Rich text values, workbook properties, `workbook.definedNames`, column styles, and dates written with built-in format 14 as exceljs writes them.
