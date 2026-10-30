# Changelog

## 1.0.0

First release of `xlvane` and `@xlvane/exceljs`.

### xlvane

- **Reading** XLSX, XLSM, XLTX and XLTM files from untrusted sources:
  - ZIP and XML limits; no DTDs or external entities; macros are never run and links are never followed;
  - every error has an `XLV_E_*` code, location and hint (docs/errors.md);
  - damaged rows and cell references are skipped with a warning, as LibreOffice does.
- **`inspect()`**: sheets, used ranges, header rows, column types and samples, merges, formulas, warnings (docs/inspect-json.md).
- **Writing** with styles (fonts, fills, borders, alignment, number formats), column widths, row heights, merges, frozen panes, autofilter and hidden sheets; deterministic output.
- **Task API**:
  - `readTable` with Standard Schema validation, `writeTable`;
  - `toCSV`/`fromCSV` with CSV-injection escaping, `toJSON`/`fromJSON`;
  - `editCells`/`appendRows`, which keep everything else in the file unchanged;
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
