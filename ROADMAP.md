# Roadmap

This file lists only work that the maintainer has planned, deferred or decided. It gives no dates. What each released version changed is in [CHANGELOG.md](CHANGELOG.md).

## Next

1. **`@xlvane/xlsx-compat`**: a package with a SheetJS-compatible API on top of xlvane, for projects that use the `xlsx` package from npm. It is the phase after charts (maintainer's decision of 2026-11-03). The package does not exist yet; until then, [docs/xlsx-cves.md](docs/xlsx-cves.md) shows the xlvane calls that correspond to the common SheetJS calls.

## Done

- **Charts** (phase K, version 1.3.0): writing and reading column, bar, line, pie, scatter and area charts in `xlvane`. Moving charts in `insertRows`/`deleteRows` is deferred (below).

## Deferred

These items were set aside in earlier phases. Each is recorded, none is scheduled.

- **Charts in `insertRows` and `deleteRows`**: moving charts with the rows (deferred from phase K). Until then, row operations refuse a sheet whose charts would have to move, with `XLV_E_NOT_SUPPORTED`.
- **Reading standard input in the command line**: the commands read files only.
- **Removing members marked `@internal` from the published `.d.ts` files**.
- **Checking `activeTab` when it is set**: today it is checked when the workbook is written.
- **Agent benchmark** (`bench/agents/`): skipped by the maintainer's decision, because every run costs API tokens.
- **Corpus files from Google Sheets, Numbers and Excel with the shared content**: requested from the maintainer (tests/corpus/REQUESTED.md). The corpus already has 340 Apache POI files and files saved by LibreOffice.
- **An entry in the e18e `module-replacements` list** for `exceljs`.

## Not planned, or not done

- **Formula calculation.** xlvane reads and writes formula text and the cached results stored in the file; it does not calculate formulas. It is not planned for the packages in this repository.
- **Legacy `.xls` files, and formats other than XLSX (with XLSM, XLTX and XLTM on reading), CSV and JSON.** They are rejected or not offered, and no phase plans them.
- **Encrypted (password-protected) workbooks.** They are rejected with `XLV_E_UNSUPPORTED_CONTAINER`; they are not planned for this repository. Sheet and workbook protection passwords, which are not encryption, are supported.
- **Pivot tables, slicers, shapes and form controls.** They are not modelled; `editCells` keeps them unchanged. No phase plans them.
