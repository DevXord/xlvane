# Roadmap

This file lists only work that is recorded as planned or deferred in [docs/PLAN.md](docs/PLAN.md) or [docs/decisions.md](docs/decisions.md), or decided by the maintainer. It gives no dates. What each released version changed is in [CHANGELOG.md](CHANGELOG.md).

## Next

1. **`@xlvane/xlsx-compat`**: a package with a SheetJS-compatible API on top of xlvane, for projects that use the `xlsx` package from npm. It is the phase after charts (maintainer's decision of 2026-11-03). The package does not exist yet; until then, [docs/xlsx-cves.md](docs/xlsx-cves.md) shows the xlvane calls that correspond to the common SheetJS calls.

## Done

- **Charts** (phase K, version 1.3.0): writing and reading column, bar, line, pie, scatter and area charts in `xlvane` (docs/decisions.md D-K01). Moving charts in `insertRows`/`deleteRows` is deferred (below).

## Deferred

These items were set aside in earlier phases. Each is recorded, none is scheduled.

- **Charts in `insertRows` and `deleteRows`**: moving charts with the rows (deferred from phase K). Until then, row operations refuse a sheet whose charts would have to move, with `XLV_E_NOT_SUPPORTED`.
- **Reading standard input in the command line**: the commands read files only (docs/decisions.md D-R02; docs/CHANGELOG-dev.md, phase F).
- **Removing members marked `@internal` from the published `.d.ts` files** (docs/CHANGELOG-dev.md, phase F).
- **Checking `activeTab` when it is set**: today it is checked when the workbook is written (docs/CHANGELOG-dev.md, phase F).
- **Agent benchmark** (`bench/agents/`, docs/PLAN.md section 4 and phase 8): skipped by the maintainer's decision, because every run costs API tokens (docs/decisions.md D-801, kept in D-I01).
- **Corpus files from Google Sheets, Numbers and Excel with the shared content**: requested from the maintainer (docs/decisions.md D-110 and D-902; tests/corpus/REQUESTED.md). The corpus already has 340 Apache POI files and files saved by LibreOffice.
- **An entry in the e18e `module-replacements` list** for `exceljs` (docs/PLAN.md section 10).

## Not planned, or not done

- **Formula calculation.** xlvane reads and writes formula text and the cached results stored in the file; it does not calculate formulas. docs/PLAN.md does not plan formula calculation for the packages in this repository.
- **Legacy `.xls` files, and formats other than XLSX (with XLSM, XLTX and XLTM on reading), CSV and JSON.** They are rejected or not offered, and no phase plans them.
- **Encrypted (password-protected) workbooks.** They are rejected with `XLV_E_UNSUPPORTED_CONTAINER`; docs/PLAN.md does not plan them for this repository. Sheet and workbook protection passwords, which are not encryption, are supported.
- **Pivot tables, slicers, shapes and form controls.** They are not modelled; `editCells` keeps them unchanged. No phase plans them.
