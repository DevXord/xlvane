# Corpus files requested from the maintainer

These files must come from real spreadsheet applications, which the development environment does not have. Create each file from scratch (no personal or confidential data), put it in the folder shown, and add a line to `tests/corpus/LICENSES.md`, for example:

`| excel/basic-types.xlsx | Created by the maintainer in Excel for Microsoft 365 (Windows), version …, 2026-10-30 | MIT (same as the project) |`

Then run `npx vitest run tests/corpus -u` once to create the `.inspect.json` snapshots, review them, and commit both.

The same content is used for every application so that the outputs can be compared.

## Content of "basic-types" (all applications)

Sheet `Data`:

| | A | B | C | D | E | F |
|---|---|---|---|---|---|---|
| 1 | `Text` | `Number` | `Date` | `DateTime` | `Boolean` | `Formula` |
| 2 | `Alice` | `1234.5` | 2024-01-31 (short date format) | 2024-01-31 13:45 (format `yyyy-mm-dd hh:mm`) | TRUE | `=B2*2` |
| 3 | `  spaces  ` (two leading and trailing spaces) | `-0.001` | 1900-03-01 | 2024-02-29 00:00 | FALSE | `=1/0` |
| 4 | a line break inside the cell (`first` ↵ `second`) | `1E+20` (scientific format) | 9999-12-31 | | | `=A2&"!"` |
| 5 | `zażółć 😀` | `0.1` with format `0.00%` | | | | `=SUM(B2:B4)` |

Sheet `Second`: cell A1 = `=Data!B2`.

## Excel (desktop for Windows or Mac; Excel for the web where noted)

| File | How to make it | Covers |
|---|---|---|
| `excel/basic-types.xlsx` | "basic-types" content, save as Excel Workbook (*.xlsx). Excel for the web is fine. | cell types, dates, shared strings, styles |
| `excel/formulas.xlsx` | A1:A10 = 1…10; B1 = `=A1*2` filled down to B10; C1 = `=SUM(A1:A10*B1:B10)` entered with Ctrl+Shift+Enter (array formula); D1 = `=SEQUENCE(3)` | shared and array formulas, cached values |
| `excel/merged-hidden.xlsx` | Sheet `Report`: "Title" in A1, merge A1:D1, headers in row 2, three data rows. Sheet `Hidden`: any value, then hide the sheet. Desktop only, optional: a sheet `Secret` set to xlSheetVeryHidden in the VBA editor's Properties window (save as .xlsx, not .xlsm). | merged cells, header detection, hidden and very hidden sheets |
| `excel/date1904.xlsx` | Desktop: File → Options → Advanced → "Use 1904 date system" (or Excel for Mac → Preferences → Calculation). A1 = 1904-01-01, A2 = 2024-01-31. | 1904 date system |
| `excel/strict.xlsx` | Desktop: "basic-types" content, File → Save As → "Strict Open XML Spreadsheet (*.xlsx)". | Strict conformance |
| `excel/macro.xlsm` | Desktop: any value in A1, add a VBA module containing only `Sub Hello()` / `End Sub`, save as Excel Macro-Enabled Workbook. | macro detection (never run) |
| `excel/password.xlsx` | Desktop: File → Info → Protect Workbook → Encrypt with Password, password `test`. | encrypted workbook error |
| `excel/legacy.xls` | Desktop: save "basic-types" as Excel 97-2003 Workbook (*.xls). | legacy format error |

## LibreOffice Calc (free download from libreoffice.org)

| File | How to make it | Covers |
|---|---|---|
| `libreoffice/basic-types.xlsx` | "basic-types" content, File → Save As → "Excel 2007-365 (.xlsx)", keep format. | cell types as written by LibreOffice |
| `libreoffice/formulas-merged-hidden.xlsx` | The `excel/formulas.xlsx` content on sheet `Formulas` (array formula with Ctrl+Shift+Enter, skip SEQUENCE), plus a `Report` sheet with A1:D1 merged and a hidden sheet `Hidden`. | formulas, merges, hidden sheets |
| `libreoffice/date1904.xlsx` | Tools → Options → LibreOffice Calc → Calculate → Date: 01/01/1904; A1 = 1904-01-01, A2 = 2024-01-31. | 1904 date system |

## Google Sheets

| File | How to make it | Covers |
|---|---|---|
| `google-sheets/basic-types.xlsx` | "basic-types" content, File → Download → Microsoft Excel (.xlsx). | cell types as written by Google Sheets |
| `google-sheets/formulas-merged-hidden.xlsx` | B1:B10 = `=A1*2` filled down, an `=ARRAYFORMULA(A1:A3*10)`, A1:D1 merged on another sheet, one hidden sheet; download as .xlsx. | formulas, merges, hidden sheets |
