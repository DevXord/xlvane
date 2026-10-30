# Ported exceljs specs

The tests in this directory (except `alias.test.ts`, `cross-check.test.ts`, `stream-layout.test.ts`,
`hyperlinks-views.test.ts` and their helpers) are ported from the exceljs test suite. They check that `@xlvane/exceljs` behaves like exceljs for the
supported subset described in [docs/exceljs-compat.md](../../docs/exceljs-compat.md).

## Source

- Repository: https://github.com/exceljs/exceljs, tag `v4.4.0`, commit
  `d911329c2c3ca7a4d5e979f2539a3cd868fc09d4`.
- Files were downloaded from
  `https://raw.githubusercontent.com/exceljs/exceljs/v4.4.0/<path>`.
- License: MIT, Copyright (c) 2014-2019 Guyon Roche. The full text is in
  [LICENSE-exceljs.md](LICENSE-exceljs.md). Each ported file starts with a header naming its source
  file and the changes made.
- `helpers.ts` holds the style fixtures from `spec/utils/data/styles.json` (only the entries used)
  and file helpers that replace `spec/utils/index.js`.

Paths checked at the tag (HTTP status of the raw URL):

| Status | Path |
|---|---|
| 200 | `spec/unit/doc/workbook.spec.js`, `row.spec.js`, `cell.spec.js`, `column.spec.js`, `range.spec.js`, `worksheet.merge.spec.js`, `worksheet.values.spec.js`, `defined-names.spec.js` |
| 200 | `spec/unit/utils/col-cache.spec.js` |
| 200 | `spec/integration/worksheet.spec.js`, `spec/integration/workbook/workbook.spec.js`, `spec/integration/workbook/styles.spec.js`, `spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js`, `spec/integration/workbook-xlsx-reader.spec.js` |
| 200 | `spec/utils/index.js`, `spec/utils/data/styles.json` |
| 200 | `spec/integration/workbook/images.spec.js`, `spec/integration/data/image.png` (phase D) |
| 404 | `spec/integration/workbook/tables.spec.js`, `spec/unit/doc/table.spec.js`, `spec/utils/test-table-sheet.js` (phase D: no table spec was found at the tag) |
| 404 | `spec/unit/doc/worksheet.spec.js`, `spec/unit/doc/worksheet.dimensions.spec.js`, `spec/integration/workbook-xlsx-reader/workbook-xlsx-reader.spec.js`, `spec/integration/workbook/xlsx.spec.js`, `spec/integration/workbook/workbook-xlsx.spec.js`, `spec/utils/test-utils.js` |

`range.spec.js` and `col-cache.spec.js` test internal classes (`Range`, `colCache`) that are
not part of the exceljs public API, so they were not ported. `defined-names.spec.js` was ported in
phase A1: its class is the object behind the public `workbook.definedNames`.

One fixture was copied in phase A1: `spec/integration/data/rich-text-sample.js` (a base64 XLSX
written by Excel for Mac), decoded unchanged to
`tests/corpus/excel/exceljs-rich-text-sample.xlsx`, and `spec/integration/data/rich-text-sample-a1.json`
as `rich-text-sample-a1.json` (both listed in tests/corpus/LICENSES.md). No other binary
fixtures (`spec/integration/data/*`) were copied in phase A1; tests that read them are skipped.
Phase D copied `spec/integration/data/image.png` unchanged to `tests/exceljs/data/image.png`
(listed in tests/corpus/LICENSES.md).

## Files

| Test file | Source | Passing | Skipped |
|---|---|---|---|
| `workbook.test.ts` | `spec/unit/doc/workbook.spec.js` | 10 (2 added) | 1 |
| `row.test.ts` | `spec/unit/doc/row.spec.js` | 13 (2 added) | 3 |
| `cell.test.ts` | `spec/unit/doc/cell.spec.js` | 8 | 4 |
| `column.test.ts` | `spec/unit/doc/column.spec.js` | 7 | 2 |
| `worksheet-merge.test.ts` | `spec/unit/doc/worksheet.merge.spec.js` | 7 (2 added) | 0 |
| `worksheet-values.test.ts` | `spec/unit/doc/worksheet.values.spec.js` and `spec/utils/test-spliced-sheet.js` (as `spliced-sheets.ts`) | 40 | 0 |
| `defined-names.test.ts` | `spec/unit/doc/defined-names.spec.js` | 8 (2 added) | 0 |
| `integration-workbook.test.ts` | `spec/integration/workbook/workbook.spec.js` and `spec/utils/test-data-validation-sheet.js` (as `sheet-rules-sheets.ts`) | 34 (1 added; 4 compression tests with an own sheet) | 2 |
| `integration-worksheet.test.ts` | `spec/integration/worksheet.spec.js` | 20 | 3 |
| `integration-styles.test.ts` | `spec/integration/workbook/styles.spec.js` | 2 | 3 |
| `workbook-writer.test.ts` | `spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js`, `spec/utils/test-data-validation-sheet.js` and `spec/utils/test-conditional-formatting-sheet.js` (as `sheet-rules-sheets.ts`) | 18 (3 added; the zip-option test with an own sheet) | 2 |
| `workbook-reader.test.ts` | `spec/integration/workbook-xlsx-reader.spec.js` | 2 (added, not upstream) | 29 |
| `tables-images.test.ts` | `spec/integration/workbook/images.spec.js` and the two background-image tests of `workbook-xlsx-writer.spec.js` | 16 (8 ported, 8 added) | 0 |
| `csv.test.ts` | added for xlvane (exceljs `lib/csv/csv.js` behaviour, X4-23, X4-24) | 16 | 0 |
| `phase-e-api.test.ts` | added for xlvane (phase E members, styles, formulas, views; X4-27 … X4-30) | 17 | 0 |
| `options-reader.test.ts` | added for xlvane (xlsx options, WorkbookReader options and events, enums; X4-25, X4-26) | 12 | 0 |

"Passing" includes tests that expect `XLV_E_NOT_SUPPORTED` for an unsupported feature.

## Why tests are skipped or changed

Each skipped test names its reason in the title. The reasons are:

- **Phase E**: shared formulas (`fillFormula`, streamed `{ formula, ref }` masters), workbook views,
  zip options and multi-row column headers are ported. `multiple book views` expects the split
  view back with `workbookViewId` 0 and without `topLeftCell`/`activeCell` (xlvane writes one
  sheet view, docs/exceljs-compat.md). The zip-compression tests write their own sheet instead
  of `createTestBook`. The two CSV tests of `workbook.spec.js` run: they
  build the CSV part of `createTestBook` (row 1 and the merged A2:B3 of
  `spec/utils/test-values-sheet.js`, downloaded once at tag v4.4.0) themselves.
- **Images** (phase D): the ported image tests pass the image as `buffer`, because @xlvane/exceljs
  does not read files by name (`filename` throws `XLV_E_NOT_SUPPORTED`). One expectation differs
  on purpose and is marked `xlvane:` (a sized picture with `editAs: 'absolute'` reads back as
  `oneCell` without the tooltip). No table spec exists at the tag; `tables-images.test.ts` has
  xlvane tests for tables.
- **Data validation, conditional formatting and printing** (phase C1):
  `spec/utils/test-data-validation-sheet.js`, `spec/utils/test-conditional-formatting-sheet.js` and
  their data `spec/utils/data/data-validations.json` and `conditional-formatting.json` were
  downloaded once from the exceljs repository at tag v4.4.0 and ported as `sheet-rules-sheets.ts`.
  Two validations of the data are changed in place and marked `xlvane:` (an `errorStyle` that is not
  a schema value, a decimal validation without bounds). The #1814 test expects the number format
  of a rule as a code string. `sheet-rules.test.ts` adds xlvane tests for the differences.
- **Row and column splicing** (phase B): `spec/utils/test-spliced-sheet.js` was downloaded once
  from the exceljs repository at tag v4.4.0 and ported as `spliced-sheets.ts`. Where xlvane moves
  something as a spreadsheet application does and exceljs does not (cells inserted inside a
  defined name join it), the expectation is changed in place and marked `xlvane:`. The
  duplicateRow integration tests build their sheet in memory instead of reading
  `duplicateRowTest.xlsx`, and compare styles by value (xlvane copies styles). Formulas follow
  through xlvane's reference engine (`splice.test.ts`).
- **Internal API**: tests of internal classes and fields (`new Cell`, `Column.toModel`,
  `row.model`, `_upgradeToHyperlink`, sheet mocks) are skipped or rewritten
  with the public API.
- **Fixtures**: `createTestBook`/`checkTestBook` from `spec/utils/index.js` and the sheets they
  combine are not ported; tests based on them (except the compression tests, which write their
  own sheet), and tests reading `spec/integration/data/*` files that were not copied, are skipped.
- **What xlvane does not keep after writing and reading a file**: cells that have a style but
  no value, row default styles (docs/decisions.md D-206), sheet views other than the first,
  `topLeftCell` and `activeCell` of a split view, and 6-digit `argb` colours (read back with
  alpha `FF`). Round-trip tests use one view per worksheet and do not read back row styles.
- **Messages**: worksheet-name errors are checked with `toThrow()`, without the exceljs message
  text.
