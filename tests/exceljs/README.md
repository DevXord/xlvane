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
| 404 | `spec/unit/doc/worksheet.spec.js`, `spec/unit/doc/worksheet.dimensions.spec.js`, `spec/integration/workbook-xlsx-reader/workbook-xlsx-reader.spec.js`, `spec/integration/workbook/xlsx.spec.js`, `spec/integration/workbook/workbook-xlsx.spec.js`, `spec/utils/test-utils.js` |

`range.spec.js` and `col-cache.spec.js` test internal classes (`Range`, `colCache`) that are
not part of the exceljs public API, so they were not ported. `defined-names.spec.js` was ported in
phase A1: its class is the object behind the public `workbook.definedNames`.

One fixture was copied in phase A1: `spec/integration/data/rich-text-sample.js` (a base64 XLSX
written by Excel for Mac), decoded unchanged to
`tests/corpus/excel/exceljs-rich-text-sample.xlsx`, and `spec/integration/data/rich-text-sample-a1.json`
as `rich-text-sample-a1.json` (both listed in tests/corpus/LICENSES.md). No other binary
fixtures (`spec/integration/data/*`) were copied; tests that read them are skipped.

## Files

| Test file | Source | Passing | Skipped |
|---|---|---|---|
| `workbook.test.ts` | `spec/unit/doc/workbook.spec.js` | 10 (2 added) | 1 |
| `row.test.ts` | `spec/unit/doc/row.spec.js` | 13 (2 added) | 3 |
| `cell.test.ts` | `spec/unit/doc/cell.spec.js` | 6 | 6 |
| `column.test.ts` | `spec/unit/doc/column.spec.js` | 6 | 3 |
| `worksheet-merge.test.ts` | `spec/unit/doc/worksheet.merge.spec.js` | 7 (2 added) | 0 |
| `worksheet-values.test.ts` | `spec/unit/doc/worksheet.values.spec.js` and `spec/utils/test-spliced-sheet.js` (as `spliced-sheets.ts`) | 38 | 2 |
| `defined-names.test.ts` | `spec/unit/doc/defined-names.spec.js` | 8 (2 added) | 0 |
| `integration-workbook.test.ts` | `spec/integration/workbook/workbook.spec.js` | 23 (1 added) | 13 |
| `integration-worksheet.test.ts` | `spec/integration/worksheet.spec.js` | 19 | 4 |
| `integration-styles.test.ts` | `spec/integration/workbook/styles.spec.js` | 2 | 3 |
| `workbook-writer.test.ts` | `spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js` | 14 (3 added) | 8 |
| `workbook-reader.test.ts` | `spec/integration/workbook-xlsx-reader.spec.js` | 2 (added, not upstream) | 29 |

"Passing" includes tests that expect `XLV_E_NOT_SUPPORTED` for an unsupported feature.

## Why tests are skipped or changed

Each skipped test names its reason in the title. The reasons are:

- **Unsupported features** (they throw `XLV_E_NOT_SUPPORTED`; where cheap, one test checks that):
  images, conditional formatting, data validation, CSV, print settings, shared-formula filling,
  zip compression options and workbook views.
- **Row and column splicing** (phase B): `spec/utils/test-spliced-sheet.js` was downloaded once
  from the exceljs repository at tag v4.4.0 and ported as `spliced-sheets.ts`. Where xlvane moves
  something as a spreadsheet application does and exceljs does not (cells inserted inside a
  defined name join it), the expectation is changed in place and marked `xlvane:`. The
  duplicateRow integration tests build their sheet in memory instead of reading
  `duplicateRowTest.xlsx`, and compare styles by value (xlvane copies styles). Formula rewriting
  waits for the reference engine (`splice.test.ts`, `it.todo`).
- **Internal API**: tests of internal classes and fields (`new Cell`, `Column.toModel`,
  `row.model`, `_upgradeToHyperlink`, `effectiveType`, sheet mocks) are skipped or rewritten
  with the public API.
- **Fixtures**: `createTestBook`/`checkTestBook` from `spec/utils/index.js` build workbooks with
  unsupported features; tests based on them, and tests reading `spec/integration/data/*`, are
  skipped.
- **What xlvane does not keep after writing and reading a file**: font `family` and `outline`,
  gradient fills, cells that have a style but no value, row default styles
  (docs/decisions.md D-206), sheet views other than the first, `topLeftCell` of a split view and
  `activeCell`, rich-text run `family`/`scheme`, and 6-digit `argb` colours (read back with alpha
  `FF`). Round-trip tests use the values xlvane keeps (`styles.roundTripFonts`, a pattern fill,
  one view per worksheet) and do not read back row styles.
- **Messages**: worksheet-name errors are checked with `toThrow()`, without the exceljs message
  text.
