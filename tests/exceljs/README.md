# Ported exceljs specs

The tests in this directory (except `alias.test.ts`, `cross-check.test.ts` and their helpers) are
ported from the exceljs test suite. They check that `@xlvane/exceljs` behaves like exceljs for the
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

`range.spec.js`, `defined-names.spec.js` and `col-cache.spec.js` test internal classes (`Range`,
`DefinedNames`, `colCache`) that are not part of the exceljs public API, so they were not ported.

No binary fixtures (`spec/integration/data/*`) were copied. Tests that read them are skipped.

## Files

| Test file | Source | Passing | Skipped |
|---|---|---|---|
| `workbook.test.ts` | `spec/unit/doc/workbook.spec.js` | 9 | 1 |
| `row.test.ts` | `spec/unit/doc/row.spec.js` | 5 | 9 |
| `cell.test.ts` | `spec/unit/doc/cell.spec.js` | 5 | 7 |
| `column.test.ts` | `spec/unit/doc/column.spec.js` | 6 | 3 |
| `worksheet-merge.test.ts` | `spec/unit/doc/worksheet.merge.spec.js` | 5 | 0 |
| `worksheet-values.test.ts` | `spec/unit/doc/worksheet.values.spec.js` | 16 | 12 |
| `integration-workbook.test.ts` | `spec/integration/workbook/workbook.spec.js` | 12 | 23 |
| `integration-worksheet.test.ts` | `spec/integration/worksheet.spec.js` | 15 | 8 |
| `integration-styles.test.ts` | `spec/integration/workbook/styles.spec.js` | 1 | 4 |
| `workbook-writer.test.ts` | `spec/integration/workbook-xlsx-writer/workbook-xlsx-writer.spec.js` | 9 | 10 |
| `workbook-reader.test.ts` | `spec/integration/workbook-xlsx-reader.spec.js` | 1 (added, not upstream) | 29 |

"Passing" includes tests that expect `XLV_E_NOT_SUPPORTED` for an unsupported feature.

## Why tests are skipped or changed

Each skipped test names its reason in the title. The reasons are:

- **Unsupported features** (they throw `XLV_E_NOT_SUPPORTED`; where cheap, one test checks that):
  rich text and hyperlink values, notes, images, conditional formatting, data validation, CSV,
  defined names, print settings, `insertRow`, `spliceRows`, `spliceColumns`, `duplicateRow`,
  `addRow(values, 'i')`, shared-formula filling, zip compression options, split panes and
  workbook views.
- **Internal API**: tests of internal classes and fields (`new Cell`, `Column.toModel`,
  `row.model`, `_upgradeToHyperlink`, `effectiveType`, sheet mocks) are skipped or rewritten
  with the public API.
- **Fixtures**: `createTestBook`/`checkTestBook` from `spec/utils/index.js` build workbooks with
  unsupported features; tests based on them, and tests reading `spec/integration/data/*`, are
  skipped.
- **What xlvane does not keep after writing and reading a file**: font `family` and `outline`,
  gradient fills, cells that have a style but no value, row and column default styles
  (docs/decisions.md D-206), sheet views other than the first frozen pane and document
  properties (not written in phase 5). Round-trip tests use the values xlvane keeps
  (`styles.roundTripFonts`, a pattern fill, one frozen view) and do not read back row or column
  styles.
- **Messages**: worksheet-name errors are checked with `toThrow()`, without the exceljs message
  text.
