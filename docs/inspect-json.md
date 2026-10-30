# `inspect()` JSON schema, version 1

`inspect(input, options?)` (and `inspectFile(path, options?)` from `xlvane/node`) returns plain JSON data describing a workbook. `JSON.stringify` of the result is lossless. The field `schemaVersion` changes only when a field is removed or changes meaning; new fields may be added within a version.

## Options

| Option | Default | Meaning |
|---|---|---|
| `limits` | see `DEFAULT_READ_LIMITS` | resource limits used while reading the file |
| `maxColumns` | 200 | maximum number of columns described per sheet; `columnsTruncated` says whether more exist |
| `headerScanRows` | 20 | number of leading non-empty rows searched for a header row |
| `formulaSamples` | 10 | maximum number of example formulas per sheet |

## Workbook

| Field | Type | Meaning |
|---|---|---|
| `schemaVersion` | `1` | version of this schema |
| `format` | `"xlsx" \| "xlsm" \| "xltx" \| "xltm" \| "unknown"` | from the content type of the workbook part |
| `conformance` | `"transitional" \| "strict"` | Office Open XML conformance class of the package |
| `dateSystem` | `"1900" \| "1904"` | date system used for serial dates |
| `hasMacros` | boolean | the package contains a VBA project or Excel 4.0 macro sheets (never run) |
| `externalLinkCount` | number | links to other workbooks (never followed) |
| `properties` | object | document properties that are set (`title`, `creator`, `created`, …, from `docProps/core.xml`); all values are strings, dates as ISO 8601 in UTC (`"2026-10-01T00:00:00.000Z"`); `{}` when the file has none |
| `definedNames` | `{ name, refersTo, sheet, hidden }[]` | defined names in file order, at most 100; `refersTo` is the formula text without "=", cut to 100 characters; `sheet` is `null` for a workbook-level name; built-in names such as `_xlnm._FilterDatabase` are included |
| `definedNamesTruncated` | boolean | the workbook has more than 100 defined names |
| `sheetCount` | number | number of sheets |
| `sheets` | array of sheet objects | in workbook order |
| `warnings` | array of warning objects | see below |

## Sheet

| Field | Type | Meaning |
|---|---|---|
| `name` | string | sheet name |
| `index` | number | 0-based position |
| `kind` | `"worksheet" \| "chartsheet" \| "dialogsheet" \| "macrosheet"` | only worksheets have cells |
| `state` | `"visible" \| "hidden" \| "veryHidden"` | visibility |
| `dimension` | string or `null` | range declared in the file, e.g. `"A1:D20"` |
| `usedRange` | string or `null` | smallest range containing every cell with a value or formula; `null` for an empty sheet |
| `rowCount`, `columnCount` | number | size of `usedRange` |
| `cellCount` | number | cells with a value or formula |
| `formulaCount` | number | cells with a formula |
| `mergedRanges` | string[] | merged ranges, e.g. `["A1:D1"]` |
| `headerRow` | `{ row, confidence }` or `null` | detected header row (1-based) and `"high"` or `"low"` confidence |
| `columns` | array of column objects | one per column of `usedRange`, at most `maxColumns` |
| `columnsTruncated` | boolean | `usedRange` has more columns than were described |
| `formulas` | `{ cell, formula }[]` | up to `formulaSamples` examples; `formula` is the text without "=" or `null` for cells that reuse a shared formula |
| `richTextCellCount` | number | cells whose text has formatted runs (`Cell.richText`) |
| `richTextCells` | string[] | up to ten of those cells |
| `styledColumns` | string[] | columns with a default style (`col@style`, `ColumnInfo.style`), e.g. `["B", "D:F"]`; at most `maxColumns` entries |
| `hyperlinkCount` | number | hyperlinks of the sheet (`Sheet.hyperlinks`) |
| `hyperlinks` | object[] | up to ten of them: `ref`, `target` (external URL or path as stored, or `null`), `location` (place in the workbook, or `null`); targets are never opened |

**Header detection.** Among the first `headerScanRows` non-empty rows, the header row is the first one whose cells are all distinct, non-blank strings and cover at least half of the used width. The confidence is `"high"` when the next non-empty row contains a value that is not a string (number, date, boolean or error), otherwise `"low"`. This is a heuristic; check `confidence` before relying on it.

## Column

| Field | Type | Meaning |
|---|---|---|
| `index` | number | 1-based column number |
| `letter` | string | column letters, e.g. `"C"` |
| `header` | string or `null` | text of the header cell, `null` without a header row |
| `dominantType` | `"number" \| "string" \| "date" \| "boolean" \| "error" \| "empty"` | most frequent non-empty type in the data rows (ties: number, string, date, boolean, error) |
| `types` | `{ number, string, boolean, date, error, empty }` | counts over the data rows: the rows below the header row (or all rows without one) down to the end of `usedRange` |
| `samples` | array | up to three distinct example values from the data rows |

Values in `samples` are JSON-safe: numbers, booleans, strings (cut to 100 characters) and dates as ISO 8601 strings **without a time zone** (`"2024-01-31T13:45:00.000"`), because spreadsheet dates have no time zone. Error cells appear as their code, e.g. `"#DIV/0!"`.

## Warning

| Field | Type | Meaning |
|---|---|---|
| `code` | string | `XLV_W_*` code, see docs/errors.md |
| `message` | string | what was found and what xlvane did |
| `location` | object | `sheet` and/or `part` where it applies |
| `count` | number | number of occurrences (warnings are aggregated per sheet) |
| `cells` | string[] (optional) | up to ten example cells |

## Example

```json
{
  "schemaVersion": 1,
  "format": "xlsx",
  "conformance": "transitional",
  "dateSystem": "1900",
  "hasMacros": false,
  "externalLinkCount": 0,
  "properties": { "title": "Sales 2024", "creator": "Finance" },
  "definedNames": [{ "name": "Total", "refersTo": "Sales!$C$3", "sheet": null, "hidden": false }],
  "definedNamesTruncated": false,
  "sheetCount": 1,
  "sheets": [
    {
      "name": "Sales",
      "index": 0,
      "kind": "worksheet",
      "state": "visible",
      "dimension": "A1:C3",
      "usedRange": "A1:C3",
      "cellCount": 9,
      "formulaCount": 1,
      "mergedRanges": [],
      "formulas": [{ "cell": "C3", "formula": "SUM(C2:C2)" }],
      "richTextCellCount": 0,
      "richTextCells": [],
      "styledColumns": [],
      "hyperlinkCount": 0,
      "hyperlinks": [],
      "rowCount": 3,
      "columnCount": 3,
      "headerRow": { "row": 1, "confidence": "high" },
      "columns": [
        { "index": 1, "letter": "A", "header": "Date", "dominantType": "date", "types": { "number": 0, "string": 0, "boolean": 0, "date": 2, "error": 0, "empty": 0 }, "samples": ["2024-01-31T00:00:00.000", "2024-02-01T00:00:00.000"] },
        { "index": 2, "letter": "B", "header": "Customer", "dominantType": "string", "types": { "number": 0, "string": 1, "boolean": 0, "date": 0, "error": 0, "empty": 1 }, "samples": ["ACME"] },
        { "index": 3, "letter": "C", "header": "Amount", "dominantType": "number", "types": { "number": 2, "string": 0, "boolean": 0, "date": 0, "error": 0, "empty": 0 }, "samples": [100.5] }
      ],
      "columnsTruncated": false
    }
  ],
  "warnings": []
}
```
