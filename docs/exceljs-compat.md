# exceljs compatibility (`@xlvane/exceljs`)

`@xlvane/exceljs` offers the commonly used part of the exceljs 4.4.0 API on top of xlvane, so that code written for exceljs keeps working after one change in `package.json`:

```json
"dependencies": {
  "exceljs": "npm:@xlvane/exceljs@^1"
}
```

npm installs an alias `<alias>@npm:<name>` under the alias name in `node_modules` (npm/cli `docs/lib/content/using-npm/package-spec.md`, "Aliases"), so `require('exceljs')` and `import ExcelJS from 'exceljs'` load `@xlvane/exceljs`. The package is not published yet.

- **ESM:** `import ExcelJS from 'exceljs'` works as with exceljs. `@xlvane/exceljs` is an ES module, so named imports (`import { Workbook } from 'exceljs'`) work too; exceljs 4.4.0 rejects them because it is a CommonJS package.
- **CommonJS:** `const ExcelJS = require('exceljs')` needs **Node.js 22.12 or later**. The package is ESM-only and exports the `ExcelJS` object under the name `"module.exports"`, which `require()` of an ES module returns directly from Node.js 22.12 on (Node.js `doc/api/modules.md`, "Loading ECMAScript modules using `require()`"). `engines.node` is `>=22.12`.
- **Unsupported API** throws an `XlvError` with code `XLV_E_NOT_SUPPORTED`, naming the method (docs/errors.md). Nothing is dropped silently.

Status: **supported** means the same behaviour as exceljs for the cases covered by the ported tests (`tests/exceljs/`). **Partly** means supported with the difference given in the notes. **Not supported** means it throws `XLV_E_NOT_SUPPORTED`. The source column points to the exceljs 4.4.0 implementation (paths relative to the package root, as installed in `bench/node_modules/exceljs`).

## Workbook

| API | Status | Notes | Source in exceljs |
|---|---|---|---|
| `new Workbook()` | supported | | `lib/doc/workbook.js` |
| `creator`, `lastModifiedBy`, `created`, `modified` | partly | Stored on the object, not written to the file (xlvane writes no `docProps` part yet, W2-33). | `lib/doc/workbook.js` L18, L22, L153–158 |
| `addWorksheet(name, { state, views, properties })` | partly | `properties.tabColor` and other sheet properties are not written. | `lib/doc/workbook.js` L53 |
| `getWorksheet(name \| id)` | supported | | `lib/doc/workbook.js` L106 |
| `worksheets`, `eachSheet(cb)` | supported | | `lib/doc/workbook.js` L119, L127 |
| `removeWorksheet(id)` | supported | | `lib/doc/workbook.js` L99 |
| `xlsx.load(buffer)` | supported | Reading goes through xlvane's limits for untrusted files. | `lib/xlsx/xlsx.js` L257 |
| `xlsx.readFile(path)` | supported | | `lib/xlsx/xlsx.js` L49 |
| `xlsx.read(stream)` | supported | | `lib/xlsx/xlsx.js` L244 |
| `xlsx.writeBuffer()` | supported | | `lib/xlsx/xlsx.js` L685 |
| `xlsx.writeFile(path)` | supported | | `lib/xlsx/xlsx.js` L666 |
| `xlsx.write(stream)` | supported | | `lib/xlsx/xlsx.js` L643 |
| `csv` (`workbook.csv.*`) | not supported | Use `toCSV` / `fromCSV` from `xlvane`. | `lib/doc/workbook.js` L38 |
| `definedNames` | not supported | | `lib/doc/workbook.js` L133 |
| `addImage(image)` | not supported | | `lib/doc/workbook.js` L142 |

## Worksheet

| API | Status | Notes | Source in exceljs |
|---|---|---|---|
| `name`, `id`, `state` | supported | | `lib/doc/worksheet.js` L33 |
| `columns` (get / set `{ header, key, width, style }`) | supported | | `lib/doc/worksheet.js` L203, L209 |
| `getColumn(key \| letter \| number)` | supported | | `lib/doc/worksheet.js` L243 |
| `addRow(values \| object, style?)`, `addRows` | supported | | `lib/doc/worksheet.js` L406, L414 |
| `getRow(n)`, `getRows(start, length)` | supported | | `lib/doc/worksheet.js` L388, L397 |
| `getCell(ref \| row, col)` | supported | | `lib/doc/worksheet.js` L592 |
| `eachRow({ includeEmpty }, cb)` | supported | | `lib/doc/worksheet.js` L551 |
| `rowCount`, `actualRowCount`, `columnCount` | supported | | `lib/doc/worksheet.js` L374, L378, L313 |
| `mergeCells(range \| t, l, b, r)`, `unMergeCells` | supported | | `lib/doc/worksheet.js` L602, L656 |
| `views` (`state: 'frozen'`, `xSplit`, `ySplit`) | partly | Only frozen panes are written and read; split panes, zoom and other view settings are not. | `lib/doc/worksheet.js` L114; `index.d.ts` L84 |
| `autoFilter` (range string) | supported | | `lib/doc/worksheet.js` L116 |
| `insertRow`, `insertRows` | not supported | | `lib/doc/worksheet.js` L422 |
| `spliceRows` | not supported | | `lib/doc/worksheet.js` L483 |
| `duplicateRow` | not supported | | `lib/doc/worksheet.js` L463 |
| `addImage` | not supported | | `lib/doc/worksheet.js` L725 |
| `addTable` | not supported | | `lib/doc/worksheet.js` L791 |
| `addConditionalFormatting` | not supported | | `lib/doc/worksheet.js` L811 |
| `dataValidations` | not supported | | `lib/doc/worksheet.js` L111 |
| `protect`, `unprotect` | not supported | | `lib/doc/worksheet.js` L753, L785 |
| `pageSetup`, `headerFooter` (print settings) | not supported | | `lib/doc/worksheet.js` L64, L96 |

## Row

| API | Status | Notes | Source in exceljs |
|---|---|---|---|
| `number` | supported | | `lib/doc/row.js` L19 |
| `values` (get / set; sparse, 1-based) | supported | | `lib/doc/row.js` L166, L177 |
| `getCell(col)` | supported | | `lib/doc/row.js` L56 |
| `eachCell({ includeEmpty }, cb)` | supported | As in exceljs, cells of a merged range give the master cell's value. | `lib/doc/row.js` L130 |
| `height`, `hidden` | supported | | `lib/doc/row.js` L310 |
| `font`, `fill`, `border`, `alignment`, `numFmt` setters | supported | Apply to the row's cells. | `lib/doc/row.js` L274, L306 |
| `commit()` | supported | Meaningful for `WorkbookWriter`; does nothing in a normal workbook, as in exceljs. | `lib/doc/row.js` L29 |
| `cellCount`, `actualCellCount` | supported | | `lib/doc/row.js` L216, L220 |

## Column

| API | Status | Notes | Source in exceljs |
|---|---|---|---|
| `key`, `header`, `width`, `hidden` | supported | | `lib/doc/column.js` L77, L92, L108 |
| `letter`, `number` | supported | | `lib/doc/column.js` L23, L31 |
| `values` | supported | | `lib/doc/column.js` L179 |
| `eachCell({ includeEmpty }, cb)` | supported | | `lib/doc/column.js` L168 |
| `font`, `fill`, `border`, `alignment`, `numFmt` setters | partly | Applied to the column's existing cells; xlvane has no column default style yet (D-206), so cells added later do not inherit it. | `lib/doc/column.js` L226 |

## Cell values

| Value | Status | Notes | Source in exceljs |
|---|---|---|---|
| `null`, number, string, boolean | supported | | `lib/doc/cell.js` L206, L211 |
| `Date` | partly | Written with the format `yyyy-mm-dd` (or `yyyy-mm-dd hh:mm:ss`) unless `numFmt` is set; exceljs uses built-in format 14, which applications show in the local date order. | `lib/doc/cell.js` L538 |
| `{ formula, result }` | supported | xlvane does not calculate; `result` is stored as given. | `lib/doc/cell.js` L734 |
| `{ sharedFormula, result }` | partly | Written as an ordinary formula. When reading, xlvane translates shared formulas, so such cells come back as `{ formula, result }`. | `lib/doc/cell.js` L744, L811; `index.d.ts` L357 |
| `{ error: '#N/A' }` | supported | | `lib/doc/cell.js` L970 |
| `{ richText: [...] }` | partly | Writing throws `XLV_E_NOT_SUPPORTED`; reading returns the plain text. | `lib/doc/cell.js` L494; `index.d.ts` L341 |
| `{ text, hyperlink }` | not supported | Writing throws; hyperlinks in read files are not reported. | `lib/doc/cell.js` L582; `index.d.ts` L345 |
| `note` | not supported | | `lib/doc/cell.js` L224 |
| `type` (`ValueType` numbers) | supported | Same numbers as exceljs: Null 0, Merge 1, Number 2, String 3, Date 4, Hyperlink 5, Formula 6, SharedString 7, RichText 8, Boolean 9, Error 10. | `lib/doc/enums.js` L4–16; `lib/doc/cell.js` L142 |
| `text`, `address`, `row`, `col` | supported | | `lib/doc/cell.js` L123, L232 |
| `isMerged`, `master` | supported | | `lib/doc/cell.js` L165, L190 |

## Styles

| API | Status | Notes | Source in exceljs |
|---|---|---|---|
| `numFmt` | supported | | `lib/doc/cell.js` L51 |
| `font` (`name`, `size`, `bold`, `italic`, `underline`, `strike`, `color`) | partly | `family`, `scheme`, `charset`, `outline`, `vertAlign` are not written. | `lib/doc/cell.js` L59; `index.d.ts` L217 |
| `fill` (`type: 'pattern'`, `pattern`, `fgColor`, `bgColor`) | partly | Gradient fills are not supported (`XLV_E_NOT_SUPPORTED`). | `lib/doc/cell.js` L83; `index.d.ts` L161–215 |
| `border` (`top`, `left`, `bottom`, `right`) | partly | `diagonal` is not supported. | `lib/doc/cell.js` L75; `index.d.ts` L258 |
| `alignment` (`horizontal`, `vertical`, `wrapText`, `indent`, `textRotation`) | partly | `shrinkToFit`, `readingOrder` are not written. | `lib/doc/cell.js` L67; `index.d.ts` L280 |
| Colours `{ argb }`, `{ theme, tint }`, `{ indexed }` | supported | | `index.d.ts` |

## Streaming

| API | Status | Notes | Source in exceljs |
|---|---|---|---|
| `stream.xlsx.WorkbookWriter({ filename \| stream, useStyles, useSharedStrings })` | supported | Built on `writeWorkbookStream`. | `lib/stream/xlsx/workbook-writer.js` L22, L33, L37 |
| `writer.addWorksheet`, `worksheet.addRow(...).commit()`, `worksheet.commit()`, `await writer.commit()` | supported | | `lib/stream/xlsx/workbook-writer.js` L97, L136; `lib/stream/xlsx/worksheet-writer.js` L215, L407 |
| Sheet layout (`columns`, `views`, `autoFilter`, merges) of a streamed worksheet | partly | Must be set before the worksheet's first row is committed. A change made later throws `XLV_E_NOT_SUPPORTED` (at `mergeCells()` or at `worksheet.commit()`), because xlvane writes the layout before the rows. exceljs writes `autoFilter` and merges when the worksheet is committed. | `lib/stream/xlsx/worksheet-writer.js` |
| `stream.xlsx.WorkbookReader` with `for await` over worksheets and rows | supported | Built on `openWorkbookStream`. | `lib/stream/xlsx/workbook-reader.js` L19, L71; `lib/stream/xlsx/worksheet-reader.js` L105 |

## Other

| API | Status | Notes |
|---|---|---|
| Module shape: default export, `require()` | supported | See above; `require()` needs Node.js 22.12+. |
| Named ESM imports | extra | Not available with exceljs 4.4.0. |
| `exceljs/dist/es5` and other deep imports | not supported | Only the package's main entry is offered. |

## Known differences found by the cross-check

`tests/exceljs/cross-check.test.ts` runs `tests/exceljs/sample-project` against exceljs 4.4.0 and stores its output in `expected-exceljs-4.4.0.json`. `tests/exceljs/alias.test.ts` requires @xlvane/exceljs to produce the same output, except for the entries of `KNOWN_DIFFERENCES` in `tests/exceljs/project-helper.ts`, which are listed here:

- None so far for what the sample reads back.
- Not covered by the sample's summary: the display format of dates written without `numFmt` (see "Cell values").
