# exceljs compatibility (`@xlvane/exceljs`)

`@xlvane/exceljs` offers the commonly used part of the exceljs 4.4.0 API on top of xlvane, so that code written for exceljs keeps working after one change in `package.json`:

```json
"dependencies": {
  "exceljs": "npm:@xlvane/exceljs@^1"
}
```

npm installs an alias `<alias>@npm:<name>` under the alias name in `node_modules` (npm/cli `docs/lib/content/using-npm/package-spec.md`, "Aliases"), so `require('exceljs')` and `import ExcelJS from 'exceljs'` load `@xlvane/exceljs`.

- **ESM:** `import ExcelJS from 'exceljs'` works as with exceljs. `@xlvane/exceljs` is an ES module, so named imports (`import { Workbook } from 'exceljs'`) work too; exceljs 4.4.0 rejects them because it is a CommonJS package.
- **CommonJS:** `const ExcelJS = require('exceljs')` needs **Node.js 22.12 or later**. The package is ESM-only and exports the `ExcelJS` object under the name `"module.exports"`, which `require()` of an ES module returns directly from Node.js 22.12 on (Node.js `doc/api/modules.md`, "Loading ECMAScript modules using `require()`"). `engines.node` is `>=22.12`.
- **Unsupported API** throws an `XlvError` with code `XLV_E_NOT_SUPPORTED`, naming the method (docs/errors.md). Nothing is dropped silently.

Status: **supported** means the same behaviour as exceljs for the cases covered by the ported tests (`tests/exceljs/`). **Partly** means supported with the difference given in the notes. **Not supported** means it throws `XLV_E_NOT_SUPPORTED`. The source column points to the exceljs 4.4.0 implementation (paths relative to the package root, as installed in `bench/node_modules/exceljs`).

## Workbook

| API | Status | Notes | Source in exceljs |
|---|---|---|---|
| `new Workbook()` | supported | | `lib/doc/workbook.js` |
| `creator`, `lastModifiedBy`, `created`, `modified`, `lastPrinted`, `title`, `subject`, `keywords`, `category`, `description`, `language`, `revision`, `contentStatus` | supported | Written to `docProps/core.xml` with exceljs's defaults (`creator` and `lastModifiedBy` "Unknown", `created` and `modified` the time the workbook was made), and read back; `revision` reads back as a number. Dates are stored in whole seconds (UTC). | `lib/doc/workbook.js` L16–27, L153–180; `lib/xlsx/xform/core/core-xform.js` |
| `company`, `manager` | supported | Written to `docProps/app.xml` as exceljs writes them: `Company` always (`''` when not set), `Manager` when not `undefined`; read back from the file (`undefined` when missing). exceljs also writes `Application` "Microsoft Excel" and `AppVersion` "16.0300"; xlvane does not claim an application (D-A10), which exceljs does not read back. | `lib/xlsx/xform/core/app-xform.js` L24–37; `lib/xlsx/xlsx.js` L345–349 |
| `addWorksheet(name, { state, views, properties })` | partly | `properties.tabColor` (and the deprecated `worksheet.tabColor`) is written and read back; other sheet properties (`defaultRowHeight`, outline levels and properties, `dyDescent`) are not written. | `lib/doc/workbook.js` L53; `lib/xlsx/xform/sheet/sheet-properties-xform.js` |
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
| `definedNames` (`add`, `addEx`, `remove`, `removeEx`, `removeAllNames`, `forEach`, `getNames`, `getNamesEx`, `getRanges`, `getMatrix`, `model`) | supported | As in exceljs, a name is a set of cells, written as rectangles with absolute references and no sheet scope; on reading, names that are not cell ranges (constants, formulas, `#REF!`) and the sheet scope (`localSheetId`) are dropped. Built-in `_xlnm.` names are not listed: print areas and titles (exceljs `pageSetup`) are not supported yet, and xlvane writes the autofilter name itself. Names exceljs accepts but Excel does not (for example `A1` or names starting with `_xl`) throw `XLV_E_INVALID_DEFINED_NAME` when the workbook is written. | `lib/doc/defined-names.js`; `lib/xlsx/xform/book/defined-name-xform.js`; `lib/xlsx/xform/book/workbook-xform.js` L187 |
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
| `views` | partly | The first view is written: frozen panes, split panes (`xSplit`/`ySplit` in 1/20 point), `zoomScale` (10–400), `showGridLines`, `showRowColHeaders`, `rightToLeft`, `style` (`pageBreakPreview`, `pageLayout`). Not written: further views, `topLeftCell` (xlvane writes the default cell for frozen panes and none for a split), `activeCell`, `activePane` (xlvane writes the pane the split implies), `showRuler`, `zoomScaleNormal`, `tabSelected`. Every worksheet reads back with one view shaped as exceljs reads it, with `showRuler: true` and `zoomScaleNormal: 100`. | `lib/doc/worksheet.js` L114; `lib/xlsx/xform/sheet/sheet-view-xform.js`; `index.d.ts` L33–149 |
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
| `font`, `fill`, `border`, `alignment`, `numFmt` setters, `style` | supported | Applied to the column's existing cells, inherited by cells added later (the row style wins), and written as the column's default style (`col@style`). As in exceljs, a styled column without a width is written with width 9, and a column with width 9 and nothing else is not written. The streaming writer writes the column style only with `useStyles: true`. | `lib/doc/column.js` L147, L206, L265 |

## Cell values

| Value | Status | Notes | Source in exceljs |
|---|---|---|---|
| `null`, number, string, boolean | supported | | `lib/doc/cell.js` L206, L211 |
| `Date` | supported | Without `numFmt`, a date (also a formula result that is a date) is written with built-in format 14, as exceljs does, also by the streaming writer without `useStyles`; applications show it in the reader's date order, and reading gives `numFmt` `"mm-dd-yy"`. | `lib/doc/cell.js` L538; `lib/xlsx/xform/style/styles-xform.js` L253–266, L498–515 |
| `{ formula, result }` | supported | xlvane does not calculate; `result` is stored as given. | `lib/doc/cell.js` L734 |
| `{ sharedFormula, result }` | partly | Written as an ordinary formula. When reading, xlvane translates shared formulas, so such cells come back as `{ formula, result }`. | `lib/doc/cell.js` L744, L811; `index.d.ts` L357 |
| `{ error: '#N/A' }` | supported | | `lib/doc/cell.js` L970 |
| `{ richText: [...] }` | partly | Written and read back with the run fonts' `name`, `size`, `bold`, `italic`, `underline`, `strike` and `color`; `cell.text` joins the runs and `type` is `RichText`. `family`, `scheme` and `charset` are dropped (not written, D-A01), and `vertAlign` (superscript, subscript) and `outline` in a run throw `XLV_E_NOT_SUPPORTED`. A 6-digit `argb` colour reads back with alpha `FF` (8 digits). | `lib/doc/cell.js` L494; `lib/xlsx/xform/strings/rich-text-xform.js`; `index.d.ts` L336–343 |
| `{ text, hyperlink, tooltip }` | partly | Written and read back; `type` is `Hyperlink`, `cell.hyperlink` the target and `cell.text` the text (also rich text). Differences, because xlvane follows the specification (D-A08): a `hyperlink` starting with `#` is written as a place in the workbook (`location`, no relationship), where exceljs writes every link as an external relationship; the `tooltip` is read back, which exceljs drops; links stored only as a location (as Excel writes links inside a workbook) are read as `#` + location, which exceljs drops. Limits of [MS-OI29500] apply (`XLV_E_INVALID_HYPERLINK`, e.g. a tooltip over 255 characters). The streaming `WorkbookWriter` writes links with the row they are in; the streaming `WorkbookReader` ignores them, as exceljs does by default. | `lib/doc/cell.js` L582–666, L1078; `lib/xlsx/xform/sheet/hyperlink-xform.js`; `index.d.ts` L345 |
| `note` (string or `{ texts, margins, protection, editAs }`) | partly | Written and read back, also by the streaming `WorkbookWriter`; every note is written with the author "Author", as exceljs does, and a note of one run without a font reads back as a string. Differences: `margins`, `protection` and `editAs` are kept on the object but not written (xlvane writes every note box in the form of the ECMA-376 examples, D-A12), so they are absent after reading, where exceljs reads them back from the VML; run fonts lose `family`, `scheme` and `charset` as for rich text; a note on a cell without a value is kept, where exceljs loses it on reading. The streaming `WorkbookReader` ignores notes, as exceljs does. | `lib/doc/cell.js` L224–230; `lib/doc/note.js`; `lib/xlsx/xform/comment/comments-xform.js` L29–31 |
| `type` (`ValueType` numbers) | supported | Same numbers as exceljs: Null 0, Merge 1, Number 2, String 3, Date 4, Hyperlink 5, Formula 6, SharedString 7, RichText 8, Boolean 9, Error 10. | `lib/doc/enums.js` L4–16; `lib/doc/cell.js` L142 |
| `text`, `address`, `row`, `col` | supported | | `lib/doc/cell.js` L123, L232 |
| `isMerged`, `master` | supported | | `lib/doc/cell.js` L165, L190 |
| `name`, `names`, `addName`, `removeName`, `removeAllNames`, `fullAddress` | supported | Through `workbook.definedNames`. | `lib/doc/cell.js` L270–310 |

## Styles

| API | Status | Notes | Source in exceljs |
|---|---|---|---|
| `numFmt` | supported | Codes that exceljs maps to built-in formats (`'0.00'`, `'mm-dd-yy'`, `'@'`, …) are written as the built-in id, as exceljs does. | `lib/doc/cell.js` L51; `lib/xlsx/defaultnumformats.js`; `lib/xlsx/xform/style/styles-xform.js` L374 |
| `font` (`name`, `size`, `bold`, `italic`, `underline`, `strike`, `color`) | partly | `family`, `scheme`, `charset`, `outline`, `vertAlign` are not written. A font without `name` or `size` is written with the workbook's default font for the missing ones (Calibri 11), so it reads back with them; exceljs writes only what is given. | `lib/doc/cell.js` L59; `index.d.ts` L217 |
| `fill` (`type: 'pattern'`, `pattern`, `fgColor`, `bgColor`) | partly | Gradient fills are not supported (`XLV_E_NOT_SUPPORTED`). | `lib/doc/cell.js` L83; `index.d.ts` L161–215 |
| `border` (`top`, `left`, `bottom`, `right`) | partly | `diagonal` is not supported. | `lib/doc/cell.js` L75; `index.d.ts` L258 |
| `alignment` (`horizontal`, `vertical`, `wrapText`, `indent`, `textRotation`) | partly | `shrinkToFit`, `readingOrder` are not written. | `lib/doc/cell.js` L67; `index.d.ts` L280 |
| Colours `{ argb }`, `{ theme, tint }`, `{ indexed }` | supported | | `index.d.ts` |

## Streaming

| API | Status | Notes | Source in exceljs |
|---|---|---|---|
| `stream.xlsx.WorkbookWriter({ filename \| stream, useStyles, useSharedStrings })` | supported | Built on `writeWorkbookStream`. | `lib/stream/xlsx/workbook-writer.js` L22, L33, L37 |
| `writer.addWorksheet`, `worksheet.addRow(...).commit()`, `worksheet.commit()`, `await writer.commit()` | supported | | `lib/stream/xlsx/workbook-writer.js` L97, L136; `lib/stream/xlsx/worksheet-writer.js` L215, L407 |
| Sheet layout (`columns`, `views`, `properties.tabColor`, `autoFilter`, merges) of a streamed worksheet | partly | Must be set before the worksheet's first row is committed. A change made later throws `XLV_E_NOT_SUPPORTED` (at `mergeCells()` or at `worksheet.commit()`), because xlvane writes the layout before the rows. exceljs writes `autoFilter` and merges when the worksheet is committed. | `lib/stream/xlsx/worksheet-writer.js` |
| `WorkbookWriter` options `created`, `modified`, `creator`, `lastModifiedBy`, `lastPrinted`; `definedNames` | partly | Written with exceljs's defaults (`creator` and `lastModifiedBy` "ExcelJS"). Names and these properties are written when the first row is committed (writing starts then); a change made later throws `XLV_E_NOT_SUPPORTED` at `writer.commit()`. exceljs writes them at the end. | `lib/stream/xlsx/workbook-writer.js` L22–40, L64, L264 |
| `stream.xlsx.WorkbookReader` with `for await` over worksheets and rows | supported | Built on `openWorkbookStream`. Rich text cells come back as `{ richText }`. | `lib/stream/xlsx/workbook-reader.js` L19, L71; `lib/stream/xlsx/worksheet-reader.js` L105 |

## Other

| API | Status | Notes |
|---|---|---|
| Module shape: default export, `require()` | supported | See above; `require()` needs Node.js 22.12+. |
| Named ESM imports | extra | Not available with exceljs 4.4.0. |
| `exceljs/dist/es5` and other deep imports | not supported | Only the package's main entry is offered. |

## Known differences found by the cross-check

`tests/exceljs/cross-check.test.ts` runs `tests/exceljs/sample-project` against exceljs 4.4.0 and stores its output in `expected-exceljs-4.4.0.json`. `tests/exceljs/alias.test.ts` requires @xlvane/exceljs to produce the same output, except for the entries of `KNOWN_DIFFERENCES` in `tests/exceljs/project-helper.ts`, which are listed here:

- None so far for what the sample reads back. The sample covers document properties (including `company` and `manager`), rich text, a column style, defined names (including `cell.name`), the built-in date format, a tab colour, a zoomed frozen view and external, `mailto:` and `#` links (without tooltips, which exceljs does not read back); `alias.test.ts` also checks that LibreOffice shows the dates as exceljs's files do (`1/31/2024`).
