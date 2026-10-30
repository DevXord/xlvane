# Errors and warnings

Every problem xlvane reports has a stable code. Errors are thrown as `XlvError` (`code`, `message`, `location`, `hint`, `docsUrl`, `details`; `error.toJSON()` returns the same as data). Warnings never stop reading; they are listed in `workbook.warnings` and in the `warnings` array of `inspect()`, aggregated per sheet with a `count` and up to ten example `cells`.

`location` contains only the fields that are known: `part` (part name inside the package, e.g. `/xl/worksheets/sheet1.xml`), `sheet`, `cell` (A1 reference), `line` and `column` (inside an XML part) and `offset` (byte offset in the file, for ZIP structures).

Codes are never reused for a different condition. New codes are added here first.

## Errors

### XLV_E_INVALID_INPUT
The input is not what the function expects: `readWorkbook` and `inspect` accept a `Uint8Array`, an `ArrayBuffer` (or another typed-array view) and, for `inspect` and the table functions, a workbook returned by `readWorkbook`. `editCells` and `appendRows` need the original bytes, because unchanged parts are copied from them. `fromCSV` needs text or UTF-8 bytes.
**Fix:** pass the bytes of the file. A path string is the most common mistake; in Node.js use `readWorkbookFile()` or `inspectFile()` from `xlvane/node`.

### XLV_E_INVALID_OPTION
An option is unknown, has the wrong type or is out of range. The hint lists the known options and suggests the closest name. `details.option` names the option.
**Fix:** correct the option name or value.

### XLV_E_LIMIT_EXCEEDED
The file exceeds one of the resource limits (`ReadLimits`): input size, number of ZIP entries, decompressed size of one entry or of all entries, compression ratio, XML nesting depth, attribute count, name or text length, number of elements in a tree-parsed part, or number of cells. `details.limit` names the limit and `details.value` its current value. Sizes are measured on the bytes actually produced by decompression, so a file cannot avoid a limit by declaring false sizes.
**Fix:** if the file comes from a trusted source and is simply large, raise the named limit, e.g. `{ limits: { maxCells: 50_000_000 } }`. A compression-ratio error on a small file usually means a zip bomb.

### XLV_E_NOT_ZIP
The input is not a ZIP archive, so it cannot be an XLSX file. CSV, HTML or text files saved with an `.xlsx` extension produce this error.
**Fix:** check what the file really is; use a CSV reader for CSV data.

### XLV_E_UNSUPPORTED_CONTAINER
The input is a Compound File Binary container (signature `D0 CF 11 E0 A1 B1 1A E1`). This is either a legacy `.xls` workbook or a password-protected `.xlsx` workbook, which Office stores encrypted inside such a container.
**Fix:** open the file in a spreadsheet application and save it as an unencrypted `.xlsx` workbook.

### XLV_E_ZIP_CORRUPT
The ZIP structure is damaged or ambiguous. `details.reason` says what was found:

| reason | meaning |
|---|---|
| `eocd-missing` | the end of central directory record is missing (usually a truncated download) |
| `trailing-data` | bytes follow the end of central directory record |
| `leading-data` | bytes precede the first entry (self-extracting archive or a file disguised as another format) |
| `central-directory-bounds`, `central-directory-entry`, `central-directory-size` | the central directory is not where it should be or does not match its entries |
| `zip64-locator`, `zip64-record`, `zip64-extra`, `zip64-value` | ZIP64 structures are missing or inconsistent |
| `extra-field` | an extra field runs past its declared length |
| `name-encoding` | an entry name flagged as UTF-8 is not valid UTF-8 |
| `duplicate-name` | two entries have the same name, also when the names differ only in letter case |
| `local-header`, `local-header-mismatch`, `entry-bounds` | a local file header is missing, disagrees with the central directory, or its data runs into the central directory |
| `overlapping-entries` | two entries share bytes — the technique used by non-recursive zip bombs |
| `size-mismatch` | an entry decompresses to a different size than its header declares |
| `crc-mismatch` | an entry fails its CRC-32 check |
| `deflate` | an entry is not a valid Deflate stream or ends early |

`location.offset` points at the structure in the file; `details.name` names the entry when known.
**Fix:** re-save the file in a spreadsheet application or ask the sender for a new copy.

### XLV_E_ZIP_UNSUPPORTED
The archive uses a ZIP feature that Office Open XML packages do not allow and xlvane does not read. `details.reason` is `encrypted` (ZIP encryption), `compression-method` (anything other than stored or Deflate) or `multi-disk` (split or spanned archives).
**Fix:** re-save the workbook in a spreadsheet application without ZIP encryption and as a single file.

### XLV_E_XML_MALFORMED
An XML part is not well-formed: for example a mismatched end tag, a missing root element, an unescaped `&` or `<`, a duplicate attribute, a character that XML does not allow, or an undeclared namespace prefix. `location` gives the part, line and column.
**Fix:** re-save the file in a spreadsheet application; if a program generated it, report the problem to its authors.

### XLV_E_XML_DTD_FORBIDDEN
An XML part contains a document type declaration (`<!DOCTYPE …>`). xlvane rejects every DTD, which rules out external entities (XXE) and entity-expansion attacks ("billion laughs"). This is a deliberate deviation from XML 1.0, which requires processors to read the internal DTD subset; Office Open XML parts do not use DTDs.
**Fix:** none needed for genuine spreadsheets. A file that triggers this error was produced by an unusual tool or crafted on purpose.

### XLV_E_XML_ENTITY_FORBIDDEN
An XML part refers to an entity other than the five predefined ones (`&lt;` `&gt;` `&amp;` `&apos;` `&quot;`). Without a DTD such a reference is not well-formed XML. `details.entity` holds the entity name.
**Fix:** re-save the file in a spreadsheet application. Character references such as `&#160;` are always allowed.

### XLV_E_XML_ENCODING
An XML part is not valid UTF-8 or UTF-16, or declares another encoding. `details.encoding` holds a declared encoding.
**Fix:** re-save the file in a spreadsheet application.

### XLV_E_PART_MISSING
A part that the workbook needs does not exist: the content types stream (`[Content_Types].xml`; `details.reason` is `missing`, or `interleaved` when it is stored in pieces), the package relationships, the workbook part, or an entry that a relationship points to.
**Fix:** the file is incomplete; re-save it in a spreadsheet application. Check also that the file is a spreadsheet and not another Office document.

### XLV_E_PART_INVALID
A part exists but has the wrong structure: for example the workbook part's root element is not a SpreadsheetML `workbook`, or the workbook has no `sheets` element. `details` names the element and namespace when relevant. For the packaging parts, `details.reason` says what is wrong:

| reason | meaning |
|---|---|
| `root-element` | the content types or relationships part has the wrong root element or namespace |
| `missing-attribute` | a required attribute (e.g. `Id`, `Type`, `Target`, `PartName`, `ContentType`) is missing |
| `duplicate-default`, `duplicate-override` | the content types stream declares the same extension or part twice |
| `duplicate-id` | two relationships in one part have the same `Id` |
| `target-mode` | `TargetMode` is neither `Internal` nor `External` |
| `internal-target` | an internal relationship target cannot be resolved to a part name |
| `equivalent-part-names` | two ZIP entries map to equivalent part names (e.g. a raw UTF-8 name and its percent-encoded form) |
**Fix:** check that the file is an Excel workbook (not a Word or PowerPoint document) and re-save it in a spreadsheet application.

### XLV_E_INVALID_CELL_REF
A cell or row reference is not valid or lies outside the sheet grid (A1 to XFD1048576). Thrown while reading a damaged worksheet, by `sheet.cell(ref)` when `ref` is not an A1-style reference, and by the writer's `set`, `style`, `column`, `row`, `merge` and `autoFilter` for references, ranges, column letters or row numbers outside the grid. `details.ref` holds the reference.
**Fix:** for files, re-save them in a spreadsheet application. For `sheet.cell()`, pass a reference such as `"B7"`.

### XLV_E_SHEET_NOT_FOUND
`workbook.getSheet(name)` (of a read or a new workbook) found no sheet with that name. Sheet names are compared exactly (case-sensitively). The hint suggests close matches and lists the available sheets; `details.available` holds the names.
**Fix:** use one of the listed names, or iterate over `workbook.sheets`.

### XLV_E_FILE_READ
`xlvane/node` could not read the file: it does not exist, is a directory, or cannot be read. `details.path` holds the path and `details.code` the operating-system error code (e.g. `ENOENT`).
**Fix:** check the path and the file permissions.

### XLV_E_INVALID_VALUE
A value given to the writer cannot be stored in a cell: `NaN` or an infinite number, text longer than 32 767 characters (the Excel limit) or with an unpaired UTF-16 surrogate, an invalid `Date` or a date outside the workbook's date system (the 1900 system covers 1899-12-31 to 9999-12-31, the 1904 system 1904-01-01 to 9999-12-31), an unknown error value, an empty formula, a column width outside 0–255 or a row height outside 0–409. `location` names the sheet and cell.
**Fix:** convert the value first, for example write it as text, or leave the cell empty.

### XLV_E_INVALID_STYLE
A style object has an unknown property (including `__proto__`), a wrong type or a value outside the allowed set, for example a colour that is not 6 or 8 hex digits or a border style Excel does not know. `details.property` is the path of the property, e.g. `font.color`; the hint lists the allowed values and suggests the closest one.
**Fix:** correct the property; the allowed values are those of the `CellStyle` type.

### XLV_E_INVALID_SHEET_NAME
A sheet name is empty, longer than 31 characters, contains one of `\ / ? * : [ ]` or a control character, or starts or ends with an apostrophe.
**Fix:** choose another name.

### XLV_E_DUPLICATE_SHEET_NAME
The workbook already has a sheet with this name. Names are compared ignoring upper and lower case. `details.existing` is the name of the existing sheet.
**Fix:** choose another name, or get the existing sheet with `getSheet()`.

### XLV_E_OVERLAPPING_MERGE
A range passed to `merge()` shares cells with a range merged before. `details.other` is that range.
**Fix:** merge ranges that do not overlap.

### XLV_E_FILE_EXISTS
`writeWorkbookFile` was asked to write to a path where a file already exists, and `overwrite` was not set. Nothing was written.
**Fix:** choose another path, or pass `{ overwrite: true }` if the file should be replaced.

### XLV_E_FILE_WRITE
`writeWorkbookFile` could not write the file. `details.code` is the operating-system error code, for example `ENOENT` when the directory does not exist or `EACCES` when writing is not permitted.
**Fix:** check that the directory exists and is writable.

### XLV_E_HEADER_NOT_FOUND
`readTable` was given a header row number that holds no cells (or lies outside the range), or `appendRows` received object rows for a sheet without a detectable header row. `details.row` is the row.
**Fix:** pass the row that holds the column names, `header: "auto"`, or `header: false` to use column letters; for `appendRows`, pass rows as arrays.

### XLV_E_SCHEMA_INVALID
The `schema` option is not a Standard Schema (version 1): it has no `~standard` property with `version: 1` and a `validate` function.
**Fix:** pass a schema object from a library that implements Standard Schema, such as Zod, Valibot or ArkType.

### XLV_E_ASYNC_SCHEMA
The schema returned a Promise from `validate`, which the synchronous `readTable` cannot wait for.
**Fix:** use `await readTableAsync(input, options)`.

### XLV_E_TABLE_INVALID
Rows given to `writeTable`, `fromJSON` or `appendRows` have the wrong shape: not an array, rows mixing objects and arrays, a value that is not a cell value (for example a nested object), or, for `appendRows`, a key that is not a column of the sheet. `details.row` and `details.key` say where; the hint lists the columns and suggests close matches.
**Fix:** pass an array of flat objects (or of arrays) whose values are numbers, strings, booleans, Dates or null.

### XLV_E_JSON_INVALID
The text given to `fromJSON` is not valid JSON. The message contains the parser's description, which includes the position.
**Fix:** correct the JSON, or pass the already parsed array.

### XLV_E_CSV_MALFORMED
The CSV text given to `fromCSV` has a quoted field that is never closed, text right after a closing quote, or bytes that are not valid UTF-8. `details.line` and `details.column` give the position (1-based; 0 for an encoding error).
**Fix:** close the quoted field or double the quotes inside it (`""`), or convert the file to UTF-8.

### XLV_E_NOT_SUPPORTED
`@xlvane/exceljs` was asked for an exceljs feature it does not implement, for example a rich-text or hyperlink value, images, tables, conditional formatting, data validation, notes, CSV, protection or `insertRow`. `details.feature` names it. The error is thrown where the feature is used, so nothing is dropped silently.
**Fix:** see docs/exceljs-compat.md for the supported subset; remove the call, or keep using exceljs for files that need the feature.

## Warnings

### XLV_W_DIMENSION_MISMATCH
The sheet's declared range (`dimension`) does not contain all of its cells. Programs that trust the declared range may miss data. A declared range that is larger than the cells is normal (formatting without values) and is not reported.

### XLV_W_SHARED_FORMULA_NOT_EXPANDED
Some cells reuse a shared formula written in another cell, and xlvane could not translate it to them: the master cell is missing, the cell lies outside the master's range, a moved reference would leave the sheet, or the formula contains something the tokenizer does not recognise (docs/facts/formulas.md). Their `formula.text` is `null` and `formula.sharedIndex` names the group; their cached values are available. Other cells of shared formulas get their translated text and no warning.

### XLV_W_FORMULA_NO_CACHED_VALUE
Some formula cells have no cached result, so their `value` is `null`. xlvane does not calculate formulas; open and save the file in a spreadsheet application to store the results.

### XLV_W_VALUES_IN_MERGED_CELLS
Some merged ranges have values in cells other than the top-left one. Spreadsheet applications do not display those values, so the data may surprise a reader.

### XLV_W_VERY_HIDDEN_SHEET
A sheet has the state `veryHidden`: spreadsheet applications offer no menu command to unhide it, so users usually do not know it exists.

### XLV_W_MACROS_PRESENT
The workbook contains a VBA project or Excel 4.0 macro sheets. xlvane never runs macros; the warning exists so that callers can decide how to treat such files.

### XLV_W_EXTERNAL_LINKS
The workbook links to other workbooks. xlvane never follows the links; values that come from them are the copies cached in the file.

### XLV_W_DUPLICATE_CELL
A sheet defines the same cell more than once. The last definition is used.

### XLV_W_INVALID_VALUE
A value does not match its declared type and is returned in a safe form, for example a numeric cell whose text is not a number (returned as text), a boolean cell that is neither 0 nor 1, a shared-string index that does not exist (the cell is treated as empty), a date serial outside the date system (returned as a number), an unknown sheet state (treated as visible), or an invalid style index or range. The message says what was found and what was done.

### XLV_W_SHEET_PART_MISSING
A sheet listed in the workbook has no part (the relationship or the part itself is missing). The sheet is returned with `part: null` and no cells.

### XLV_W_PART_MISSING
The workbook points to a styles or shared strings part that does not exist. Reading continues without it: without styles, date-formatted numbers are returned as numbers; without shared strings, shared-string cells are treated as empty.

### XLV_W_FORMULAS_NOT_RECALCULATED
Reported by `editCells` and `appendRows` when the workbook contains formulas. xlvane does not calculate formulas, so cached results that depend on edited cells may be out of date. The edited file asks spreadsheet applications to recalculate when it is opened (`calcPr fullCalcOnLoad`), and the calculation chain is removed when formulas themselves changed. Reading the file with xlvane still returns the old cached results.
**Fix:** nothing, if the file will be opened in a spreadsheet application; otherwise do not rely on cached results of formulas that depend on edited cells.
