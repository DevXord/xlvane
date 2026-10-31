# Errors and warnings

Every problem xlvane reports has a stable code. Errors are thrown as `XlvError` (`code`, `message`, `location`, `hint`, `docsUrl`, `details`; `error.toJSON()` returns the same as data). Warnings never stop reading; they are listed in `workbook.warnings` and in the `warnings` array of `inspect()`, aggregated per sheet with a `count` and up to ten example `cells`.

`location` contains only the fields that are known: `part` (part name inside the package, e.g. `/xl/worksheets/sheet1.xml`), `sheet`, `cell` (A1 reference), `line` and `column` (inside an XML part) and `offset` (byte offset in the file, for ZIP structures).

Codes are never reused for a different condition. New codes are added here first.

## Errors

### XLV_E_INVALID_INPUT
The input is not what the function expects: `readWorkbook` and `inspect` accept a `Uint8Array`, an `ArrayBuffer` (or another typed-array view) and, for `inspect` and the table functions, a workbook returned by `readWorkbook`. `editCells` and `appendRows` need the original bytes, because unchanged parts are copied from them. `fromCSV` needs text or UTF-8 bytes.
**Fix:** pass the bytes of the file. A path string is the most common mistake; in Node.js use `readWorkbookFile()` or `inspectFile()` from `xlvane/node`.

### XLV_E_INVALID_OPTION
An option is unknown, has the wrong type or is out of range. The hint lists the known options and suggests the closest name. `details.option` names the option. Ranges checked include the sheet view zoom (10 to 400 percent, ECMA-376 Part 1 §18.3.1.87), a split window that needs a position above 0 in at least one direction ([MS-OI29500] §2.1.640 e), the document property `appVersion` (one or two digits, a dot, one to four digits, [MS-OI29500] §2.1.1717), and the printing limits of Office ([MS-OI29500] §2.1.589–§2.1.645): page margins from 0 to below 49 inches, print scale 10 to 400 or 0 (automatic), copies 1 to 32767, fitToWidth and fitToHeight at most 32767, header and footer texts of at most 255 characters, at most 1022 row and 1023 column page breaks (docs/facts/sheet-rules.md SR-23 … SR-28). Data validations and conditional formatting (`NewSheet.validate`, `NewSheet.conditionalFormat`) are checked against the schema enumerations and the Office rules: a data validation needs `formula1` unless its type is `none`, and `formula2` exactly with `between`/`notBetween`; `prompt` at most 255 characters, `error` 224, `promptTitle` and `errorTitle` 32 ([MS-OI29500] §2.1.613); formulas at most 8192 characters (§2.1.1085); `cellIs` rules need an `operator` and `timePeriod` rules a `timePeriod` (§2.1.595); colour scales 2 or 3 value objects with as many colours (§2.1.601); data bar lengths at most 100 (§2.1.611); icon set thresholds not `min` or `max` (§2.1.596); rule priorities unique on the sheet and at least 1; at most 32767 ranges in a validation's range list and 8192 in a conditional format's (§2.1.613, §2.1.1092). `details.option` names the property, for example `rules[0].operator` (docs/facts/sheet-rules.md SR-01 … SR-16). `NewSheet.addImage` and `NewSheet.background` also throw it for an image number that `addImage` did not return, a picture position that is not a range, cell or 0-based `{ col, row }` within the sheet (columns below 16384, rows below 1048576, [MS-OI29500]), a `to` corner above or left of `from`, a size that rounds to less than 1 EMU, `editAs` with a size, and unknown or invalid picture options (name, description or hyperlink longer than 2084 characters, with characters XML cannot carry — control characters other than tab, line feed and carriage return, U+FFFE, U+FFFF, unpaired surrogates — or, for the hyperlink, any control character). `autoFilter` throws it for a range that overlaps a table of the sheet, and `conditionalFormat` for a rule `text` with characters XML cannot carry (it is xsd:string, without the `_x` escape).
**Fix:** correct the option name or value.

### XLV_E_INVALID_ARGUMENT
An argument of a method of `@xlvane/exceljs` is not valid, in a case where exceljs 4.4.0 accepts it and loses or moves data without an error: for example `insertRow(0, …)`, a row position given as a string, a negative count in `spliceRows` or `spliceColumns`, an outline level outside 0–7 set on a row or column (exceljs writes it), `duplicateRow` of a row that does not exist, `insertRows` without an array of values, or rows or columns pushed beyond the sheet (row 1048576, column XFD). `details.method`, `details.argument` and `details.value` name the method, the argument and its value (docs/exceljs-compat.md).
**Fix:** pass a whole number in the range the hint gives (rows from 1, columns from 1), or remove rows at the end of the sheet before inserting.

### XLV_E_LIMIT_EXCEEDED
The file exceeds one of the resource limits (`ReadLimits`): input size, number of ZIP entries, decompressed size of one entry or of all entries, compression ratio, XML nesting depth, attribute count, name or text length, number of elements in a tree-parsed part, or number of cells. `details.limit` names the limit and `details.value` its current value. Sizes are measured on the bytes actually produced by decompression, so a file cannot avoid a limit by declaring false sizes.

`NewSheet.validate` throws it when a sheet would get more than 65534 data validations ([MS-OI29500] §2.1.614 g; `details.limit` `dataValidations`). When reading, the conditional formatting and data validations of one sheet count against `maxXmlTreeElements`.

The writers and `editCells` also throw this code when the styles part would hold more records than Office accepts ([MS-OI29500] rev. 25.0): more than 206 custom number formats (`details.limit` `customNumFmts`, §2.1.713 g), 65430 cell formats (`cellXfs`, §2.1.700), 65491 fonts (`fonts`, §2.1.707), 65430 fills (`fills`, §2.1.705) or 65430 borders (`borders`, §2.1.696). `editCells` counts the records already in the file plus the ones it adds, and checks only when it adds one. The writers also throw it for more rows than a sheet holds.
**Fix:** if the file comes from a trusted source and is simply large, raise the named limit, e.g. `{ limits: { maxCells: 50_000_000 } }`. A compression-ratio error on a small file usually means a zip bomb. For a styles limit, use fewer distinct styles: reuse one style object (or one number format code) for all cells that should look the same.

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
A cell or row reference is not valid or lies outside the sheet grid (A1 to XFD1048576). Thrown by `editCells` and `appendRows` for a sheet that contains such references (reading skips them with `XLV_W_INVALID_CELL_REF`), by `sheet.cell(ref)` when `ref` is not an A1-style reference, and by the writer's `set`, `style`, `column`, `row`, `merge` and `autoFilter` for references, ranges, column letters or row numbers outside the grid. `details.ref` holds the reference.
**Fix:** for files, re-save them in a spreadsheet application, which drops the damaged cells. For `sheet.cell()`, pass a reference such as `"B7"`.

### XLV_E_SHEET_NOT_FOUND
`workbook.getSheet(name)` (of a read or a new workbook) found no sheet with that name. Sheet names are compared exactly (case-sensitively). The hint suggests close matches and lists the available sheets; `details.available` holds the names.
**Fix:** use one of the listed names, or iterate over `workbook.sheets`.

### XLV_E_TABLE_NOT_FOUND
`readTable(file, { table })` found no table with that name. Table names are compared ignoring case, against both the name used in formulas (`displayName`) and the `name` of the table. The hint suggests close matches and lists the tables of the workbook; `details.available` holds them.
**Fix:** use one of the listed names, or look at `Sheet.tables` or `xlvane inspect`.

### XLV_E_FILE_READ
`xlvane/node` could not read the file: it does not exist, is a directory, or cannot be read. `details.path` holds the path and `details.code` the operating-system error code (e.g. `ENOENT`).
**Fix:** check the path and the file permissions.

### XLV_E_INVALID_VALUE
A value given to the writer cannot be stored in a cell: `NaN` or an infinite number, text longer than 32 767 characters (the Excel limit) or with an unpaired UTF-16 surrogate, an invalid `Date` or a date outside the workbook's date system (the 1900 system covers 1899-12-31 to 9999-12-31, the 1904 system 1904-01-01 to 9999-12-31), an unknown error value, an empty formula, rich text that is not `{ richText: [{ text, font? }, …] }` with 1 to 32 767 runs and at most 32 767 characters in all ([MS-OI29500] §2.1.675), a column width outside 0–255 or a row height outside 0–409. `location` names the sheet and cell. An invalid font in a rich text run is `XLV_E_INVALID_STYLE`, with `details.run` giving the run's index.
Array and shared formulas (D-E02, F2-13, F2-15) are checked too: `array` or `shared` must be a range whose top-left cell is the cell itself, not both; when the sheet is written, a formula inside another cell of an array range, overlapping array or shared ranges, a `{ sharedFormula }` cell outside its master's range or whose master holds no shared formula give this error with `details.cell`.
**Fix:** convert the value first, for example write it as text, or leave the cell empty.

### XLV_E_INVALID_STYLE
A style object has an unknown property (including `__proto__`), a wrong type or a value outside the allowed set, for example a colour that is not 6 or 8 hex digits or a border style Excel does not know. The ranges are those Office accepts ([MS-OI29500] rev. 25.0): `font.size` 1 to 409.55 points (§2.1.677 c), `font.name` 1 to 31 characters (§2.1.711), `numFmt` 1 to 254 characters (shorter than 255, §2.1.712 b) or a built-in format id of ECMA-376 Part 1 §18.8.30 that every language has (0–4, 9–22, 37–40, 45–49; docs/facts/sml-write.md W2-41), `alignment.indent` 0 to 255 (§2.1.693 b), `alignment.textRotation` 0 to 180 or 255 (§2.1.693 a), `font.family` 0 to 5 (§2.1.703). Also checked: `font.charset` 0 to 255 (ECMA-376 Part 1 §18.4.1); a gradient fill needs 2 to 256 stops with positions from 0 to 1, path sides from 0 to 1 and a `degree` from −360 to 360 (docs/facts/sml-write.md W2-60); `font.vertAlign`, `font.scheme` and `alignment.readingOrder` take the values of their types. Styles already in a file changed by `editCells` are kept even when they are outside these ranges; only the new style is checked. `details.property` is the path of the property, e.g. `font.color`; the hint names the allowed values or range with a valid example, and suggests the closest name.
**Fix:** correct the property; the allowed values are those of the `CellStyle` type.

### XLV_E_INVALID_SHEET_NAME
A sheet name is empty, longer than 31 characters, contains one of `\ / ? * : [ ]` or a control character (including U+0003, which Office excludes, [MS-OI29500] §2.1.578 a), or starts or ends with an apostrophe.
**Fix:** choose another name.

### XLV_E_DUPLICATE_SHEET_NAME
The workbook already has a sheet with this name. Names are compared ignoring upper and lower case. `details.existing` is the name of the existing sheet.
**Fix:** choose another name, or get the existing sheet with `getSheet()`.

### XLV_E_INVALID_DEFINED_NAME
`defineName` (or `definedNames` of `writeWorkbookStream`) was given a name that spreadsheet applications do not accept, or a duplicate. Rules ([MS-OI29500] §2.1.568, ECMA-376 Part 1 §18.17.2.5): 1 to 255 characters; letters, digits, `_`, `.`, `\` and `?`, starting with a letter, `_` or `\`; not `TRUE`/`FALSE`, not a cell reference (`A1`, `R1C1`); not starting with `_xl` (reserved for built-in names); unique in its scope, ignoring case. The same code is used for an empty or too long formula (`refersTo`, at most 8192 characters) and a comment longer than 255 characters. `details.name` is the name.
**Fix:** choose another name, for example `Sales_2026` instead of `Sales 2026` or `A1`, or make the name local to a sheet with `{ sheet: "Data" }`.

### XLV_E_INVALID_HYPERLINK
`link` (or `hyperlinks` of a sheet of `writeWorkbookStream`) was given a hyperlink that cannot be stored: not an object, an unknown property, a property that is not a string, neither `url` nor `location`, an empty `url` or one with control characters (including tab, line feed and carriage return) or the non-characters U+FFFE/U+FFFF, which a relationship target cannot carry, or a property longer than Office accepts ([MS-OI29500] §2.1.626: `location` and `display` at most 2084 characters, `tooltip` at most 255; `url` at most 65 536, an xlvane limit). `location.cell` is the linked cell or range and `details.property` the property.
**Fix:** pass `{ url: "https://example.com" }` for a web page or file, or `{ location: "Sheet1!A1" }` for a place in the workbook, and shorten long texts.

### XLV_E_INVALID_IMAGE
`NewWorkbook.addImage` (or `images` of `writeWorkbookStream`) was given something that is not the bytes of a PNG, JPEG or GIF file. The format is recognised from the first bytes of the file (PNG `89 50 4E 47 0D 0A 1A 0A`, JPEG `FF D8 FF`, GIF `GIF87a` or `GIF89a`), not from a file name. `details.size` is the number of bytes given.
**Fix:** pass the file contents as a `Uint8Array`, for example `new Uint8Array(await file.arrayBuffer())`, and convert other formats (SVG, EMF, WMF, BMP, TIFF, WebP) to PNG or JPEG first.

### XLV_E_INVALID_TABLE
`addTable` (or `tables` of a sheet of `writeWorkbookStream`) was given a table that spreadsheet applications do not accept: a name or display name that is not a valid name (letters, digits, `_`, `.` and `\`, starting with a letter or `_`; no spaces; not a cell reference such as `A1` or `R1C1`; at most 255 characters; ECMA-376 Part 1 §18.5.1.2, §18.17.2.5, [MS-OI29500] §2.1.680), a name already used by another table or a defined name of the workbook (ignoring case), no columns or two columns with the same name (ignoring case), a column name longer than 255 characters, a totals label together with a totals function, `totalsRowFormula` without `totalsRowFunction: "custom"` or the reverse, a style that is not one of the 60 built-in table styles, a range that is not as wide as the columns, leaves no data row or reaches beyond the sheet, a range that overlaps another table, a merged range or the sheet filter ([MS-OI29500] §2.1.680 h; `autoFilter: true` of the streaming writer covers every used cell, so it cannot be combined with tables), or a name equal to a built-in name such as `Print_Area`, `Print_Titles`, `Criteria` or `Database` (with or without `_xlnm.`) or starting with `_xl`. `merge` also throws it for a range that overlaps a table, and so does `defineName` (as `XLV_E_INVALID_DEFINED_NAME`) for a table's name. `details.name` is the table name.
**Fix:** choose another name, for example `Sales_2026` instead of `Sales 2026`, give every column a unique name, or move the table so that it does not touch other tables or merged cells.

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
`hashPassword` and `verifyPassword` throw it for a hash algorithm other than SHA-1, SHA-256, SHA-384 and SHA-512, and `hashPassword` without a `salt` throws it in a runtime that has no `crypto.getRandomValues` (pass your own random salt there; SR-50, D-C08).
`editCells` throws it for an array or shared formula in an edit (it writes plain formulas only), and for a cell inside an array formula's range or the master of a shared formula with other cells in its range (D-E02); `details.cell` names it.
`insertRows` and `deleteRows` (and `xlvane edit --insert-rows` / `--delete-rows`) also throw it when the sheet has content whose cell references xlvane cannot move yet: tables, charts, shapes and other drawing objects (pictures alone are moved, and deleting rows that would cut a picture is refused), pivot tables, form controls, embedded objects, custom sheet views, threaded comments, extended content such as data bars or sparklines, a chart or pivot cache elsewhere that names the sheet, a data table, an array formula the change would split, or a formula or defined name that cannot be adjusted safely (for example a 3-D reference such as `Jan:Mar!A1` that includes the sheet, D-B01). The message names the part or the cell, and `details` gives the element, relationship type or formula.
**Fix:** for `@xlvane/exceljs`, see docs/exceljs-compat.md for the supported subset; remove the call, or keep using exceljs for files that need the feature.
For `insertRows`/`deleteRows`, insert or delete the rows in a spreadsheet application, or remove the named content first.

### XLV_E_USAGE
The `xlvane` command line was used incorrectly: an unknown command or option, a missing or extra argument, an option without its value, or an option value that is not allowed. `details` names the command and the option or argument; `details.suggestions` lists close matches ("Did you mean ...?"). Exit code 2.
**Fix:** run `xlvane help <command>` for the usage and examples of the command.

### XLV_E_INTERNAL
The `xlvane` command line caught an error that is not an `XlvError`. This is a bug in xlvane. `details.name` is the error's name; `--verbose` prints the stack trace.
**Fix:** report it with the command and, if possible, the input file (see SECURITY.md for files that must stay private).

## Warnings

### XLV_W_DIMENSION_MISMATCH
The sheet's declared range (`dimension`) does not contain all of its cells. Programs that trust the declared range may miss data. A declared range that is larger than the cells is normal (formatting without values) and is not reported.

### XLV_W_SHARED_FORMULA_NOT_EXPANDED
Some cells reuse a shared formula written in another cell, and xlvane could not translate it to them: the master cell is missing, the cell lies outside the master's range, a moved reference would leave the sheet, or the formula contains something the tokenizer does not recognise (docs/facts/formulas.md). Formulas that use the implicit-intersection operator `@` or a spill reference `#` (such as `A1#`) are never moved, because ECMA-376 Part 1 §18.17, [MS-OI29500] and [MS-XLSX] §2.2.2 do not define these operators (docs/decisions.md D-1102). Their `formula.text` is `null` and `formula.sharedIndex` names the group; their cached values are available. Other cells of shared formulas get their translated text and no warning.

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
A sheet defines the same cell more than once. The last definition is used, also when one definition holds only a style (D-E06): a later cell with a value replaces a style-only one, and a later style-only cell replaces one with a value.

### XLV_W_INVALID_VALUE
A value does not match its declared type and is returned in a safe form, for example a numeric cell whose text is not a number (returned as text), a boolean cell that is neither 0 nor 1, a shared-string index that does not exist (the cell is treated as empty), a date serial outside the date system (returned as a number), an unknown sheet state (treated as visible), or an invalid style index or range. The message says what was found and what was done.

### XLV_W_UNUSUAL_HYPERLINK
A sheet has hyperlinks whose external target uses a scheme other than `http`, `https` or `mailto`, for example `file:`, `javascript:` or a relative file path. xlvane never opens or follows link targets; the warning is there so that an application which shows or opens links can decide what to do. The message names one such target; `Sheet.hyperlinks` lists all of them.

### XLV_W_DRAWING_CONTENT_NOT_READ
The drawing of a sheet holds objects other than pictures: charts, shapes, connectors, group shapes, content parts or alternate content. xlvane reads only the pictures (`Sheet.images`); the message gives the number of other objects. `editCells` keeps the whole drawing unchanged; `insertRows` and `deleteRows` refuse such sheets.

### XLV_W_THREADED_COMMENTS
A sheet has threaded comments, the conversations of current Excel versions ([MS-XLSX] §2.3.7). xlvane reads only the notes (legacy comments, ECMA-376 Part 1 §18.7). Excel stores a placeholder note next to each conversation; its author is `tc={id}` and its text may not contain the replies. `Sheet.notes` lists these placeholders as they are stored.

### XLV_W_SHEET_PART_MISSING
A sheet listed in the workbook has no part (the relationship or the part itself is missing). The sheet is returned with `part: null` and no cells.

### XLV_W_PART_MISSING
The workbook points to a styles or shared strings part that does not exist. Reading continues without it: without styles, date-formatted numbers are returned as numbers; without shared strings, shared-string cells are treated as empty.

### XLV_W_FORMULAS_NOT_RECALCULATED
Reported by `editCells` and `appendRows` when the workbook contains formulas. xlvane does not calculate formulas, so cached results that depend on edited cells may be out of date. The edited file asks spreadsheet applications to recalculate when it is opened (`calcPr fullCalcOnLoad`), and the calculation chain is removed when formulas themselves changed. Reading the file with xlvane still returns the old cached results.
**Fix:** nothing, if the file will be opened in a spreadsheet application; otherwise do not rely on cached results of formulas that depend on edited cells.

### XLV_W_INVALID_CELL_REF
A worksheet contains a row whose number is not a whole number from 1 to 1048576, a cell whose `r` reference is not a valid A1 reference inside the sheet grid (A1 to XFD1048576), or a cell without a reference that would lie past column XFD. Such rows and cells are skipped and the rest of the sheet is read, which is what LibreOffice does with them. A cell without a reference (or with an empty one) is placed after the last valid cell of its row. `count` is the number of skipped rows and cells.
**Fix:** nothing, if the skipped cells are not needed; otherwise re-save the file in the application that made it. `editCells` and `appendRows` refuse such sheets (`XLV_E_INVALID_CELL_REF`).
