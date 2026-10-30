# Facts for the XLSX writer (phase 2)

Ids `W2-*` are cited from code comments. Each fact says where it was checked.

**Status of the sources (2026-10-30).** The cloud session that wrote phase 2 could not reach ecma-international.org or learn.microsoft.com (the network proxy refused both), so the ECMA-376 PDFs were not available. Facts marked **PROVISIONAL** come from these secondary sources and must be checked against ECMA-376 Part 1 and Part 4 (5th edition) once the maintainer supplies them (see "To verify" at the end):

| Id | Source | Version |
|---|---|---|
| S-XSD | `ref/xsd/sml.xsd` in github.com/python-openxml/python-docx: a copy of the Transitional SpreadsheetML schema. Its line numbers differ from the 5th-edition `xsd-trans/sml.xsd` cited in ecma-376.md (CT_Stylesheet L3387 here vs L3386, CT_Xf L3584 vs L3598, CT_Sheet L4202 vs L4216), so it is from a different edition. | branch `master` at e45454602b53e8e572b179ccf1c91093ec9f4ed7, SHA-256 of the file `8b7dcfb0…9a1c` |
| S-SDK | Open XML SDK part definitions `data/parts/*.json` (github.com/OfficeDev/Open-XML-SDK, MIT, .NET Foundation) | branch `main` at 431ab05cf160248cc3885a4a766026d4f8243792 |
| S-LO | LibreOffice help, `source/text/shared/guide/csv_params.xhp` (github.com/LibreOffice/help) | branch `master` at 9c739b197574ad9b8e9bb0f599310a91f02d2cb1 |
| S-NODE | Node.js documentation `doc/api/fs.md`, `doc/api/child_process.md` (github.com/nodejs/node) | branch `v22.x`, read 2026-10-30 |

Schema facts that the phase 1 facts already established from the 5th edition (ecma-376.md, sst-styles.md) are cited by their ids and are not provisional.

## Worksheet part

### W2-01 — CT_Worksheet child order — PROVISIONAL
Sequence: `sheetPr`, `dimension`, `sheetViews`, `sheetFormatPr`, `cols`*, `sheetData` (required), `sheetCalcPr`, `sheetProtection`, `protectedRanges`, `scenarios`, `autoFilter`, `sortState`, `dataConsolidate`, `customSheetViews`, `mergeCells`, `phoneticPr`, `conditionalFormatting`*, `dataValidations`, `hyperlinks`, `printOptions`, `pageMargins`, `pageSetup`, `headerFooter`, … `extLst`. Source: S-XSD CT_Worksheet L2170–2213. The writer emits `dimension`, `sheetViews`, `sheetFormatPr`, `cols`, `sheetData`, `autoFilter`, `mergeCells` in this order.

### W2-02 — Cells and formulas
CT_Cell is the sequence `f`, `v`, `is`, `extLst` with attributes `r`, `s` (default 0), `t` (default `n`) (S-XSD L2274; agrees with E-20 … E-23). CT_CellFormula extends `ST_Formula`, which restricts `s:ST_Xstring` (S-XSD L216–218, L2751), so formula text uses the `_xHHHH_` escape like other ST_Xstring values (F-XS-01). PROVISIONAL for ST_Formula only; the reader now decodes formula text accordingly.

### W2-03 — `cols` / `col` — PROVISIONAL
CT_Cols holds 1..n `col`. CT_Col attributes: `min`, `max` (unsignedInt, required), `width` (double), `style` (default 0), `hidden`, `bestFit`, `customWidth`, `phonetic` (boolean, default false), `outlineLevel`, `collapsed`. Source: S-XSD L2233–2252. **Not checked:** the unit of `width` (the writer passes the caller's number through and documents it as "characters", as in Excel's column width box) and whether `customWidth="1"` is needed for a width to apply.

### W2-04 — `row` attributes — PROVISIONAL
CT_Row attributes include `r`, `s`, `customFormat`, `ht` (double), `hidden`, `customHeight` (boolean, default false), `outlineLevel`, `collapsed`, `thickTop`, `thickBot`, `ph`. Source: S-XSD L2256–2273. **Not checked:** that `ht` is in points.

### W2-05 — `sheetViews` / `sheetView` / `pane` / `selection` — PROVISIONAL
CT_SheetViews: 1..n `sheetView`. CT_SheetView: `pane`?, `selection`{0,4}, `pivotSelection`{0,4}, `extLst`?; attribute `workbookViewId` is required, `tabSelected` defaults to false. CT_Pane: `xSplit`, `ySplit` (double, default 0), `topLeftCell`, `activePane` (ST_Pane: `bottomRight`, `topRight`, `bottomLeft`, `topLeft`; default `topLeft`), `state` (ST_PaneState: `split`, `frozen`, `frozenSplit`; default `split`). CT_Selection: `pane`, `activeCell`, `activeCellId`, `sqref`. Source: S-XSD L2326–2398, L3172. **Not checked:** that in the frozen state `xSplit`/`ySplit` count the frozen columns/rows (the LibreOffice test confirms LibreOffice reads them that way), and which `activePane` belongs to which split.

### W2-06 — `autoFilter` — PROVISIONAL
CT_AutoFilter: `filterColumn`*, `sortState`?, `extLst`?; attribute `ref` (ST_Ref). Source: S-XSD L16. **Not checked:** the built-in defined name `_xlnm._FilterDatabase` that the writer adds for each filter (it follows what Excel files contain; LibreOffice keeps the filter either way).

### W2-07 — `t` and `xml:space`
White space handling of `t` (F-SST-05, D-SST-1). The writer adds `xml:space="preserve"` when a string starts or ends with white space or contains tab, LF or CR, so that other readers keep it.

### W2-08 — Sheet names in formulas — PROVISIONAL
The `_xlnm._FilterDatabase` value quotes the sheet name with apostrophes and doubles apostrophes inside it. **Not checked** in the formula grammar (Part 1 §18.17); LibreOffice accepted names with `&`, `<`, spaces and apostrophes in the test.

### W2-09 — `sheetFormatPr` — PROVISIONAL
CT_SheetFormatPr requires `defaultRowHeight` (double). Source: S-XSD L2222–2232. The writer uses 15 (the value Excel uses for an 11 pt Calibri default font; **not checked**).

### W2-09a — `mergeCells`
CT_MergeCells: 1..n `mergeCell` with required `ref`, and an optional `count` (S-XSD L2470–2478; agrees with E-35).

## Styles part

### W2-10 — CT_Stylesheet order
`numFmts`, `fonts`, `fills`, `borders`, `cellStyleXfs`, `cellXfs`, `cellStyles`, `dxfs`, `tableStyles`, `colors`, `extLst` (F-STY-02, 5th edition).

### W2-11 — CT_Font — PROVISIONAL
A choice (any order, each at most once) of `name` (val ST_Xstring), `charset`, `family`, `b`, `i`, `strike`, `outline`, `shadow`, `condense`, `extend` (CT_BooleanProperty, `val` default **true**), `color`, `sz` (val double), `u` (CT_UnderlineProperty, `val` default **single**; ST_UnderlineValues `single`, `double`, `singleAccounting`, `doubleAccounting`, `none`), `vertAlign`, `scheme`. Source: S-XSD L3730–3786.

### W2-12 — Font 0 is the default font — PROVISIONAL (convention)
The writer puts the default font (Calibri 11) at index 0 and the reader treats `fontId="0"` as "default". This is how Excel and LibreOffice files are laid out (the cell style "Normal" uses `xfId 0` → font 0), **not checked** in §18.8.

### W2-13 — CT_Xf — PROVISIONAL
Sequence `alignment`?, `protection`?, `extLst`?; attributes `numFmtId`, `fontId`, `fillId`, `borderId`, `xfId`, `quotePrefix`, `pivotButton`, `applyNumberFormat`, `applyFont`, `applyFill`, `applyBorder`, `applyAlignment`, `applyProtection` (all optional). Source: S-XSD L3584–3603 (agrees with F-STY-06 for the attributes covered there).

### W2-14 — Fills; fills 0 and 1 — PROVISIONAL
CT_Fill is a choice of `patternFill` or `gradientFill`. CT_PatternFill: `fgColor`?, `bgColor`?, attribute `patternType` (ST_PatternType: `none`, `solid`, `mediumGray`, `darkGray`, `lightGray`, `darkHorizontal`, `darkVertical`, `darkDown`, `darkUp`, `darkGrid`, `darkTrellis`, `lightHorizontal`, `lightVertical`, `lightDown`, `lightUp`, `lightGrid`, `lightTrellis`, `gray125`, `gray0625`). Source: S-XSD L3469–3518. **Convention, not checked:** fills 0 and 1 are `none` and `gray125` in files written by Excel; the writer reserves them the same way. To check in [MS-OI29500].

### W2-15 — Borders — PROVISIONAL
CT_Border sequence: `start`, `end`, `left`, `right`, `top`, `bottom`, `diagonal`, `vertical`, `horizontal` (all optional); CT_BorderPr: `color`?, attribute `style` (ST_BorderStyle, default `none`: `none`, `thin`, `medium`, `dashed`, `dotted`, `thick`, `double`, `hair`, `mediumDashed`, `dashDot`, `mediumDashDot`, `dashDotDot`, `mediumDashDotDot`, `slantDashDot`). Source: S-XSD L3413–3458. The writer emits `left`, `right`, `top`, `bottom`, `diagonal`.

### W2-16 — CT_Color — PROVISIONAL
Attributes `auto` (boolean), `indexed` (unsignedInt), `rgb` (ST_UnsignedIntHex = hexBinary of length 4, i.e. 8 hex digits), `theme` (unsignedInt), `tint` (double, default 0). Source: S-XSD L219–223, L3488–3494. **Not checked:** that the 4 bytes are in ARGB order (the writer writes alpha first; LibreOffice read `FFFF0000` as red in the test).

### W2-17 — CT_CellAlignment — PROVISIONAL
Attributes `horizontal` (ST_HorizontalAlignment: `general`, `left`, `center`, `right`, `fill`, `justify`, `centerContinuous`, `distributed`), `vertical` (ST_VerticalAlignment: `top`, `center`, `bottom`, `justify`, `distributed`), `textRotation` (unsignedInt), `wrapText`, `indent` (unsignedInt), `relativeIndent`, `justifyLastLine`, `shrinkToFit`, `readingOrder`. Source: S-XSD L3402–3412, L3541–3561. **Not checked:** the meaning and range of `textRotation` (the writer accepts 0–180 and 255) and the maximum `indent` (the writer accepts 0–250).

### W2-18 — Cell styles — PROVISIONAL
CT_CellStyleXfs and CT_CellXfs hold 1..n `xf`; CT_CellStyles holds 1..n `cellStyle` with required `xfId` and optional `name`, `builtinId` (S-XSD L3572–3617). The writer writes one cell style xf and one cell style `Normal` with `builtinId="0"`; **not checked** that built-in id 0 is "Normal".

### W2-19 — Font size limit — PROVISIONAL
The writer accepts font sizes greater than 0 and at most 409 points. 409 is the Excel row-height limit (W-04); a font-size limit is **not checked** in the Excel specifications page.

## Workbook

### W2-20 — Sheet names — PROVISIONAL (Excel behaviour)
The writer rejects names longer than 31 characters, names containing `\ / ? * : [ ]`, names starting or ending with an apostrophe, and duplicates ignoring case. ECMA-376 states only that names are unique (E-09). These rules are what Excel enforces and are **not checked** against Microsoft documentation in this session.

### W2-21 — Formula length — PROVISIONAL
The writer rejects formula text longer than 8 192 characters ("Length of formula contents", Excel specifications and limits; **not checked** in this session).

### W2-22 — CT_Workbook order — PROVISIONAL
`fileVersion`, `fileSharing`, `workbookPr`, `workbookProtection`, `bookViews`, `sheets` (required), `functionGroups`, `externalReferences`, `definedNames`, `calcPr`, … Source: S-XSD L4083–4117. CT_BookViews: 1..n `workbookView`; CT_BookView `activeTab` and `firstSheet` default 0 (S-XSD L4118–4141). CT_DefinedName: text content ST_Formula, attributes `name` (required), `localSheetId`, `hidden` (S-XSD L4298–4322).

## Package

### W2-30 — Content types — PROVISIONAL
Worksheet `application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`, styles `…spreadsheetml.styles+xml`, shared strings `…spreadsheetml.sharedStrings+xml` (S-SDK WorksheetPart.json, WorkbookStylesPart.json, SharedStringTablePart.json). Workbook: `…spreadsheetml.sheet.main+xml` (E-07). Relationships parts: `application/vnd.openxmlformats-package.relationships+xml` (E-03).

### W2-31 — Relationship types
Transitional base URI plus `officeDocument`, `worksheet`, `styles`, `sharedStrings` (E-05).

### W2-32 — Part names — PROVISIONAL (convention)
`/xl/workbook.xml`, `/xl/worksheets/sheetN.xml`, `/xl/styles.xml`, `/xl/sharedStrings.xml`. ECMA-376 does not require particular names (parts are found through relationships, E-08); these are the names Excel uses and the Open XML SDK's default paths (`xl`, `worksheets`, S-SDK).

### W2-33 — Optional parts
The writer writes no `docProps/core.xml` or `docProps/app.xml`, and no theme. **Not checked** in Part 2 §8 that core properties are optional; LibreOffice opens the files without them.

### W2-40 — ZIP date and time
APPNOTE says only "standard MS-DOS format" (Z-19). The bit layout used, `(year − 1980) << 9 | month << 5 | day` and `hour << 11 | minute << 5 | second / 2`, is the one in fflate's ZIP writer (`node_modules/fflate/esm/index.mjs`, the `wzh` header writer, fflate 0.8.3).

## Tools

### F-LO-01 — LibreOffice CSV filter options
`--convert-to "csv:Text - txt - csv (StarCalc):<tokens>"`: token 1 field separator (44 = comma), 2 text delimiter (34 = `"`), 3 character set (76 = UTF-8, per the help's import example), 8 "detect special numbers"/numbers as numbers (export default true), 9 "save cell contents as shown", 10 export formulas, 12 export sheets (`-1` = every sheet to `<base>-<sheet name>.csv`). Source: S-LO. LibreOffice 24.2.7.2 (Ubuntu 24.04 package `libreoffice-calc`) was used in this session.

### F-NODE-04 — fs.promises.writeFile flags
`fsPromises.writeFile(file, data[, options])` takes `options.flag` (default `'w'`); the flag `'wx'` is "Like 'w' but fails if the path exists". Source: S-NODE fs.md, `fsPromises.writeFile` and "File system flags".

### F-NODE-05 — child_process.spawnSync
`spawnSync(command, args, options)` with `timeout` and `encoding`, returning `status`, `stdout`, `stderr`. Source: S-NODE child_process.md. Used only by tests and scripts.

### F-GH-01 — GitHub Actions pins
`actions/checkout` tag v7.0.1 → commit 3d3c42e5aac5ba805825da76410c181273ba90b1; `actions/setup-node` tag v7.1.0 → commit 949feb2413d6458794dcd2491c4babbbce0c15c1 (`git ls-remote`, lightweight tags, 2026-10-30). Inputs used (`persist-credentials`, `node-version`) are declared in each action's `action.yml` at that tag.

## To verify when ECMA-376 is available
W2-01, W2-03 … W2-06, W2-08, W2-09, W2-11 … W2-22, W2-30, W2-32, W2-33: compare with Part 1 §18.3 (worksheet), §18.8 (styles), §18.2 (workbook), §12.3 (part summaries), Part 2 (core properties) and Part 4 (Transitional), and with [MS-OI29500] and the Excel specifications page for W2-14, W2-19 … W2-21.
