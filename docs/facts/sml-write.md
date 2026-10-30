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

**Check against the official documents (2026-10-30).** The maintainer supplied these ECMA-376 archives; they are kept in `spec/` (not committed):

| Id | Document | File |
|---|---|---|
| P2 | ECMA-376 Part 2, 5th edition (December 2021), Open Packaging Conventions | `ECMA-376-2_5th_edition_december_2021.zip`: `Ecma Office Open XML Part 2 - Open Packaging Conventions.pdf`, `OpenPackagingConventions-XMLSchema.zip` |
| P3 | ECMA-376 Part 3, 5th edition (December 2015), Markup Compatibility and Extensibility | `ECMA-376-3_5th_edition_december_2015.zip` |
| P4 | ECMA-376 Part 4, 5th edition (December 2016), Transitional Migration Features | `ECMA-376-4_5th_edition_december_2016.zip`: `Ecma Office Open XML Part 4 - Transitional Migration Features.pdf`, `OfficeOpenXML-XMLSchema-Transitional.zip` (contains `sml.xsd`, SHA-256 `495debc8…6ab9`) |
| P1S | ECMA-376 Part 1 supporting files (no Part 1 PDF yet) | `OfficeOpenXML-XMLSchema-Strict.zip` (Strict schemas; `sml.xsd` `targetNamespace` `http://purl.oclc.org/ooxml/spreadsheetml/main` at L7), `OfficeOpenXML-SpreadsheetMLStyles.zip` (`presetCellStyles.xml`, `presetTableStyles.xml`) |

Line numbers `sml.xsd Lnnn` below refer to that official Transitional `sml.xsd`. Every S-XSD citation was compared with it: the worksheet types are at the same lines with the same content; the styles and workbook types are shifted by 14–51 lines with the same content, except CT_CellAlignment (see W2-17). PDF locators give the printed page; `pdftotext -layout` line numbers are not repeated here. ECMA-376 Part 1 (the SpreadsheetML reference prose, §18) was not supplied, so semantics that only Part 1 states (units, meanings, defaults beyond the schema) remain PROVISIONAL.

## Worksheet part

### W2-01 — CT_Worksheet child order — VERIFIED
Sequence: `sheetPr`, `dimension`, `sheetViews`, `sheetFormatPr`, `cols`*, `sheetData` (required), `sheetCalcPr`, `sheetProtection`, `protectedRanges`, `scenarios`, `autoFilter`, `sortState`, `dataConsolidate`, `customSheetViews`, `mergeCells`, `phoneticPr`, `conditionalFormatting`*, `dataValidations`, `hyperlinks`, `printOptions`, `pageMargins`, `pageSetup`, `headerFooter`, … `extLst`. Source: S-XSD CT_Worksheet L2170–2213. The writer emits `dimension`, `sheetViews`, `sheetFormatPr`, `cols`, `sheetData`, `autoFilter`, `mergeCells` in this order.
**Verified:** `sml.xsd` (Part 4 schema archive) CT_Worksheet L2170–2213 (same sequence; `cols` 0..unbounded, `sheetData` minOccurs 1).

### W2-02 — Cells and formulas
CT_Cell is the sequence `f`, `v`, `is`, `extLst` with attributes `r`, `s` (default 0), `t` (default `n`) (S-XSD L2274; agrees with E-20 … E-23). CT_CellFormula extends `ST_Formula`, which restricts `s:ST_Xstring` (S-XSD L216–218, L2751), so formula text uses the `_xHHHH_` escape like other ST_Xstring values (F-XS-01). PROVISIONAL for ST_Formula only; the reader now decodes formula text accordingly.
**Verified (2026-10-30):** `sml.xsd` (Part 4 schema archive) ST_Formula L216–218 (`restriction base="s:ST_Xstring"`), CT_CellFormula L2751 (simpleContent extension of ST_Formula), CT_Cell L2274. ST_Formula is no longer provisional.

### W2-03 — `cols` / `col` — VERIFIED (unit corrected 2026-10-30)
CT_Cols holds 1..n `col`. CT_Col attributes: `min`, `max` (unsignedInt, required), `width` (double), `style` (default 0), `hidden`, `bestFit`, `customWidth`, `phonetic` (boolean, default false), `outlineLevel`, `collapsed`. Source: S-XSD L2233–2252. **Not checked:** the unit of `width` (the writer passes the caller's number through and documents it as "characters", as in Excel's column width box) and whether `customWidth="1"` is needed for a width to apply.
**Verified:** `sml.xsd` (Part 4 schema archive) CT_Cols L2233–2237, CT_Col L2238–2252. The unit of `width` and the role of `customWidth` are prose in Part 1 §18.3.1.13 (not supplied).
**Checked 2026-10-30 (D-1101, D-1104), corrects the unit:** `col@width` is in characters of the maximum digit width of the Normal font, including 5 px padding, truncated to 1/256: `Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256`. With Calibri 11 (maximum digit width 7 px), 8 characters are stored as 8.7109375. Source: ECMA-376 Part 1 §18.3.1.13, p.1601–1602. Office requires 0 ≤ width ≤ 255 ([MS-OI29500] rev. 25.0 §2.1.598 c, p.253). `customWidth` is not stated as required for `width` to apply. xlvane writes the caller's value unchanged, rejects values outside 0–255, and its TSDoc no longer calls the value the number in Excel's column-width box.

**Column style (2026-10-30, D-A04):** CT_Col `style` is `xsd:unsignedInt`, optional, default `0` (sml.xsd L2242); `width` is optional (L2241), so a `col` with only `style` is valid. [MS-OI29500] rev. 25.0 §2.1.598 d, p.253: Office requires `style` 0 … 65429, which the cellXfs count limit (65430, D-1101) already guarantees. The writer writes `style` only when it is not 0 and never applies a column's style to its cells. **PROVISIONAL, awaiting the second local report:** the meaning of `col@style` (that it is the format applications give to cells added to the column later) is prose in Part 1 §18.3.1.13, not in the container; `editCells` therefore does not apply it to new cells.

### W2-04 — `row` attributes — VERIFIED
CT_Row attributes include `r`, `s`, `customFormat`, `ht` (double), `hidden`, `customHeight` (boolean, default false), `outlineLevel`, `collapsed`, `thickTop`, `thickBot`, `ph`. Source: S-XSD L2256–2273. **Not checked:** that `ht` is in points.
**Verified:** `sml.xsd` (Part 4 schema archive) CT_Row L2256–2273. That `ht` is in points needs Part 1 §18.3.1.73.
**Checked 2026-10-30 (D-1104):** ECMA-376 Part 1, 5th ed. (2016) §18.3.1.73 row, attribute `ht`, p.1678: row height in points, no margin padding; `customHeight` = 1 if set manually. [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.644 a, p.279: `row@r` 1 … 1048576, non-decreasing; outlineLevel 0–7; `s` 0–65490; no `ht` limit.

### W2-05 — `sheetViews` / `sheetView` / `pane` / `selection` — VERIFIED
CT_SheetViews: 1..n `sheetView`. CT_SheetView: `pane`?, `selection`{0,4}, `pivotSelection`{0,4}, `extLst`?; attribute `workbookViewId` is required, `tabSelected` defaults to false. CT_Pane: `xSplit`, `ySplit` (double, default 0), `topLeftCell`, `activePane` (ST_Pane: `bottomRight`, `topRight`, `bottomLeft`, `topLeft`; default `topLeft`), `state` (ST_PaneState: `split`, `frozen`, `frozenSplit`; default `split`). CT_Selection: `pane`, `activeCell`, `activeCellId`, `sqref`. Source: S-XSD L2326–2398, L3172. **Not checked:** that in the frozen state `xSplit`/`ySplit` count the frozen columns/rows (the LibreOffice test confirms LibreOffice reads them that way), and which `activePane` belongs to which split.
**Verified:** `sml.xsd` (Part 4 schema archive) CT_SheetViews L2326 (also an optional `extLst`), CT_SheetView L2332, CT_Pane L2359, CT_Selection L2388 (`sqref` default `A1`), ST_Pane L2394, ST_PaneState L3172. The meaning of `xSplit`/`ySplit` in the frozen state needs Part 1 §18.3.1.66.
**Checked 2026-10-30 (D-1104):** ECMA-376 Part 1, 5th ed. (2016) §18.3.1.66 pane, p.1665: when the pane is frozen, `xSplit` is the number of columns and `ySplit` the number of rows that stay visible (otherwise 1/20 point positions). §18.18.52 ST_Pane, p.2468–2469: rows frozen only → `bottomLeft`; columns only → `topRight`; both → `bottomRight`. §18.18.53 ST_PaneState, p.2469–2470: `frozen`, `frozenSplit`, `split`. (P1's wording names the panes the other way round, "top pane"/"left pane".)

### W2-06 — `autoFilter` — VERIFIED
CT_AutoFilter: `filterColumn`*, `sortState`?, `extLst`?; attribute `ref` (ST_Ref). Source: S-XSD L16. **Not checked:** the built-in defined name `_xlnm._FilterDatabase` that the writer adds for each filter (it follows what Excel files contain; LibreOffice keeps the filter either way).
**Verified:** `sml.xsd` (Part 4 schema archive) CT_AutoFilter L16. Part 4 §15.4.2 (p.218) changes only `dynamicFilter` descriptions. The `_xlnm._FilterDatabase` name is not in P2/P4; it needs Part 1 §18.2.6 or [MS-OI29500].
**Checked 2026-10-30 (D-1104):** ECMA-376 Part 1, 5th ed. (2016) §18.2.5 definedName, p.1556: `_xlnm._FilterDatabase` is the built-in name for the source range of an advanced filter or a range with an AutoFilter; names beginning `_xlnm.` are reserved. [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.568, p.238–239: names beginning `_xl` are restricted; `name` 1–255 characters; `localSheetId` ≤ 32766; uniqueness per sheet or per workbook. No document says the name is required for an AutoFilter.

### W2-07 — `t` and `xml:space`
White space handling of `t` (F-SST-05, D-SST-1). The writer adds `xml:space="preserve"` when a string starts or ends with white space or contains tab, LF or CR, so that other readers keep it.

### W2-08 — Sheet names in formulas — VERIFIED
The `_xlnm._FilterDatabase` value quotes the sheet name with apostrophes and doubles apostrophes inside it. **Not checked** in the formula grammar (Part 1 §18.17); LibreOffice accepted names with `&`, `<`, spaces and apostrophes in the test.
**Checked 2026-10-30 (D-1104):** ECMA-376 Part 1, 5th ed. (2016) §18.17.2.3, p.2049–2051: `work-sheet-prefix-special= apostrophe, sheet-name-special, apostrophe, "!" …`, `sheet-name-character-special= apostrophe, apostrophe | character – (apostrophe | "*" | "[" | "]" | "\" | ":" | "/" | "?")`; unquoted names exclude operators, apostrophe, `[`, `]`, `\`, `?`. Quoting any name is valid. [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.863 a, p.392 and [MS-XLSX] rev. 30.0 (2026-10-16) §2.2.2, p.39 agree and also exclude U+0003 in quoted names; a quoted 3-D prefix has one pair of apostrophes (`'Sheet 1:Sheet 3'!A1`).

### W2-09 — `sheetFormatPr` — VERIFIED (schema); value 15 is a convention
CT_SheetFormatPr requires `defaultRowHeight` (double). Source: S-XSD L2222–2232. The writer uses 15 (the value Excel uses for an 11 pt Calibri default font; **not checked**).
**Verified:** `sml.xsd` (Part 4 schema archive) CT_SheetFormatPr L2222–2232 (`defaultRowHeight` xsd:double, required).
**Checked 2026-10-30 (D-1104):** ECMA-376 Part 1, 5th ed. (2016) §18.3.1.81 sheetFormatPr, p.1685–1686: `defaultRowHeight` in points (an optimisation). [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.650, p.281: `baseColWidth` ≤ 255, `defaultColWidth` 0 … < 65536. The value 15 for Calibri 11 is not in P1, OI or [MS-XLSX] rev. 30.0 (2026-10-16): it stays a convention.

### W2-09a — `mergeCells`
CT_MergeCells: 1..n `mergeCell` with required `ref`, and an optional `count` (S-XSD L2470–2478; agrees with E-35).

## Styles part

### W2-10 — CT_Stylesheet order
`numFmts`, `fonts`, `fills`, `borders`, `cellStyleXfs`, `cellXfs`, `cellStyles`, `dxfs`, `tableStyles`, `colors`, `extLst` (F-STY-02, 5th edition).

### W2-11 — CT_Font — VERIFIED
A choice (any order, each at most once) of `name` (val ST_Xstring), `charset`, `family`, `b`, `i`, `strike`, `outline`, `shadow`, `condense`, `extend` (CT_BooleanProperty, `val` default **true**), `color`, `sz` (val double), `u` (CT_UnderlineProperty, `val` default **single**; ST_UnderlineValues `single`, `double`, `singleAccounting`, `doubleAccounting`, `none`), `vertAlign`, `scheme`. Source: S-XSD L3730–3786.
**Verified:** `sml.xsd` (Part 4 schema archive) CT_Font L3781–3800 (choice, maxOccurs unbounded, each element 0..1), CT_BooleanProperty L3744 (`val` default true), CT_FontSize L3747 (`val` double, required), CT_FontName (`val` s:ST_Xstring), CT_UnderlineProperty L3769 (`val` default `single`), ST_UnderlineValues L3772.

### W2-12 — Font 0 is the default font — VERIFIED
The writer puts the default font (Calibri 11) at index 0 and the reader treats `fontId="0"` as "default". This is how Excel and LibreOffice files are laid out (the cell style "Normal" uses `xfId 0` → font 0), **not checked** in §18.8.
**Supporting evidence (not a normative rule):** in `presetCellStyles.xml` (Part 1 supporting archive) the `normal` preset's style sheet has the default font first and `cellStyle Normal` with `xfId="0"`, whose xf uses font 0. The rule itself still needs Part 1 §18.8 or [MS-OI29500].
**Checked 2026-10-30 (D-1104):** ECMA-376 Part 1, 5th ed. (2016) §18.4.3 phoneticPr, attribute `fontId`, p.1722: if the index is out of bounds the default font of the Normal style should be used, and this default font should be at index 0.

### W2-13 — CT_Xf — VERIFIED
Sequence `alignment`?, `protection`?, `extLst`?; attributes `numFmtId`, `fontId`, `fillId`, `borderId`, `xfId`, `quotePrefix`, `pivotButton`, `applyNumberFormat`, `applyFont`, `applyFill`, `applyBorder`, `applyAlignment`, `applyProtection` (all optional). Source: S-XSD L3584–3603 (agrees with F-STY-06 for the attributes covered there).
**Verified:** `sml.xsd` (Part 4 schema archive) CT_Xf L3598–3617.

### W2-14 — Fills; fills 0 and 1 — VERIFIED (reserved); values a convention
CT_Fill is a choice of `patternFill` or `gradientFill`. CT_PatternFill: `fgColor`?, `bgColor`?, attribute `patternType` (ST_PatternType: `none`, `solid`, `mediumGray`, `darkGray`, `lightGray`, `darkHorizontal`, `darkVertical`, `darkDown`, `darkUp`, `darkGrid`, `darkTrellis`, `lightHorizontal`, `lightVertical`, `lightDown`, `lightUp`, `lightGrid`, `lightTrellis`, `gray125`, `gray0625`). Source: S-XSD L3469–3518. **Convention, not checked:** fills 0 and 1 are `none` and `gray125` in files written by Excel; the writer reserves them the same way. To check in [MS-OI29500].
**Verified:** `sml.xsd` (Part 4 schema archive) CT_Fill L3489 (choice of one), CT_PatternFill L3495, ST_PatternType L3509 (same 19 values). Fills 0 and 1 remain a convention to check in [MS-OI29500].
**Supporting evidence (not a normative rule):** the style sheets in `presetCellStyles.xml` (Part 1 supporting archive `OfficeOpenXML-SpreadsheetMLStyles.zip`) all start with `<fills count="2">` holding `patternFill patternType="none"` and then `patternType="gray125"`, the layout the writer uses. The rule itself still needs Part 1 §18.8 or [MS-OI29500].
**Checked 2026-10-30 (D-1104):** [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.705 b (P1 §18.8.21 fills), p.298–299: in Excel the first two fills are reserved and deviations from their predefined values are not respected; a: 1 … 65430 fills. The predefined values (`none`, `gray125`) are not stated in P1 or OI; the preset style sheets of the Part 1 supporting files use them.

### W2-15 — Borders — VERIFIED
CT_Border sequence: `start`, `end`, `left`, `right`, `top`, `bottom`, `diagonal`, `vertical`, `horizontal` (all optional); CT_BorderPr: `color`?, attribute `style` (ST_BorderStyle, default `none`: `none`, `thin`, `medium`, `dashed`, `dotted`, `thick`, `double`, `hair`, `mediumDashed`, `dashDot`, `mediumDashDot`, `dashDotDot`, `mediumDashDotDot`, `slantDashDot`). Source: S-XSD L3413–3458. The writer emits `left`, `right`, `top`, `bottom`, `diagonal`.
**Verified:** `sml.xsd` (Part 4 schema archive) CT_Border L3451–3466 (also attributes `diagonalUp`, `diagonalDown`, `outline` default true), CT_BorderPr L3467 (`style` default `none`), ST_BorderStyle L3427 (same 14 values). P4 §15.5.1 and §15.5.2 (p.219–220): in Transitional documents `left` is semantically equivalent to `start` and `right` to `end`.

### W2-16 — CT_Color — VERIFIED
Attributes `auto` (boolean), `indexed` (unsignedInt), `rgb` (ST_UnsignedIntHex = hexBinary of length 4, i.e. 8 hex digits), `theme` (unsignedInt), `tint` (double, default 0). Source: S-XSD L219–223, L3488–3494. **Not checked:** that the 4 bytes are in ARGB order (the writer writes alpha first; LibreOffice read `FFFF0000` as red in the test).
**Verified:** `sml.xsd` (Part 4 schema archive) CT_Color L3502 (`tint` default `0.0`), ST_UnsignedIntHex L219–223 (hexBinary, length 4). The ARGB byte order needs Part 1 §18.8.19.
**Checked 2026-10-30 (D-1104):** ECMA-376 Part 1, 5th ed. (2016) §18.8.19 fgColor `rgb`, p.1766 (also §18.8.3 bgColor, §18.3.1.15 color): "Standard Alpha Red Green Blue color value (ARGB)" — alpha first. [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.704 a, p.298: Office ranks `auto` < `indexed` < `rgb` < `theme` and keeps only the highest one.

### W2-17 — CT_CellAlignment — VERIFIED (with a schema difference; indent corrected 2026-10-30)
Attributes `horizontal` (ST_HorizontalAlignment: `general`, `left`, `center`, `right`, `fill`, `justify`, `centerContinuous`, `distributed`), `vertical` (ST_VerticalAlignment: `top`, `center`, `bottom`, `justify`, `distributed`), `textRotation` (unsignedInt), `wrapText`, `indent` (unsignedInt), `relativeIndent`, `justifyLastLine`, `shrinkToFit`, `readingOrder`. Source: S-XSD L3402–3412, L3541–3561. **Not checked:** the meaning and range of `textRotation` (the writer accepts 0–180 and 255) and the maximum `indent` (the writer accepts 0–250).
**Verified:** `sml.xsd` (Part 4 schema archive) CT_CellAlignment L3402–3412, ST_TextRotation L3413–3426, ST_HorizontalAlignment L3555, ST_VerticalAlignment L3567. **Difference between official sources:** the electronic `sml.xsd` types `textRotation` as ST_TextRotation, a union of 0–180 and the value 255, and gives `vertical` the default `bottom`; the schema printed in P4 Annex A (printed page 1003) and S-XSD type `textRotation` as xsd:unsignedInt with no `vertical` default. The writer's 0–180 or 255 satisfies both. The maximum `indent` (writer: 250) is not in the schema (xsd:unsignedInt) and stays PROVISIONAL (Excel behaviour).
**Checked 2026-10-30 (D-1101, D-1104):** `textRotation` 0–180 or 255 for vertical text ([MS-OI29500] rev. 25.0 §2.1.693 a, p.295–296). `indent` at most 255 (§2.1.693 b); ECMA-376 Part 1 §18.8.1 sets no maximum. The writer's maximum changed from 250 to 255.

### W2-18 — Cell styles — VERIFIED
CT_CellStyleXfs and CT_CellXfs hold 1..n `xf`; CT_CellStyles holds 1..n `cellStyle` with required `xfId` and optional `name`, `builtinId` (S-XSD L3572–3617). The writer writes one cell style xf and one cell style `Normal` with `builtinId="0"`; **not checked** that built-in id 0 is "Normal".
**Verified:** `sml.xsd` (Part 4 schema archive) CT_CellStyleXfs L3586, CT_CellXfs L3592 (1..n `xf`), CT_CellStyles L3618, CT_CellStyle L3624 (`xfId` required). That built-in id 0 is "Normal" needs Part 1 §18.8.7.
**Verified (builtinId 0):** `presetCellStyles.xml` in the Part 1 supporting archive `OfficeOpenXML-SpreadsheetMLStyles.zip` (supplied by the maintainer, 2026-10-30) defines the preset style `<normal builtinId="0">`, whose style sheet contains `<cellStyle name="Normal" xfId="0" builtinId="0"/>` (L47).
**Checked 2026-10-30 (D-1104):** ECMA-376 Part 1, 5th ed. (2016) Annex G.2 Built-in Cell Styles, p.4472–4475: builtinId 0 = Normal (list 0–53); §18.8.7 cellStyle, p.1760: for built-in styles `builtinId` determines the style. [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.697, p.296–297: builtinId 0–53; 12–14 = Emphasis 1–3; `xfId` and `name` unique.

### W2-19 — Font size and related style limits — VERIFIED (corrected 2026-10-30)
The writer accepts font sizes greater than 0 and at most 409 points. 409 is the Excel row-height limit (W-04); a font-size limit is **not checked** in the Excel specifications page.
**Checked 2026-10-30 (D-1101, D-1104), corrects the range:** `sz@val` is a double from 1 to 409.55 inclusive ([MS-OI29500] rev. 25.0 §2.1.677 c, p.290); Office converts points to twips, so rounding may happen on save (§2.1.677 b). The writer now accepts 1–409.55 instead of "> 0 … 409". Related limits now enforced: font name 1–31 characters (§2.1.711); `formatCode` shorter than 255 characters (§2.1.712 b); at most 206 custom number formats (§2.1.713 g); at most 65430 cellXfs (§2.1.700 a, c), 65491 fonts (§2.1.707 a), 65430 fills (§2.1.705 a) and 65430 borders (§2.1.696). Style rules checked and tested: applyX (§2.1.721 a–f), bookViews always written (§2.1.584 a), font child order (§2.1.706 a, §2.1.674 a; CT_Font in sml.xsd is an unbounded choice).

## Workbook

### W2-20 — Sheet names — VERIFIED
The writer rejects names longer than 31 characters, names containing `\ / ? * : [ ]`, names starting or ending with an apostrophe, and duplicates ignoring case. ECMA-376 states only that names are unique (E-09). These rules are what Excel enforces and are **not checked** against Microsoft documentation in this session.
**Checked 2026-10-30 (D-1104):** [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.578 (P1 §18.2.19 sheet), p.243: `name` at most 31 characters, following the character rules of `sheet-name`/`sheet-name-special` (P1 §18.17); unique case-insensitively in any locale; `sheetId` 1 … 65534; §2.1.579: at most 32767 sheets. Forbidden characters per P1 p.2050–2051 and OI §2.1.863: `'` at the start or end, `*`, `[`, `]`, `\`, `:`, `/`, `?`, and U+0003 (OI, [MS-XLSX] rev. 30.0 (2026-10-16)).

### W2-21 — Formula length — PROVISIONAL (not in P1, OI or [MS-XLSX])
The writer rejects formula text longer than 8 192 characters ("Length of formula contents", Excel specifications and limits; **not checked** in this session).
**Checked 2026-10-30 (D-1104):** searched ECMA-376 Part 1, 5th ed. (2016), [MS-OI29500] rev. 25.0 (2026-08-18) and [MS-XLSX] rev. 30.0 (2026-10-16) for a formula-length limit: none (OI §2.1.620 f has no length rule; the only 8192 in OI is `oleObject@link`). [MS-XLSX] rev. 30.0 (2026-10-16) §2.2.2, p.39: at most 254 arguments. 8192 stays from the Excel limits page (W-04).

### W2-22 — CT_Workbook order — VERIFIED
`fileVersion`, `fileSharing`, `workbookPr`, `workbookProtection`, `bookViews`, `sheets` (required), `functionGroups`, `externalReferences`, `definedNames`, `calcPr`, … Source: S-XSD L4083–4117. CT_BookViews: 1..n `workbookView`; CT_BookView `activeTab` and `firstSheet` default 0 (S-XSD L4118–4141). CT_DefinedName: text content ST_Formula, attributes `name` (required), `localSheetId`, `hidden` (S-XSD L4298–4322).
**Verified:** `sml.xsd` (Part 4 schema archive) CT_Workbook L4097–4131, CT_BookViews L4132, CT_BookView L4137 (`firstSheet`, `activeTab` default 0), CT_DefinedName L4317 (simpleContent extension of ST_Formula; `name` required).

### W2-23 — Defined names: element and scope — VERIFIED (schema), PROVISIONAL (localSheetId meaning)
CT_DefinedNames holds 0..n `definedName`; CT_DefinedName is a simple-content extension of ST_Formula with `name` (ST_Xstring, required), `comment`, `customMenu`, `description`, `help`, `statusBar`, `shortcutKey` (ST_Xstring), `localSheetId` (unsignedInt), `hidden` (boolean, default false) and the flags `function`, `vbProcedure`, `xlm`, `publishToServer`, `workbookParameter`, plus `functionGroupId`. Source: `sml.xsd` (Part 4 schema archive) L4312–4337. `definedNames` follows `externalReferences` in CT_Workbook (W2-22). [MS-OI29500] rev. 25.0 §2.1.568 l (p.238–239, from the D-1104 report): `localSheetId` is the sheet on which a sheet-local name is defined; r: at most 32766; b–i, p: `customMenu`, `description`, `help`, `statusBar`, `shortcutKey`, `xlm`, `vbProcedure`, `functionGroupId`, `function` are reserved and unused by Office. **PROVISIONAL:** that `localSheetId` is the 0-based position of the sheet in `sheets` (Part 1 §18.2.5 text not in the container). Observed in the corpus: in 23 of 24 `_xlnm._FilterDatabase`/`_xlnm.Print_Area` names with `localSheetId` in tests/corpus (POI and LibreOffice files), the sheet at that 0-based position is the one the formula names; the 24th refers to `#REF!`. The writer already used this for `_xlnm._FilterDatabase` (W2-06).
**Code:** `defined-names.ts` (`definedNameXml`), `sml/workbook.ts` (reader). Only `name`, `comment`, `localSheetId` and `hidden` are written; the reader keeps those and the formula text.

### W2-24 — Defined names: name rules — VERIFIED
From the D-1104 report:
- P1 §18.17.2.5 (p.2057): `name-start-character = letter | underscore | backslash`, `name-character = letter | decimal-digit | underscore | full-stop`; names are not case-sensitive; a name shall not be TRUE/FALSE or a cell reference.
- [MS-OI29500] §2.1.867 a (p.394–395) and [MS-XLSX] §2.2.2 (p.40): Office also allows non-ASCII letters. [MS-XLSX] p.40 also allows `?` after the first character and forbids names of the form of an `R1C1-cell-reference`. [MS-OI29500] §2.1.867 c: Office does not create or load names of a backslash and one name character.
- [MS-OI29500] §2.1.568 (p.238–239): k `name` 1–255 characters; j `comment` at most 255 characters; m names beginning `_xl` are restricted; n a name with `localSheetId` is unique among that sheet's names, one without it among the workbook-level names. P1 §18.2.5 (p.1556, W2-06): names beginning `_xlnm.` are reserved for built-in names.
**Code:** `checkDefinedName` and `checkUniqueName` (case-insensitive, since names are not case-sensitive). Stricter than required, to stay safe: every token of 1–3 letters followed by digits is refused as a cell reference, also beyond the grid (`XFE1`); every `_xl` prefix is refused, ignoring case.

### W2-25 — Defined names in `docProps/app.xml` — not applicable yet
P1 §18.17.6.5 (p.2070, D-1104 report): each name shall also be an `lpstr` in the Application-Defined File Properties part. xlvane writes no `docProps/app.xml` (W2-33), so this rule does not apply yet; it will when app.xml is written (phase A2).

## Package

### W2-30 — Content types — VERIFIED
Worksheet `application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`, styles `…spreadsheetml.styles+xml`, shared strings `…spreadsheetml.sharedStrings+xml` (S-SDK WorksheetPart.json, WorkbookStylesPart.json, SharedStringTablePart.json). Workbook: `…spreadsheetml.sheet.main+xml` (E-07). Relationships parts: `application/vnd.openxmlformats-package.relationships+xml` (E-03).
P4 §10.2.15, §10.2.20 and §10.2.24 refer to Part 1 for these parts and give only the Transitional namespaces; the content types are in Part 1.
**Checked 2026-10-30 (D-1104):** ECMA-376 Part 1, 5th ed. (2016) §12.3.15 Shared String Table Part, p.88: `application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml`; §12.3.20 Styles Part, p.94: `…spreadsheetml.styles+xml`; §12.3.24 Worksheet Part, p.100: `…spreadsheetml.worksheet+xml`; §12.3.23 Workbook (E-07).

### W2-31 — Relationship types
Transitional base URI plus `officeDocument`, `worksheet`, `styles`, `sharedStrings` (E-05).

### W2-32 — Part names — PROVISIONAL (convention)
`/xl/workbook.xml`, `/xl/worksheets/sheetN.xml`, `/xl/styles.xml`, `/xl/sharedStrings.xml`. ECMA-376 does not require particular names (parts are found through relationships, E-08); these are the names Excel uses and the Open XML SDK's default paths (`xl`, `worksheets`, S-SDK).
**Checked 2026-10-30 (D-1104):** no rule found in ECMA-376 Part 1, 5th ed. (2016) §12.2–§12.3 or [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.6, §2.1.11, §2.1.12: parts are found through relationships (E-08); the names stay a convention.

### W2-33 — Optional parts — VERIFIED
The writer writes no `docProps/core.xml` or `docProps/app.xml`, and no theme. **Not checked** in Part 2 §8 that core properties are optional; LibreOffice opens the files without them.
**Verified:** P2 §8.2 Core Properties part (printed p.38): a package contains *at most one* Core Properties part, so it is optional. The optionality of `docProps/app.xml` and the theme part needs Part 1.
**Checked 2026-10-30 (D-1104):** ECMA-376 Part 1, 5th ed. (2016) §15.2.12.3 Extended File Properties Part, p.154: at most one per package; §14.2.7 Theme Part, p.135–136: zero or one per SpreadsheetML package — both optional. §18.17.6.5, p.2070: when `docProps/app.xml` is written, each defined name shall also be listed there (relevant only if xlvane writes that part). [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.30: Office's core-properties relationship type is `http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties`.

### W2-41 — Built-in number format ids on write — VERIFIED
An `xf` may use a built-in ("implied") `numFmtId` without a `numFmt` element; ids not in the "All Languages" list need a `formatCode` (E-41, V-22: ECMA-376 Part 1 §18.8.30, p.1776–1777). Some ids may be shown differently depending on the UI language, and Excel's own codes for 14, 22, 37–40 and 47 differ from the specification's text (V-22, [MS-OI29500] §2.1.712 a, p.299: 14 = `m/d/yyyy`). The writer accepts only the "All Languages" ids (`BUILTIN_NUMBER_FORMATS`: 0–4, 9–22, 37–40, 45–49) as `numFmt: <id>`, writes them as `xf@numFmtId` with no `numFmt` element, and treats 0 as General (no format). The locale-dependent ids 27–36 and 50–81 are not accepted because their codes differ per language (E-42). In `editCells` an id refers to whatever the file gives that id (implied, or declared in `numFmts`). D-A05.

### W2-40 — ZIP date and time
APPNOTE says only "standard MS-DOS format" (Z-19). The bit layout used, `(year − 1980) << 9 | month << 5 | day` and `hour << 11 | minute << 5 | second / 2`, is the one in fflate's ZIP writer (`node_modules/fflate/esm/index.mjs`, the `wzh` header writer, fflate 0.8.3).

## Tools

### F-LO-01 — LibreOffice CSV filter options
`--convert-to "csv:Text - txt - csv (StarCalc):<tokens>"`: token 1 field separator (44 = comma), 2 text delimiter (34 = `"`), 3 character set (76 = UTF-8, per the help's import example), 8 "detect special numbers"/numbers as numbers (export default true), 9 "save cell contents as shown", 10 export formulas, 12 export sheets (`-1` = every sheet to `<base>-<sheet name>.csv`). Source: S-LO. LibreOffice 24.2.7.2 (Ubuntu 24.04 package `libreoffice-calc`) was used in this session.

### F-NODE-04 — fs.promises.writeFile flags
`fsPromises.writeFile(file, data[, options])` takes `options.flag` (default `'w'`); the flag `'wx'` is "Like 'w' but fails if the path exists". Source: S-NODE fs.md, `fsPromises.writeFile` and "File system flags".

### F-NODE-05 — child_process.spawnSync
`spawnSync(command, args, options)` with `timeout` and `encoding`, returning `status`, `stdout`, `stderr`. Source: S-NODE child_process.md. Used only by tests and scripts.

### F-NODE-11 — A FileHandle closed by garbage collection is an error from Node.js 25
DEP0137 "Closing fs.FileHandle on garbage collection" reached End-of-Life in v25.0.0 (nodejs/node#58536): "Allowing a `fs.FileHandle` object to be closed on garbage collection used to be allowed, but now throws an error. Please ensure that all `fs.FileHandle` objects are explicitly closed using `FileHandle.prototype.close()`". fs.md, Class FileHandle: "always explicitly close {FileHandle}s". Source: https://raw.githubusercontent.com/nodejs/node/v25.x/doc/api/deprecations.md and fs.md (v25.x), read 2026-10-30. Reproduced on Node.js 25.8.2: an `fsPromises.writeFile(path, asyncIterable)` whose iterable never ends exits with `ERR_INVALID_STATE` after `gc()`.

### F-NODE-12 — filehandle.write may write fewer bytes
`filehandle.write(buffer, offset[, length[, position]])`: `length` defaults to `buffer.byteLength - offset`; the promise is fulfilled with `{ bytesWritten, buffer }`; "It is unsafe to use `filehandle.write()` multiple times on the same file without waiting for the promise". Source: fs.md (v25.x), read 2026-10-30.

### F-NODE-13 — fsPromises.open and writeFile with an AsyncIterable
`fsPromises.open(path, flags[, mode])` (added v10.0.0) takes the same flags as F-NODE-04. `fsPromises.writeFile` accepts `data` of type `AsyncIterable`. Source: fs.md (v25.x), read 2026-10-30. That `writeFile` opens the file before reading the iterable was observed in the reproduction of F-NODE-11, not read in the documentation.

### F-GH-01 — GitHub Actions pins
`actions/checkout` tag v7.0.1 → commit 3d3c42e5aac5ba805825da76410c181273ba90b1; `actions/setup-node` tag v7.1.0 → commit 949feb2413d6458794dcd2491c4babbbce0c15c1 (`git ls-remote`, lightweight tags, 2026-10-30). Inputs used (`persist-credentials`, `node-version`) are declared in each action's `action.yml` at that tag.

## Still to verify
The open points of phase 2 were checked on 2026-10-30 against ECMA-376 Part 1, [MS-OI29500] and [MS-XLSX] (D-1104); each entry above carries the result. Still PROVISIONAL: the formula length (W2-21, in none of the three documents) and the part names (W2-32, a convention).
