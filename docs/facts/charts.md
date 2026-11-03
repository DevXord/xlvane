# Facts: charts (DrawingML charts in XLSX)

Compiled on 2026-11-03. Ids CH-01 … CH-34; CH-35 … CH-37 were added with the implementation (phase K). Status words as in tables-drawings.md: VERIFIED (schema or a cited document states it), OBSERVED (seen in files or in a program's behaviour), PROVISIONAL (inferred), NOT FOUND (no available source states it). The prose of ECMA-376 Part 1 §21.2 was **not** available in this session (the ECMA website is blocked from the container); every point that needs it was sent to the maintainer as a question for a local review with the official documents (R6, as for the earlier reports R1 to R5) and is marked "R6-n" below; those points stay PROVISIONAL until the answers are recorded here.

## Sources

**S-XSD (schemas).** Third-party copies of the ISO/IEC 29500-4:2016 (Transitional) schemas, not the official archive:
- `dml-chart.xsd`, `dml-main.xsd`, `dml-spreadsheetDrawing.xsd` from github.com/ThinkInAIXYZ/deepchat, branch main, path `resources/skills/pptx/ooxml/schemas/ISO-IEC29500-4_2016/` (sha256 41b93bd8…bb354, 5375417f…0e9, b4532b6d…1e88).
- The same three files from github.com/openakita/openakita, branch main, path `skills/docx/ooxml/schemas/ISO-IEC29500-4_2016/`. `diff` shows that each differs from the deepchat copy only by one extra final newline.
- A third copy, byte-identical to the deepchat files, was found in a tool bundle of the development environment (directory `schemas/ISO-IEC29500-4_2016/`, not part of this repository). That folder also holds the imported `shared-commonSimpleTypes.xsd`, `shared-relationshipReference.xsd` and `dml-chartDrawing.xsd`, which were used for `xmllint --schema` validation and for ST_HexColorRGB (CH-21).
- Line numbers ("L") refer to the deepchat `dml-chart.xsd` unless another file is named. "SD" means `dml-spreadsheetDrawing.xsd` and "DM" means `dml-main.xsd`.
- **To check (R6-1):** compare these files with the xsd in the official `ECMA-376-4_5th_edition_december_2016.zip` (the `spec/` copy the TD facts cite).

**S-DX (cross-check only).** `dml-chart.xsd` from plutext/docx4j (`xsd/dml`), an older, annotated copy. After annotations, whitespace and `minOccurs/maxOccurs="1"` are normalised, it has the same element sequences as S-XSD for every type this file uses (CT_ChartSpace, CT_Chart, CT_PlotArea, CT_BarChart, CT_LineChart, CT_PieChart, CT_ScatterChart, CT_AreaChart, EG_SerShared, the series types, EG_AxShared, CT_CatAx, CT_ValAx, CT_Scaling, CT_Legend, CT_Title). Differences that matter here:
- ST_GapAmount, ST_Overlap, ST_HoleSize, ST_LblOffset, ST_BubbleScale and ST_SecondPieSize are plain integers in S-DX. S-XSD also allows a percent string such as `150%`.
- ST_MarkerStyle has no `auto` in S-DX.
- ST_Skip is unsignedShort in S-DX and unsignedInt in S-XSD.
- The S-DX `dml-main.xsd` download was a 404 page and was not used.

**S-POI (Excel files).** Apache POI corpus `tests/corpus/poi/` (Apache-2.0, tests/corpus/LICENSES.md), unzipped in the session. Application and AppVersion are from `docProps/app.xml`:

| File | app.xml | Charts |
|---|---|---|
| WithChart.xlsx | Microsoft Excel 12.0000 | line, standard |
| WithTwoCharts.xlsx | Microsoft Excel 12.0000 | line; area |
| WithThreeCharts.xlsx | Microsoft Excel 12.0000 | line; pie; area |
| chart_sheet.xlsx | Microsoft Excel 12.0000 | column clustered, on a chart sheet |
| WithChartSheet.xlsx | Microsoft Excel 12.0000 | column stacked, pivot chart, on a chart sheet |
| SimpleScatterChart.xlsx | Microsoft Excel 15.0300 | scatter, embedded plus chart sheet |
| chartTitle_noTitle.xlsx | Microsoft Excel 15.0300 | scatter |
| chartTitle_withTitle.xlsx | Microsoft Excel 15.0300 | scatter |
| chartTitle_withTitleFormula.xlsx | Microsoft Excel 16.0300 | scatter |
| 123233_charts.xlsx | "Microsoft Excel 12.0000", Company "IBM Incorporated" | bar ×2, pie, line; **not Excel output**, see CH-29 |

"Excel files" below means the first nine. Every chart part they contain validates against S-XSD with `xmllint --schema dml-chart.xsd` after `mc:AlternateContent` is replaced by its `mc:Fallback` content (with a scratch script that replaces each `mc:AlternateContent` by the content of its `mc:Fallback`; not kept in the repository). Every drawing part validates against SD with dml-chart imported.

**S-LO.** LibreOffice 24.2.7.2 (build 420) in the container. The experiments are in CH-20 and CH-21.

**S-IMPL (observed implementations, not specifications).**
- excelize: github.com/qax-os/excelize at commit c4331ec2d7aab0546515906032d54b81952bbfa0 (master on 2026-11-03), BSD-3-Clause. Files `chart.go`, `drawing.go`, `xmlChart.go`, `xmlDrawing.go`.
- Apache POI: github.com/apache/poi at commit 81ab52621cb3ecaf9cb96a5a92edcb2e5fd60408 (trunk on 2026-11-03), Apache-2.0. Files in `poi-ooxml/src/main/java/org/apache/poi/xddf/usermodel/chart/` and `xssf/usermodel/XSSFChart.java`.
- exceljs 4.4.0 in `bench/node_modules/exceljs`.

## Package

### CH-01 - Chart part - OBSERVED (content type, names) / VERIFIED (root element)
Content type `application/vnd.openxmlformats-officedocument.drawingml.chart+xml`. Every Excel file declares it with an `Override` per part, and so does LibreOffice. Excel names the parts `/xl/charts/chartN.xml`, numbered across the workbook (WithThreeCharts: chart1 and chart2 in drawing1, chart3 in drawing2).
- Root element `chartSpace` (CT_ChartSpace) in namespace `http://schemas.openxmlformats.org/drawingml/2006/chart`, which Excel binds to prefix `c` (S-XSD L8, L1496). Excel declares `xmlns:c`, `xmlns:a` (`…/drawingml/2006/main`) and `xmlns:r` (`…/officeDocument/2006/relationships`) on the root.
- The Part 1 clause for the Chart part (content type, relationship, allowed relationships) was not checked: R6-2.

### CH-02 - Drawing to chart relationship - OBSERVED
The Drawings part (TD-30) references each chart through a relationship. Type: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart`. Target: `../charts/chartN.xml`. Internal. One relationship per chart.
- The `graphicFrame` names the relationship with `c:chart r:id` (CH-04). S-XSD L1498 declares the global element `chart` of type CT_RelId (L27-29: required `r:id`).
- Observed in all ten corpus files and in LibreOffice output.
- The worksheet still references the drawing with `drawing r:id` (TD-30, TD-34).

### CH-03 - Chart style and colour parts - OBSERVED (optional in practice)
Excel 15/16 files give each chart part two relationships:
- type `http://schemas.microsoft.com/office/2011/relationships/chartStyle`, target `style1.xml`, content type `application/vnd.ms-office.chartstyle+xml`;
- type `http://schemas.microsoft.com/office/2011/relationships/chartColorStyle`, target `colors1.xml`, content type `application/vnd.ms-office.chartcolorstyle+xml`.

The Excel 12 files and LibreOffice output have neither. Proposed writer rule: write neither. Their format ([MS-ODRAWXML]) was NOT CHECKED.

### CH-04 - graphicFrame in the drawing - VERIFIED (schema) / OBSERVED (Excel's form)
- **Schema.** CT_GraphicalObjectFrame (SD L77-86) has `nvGraphicFramePr`, `xfrm` and `a:graphic`, all required, plus attributes `macro` (optional) and `fPublished` (default false).
  - CT_GraphicalObjectFrameNonVisual (SD L70-76) holds `cNvPr` and `cNvGraphicFramePr`, both required.
  - CT_NonVisualDrawingProps (DM L810-821) has required `id` and `name`.
  - CT_NonVisualGraphicFrameProperties (DM L850-856) has `graphicFrameLocks?` and `extLst?`.
  - CT_GraphicalObject (DM L870-875) has one `graphicData`. CT_GraphicalObjectData (DM L864-869) has a required `uri` and strictly processed child content.
  - CT_Transform2D (DM L635-643) has `off?` and `ext?`.
  - Anchors are as in TD-32 (CT_TwoCellAnchor SD L147-155).
- **Excel (all embedded charts; WithChart drawing1.xml shown):**
  - `<xdr:twoCellAnchor>` with no `editAs` attribute;
  - `from` and `to` markers;
  - `<xdr:graphicFrame macro="">` containing `<xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` and `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>`;
  - `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="…/drawingml/2006/chart" xmlns:r="…/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic>`;
  - `<xdr:clientData/>`.
- Excel writes `xfrm` with zeros. The anchor carries the position (the meaning of xfrm in a spreadsheet frame is R6-3).
- Variations seen:
  - Excel 16 adds `a:extLst/a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"` with `a16:creationId` inside `cNvPr` (chartTitle_withTitleFormula).
  - WithThreeCharts has `cNvPr id="3073"` / `"2049"` and an empty `<a:graphicFrameLocks/>`.
  - LibreOffice writes `twoCellAnchor editAs="oneCell"`, `cNvPr id="0"`, a non-zero xfrm in EMU and no `macro`.
  - The IBM file uses `oneCellAnchor` with xfrm equal to the anchor size.
  - excelize writes `twoCellAnchor` with `editAs` from an option, `cNvPr name="Chart N"` and an empty `xdr:xfrm` (drawing.go L1598-1650; xmlDrawing.go L284-290).

### CH-05 - Charts on chart sheets (reading) - OBSERVED
In chart_sheet.xlsx, WithChartSheet.xlsx and SimpleScatterChart.xlsx (sheet "Chart1"):
- the workbook lists a sheet whose part has content type `application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml` (`xl/chartsheets/sheet1.xml`, root `chartsheet`);
- that part references a drawing with `<drawing r:id>` and the `…/relationships/drawing` relationship;
- the drawing holds one `xdr:absoluteAnchor` (`pos`, `ext`) with `cNvGraphicFramePr/a:graphicFrameLocks noGrp="1"` and the same `graphicData`/`c:chart` as CH-04.

The chartsheet part is in the ECMA-376 facts table (ecma-376.md, Part 1 §12.3.2). A reader that lists charts should also follow chartsheets. Writing chartsheets is out of scope.

## chartSpace and chart

### CH-06 - CT_ChartSpace and what Excel writes - VERIFIED (order) / OBSERVED
- **Sequence (L1478-1495):** `date1904?`, `lang?`, `roundedCorners?`, `style?`, `clrMapOvr?`, `pivotSource?`, `protection?`, `chart` (required), `spPr?`, `txPr?`, `externalData?`, `printSettings?`, `userShapes?`, `extLst?`.
- **Excel 12** (WithChart, WithTwoCharts, WithThreeCharts, chart_sheet):
  - `<c:lang val="en-US"/>`, sometimes `<c:style val="18"/>`, `chart`;
  - `printSettings` for embedded charts: `<c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/>`;
  - no `date1904` and no `roundedCorners`.
- **Excel 15/16:**
  - `<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>`;
  - then `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent>`;
  - then `chart`, `spPr`, `txPr`, `printSettings`.
- `mc:AlternateContent` is not in S-XSD. The parts validate only after Markup Compatibility processing (CH-29), so a reader must accept it in this position.
- ST_Style is 1 … 48 (L1397-1405).
- LibreOffice writes `lang`, `roundedCorners val="0"`, `chart` and `spPr`.
- Semantics of `date1904`, `roundedCorners` and `lang`: R6-6.

### CH-07 - CT_Boolean - VERIFIED (schema) / R6-5 (absent element)
CT_Boolean (L18-20) has an optional attribute `val` of xsd:boolean with **default true**. `<c:varyColors/>` therefore means true.
- Every corpus and LibreOffice file writes `val` explicitly (`"0"`/`"1"`). excelize does the same through `attrValBool`.
- What an **absent** element means (for example no `roundedCorners`, no `varyColors`, no `autoTitleDeleted`) is not in the schema: R6-5.
- Proposed writer rule: always write `val` explicitly.
- Reader rule: an element without `val` is true. An absent element takes the meaning that R6-5 establishes (PROVISIONAL until then).

### CH-08 - CT_Chart - VERIFIED (order) / OBSERVED
- **Sequence (L1380-1396):** `title?`, `autoTitleDeleted?`, `pivotFmts?`, `view3D?`, `floor?`, `sideWall?`, `backWall?`, `plotArea` (required), `legend?`, `plotVisOnly?`, `dispBlanksAs?`, `showDLblsOverMax?`, `extLst?`.
- **Excel:**
  - `plotVisOnly val="1"` in every file;
  - `autoTitleDeleted`: absent in the Excel 12 files; Excel 15/16 write `val="1"` without a title (chartTitle_noTitle) and `val="0"` with a title (chartTitle_withTitle, …WithTitleFormula, SimpleScatterChart);
  - `dispBlanksAs`: `gap` in most files, `zero` in WithThreeCharts chart3 (area), absent in WithChart, WithTwoCharts and WithThreeCharts chart2 (pie);
  - Excel 15/16 add `showDLblsOverMax val="0"`.
- ST_DispBlanksAs: `span`, `gap`, `zero`, default `zero` (L1370-1379).
- excelize writes `plotVisOnly val="0"` (drawing.go L84), unlike Excel.
- Semantics of autoTitleDeleted, plotVisOnly and dispBlanksAs: R6-7, R6-8.

### CH-09 - CT_PlotArea and several chart groups - VERIFIED
**Sequence (L1285-1316):**
- `layout?`;
- then a choice, 1 … unbounded, of `areaChart`, `area3DChart`, `lineChart`, `line3DChart`, `stockChart`, `radarChart`, `scatterChart`, `pieChart`, `pie3DChart`, `doughnutChart`, `barChart`, `bar3DChart`, `ofPieChart`, `surfaceChart`, `surface3DChart`, `bubbleChart`;
- then a choice, 0 … unbounded, of `valAx`, `catAx`, `dateAx`, `serAx`;
- then `dTable?`, `spPr?`, `extLst?`.

A plot area may hold several chart-type elements (a combination chart), and all of them come before the axes. A reader therefore returns a list of chart groups, each with its own type, series and `axId` pair. Every Excel file writes `<c:layout/>` first. The IBM file does not (CH-29).

## Chart types

### CH-10 - Chart type elements and enumerations - VERIFIED
| Requested type | Element | Distinguishing children (S-XSD) |
|---|---|---|
| column / bar | `barChart` (CT_BarChart L898-907) | `barDir` **required**: ST_BarDir `bar` \| `col`, default `col` (L867-875). `grouping` **optional**: ST_BarGrouping `percentStacked` \| `clustered` \| `standard` \| `stacked`, default `clustered` (L856-866). EG_BarChartShared L889-897. |
| line | `lineChart` (CT_LineChart L783-793) | `grouping` **required**: ST_Grouping `percentStacked` \| `standard` \| `stacked`, default `standard` (L759-768). EG_LineChartShared L774-782. |
| pie | `pieChart` (CT_PieChart L949-955) | no grouping, no axes. EG_PieChartShared L942-948. |
| scatter | `scatterChart` (CT_ScatterChart L826-835) | `scatterStyle` **required**: ST_ScatterStyle `none` \| `line` \| `lineMarker` \| `marker` \| `smooth` \| `smoothMarker`, default `marker` (L813-825). |
| area | `areaChart` (CT_AreaChart L927-933) | `grouping` **optional**: ST_Grouping (EG_AreaChartShared L918-926). |

- Column vs bar is only `barDir`.
- `standard` in ST_BarGrouping has no observed use: R6-12.
- Observed in Excel: `barDir col` + `grouping clustered` (chart_sheet), `col` + `stacked` (WithChartSheet), `lineChart grouping standard`, `pieChart`, `scatterChart scatterStyle lineMarker`, `areaChart grouping standard`.
- Not in any Excel file of the corpus: `barDir bar`, `percentStacked`, line `stacked` / `percentStacked`, area `stacked`.
- LibreOffice 24.2.7.2 renders all of them and writes them back unchanged (CH-20). excelize maps types the same way (chart.go L432-456, L473, L476-494).

### CH-11 - CT_BarChart, gapWidth and overlap - VERIFIED / OBSERVED
**Sequence (L898-907):** EG_BarChartShared (`barDir`, `grouping?`, `varyColors?`, `ser*`, `dLbls?`), then `gapWidth?`, `overlap?`, `serLines*`, `axId` exactly twice, `extLst?`.

Value types:
- `gapWidth` is CT_GapAmount, default `150%`. ST_GapAmount (L290-306) is a union of a percent string 0%-500% and unsignedShort 0-500.
- `overlap` is CT_Overlap, default `0%`. ST_Overlap (L307-323) is a union of a percent string -100%…100% and a byte -100…100.
- S-DX has only the integer forms. Proposed writer rule: write integers, as Excel does.

Observations:
- Excel writes `<c:overlap val="100"/>` with `grouping stacked` (WithChartSheet). No Excel clustered chart writes overlap (chart_sheet).
- excelize writes overlap 100 for stacked and percentStacked bar and column charts (chart.go L240-245, drawing.go L758-770).
- LibreOffice writes `overlap 100` for stacked and percentStacked and `overlap 0` for clustered.

Whether stacking needs overlap 100 is R6-12. Proposed writer rule: write overlap 100 with stacked and percentStacked (PROVISIONAL).

### CH-12 - CT_LineChart - VERIFIED / OBSERVED
**Sequence (L783-793):** EG_LineChartShared (`grouping`, `varyColors?`, `ser*`, `dLbls?`, `dropLines?`), then `hiLowLines?`, `upDownBars?`, `marker?`, `smooth?`, `axId` exactly twice, `extLst?`.

Excel 12 writes `<c:marker val="1"/>` after the series (WithChart, WithThreeCharts chart1, WithTwoCharts chart1) and no `varyColors`.

### CH-13 - CT_PieChart - VERIFIED / OBSERVED
**Sequence (L949-955):** `varyColors?`, `ser*`, `dLbls?`, `firstSliceAng?`, `extLst?`.
- `firstSliceAng` is 0-360, default 0 (L350-358).
- There is no `axId`, and a plot area with only a pie chart has no axes (WithThreeCharts chart2).
- Excel writes `<c:varyColors val="1"/>`, one series with `<c:explosion val="25"/>`, and `<c:firstSliceAng val="0"/>`.
- With varyColors, each slice needs a colour (CH-21).

### CH-14 - CT_ScatterChart - VERIFIED / OBSERVED
**Sequence (L826-835):** `scatterStyle`, `varyColors?`, `ser*` (CT_ScatterSer L672-685), `dLbls?`, `axId` exactly twice, `extLst?`.

Scatter series use `xVal` (CT_AxDataSource) and `yVal` (CT_NumDataSource) in place of `cat`/`val`.

Excel 15:
- `scatterStyle lineMarker`, `varyColors 0`;
- per series `spPr` (an `a:ln` with `a:schemeClr accent1`) and `marker` (`symbol circle`, `size 5`);
- `xVal`/`yVal` as `numRef`, then `smooth val="0"`;
- two `valAx`: the first `axPos b`, the second `axPos l`, both `crossBetween midCat` (chartTitle_*, SimpleScatterChart).

excelize writes `crossBetween between` for scatter (chart.go L422), unlike Excel. How scatterStyle relates to the series' line and marker properties is R6-13.

### CH-15 - CT_AreaChart - VERIFIED / OBSERVED
**Sequence (L927-933):** EG_AreaChartShared (`grouping?`, `varyColors?`, `ser*`, `dLbls?`, `dropLines?`), then `axId` exactly twice, `extLst?`.

Excel 12 writes `grouping standard`, no `varyColors`, and `valAx crossBetween midCat` (WithThreeCharts chart3, WithTwoCharts chart2). excelize also uses midCat for area (chart.go L374).

## Series

### CH-16 - Series elements and order - VERIFIED / OBSERVED
EG_SerShared (L650-657) is `idx` (required), `order` (required), `tx?` (CT_SerTx) and `spPr?`. The type-specific sequences follow:

| Series type | Sequence after EG_SerShared |
|---|---|
| CT_BarSer (L697-711) | `invertIfNegative?`, `pictureOptions?`, `dPt*`, `dLbls?`, `trendline*`, `errBars?`, `cat?`, `val?`, `shape?`, `extLst?` |
| CT_LineSer (L658-671) | `marker?`, `dPt*`, `dLbls?`, `trendline*`, `errBars?`, `cat?`, `val?`, `smooth?`, `extLst?` |
| CT_PieSer (L725-735) | `explosion?`, `dPt*`, `dLbls?`, `cat?`, `val?`, `extLst?` |
| CT_AreaSer (L712-724) | `pictureOptions?`, `dPt*`, `dLbls?`, `trendline*`, `errBars?` (0-2), `cat?`, `val?`, `extLst?` |
| CT_ScatterSer (L672-685) | `marker?`, `dPt*`, `dLbls?`, `trendline*`, `errBars?` (0-2), `xVal?`, `yVal?`, `smooth?`, `extLst?` |

- Every data source is optional in the schema.
- `idx` and `order` are independent. WithThreeCharts chart3 has a series with `idx 1, order 0` before one with `idx 0, order 1`. Their meanings and uniqueness scope are R6-10.
- Excel numbers both from 0 in document order in all other files.
- Excel 16 adds `c:extLst/c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"` with `c16:uniqueId` to a series (chartTitle_withTitleFormula). Proposed writer rule: omit it.

### CH-17 - Series name - VERIFIED / OBSERVED
CT_SerTx (L132-139) is a choice of `strRef` (a cell reference) or `v` (literal text, ST_Xstring). Titles use CT_Tx (CH-22), which has no `v`. Excel writes both forms of the series name:
- **literal:** `<c:tx><c:v>1st Column</c:v></c:tx>` (WithChart, WithTwoCharts, WithThreeCharts; Excel 12);
- **reference:** `<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Y</c:v></c:pt></c:strCache></c:strRef></c:tx>` (SimpleScatterChart; chart_sheet `Sheet1!$A$1`; WithChartSheet `Sheet2!$C$3`).

Also observed:
- POI writes `v` without a cell reference and `strRef` with a one-point cache otherwise (XDDFChartData.java L173-197).
- excelize always writes `strRef` with only `f` (drawing.go L814-818).

### CH-18 - Data references (`f`) - VERIFIED (structure) / OBSERVED (form) / R6-9 (grammar)
The data source types are:
- CT_NumDataSource (L63-70): `numRef` or `numLit`.
- CT_AxDataSource (L121-131): `multiLvlStrRef`, `numRef`, `numLit`, `strRef` or `strLit`.
- CT_NumRef (L56-62): `f` (xsd:string, required), `numCache?`, `extLst?`.
- CT_StrRef (L84-90): `f`, `strCache?`, `extLst?`.
- CT_MultiLvlStrRef (L114-120): `f`, `multiLvlStrCache?`, `extLst?`.

Forms of `f` observed in Excel files: `Sheet1!$A$1:$A$6`, `Sheet1!$B$1:$E$1` (a row), `Sheet2!$C$3`, `Sheet1!$E$1` (title). All are absolute, with no `=` and with unquoted simple sheet names.

Other forms observed:
- The IBM file has relative ones (`data_Page1_1_1!A2:A2`).
- LibreOffice keeps `'Sales Data'!$B$2:$B$5` (written with `&apos;`, which is the same character data) and drops unnecessary quotes: `'Sales'!$D$2:$D$5` became `Sales!$D$2:$D$5`.
- For two-column categories (pivot chart) Excel uses `multiLvlStrRef` with `Sheet2!$A$4:$B$11` and a `multiLvlStrCache` of `lvl` elements (WithChartSheet).

The schema gives `f` only as xsd:string. That it follows the cell-formula reference grammar is an inference: R6-9. Proposed writer rule: write `'Sheet name'!$A$1:$A$12`, quoting by the rules of F2-11 (PROVISIONAL). The task's example `'Sales'!$B$2:$B$13` is accepted by LibreOffice.

No Excel file in the corpus has a plain `c:cat/c:strRef`. The form is from the schema and LibreOffice output (CH-20).

### CH-19 - Caches - VERIFIED (structure) / OBSERVED (writers)
- CT_NumData (L48-55) is `formatCode?` (ST_Xstring), `ptCount?`, `pt*`, `extLst?`. CT_NumVal (L41-47) has a required `idx` attribute, a required child `v` (ST_Xstring) and an optional `formatCode` attribute.
- CT_StrData (L77-83) is `ptCount?`, `pt*`, `extLst?` (no formatCode). CT_StrVal (L71-76) has `idx` and `v`.
- Every cache element is optional in the schema (`numCache?`, `strCache?`).

Writers:
- **Excel.** Every reference in every Excel chart part has its cache: 2 `numRef`s with 2 `numCache`s in WithChart, 1 + 1 in WithChartSheet, and so on. Excel writes `numCache` as `<c:formatCode>General</c:formatCode><c:ptCount val="N"/>` plus one `pt` per value, and `strCache` as `ptCount` plus `pt`s.
- **IBM generator.** The 123233 file has references with no caches at all.
- **excelize** writes an empty `numCache`. cNumCache has `FormatCode string` without omitempty and nil `Pt`/`PtCount` (xmlChart.go L552-556; drawing.go L1051-1058), so the cache has only an empty `formatCode`. excelize writes no `strCache` (xmlChart.go L518-521, drawing.go L814-818).
- **POI** fills caches from the data on `plot()` (XDDFChartData.java L207-221).

Whether a consumer needs or uses the caches is R6-4. LibreOffice ignores them (CH-20).

### CH-20 - LibreOffice and caches (experiment) - OBSERVED (2026-11-03)
**Files.** A scratch script (not kept in the repository) builds minimal workbooks with fflate 0.8.3. Each has:
- sheet "Sales Data" with A1:B5 = Month/Sales, Jan…Apr and 10, 20, 30, 40;
- one twoCellAnchor drawing;
- one `barChart` (`col`, `clustered`, `varyColors 0`) with one series: `tx` = `'Sales Data'!$B$1`, `cat` = `'Sales Data'!$A$2:$A$5`, `val` = `'Sales Data'!$B$2:$B$5`;
- catAx/valAx as in CH-23.

Every chart part validates against S-XSD. The variants differ as follows:

| Variant | Caches | Fill | Theme part | Note |
|---|---|---|---|---|
| A | filled correctly | none | none | |
| B | omitted (`f` only) | none | none | |
| C | stale (`STALE NAME`, S1…S4, 400…100) | none | none | |
| D | omitted | red series fill | none | plus title, axis titles, valAx min 5 / max 50, legend `t` |
| E | filled | red series fill | none | |
| F | omitted | none | Excel 12 theme copied from WithChart | |
| G | stale | red series fill | none | |
| H | filled | none | Excel 12 theme copied from WithChart | |
| I | E's caches | red series fill | none | with all sheet cells removed (by a scratch script) |

**Commands** (S = a scratch working folder):
`soffice --headless -env:UserInstallation=file://$S/lo-profile --convert-to png --outdir $S/lo-png <files>`
and the same with `--convert-to xlsx --outdir $S/lo-xlsx`.

**Results.**
- A, B and C give byte-identical PNGs (sha256 9029a35c…). So do E and G (11d02bd2…) and F and H (2b984f43…).
- The axes, category labels (Jan…Apr, not S1…S4), axis scale (0-45), bar heights (10-40) and legend text ("Sales", not "STALE NAME") come from the cells in every variant.
- With the cells removed (I), LibreOffice shows an empty plot with its own series name "Column B" and writes back empty caches. LibreOffice 24.2.7.2 therefore **does not use the caches** of a chart over cells in the same workbook.
- On save it writes fresh `strCache`/`numCache` (formatCode `General`, `ptCount`, `pt`) for all three references, whether the input had them or not. The output of A and B is identical except for the axis ids, and every output chart part validates.
- LibreOffice keeps the chart, title, axis titles (the valAx title with `a:bodyPr rot="-5400000"`), `max`/`min`, legend `t` and the red fill (written `ff0000`).
- It re-anchors the frame: `twoCellAnchor editAs="oneCell"`, from row 0 with a row offset in place of row 1.

The behaviour of Excel without caches was NOT CHECKED (Excel is not available). Proposed writer rule: write caches computed from the written cell values, as Excel and POI do, for the best chance with other consumers (PROVISIONAL, R6-4).

### CH-21 - Series colour and the theme - VERIFIED (schema) / OBSERVED (LibreOffice)
- **Schema.** Series colour goes in `c:ser/c:spPr` (a:CT_ShapeProperties, DM L2245-2257: `xfrm?`, geometry, a fill choice, `ln?`, …).
  - The fill choice EG_FillProperties (DM L1599-1608) includes `noFill` and `solidFill`.
  - `solidFill` (DM L1413-1417) takes one EG_ColorChoice (DM L689-698), for example `a:srgbClr val` (CT_SRgbColor DM L337-342).
  - The colour type is ST_HexColorRGB = xsd:hexBinary of length 3, that is 6 hex digits (shared-commonSimpleTypes.xsd L9-12, skills-bundle copy).
- **Observed forms.**
  - For bars, Excel (IBM file) and LibreOffice use `a:solidFill`. LibreOffice writes lower-case hex (`ff0000`), which hexBinary allows.
  - For line and scatter series, Excel 15 and excelize put the colour on `a:ln/a:solidFill` (excelize drawing.go L1021-1046).
  - Excel 15 uses theme colours (`a:schemeClr val="accent1"`).
- **LibreOffice** (CH-20 A vs E/F; type montage `lo-types/montage.png`):
  - Without a theme part and without `spPr`, the series are drawn with no fill and no gridline or axis colour. LibreOffice writes the series back as `<a:noFill/>`.
  - The pie (varyColors 1, no `dPt`) is invisible. LibreOffice writes four `dPt` with `noFill`.
  - With Excel's theme part added (F, H), the same file draws in accent1 `4f81bd`. With an explicit `srgbClr` (D, E, G), it draws in that colour.
- **Implication for xlvane**, which writes no theme part (`packages/xlvane/src` has no theme writer): write an explicit colour for every series and, for pie with varyColors, for every data point. The alternative is a theme part, which is a maintainer decision (PROVISIONAL). How Excel colours series without a theme was NOT CHECKED.

## Titles, axes, legend

### CH-22 - Chart title - VERIFIED / OBSERVED
CT_Title (L178-187) is `tx?`, `layout?`, `overlay?`, `spPr?`, `txPr?`, `extLst?`. Its `tx` is CT_Tx (L91-98): a choice of `strRef` or `rich`.
- `rich` is a:CT_TextBody (DM L2681-2687): `bodyPr` (required), `lstStyle?`, `p` (one or more).
- A paragraph (CT_TextParagraph DM L2562-2567) is `pPr?`, then runs (EG_TextRun DM L3068-3074: `r`, `br` or `fld`), then `endParaRPr?`.
- A run (CT_RegularTextRun DM L3075-3080) is `rPr?` and `t`.

Forms observed:
1. **Literal** (chartTitle_withTitle, Excel 15): `<c:title><c:tx><c:rich><a:bodyPr …/><a:lstStyle/><a:p><a:pPr><a:defRPr …/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>Original title set in Excel 2013</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/><c:spPr>…</c:spPr><c:txPr>…</c:txPr></c:title>`.
2. **Literal in two runs** (WithThreeCharts chart2): `Pie Chart Title` + ` Thingy`. A reader concatenates the `a:t` of all runs. Paragraph breaks are R6-11.
3. **Reference** (chartTitle_withTitleFormula, Excel 16): `<c:tx><c:strRef><c:f>Sheet1!$E$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Formula Title from Excel 2016</c:v></c:pt></c:strCache></c:strRef></c:tx>`, followed by `overlay`, `spPr`, `txPr` (no `layout`).
4. **`title` without `tx`** (SimpleScatterChart, WithChartSheet): an automatic title whose text the file does not contain (R6-7).
5. **No `title`**, with `autoTitleDeleted val="1"` (chartTitle_noTitle) or with no autoTitleDeleted (Excel 12 files).

Other implementations:
- LibreOffice writes the literal form with `a:rPr` per run and `a:bodyPr rot="0"`.
- excelize writes `rich` or `strRef` without a cache (drawing.go L1314-1345).
- POI `removeTitle` sets `autoTitleDeleted` true (XDDFChart.java L247-252, L328).

### CH-23 - Axes - VERIFIED (schema) / OBSERVED (Excel's values)
**EG_AxShared (L1220-1241)** in order:
- `axId` (required);
- `scaling` (required);
- `delete?`;
- `axPos` (required; ST_AxPos `b`, `l`, `r`, `t`, no default, L1037-1047);
- `majorGridlines?`, `minorGridlines?`;
- `title?`;
- `numFmt?` (CT_NumFmt L410-413: required `formatCode`, optional `sourceLinked`);
- `majorTickMark?`, `minorTickMark?` (ST_TickMark `cross`, `in`, `none`, `out`, default `cross`, L1067-1077);
- `tickLblPos?` (ST_TickLblPos `high`, `low`, `nextTo`, `none`, default `nextTo`, L1078-1088);
- `spPr?`, `txPr?`;
- `crossAx` (required);
- an optional choice of `crosses` (ST_Crosses `autoZero`, `max`, `min`, required val, L1048-1057) or `crossesAt` (CT_Double).

**Type-specific tails.**
- CT_CatAx (L1242-1253) adds `auto?`, `lblAlgn?` (`ctr`, `l`, `r`), `lblOffset?`, `tickLblSkip?`, `tickMarkSkip?`, `noMultiLvlLbl?`, `extLst?`.
- CT_ValAx (L1275-1284) adds `crossBetween?` (ST_CrossBetween `between`, `midCat`, required val, no default, L1058-1066), `majorUnit?`, `minorUnit?`, `dispUnits?`, `extLst?`.

**Excel 12 column, line and area charts:**
- catAx: `axId`, `scaling/orientation minMax`, `axPos b`, (`numFmt General sourceLinked=1`), `tickLblPos nextTo`, `crossAx`, `crosses autoZero`, `auto 1`, `lblAlgn ctr`, `lblOffset 100`.
- valAx: `axId`, `scaling/orientation minMax`, `axPos l`, `majorGridlines`, `numFmt General sourceLinked=1`, `tickLblPos nextTo`, `crossAx`, `crosses autoZero`, `crossBetween between` (bar, line) or `midCat` (area).

Excel 15 adds `delete val="0"`, `majorTickMark none`, `minorTickMark none`, `spPr` and `txPr`. Each axis's `crossAx` names the other's `axId`.

Horizontal bars (`barDir bar`) have no Excel sample:
- excelize writes catAx `b` and valAx `l` regardless of barDir (chart.go L554-561, drawing.go L1131-1230).
- A file written with catAx `l` and valAx `b` rendered horizontally in LibreOffice, which saved it back with catAx `b` and valAx `l`.

The meaning of `axPos` for horizontal bars is R6-14. The semantics of crosses, crossesAt and crossBetween are R6-14.

### CH-24 - Axis ids - VERIFIED (type, counts) / OBSERVED (values)
- `axId` is CT_UnsignedInt (L24-26, required `val`).
- barChart, lineChart, areaChart and scatterChart have exactly two `axId` (L904, L790, L930, L832). The 3-D variants have 2-3. pieChart has none.
- Each id matches the `axId` of an axis element in the same plot area.
- Values written:

| Writer | Ids |
|---|---|
| Excel 12 | 49660672 / 73939968, 119628544 / 119630080, … |
| Excel 15 | 327141000 / 327140608, … |
| IBM file | 0 / 2 |
| excelize | 100000000 / 100000001 (drawing.go L1809-1815) |
| LibreOffice | 26733192 / 41809333 |

Rules on uniqueness and allowed values: R6-15. Proposed writer rule: fixed ids per chart group, unique within the chart part (PROVISIONAL).

### CH-25 - Axis title - VERIFIED (schema) / NOT FOUND in Excel files
An axis title is EG_AxShared `title` (CT_Title, as CH-22), after the gridlines and before `numFmt`. No Excel file in the corpus has an axis title.
- The IBM file has a valAx title (rich text "Revenue").
- LibreOffice reads and writes `catAx/title` and `valAx/title` (CH-20 D).
- excelize writes both axis titles through the same title function (drawing.go L1152, L1209). The y-axis title gets `a:bodyPr rot="-5400000" vert="horz"` (drawing.go L1345-1347, L1358-1360).

### CH-26 - Axis minimum and maximum - VERIFIED (schema) / OBSERVED (LibreOffice)
CT_Scaling (L1194-1202) is `logBase?` (2-1000), `orientation?` (ST_Orientation `maxMin`, `minMax`, default `minMax`), `max?`, `min?`, `extLst?`. **`max` comes before `min`.** Both are CT_Double (required `val`).
- No Excel file in the corpus sets them.
- LibreOffice applies and keeps `<c:max val="50"/><c:min val="5"/>` (CH-20 D: the axis is drawn from 5).
- excelize writes them in the same order (drawing.go L1186-1206).

### CH-27 - Legend - VERIFIED / OBSERVED
CT_Legend (L1359-1369) is `legendPos?`, `legendEntry*`, `layout?`, `overlay?`, `spPr?`, `txPr?`, `extLst?`. ST_LegendPos (L1332-1342) is `b`, `tr`, `l`, `r` or `t`, **default `r`**.

What the writers produced:

| Writer | Legend markup |
|---|---|
| Excel 12 | `<c:legend><c:legendPos val="r"/><c:layout/></c:legend>`, sometimes with `txPr` |
| Excel 15 | `legendPos b`, `layout`, `overlay val="0"`, `spPr`, `txPr` (SimpleScatterChart) |
| chartTitle_* files | no `legend` element |
| LibreOffice | `t`, `b`, `l` and `r` kept on round trip; `tr` written back as `r` (bar-stacked test) |

excelize maps `bottom`, `left`, `right`, `top` and `top_right` to `b`, `l`, `r`, `t` and `tr` (chart.go L307-313).

That an absent `legend` means "no legend" is inferred from the chartTitle files (PROVISIONAL). The meaning of `tr` and of `overlay` is R6-16.

### CH-28 - varyColors - VERIFIED (schema) / OBSERVED
`varyColors` (CT_Boolean) is optional in every chart group. It is written as follows:

| Writer | varyColors |
|---|---|
| Excel 12, pie | `val="1"` |
| Excel 12, bar (chart_sheet, WithChartSheet), line and area | absent |
| Excel 15, scatter | `val="0"` |
| IBM file, bar and line | `val="0"` |
| IBM file, pie | `val="1"` |
| LibreOffice | kept; area charts written without it |

Its meaning and the meaning of its absence are R6-5. Proposed writer rule: write `val="0"` for bar, line, area and scatter, and `val="1"` for pie (PROVISIONAL).

## Conformance and other libraries

### CH-29 - Element order and Markup Compatibility - OBSERVED
- **Excel files.** Every Excel chart part validates against S-XSD, but only after `mc:AlternateContent` is resolved to `mc:Fallback`. Without that step, Excel 15/16 parts have an element the schema does not allow before `c:chart`.
- **The 123233 file.** Its app.xml says "Microsoft Excel" 12.0000 with Company "IBM Incorporated" and sheet names `data_Page1_1_N`. Its parts break the schema order:
  - catAx has `crossAx` before `axPos` and `scaling`;
  - `dispBlanksAs` comes before `legend`;
  - charts 1, 3 and 4 fail at `crossAx`, chart 2 at `ser`.
  - It also has no caches.
  - Conclusion: the file was not written by Excel's own chart writer (inference; the generator is not named in the file).
- Proposed rules: the writer follows S-XSD order exactly. The reader finds children by name and does not depend on order, so the 123233 charts can still be read (implementation choice). Whether Excel opens such files was NOT CHECKED.

### CH-30 - exceljs 4.4.0 - OBSERVED
- **No chart API.** In `bench/node_modules/exceljs` (package.json version 4.4.0), "chart" (case-insensitive) occurs in `lib/` only in a comment in `lib/xlsx/xform/book/workbook-xform.js` L174-178 ("As we don't have the infrastructure to support chartsheets, we will ignore them for now"). It does not occur in `index.d.ts`.
- **Round trip drops charts.** `readFile` then `writeFile` of WithChart.xlsx produces a file with no chart or drawing part (checked with a scratch script).

@xlvane/compat therefore has no exceljs chart API to mirror.

### CH-31 - Defaults of other writers - OBSERVED (not specification)
- **excelize** `addChart` (drawing.go L59-110) writes:
  - `date1904 0`, `lang en-US`, `roundedCorners 0`;
  - `view3D`, `floor`, `sideWall` and `backWall` even for 2-D types;
  - `plotVisOnly 0` and `showDLblsOverMax 0`;
  - chart-space `spPr` with `schemeClr bg1`;
  - `printSettings`.
- **POI** `XSSFChart.createChart` (XSSFChart.java L96-113) writes `plotArea/layout`, `plotVisOnly true` and `printSettings` with the same margins as Excel.
- Neither is a reference for the format; they are listed where they differ from Excel (CH-08 plotVisOnly, CH-14 crossBetween, CH-19 caches).

### CH-32 - Office extensions - OBSERVED
Excel 15/16 write the following. Excel 12 files have none of them, and LibreOffice writes only `c15:showLeaderLines` inside `dLbls`.
- `mc:AlternateContent` with `c14:style` (CH-06);
- `c16:uniqueId` (CH-16);
- `xmlns:c16r2` on the root (chartTitle_withTitleFormula);
- `a16:creationId` (CH-04);
- style and colour parts (CH-03).

Proposed writer rule: omit all of them (PROVISIONAL). The reader skips unknown `extLst` content and resolves `mc:AlternateContent` (CH-29).

### CH-33 - LibreOffice and the requested types - OBSERVED (2026-11-03)
A scratch script (not kept in the repository) writes one workbook per type: column clustered, stacked and percentStacked; bar clustered and stacked; line standard, stacked and percentStacked; pie; scatter; area standard and stacked. Each workbook has:
- two series with `tx`/`cat`/`val` (scatter: `xVal`/`yVal`) as `strRef`/`numRef` without caches;
- explicit `srgbClr` colours (none on pie);
- a rich title and the legend positions `r`, `b`, `t`, `l` and `tr`;
- no theme part.

All chart parts validate against S-XSD. LibreOffice 24.2.7.2 (same commands as CH-20) results:
- It renders every type except the pie, which has no colour (CH-21); see `lo-types/montage.png`.
- On save it keeps every `barDir`/`grouping`/`scatterStyle`. It writes `overlap 100` for stacked and percentStacked bars and `overlap 0` for clustered, writes caches, and changes legend `tr` to `r`.
- Every output part validates.

### CH-34 - Reading model - implementation proposal (PROVISIONAL)
For each chart reached through a worksheet's or chartsheet's drawing (CH-02, CH-05), the reader returns:
- the anchor (TD-32);
- the chart groups in plot-area order (CH-09), each with:
  - its element name mapped to a type, plus `barDir`, `grouping` and `scatterStyle` (CH-10);
  - its series in document order with `idx`, `order` and name, where the name is the `v` text, or the `f` of `strRef` plus the cached text (CH-17);
  - `cat`/`val`/`xVal`/`yVal` as `f` plus the cache values when present (CH-18, CH-19);
  - the series colour when it is a `solidFill/srgbClr` (CH-21);
- the title: literal text (concatenated runs), reference (`f` plus cache), automatic, or none (CH-22, CH-07);
- the axis titles and `scaling` min/max (CH-25, CH-26);
- the legend position (absent element = no legend; element without `legendPos` = `r`, the schema default) (CH-27).

`f` is returned as text and never evaluated against untrusted input beyond the existing reference parser.

## Added in phase K (implementation)

### CH-35 - Default series palette - OBSERVED (2026-11-03)
The writer needs explicit colours (CH-21, maintainer's decision: no theme part). The default palette is the six accent colours of the theme part of an Excel-written file of the POI corpus:
- **File:** `tests/corpus/poi/chartTitle_withTitle.xlsx` (Apache-2.0; `docProps/app.xml`: Application "Microsoft Excel", AppVersion 15.0300), part `xl/theme/theme1.xml`, root `a:theme name="Office Theme"`, `a:clrScheme name="Office"`.
- **Values** (`a:accentN/a:srgbClr val`): accent1 `5B9BD5`, accent2 `ED7D31`, accent3 `A5A5A5`, accent4 `FFC000`, accent5 `4472C4`, accent6 `70AD47`.
- The other Excel 15/16 files of the corpus (chartTitle_noTitle, chartTitle_withTitleFormula, SimpleScatterChart) have the same six values. The Excel 12 files (WithChart, WithTwoCharts, WithThreeCharts, chart_sheet, WithChartSheet) have `4F81BD`, `C0504D`, `9BBB59`, `8064A2`, `4BACC6`, `F79646`.
- Implementation: series (and pie slices) take accent1 … accent6 in order and start again after the sixth. How Excel colours a seventh series (it varies the shades) was NOT CHECKED; repeating the six is an implementation choice.
- Read with `grep -o '<a:accent[1-6]><a:srgbClr val="[0-9A-F]*"'` on the unzipped part.

### CH-36 - LibreOffice and literal series names - OBSERVED (2026-11-03)
LibreOffice 24.2.7.2 re-saves a series name written as `<c:tx><c:v>Y values</c:v></c:tx>` (CH-17) as `<c:tx><c:strRef><c:f>"Y values"</c:f><c:strCache>…<c:v>Y values</c:v>…</c:strCache></c:strRef></c:tx>`: a reference whose `f` is a quoted string (`&quot;` in the file). The xlvane reader reports such a name as `{ ref: '"Y values"' }`, since it never evaluates `f` (CH-34). Observed with `scripts/chart-samples.mjs` (scatter sample) and in `tests/libreoffice/charts.test.ts`.

### CH-37 - LibreOffice and the written charts - OBSERVED (2026-11-03)
Files written by xlvane (chart parts as in `packages/xlvane/src/charts.ts`, no theme part) validate against S-XSD (`xmllint --schema dml-chart.xsd`; drawing parts against SD with dml-chart imported), for every type and stacking variant of `scripts/chart-samples.mjs`. LibreOffice 24.2.7.2 draws every one of them in the palette colours (CH-35), including the pie (one `dPt` per slice) and the scatter markers, and the streamed variant without caches (CH-20). Axis lines and gridlines are drawn without colour, as CH-21 observed for files without a theme part; xlvane writes no axis or gridline colour (no fact on Excel's behaviour without a theme). Re-saved as XLSX, the types, grouping, titles, legend positions, axis titles, `min`/`max` and references come back (`tests/libreoffice/charts.test.ts`).
