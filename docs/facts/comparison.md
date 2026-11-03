# Facts behind the comparison table in README.md

Every cell of the table "xlvane compared with other libraries" in the root README is backed by one of the facts below. Each fact names the source it was read in, the version or commit, and the date it was checked. All checks were made on 2026-11-03 (docs/decisions.md D-K02). Performance is not compared.

## Sources

| Id | Source | Version or commit | How it was read |
|---|---|---|---|
| NPM-XLVANE | https://registry.npmjs.org/xlvane, https://registry.npmjs.org/@xlvane%2Fcompat, https://registry.npmjs.org/@xlvane%2Fmcp | registry documents as served on 2026-11-03 | `curl`; fields `dist-tags`, `time`, `versions[v].license`, `versions[v].dependencies`, `versions[v].dist.attestations` |
| NPM-EXCELJS | https://registry.npmjs.org/exceljs | as served on 2026-11-03 | `curl`; same fields |
| NPM-XLSX | https://registry.npmjs.org/xlsx | as served on 2026-11-03 | `curl`; same fields |
| EXCELJS-README | `README.md` of https://registry.npmjs.org/exceljs/-/exceljs-4.4.0.tgz | 4.4.0; the tarball's sha512 equals the registry's `dist.integrity` (`sha512-XctvKaEM…`) | downloaded, integrity compared, unpacked |
| XLSX-README | `README.md` of https://registry.npmjs.org/xlsx/-/xlsx-0.18.5.tgz | 0.18.5; the tarball's sha512 equals the registry's `dist.integrity` (`sha512-dmg3LCjB…`) | downloaded, integrity compared, unpacked |
| SHEETJS-GH | https://raw.githubusercontent.com/SheetJS/sheetjs/github/README.md | branch `github` at `3f44ddd99c704d7ca35ea76fef7b8af2d2d435cb` (`git ls-remote`) | `curl` |
| EXCELIZE | https://github.com/xuri/excelize: `README.md`, `go.mod`, `LICENSE`, `chart.go`, `calc.go`, `stream.go`, `rows.go`, `styles.go` through raw.githubusercontent.com | tag `v2.11.1` = commit `c4331ec2d7aab0546515906032d54b81952bbfa0` (`git ls-remote`; also the `HEAD` of github.com/qax-os/excelize) | `curl` at the commit |
| GOPROXY | https://proxy.golang.org/github.com/xuri/excelize/v2/@latest | as served on 2026-11-03 | `curl` |
| XLVANE-REPO | this repository: `packages/xlvane/README.md`, `SECURITY.md`, `CHANGELOG.md`, `docs/exceljs-compat.md` | the commit that adds this file | read |

Sources that are not cited: the SheetJS documentation site (docs.sheetjs.com), git.sheetjs.com and xuri.me could not be reached from the environment where the facts were collected.

## Latest version, date, license and runtime dependencies

### CMP-01 xlvane
NPM-XLVANE: `dist-tags.latest` is `1.2.2`, `time["1.2.2"]` is `2026-11-02T20:43:48.111Z`, `license` is `MIT`, `dependencies` of 1.2.2 is `{"fflate":"^0.8.3"}` (one runtime dependency). `@xlvane/compat` 1.2.2 depends on `xlvane` only; `@xlvane/mcp` 1.2.2 on `xlvane` and `@modelcontextprotocol/server` (2.4.0). `dist.attestations.provenance` (`https://slsa.dev/provenance/v1`) is present for `xlvane` and `@xlvane/compat` 1.0.1 to 1.2.2 and for `@xlvane/mcp` 1.2.2; absent for 1.0.0, and for `@xlvane/mcp` 1.2.0 and 1.2.1. Versions 1.2.3 and 1.3.0 were tagged on 2026-11-03 (CHANGELOG.md) and were not yet in the registry document when it was read; the README table gives 1.3.0 with its tag date.

### CMP-02 ExcelJS
NPM-EXCELJS: `dist-tags.latest` is `4.4.0`, `time["4.4.0"]` is `2023-10-19T19:49:55.669Z`, `license` is `MIT`, `dependencies` of 4.4.0 has 9 entries: `archiver`, `dayjs`, `fast-csv`, `jszip`, `readable-stream`, `saxes`, `tmp`, `unzipper`, `uuid`.

### CMP-03 SheetJS Community Edition on npm (`xlsx`)
NPM-XLSX: `dist-tags.latest` is `0.18.5`, `time["0.18.5"]` is `2022-03-24T14:23:09.623Z`, `license` is `Apache-2.0`, `dependencies` of 0.18.5 has 7 entries: `adler-32`, `cfb`, `codepage`, `crc-32`, `ssf`, `wmf`, `word`. SHEETJS-GH says the source repository moved to git.sheetjs.com and that "Scripts and NodeJS modules are available at https://cdn.sheetjs.com"; later versions are therefore not on npm (see docs/xlsx-cves.md for the advisories that say so).

### CMP-04 excelize
GOPROXY: `{"Version":"v2.11.1","Time":"2026-11-02T00:16:17Z", … "Hash":"c4331ec2…"}`. EXCELIZE `LICENSE`: "BSD 3-Clause License". EXCELIZE `go.mod`: module `github.com/xuri/excelize/v2`, `go 1.26.0`, 9 direct requirements (`mscfb`, `testify`, `go-deepcopy`, `efp`, `nfp`, `x/crypto`, `x/image`, `x/net`, `x/text`) and 4 indirect ones. go.mod does not separate test-only modules (testify is one).

## Formats

### CMP-05 xlvane
XLVANE-REPO: CHANGELOG.md 1.0.0 "Reading XLSX, XLSM, XLTX and XLTM files"; writing XLSX (`writeWorkbook`); CSV and JSON in both directions (`toCSV`, `fromCSV`, `toJSON`, `fromJSON`). SECURITY.md "Out of scope": encrypted workbooks and legacy `.xls` files are rejected (`XLV_E_UNSUPPORTED_CONTAINER`, packages/xlvane/README.md "Limitations"). No other spreadsheet format is read or written.

### CMP-06 ExcelJS
EXCELJS-README line 5: "Read, manipulate and write spreadsheet data and styles to XLSX and JSON." Section "File I/O" (line 2137) has the subsections XLSX (reading, writing) and CSV (reading, writing); no other file format is documented.

### CMP-07 SheetJS
XLSX-README section "File Formats" (line 3990): reads and writes XLSX/XLSM, XLSB, SpreadsheetML 2003, XLS (BIFF2 to BIFF8), CSV/TXT, DIF, SYLK, PRN, Numbers, ODS, FODS, DBF, Lotus WK1/WK3, HTML, ETH; reads only UOS, other Lotus and Quattro Pro formats, Works; writes only RTF.

### CMP-08 excelize
EXCELIZE `README.md` "Introduction": "write to and read from XLAM / XLSM / XLSX / XLTM / XLTX files". No other format is named in the README.

## Features

### CMP-09 Cell styles when writing
- xlvane: packages/xlvane/README.md "Writing a workbook" (fonts, fills, borders, alignment, protection, number formats).
- ExcelJS: EXCELJS-README section "Styles" (line 1425).
- SheetJS: XLSX-README lines 7 to 10 and SHEETJS-GH: "[SheetJS Pro] offers solutions beyond data processing: … let out your inner Picasso with styling; make custom sheets with images/graphs/PivotTables; evaluate formula expressions". The README does not describe writing fonts, fills or borders with the Community Edition: "Writing Options" (line 3409) has no style option, and `cellStyles` (line 3276) is a parsing option ("Save style/theme info to the `.s` field").
- excelize: EXCELIZE `styles.go`: `func (f *File) NewStyle(style *Style) (int, error)` (line 991) and `func (f *File) SetCellStyle(…)` (line 2368).

### CMP-10 Charts
- xlvane 1.2.2: charts are not modelled; reading skips them with a warning, `editCells` keeps them unchanged, `insertRows`/`deleteRows` refuse sheets where they would have to move (packages/xlvane/README.md "Limitations"). From 1.3.0, column, bar, line, pie, scatter and area charts are written and read (docs/decisions.md D-K01; packages/xlvane/README.md "Charts").
- ExcelJS: EXCELJS-README has no chart API (the word "chart" occurs only in the release history: "Don't break when loading an Excel file containing a chartsheet #466"); the 4.4.0 tarball has no file whose path contains "chart".
- SheetJS: CMP-09 (graphs are listed under SheetJS Pro). XLSX-README "Chartsheet Object" (line 2451): chartsheets are read as sheets holding the cached data.
- excelize: EXCELIZE `chart.go`: `AddChart` (line 1260), `AddChartSheet` (line 1294), `DeleteChart` (line 1375); README section "Add chart to spreadsheet file". No function that reads a chart into a model was found in `chart.go`.

### CMP-11 Formula calculation
- xlvane: packages/xlvane/README.md "Limitations": "xlvane does not calculate formulas"; SECURITY.md "Out of scope".
- ExcelJS: EXCELJS-README "Formula Value" (line 2623): "Note that ExcelJS cannot process the formula to generate a result, it must be supplied."
- SheetJS: XLSX-README line 2637: "If the actual results are needed in JS, SheetJS Pro offers a formula calculator component".
- excelize: EXCELIZE `calc.go` line 878 `func (f *File) CalcCellValue(sheet, cell string, opts ...Options) (result string, err error)`; its comment (lines 415 to 418): "This feature is currently in working processing. Iterative calculation, implicit intersection, explicit intersection, array formula, table formula and some other formulas are not supported currently."

### CMP-12 Streaming
- xlvane: packages/xlvane/README.md "Large files: streaming": `openWorkbookStream`, `readTableStream` (reading) and `writeWorkbookStream` (writing), on Web Streams.
- ExcelJS: EXCELJS-README "Streaming I/O" (line 2344) with "Streaming XLSX Writer" (line 2366) and "Streaming XLSX Reader" (line 2443).
- SheetJS: XLSX-README "Streaming Write" (line 1730): `XLSX.stream.to_csv`, `to_html`, `to_json` return Node.js streams; for reading, "Readable Streams in NodeJS" (line 971) says the easiest approach is to buffer the stream, and the data is then parsed as a whole.
- excelize: EXCELIZE `README.md` "Introduction": "provided streaming API for generating or reading data from a worksheet with huge amounts of data"; `stream.go` `NewStreamWriter` (line 114), `rows.go` `Rows` (line 339).

### CMP-13 What xlvane does not do
XLVANE-REPO, packages/xlvane/README.md "Limitations" and SECURITY.md "Out of scope": no formula calculation; no encrypted (password-protected) workbooks; no legacy `.xls`; pivot tables, shapes, slicers and form controls are not modelled (and charts until 1.3.0); no format other than XLSX/XLSM/XLTX/XLTM, CSV and JSON.

### CMP-14 exceljs API through @xlvane/compat
XLVANE-REPO, docs/exceljs-compat.md: `@xlvane/compat` offers "the commonly used part of the exceljs 4.4.0 API"; each method is listed as supported, partly supported (with the difference) or not supported.
