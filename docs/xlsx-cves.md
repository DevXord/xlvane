# How to fix the CVEs in the `xlsx` package from npm

xlvane is not affiliated with SheetJS or with the ExcelJS project. Their names are used only to identify the packages this page is about; the facts below are quoted from the cited advisories and registry data, and the fixes SheetJS publishes are listed first.

The `xlsx` package on npm is the SheetJS Community Edition. Its latest version on npm is 0.18.5, published on 2022-03-24, and two published advisories cover that version. The fixed versions are distributed by SheetJS outside npm. This page lists the facts with their sources and two ways to deal with them: the fixed SheetJS version, or moving to xlvane. The sources were read on 2026-11-03.

## The advisories

| CVE | GitHub advisory | What | Affected | Fixed in |
|---|---|---|---|---|
| CVE-2023-30533 | [GHSA-4r6h-8v6p-xvw6](https://github.com/advisories/GHSA-4r6h-8v6p-xvw6) | Prototype pollution when reading specially crafted files (CWE-1321) | SheetJS CE through 0.19.2, including every version of `xlsx` on npm | 0.19.3 |
| CVE-2024-22363 | [GHSA-5pgg-2g8v-p4x9](https://github.com/advisories/GHSA-5pgg-2g8v-p4x9) | Regular expression denial of service, ReDoS (CWE-1333) | SheetJS CE before 0.20.2, including every version of `xlsx` on npm | 0.20.2 |

- GitHub rates both advisories "high" (CVSS 3.1 base scores 7.8 for CVE-2023-30533 and 7.5 for CVE-2024-22363, as listed by the npm registry's advisory data).
- GHSA-4r6h-8v6p-xvw6 says: "Workflows that do not read arbitrary files (for example, exporting data to spreadsheet files) are unaffected."
- Both GitHub advisories say: "A non-vulnerable version cannot be found via npm", and that the fixed version (0.19.3, and 0.20.2) "can be downloaded via https://cdn.sheetjs.com/".
- The CVE records (published 2023-04-24 and 2024-04-05) give the same affected and fixed versions and link to SheetJS's own advisories, `https://cdn.sheetjs.com/advisories/CVE-2023-30533` and `https://cdn.sheetjs.com/advisories/CVE-2024-22363`.

### Sources

| Source | Version | What it gives |
|---|---|---|
| https://registry.npmjs.org/xlsx | registry document, 2026-11-03 | `dist-tags.latest` `0.18.5`; `time["0.18.5"]` `2022-03-24T14:23:09.623Z` |
| `POST https://registry.npmjs.org/-/npm/v1/security/advisories/bulk` with `{"xlsx":["0.18.5"]}` | answer of 2026-11-03 | both advisories for 0.18.5: GHSA-4r6h-8v6p-xvw6, `vulnerable_versions` `<0.19.3`, CVSS 7.8; GHSA-5pgg-2g8v-p4x9, `<0.20.2`, CVSS 7.5; severity "high" |
| GitHub Advisory Database, `advisories/github-reviewed/2023/04/GHSA-4r6h-8v6p-xvw6/GHSA-4r6h-8v6p-xvw6.json` | github/advisory-database commit `887dcb68f13a7d9be8bf3d604c2e32fc3d6960e8` | alias CVE-2023-30533, details, `last_known_affected_version_range` `< 0.19.3`, CWE-1321 |
| GitHub Advisory Database, `advisories/github-reviewed/2024/04/GHSA-5pgg-2g8v-p4x9/GHSA-5pgg-2g8v-p4x9.json` | same commit | alias CVE-2024-22363, details, `< 0.20.2`, CWE-1333 |
| CVE records `cves/2023/30xxx/CVE-2023-30533.json` and `cves/2024/22xxx/CVE-2024-22363.json` | CVEProject/cvelistV5 commit `cc01a172ea2819376e943af6fb3d73336eacea6b` | descriptions ("SheetJS Community Edition before 0.19.3 allows Prototype Pollution via a crafted file"; for CVE-2024-22363, SheetJS Community Edition before 0.20.2 and ReDoS), publication dates, references |
| `README.md` of github.com/SheetJS/sheetjs, branch `github` | commit `3f44ddd99c704d7ca35ea76fef7b8af2d2d435cb` | the source repository moved to https://git.sheetjs.com/sheetjs/sheetjs; "Scripts and NodeJS modules are available at https://cdn.sheetjs.com" |

The advisory files and CVE records were read through raw.githubusercontent.com at the commits given. The GitHub advisory web pages, the SheetJS sites (sheetjs.com, docs.sheetjs.com, cdn.sheetjs.com, git.sheetjs.com) and other advisory sites could not be reached from the environment where this page was written, so they are linked but not quoted.

## Route 1: the fixed SheetJS Community Edition

The fixed versions of the Community Edition (0.19.3 and later) are distributed by SheetJS from its own site, not from npm: the advisories name https://cdn.sheetjs.com/ for both fixes. The SheetJS README points to its documentation at https://docs.sheetjs.com and says "Scripts and NodeJS modules are available at https://cdn.sheetjs.com". Install the fixed version as the SheetJS documentation describes; this page does not copy an install command, because that documentation could not be read when the page was written. This route keeps the SheetJS API; check the SheetJS changelog (https://git.sheetjs.com/sheetjs/sheetjs/src/branch/master/CHANGELOG.md, linked from the CVE-2023-30533 record) for changes between 0.18.5 and the version you install.

## Route 2: move to xlvane

xlvane is a different library with a different API; it is not a drop-in replacement for `xlsx`. A package with a SheetJS-compatible API, `@xlvane/xlsx-compat`, is planned as a later phase ([ROADMAP.md](https://github.com/DevXord/xlvane/blob/main/ROADMAP.md)); it does not exist yet.

Facts about the package (npm registry, 2026-11-03; [docs/facts/comparison.md](https://github.com/DevXord/xlvane/blob/main/docs/facts/comparison.md)):

- `xlvane` is installed from npm (`npm install xlvane`), under the MIT license, with one runtime dependency, `fflate`.
- Versions from 1.0.1 are published from GitHub Actions with npm provenance ([SECURITY.md](https://github.com/DevXord/xlvane/blob/main/SECURITY.md#supported-versions)).
- It reads XLSX, XLSM, XLTX and XLTM, writes XLSX, and converts to and from CSV and JSON. It does not read legacy `.xls`, XLSB, ODS, Numbers or the other formats SheetJS supports, does not open encrypted workbooks, and does not calculate formulas.
- How it treats input files from untrusted sources (ZIP and XML limits, prototype pollution, amplification by small files, CSV formula injection) is described in [SECURITY.md](https://github.com/DevXord/xlvane/blob/main/SECURITY.md#what-the-library-defends-against).

### Equivalent calls

The SheetJS names are those of the `xlsx` 0.18.5 README ("Interface"). xlvane's functions take the file's bytes (`Uint8Array` or `ArrayBuffer`) or a workbook already read; `xlvane/node` adds file helpers for Node.js.

| SheetJS (`xlsx` 0.18.5) | xlvane |
|---|---|
| `XLSX.read(data)` | `readWorkbook(bytes)` |
| `XLSX.readFile(path)` | `await readWorkbookFile(path)` from `xlvane/node` |
| `workbook.SheetNames` | `workbook.sheets.map((sheet) => sheet.name)` |
| `XLSX.utils.sheet_to_json(worksheet)` | `readTable(bytes, { sheet }).rows`, or `toJSON(bytes, { sheet }).rows` for JSON-ready values |
| `XLSX.utils.sheet_to_csv(worksheet)` | `toCSV(bytes, { sheet })` |
| `XLSX.utils.json_to_sheet(rows)`, `book_new`, `book_append_sheet`, `XLSX.write(...)` | `writeTable(rows, { sheet })` or `fromJSON(rows)`, which return the `.xlsx` file as a `Uint8Array` |
| `XLSX.writeFile(workbook, path)` | `writeFile(path, writeTable(rows))` with `node:fs/promises`, or `writeWorkbookFile(workbook, path)` from `xlvane/node` for a workbook built with `createWorkbook` |

```ts
import { readTable, toCSV, toJSON, writeTable } from 'xlvane';
import { readWorkbookFile } from 'xlvane/node';
import { writeFile } from 'node:fs/promises';

const workbook = await readWorkbookFile('orders.xlsx');
console.log(workbook.sheets.map((sheet) => sheet.name)); // like workbook.SheetNames

// Rows as objects keyed by the header row (detected like inspect(); `header: 1` takes row 1).
const { rows } = readTable(workbook, { sheet: 'Orders' });
// The same rows with dates as ISO text, ready for JSON.stringify.
const json = toJSON(workbook, { sheet: 'Orders' }).rows;
// CSV; text that starts like a formula gets a leading ' (formulaInjection: 'off' turns this off).
const csv = toCSV(workbook, { sheet: 'Orders' });

// A new .xlsx file from objects: bold header, frozen first row, filter, fitted widths.
const bytes = writeTable(rows, { sheet: 'Orders' });
await writeFile('orders-copy.xlsx', bytes);
console.log(json.length, csv.length);
```

Differences to expect:

- `readTable` finds the header row the way `inspect()` does (`header: 'auto'`), or takes `header: 3` (a 1-based row) or `header: false` (keys are the column letters). It returns `Date` objects for date cells; `toJSON` and `toCSV` give ISO 8601 text.
- Unknown sheet names and damaged files throw an `XlvError` with a stable `XLV_E_*` code, a location and a hint ([docs/errors.md](https://github.com/DevXord/xlvane/blob/main/docs/errors.md)).
- Read limits (cells, decompressed bytes, XML size) apply by default and can be raised per call: `readWorkbook(bytes, { limits: { maxCells: 50_000_000 } })`. The defaults are in the [package README](https://github.com/DevXord/xlvane/blob/main/packages/xlvane/README.md#untrusted-input-and-limits).

The full API is described in the [package README](https://github.com/DevXord/xlvane/blob/main/packages/xlvane/README.md) (also in `node_modules/xlvane/README.md` after installing).
