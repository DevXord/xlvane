# Benchmarks

`bench/` is not an npm workspace. It has its own `package.json` and `package-lock.json` with the libraries xlvane is compared against, so the root dependencies, CI and `npm test` never install them (docs/decisions.md):

| Package | Version | License | Note |
|---|---|---|---|
| exceljs | 4.4.0 | MIT | |
| xlsx (SheetJS) | 0.18.5 | Apache-2.0 | npm release; has CVE-2023-30533 (prototype pollution when reading crafted files). Newer releases are only on the SheetJS CDN. The benchmark reads only files it generates itself. |

## Performance (`bench/perf/`)

```sh
npm run bench          # build xlvane, npm ci --prefix bench, run all variants
BENCH_RUNS=3 BENCH_ROWS=20000 node bench/perf/run.mjs   # quicker, after a build and npm ci --prefix bench
```

- `data.mjs`: deterministic data, 100 000 rows × 20 columns by default (unique id, integers, floats, short texts, dates, booleans), computed from the cell position and a fixed seed.
- `variant.mjs`: runs one variant in a fresh process and prints JSON (time, peak RSS, peak heap).
- `run.mjs`: runs every variant `BENCH_RUNS` times (default 5), validates the outputs, writes `results/<date>-<host>.json` and `RESULTS.md`.

Peak RSS comes from `process.resourceUsage().maxRSS`. The Node.js documentation describes it as "the maximum resident set size used in kibibytes (1024 bytes)" (`doc/api/process.md`, `process.resourceUsage()`, branch v22.x, read 2026-10-30), so the runner multiplies it by 1024.

## Library APIs used, and where they were checked

Paths are relative to `bench/node_modules`.

**exceljs 4.4.0**

| API | Source |
|---|---|
| `new ExcelJS.Workbook()`, `addWorksheet(name)` | `exceljs/README.md` L292 |
| `worksheet.addRow(array)` | `exceljs/README.md` L829–836 |
| `workbook.xlsx.writeBuffer()` | `exceljs/README.md` L2189 |
| `workbook.xlsx.load(buffer)` | `exceljs/README.md` L2164 |
| `getWorksheet(name)` | `exceljs/README.md` L343 |
| `eachRow`, `eachCell` | `exceljs/README.md` L794, L702 |
| `new ExcelJS.stream.xlsx.WorkbookWriter({ filename, useStyles, useSharedStrings })` | `exceljs/README.md` L2375–2392 |
| streaming `row.commit()`, `worksheet.commit()`, `await workbook.commit()` | `exceljs/README.md` L2410–2439 |
| `new ExcelJS.stream.xlsx.WorkbookReader(file, options)`, and `for await` over worksheets and rows | `exceljs/README.md` L2460–2466 and L2480–2482 (options `sharedStrings`, `styles`, `worksheets`, `hyperlinks`) |

`exceljs/lib/stream/xlsx/worksheet-reader.js` L105–113 shows that the worksheet iterator yields one row at a time. README L2468 says rows come in arrays; that note applies to the `parse()` events.

**xlsx (SheetJS) 0.18.5**

| API | Source |
|---|---|
| `XLSX.utils.aoa_to_sheet(rows)` | `xlsx/types/index.d.ts` L769 |
| `book_new()`, `book_append_sheet(wb, ws, name)` | `xlsx/types/index.d.ts` L842, L845 |
| `XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })` | `xlsx/types/index.d.ts` L34, `type` L235, `bookType` L247 |
| `XLSX.read(buffer, { type: 'buffer', cellDates: true })` | `xlsx/types/index.d.ts` L25 and L110–115 ("When reading a file, store dates as type d") |
| `ws[address].v` | cell objects of the worksheet |
