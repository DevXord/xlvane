# Benchmark results

100,000 rows × 20 columns (no header row), 5 runs per variant, median reported. Raw data: `bench/perf/results/2026-10-30-vm.json`.

## Environment

| | |
|---|---|
| Date | 2026-10-30T10:56:49.818Z |
| CPU | Intel(R) Xeon(R) Processor @ 2.10GHz, 4 logical cores |
| Memory | 15.7 GiB |
| OS | Linux 6.18.44-fc-v185 (x64) |
| Node.js | v22.22.0 |
| xlvane | 0.0.0 (this repository, unminified build) |
| exceljs | 4.4.0 |
| xlsx (SheetJS) | 0.18.5 (npm) |

## Write

| Variant | Time, s (median) | Range, s | Peak RSS, MiB | Peak heap, MiB |
|---|---|---|---|---|
| xlvane-write | 9.37 | 8.10–11.78 | 1005 | 362 |
| xlvane-write-stream | 4.91 | 4.43–5.00 | 170 | 28 |
| exceljs-write | 18.72 | 18.29–19.76 | 2102 | 1559 |
| exceljs-write-stream | 6.31 | 6.03–6.43 | 194 | 34 |
| sheetjs-write | 10.96 | 8.54–11.60 | 1002 | 379 |

## Read

Every read variant reads the same file, written by `xlvane-write` (12.1 MiB), and visits every cell.

| Variant | Time, s (median) | Range, s | Peak RSS, MiB | Peak heap, MiB |
|---|---|---|---|---|
| xlvane-read | 4.52 | 4.29–4.57 | 690 | 485 |
| xlvane-read-stream | 5.32 | 4.86–5.55 | 309 | 66 |
| exceljs-read | 6.80 | 6.34–7.17 | 889 | 706 |
| exceljs-read-stream | 4.60 | 4.03–4.71 | 309 | 128 |
| sheetjs-read | 9.04 | 8.95–9.75 | 838 | 427 |

Baseline (Node.js with the three libraries imported, no work): peak RSS 135 MiB.

## Method

- Each run is a fresh `node` process (`bench/perf/variant.mjs`) with `TZ=UTC`; the time covers the library work only (building and writing the workbook, or reading it and visiting every cell), not process start-up or imports.
- Peak RSS is `process.resourceUsage().maxRSS` (kibibytes per the Node.js documentation) at the end of the run, so it includes Node.js itself and the imported libraries (see the baseline).
- Peak heap is `process.memoryUsage().heapUsed` sampled every 10 ms and once at the end of the timed work. Synchronous work blocks the sampling timer, so for synchronous variants this is the heap at the end of the work (with the result still alive), a lower bound of the true peak.
- Every write variant's file is read back with xlvane and checked (cell count and sample values of every column type); every read variant checks the cell count and the same sample values.
- The data is deterministic (`bench/perf/data.mjs`): a unique id column, then integers, floats, short texts, dates and booleans.
- Streaming variants: xlvane streams to/from a file through Web Streams (`Readable.fromWeb`, `fs.openAsBlob`); exceljs uses its streaming WorkbookWriter (without shared strings) and WorkbookReader.

## Caveats

- SheetJS is the npm release 0.18.5 (March 2022). That version has CVE-2023-30533 (prototype pollution when reading crafted files); newer releases are distributed only from the SheetJS CDN, which this environment could not reach. The benchmark reads only files it generated itself.
- Results depend on the machine and on the data mix; they are one run on one machine, not a general ranking.
