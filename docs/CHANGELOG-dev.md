# Development changelog

One entry per plan phase (docs/PLAN.md), newest first.

## Phase 4: streaming (2026-10-30)

- **Public API** (`xlvane`):
  - `openWorkbookStream` reads `Uint8Array`, `ArrayBuffer` or `Blob` input. It returns sheets and metadata, rows through `rows(sheet)` (async iterator) or `rowStream(sheet)`, and the warnings.
  - `readTableStream` yields validated row objects.
  - `writeWorkbookStream` returns a `ReadableStream<Uint8Array>` and pulls rows from iterables.
  - `writeTableStream` writes the report preset.
  - `STREAM_LIMIT_DEFAULTS`.
- **Public API** (`xlvane/node`): `openWorkbookStreamFile` (via `fs.openAsBlob`) and `writeWorkbookStreamFile` (exclusive create; removes a partial file if writing fails).
- **Internals:**
  - an incremental XML parser (text and bytes, including UTF-16) with the same events, spans and errors as whole-text parsing (D-402);
  - streaming ZIP reading and writing, with data descriptors, limits checked on produced bytes, CRC at the end, and Blob slices (D-405);
  - the batch and streaming readers share `analyzeWorkbook` and `createWorksheetParser`;
  - the writer's value conversion is shared (`convertCellInput`).
- **Memory** (acceptance criterion): retained heap while reading and writing 400 000 rows × 5 columns is about the same as for 50 000 rows, under 1 MB in both cases (D-406, `tests/stream/memory.test.ts`).
- **Tests:**
  - chunked XML parsing compared with whole-text parsing, including security inputs, in unit and property tests;
  - streaming ZIP: unit, security (bombs, limits) and property tests;
  - streaming read and write compared with batch read and write (property);
  - Node file helpers;
  - LibreOffice opens streamed output with inline strings.
- **Benchmark:** `bench/` (separate package: exceljs 4.4.0, xlsx 0.18.5), `npm run bench`, results in bench/perf/RESULTS.md and the README (D-407).
- **Package:** size budget raised to 125 kB gzip at the maintainer's decision (101.4 kB measured, D-408).
- **Not done in this phase:**
  - ZIP64 writing (streamed archives over 4 GiB are refused);
  - streaming edits;
  - reading rows of a sheet in sorted order when the file stores them unsorted.

## Phase 3: task-oriented API (2026-10-30)

- **Public API** (`xlvane`):
  - `readTable` and `readTableAsync`, with Standard Schema validation (Zod, Valibot, ArkType) and issues tied to cell addresses;
  - `writeTable`, with the `report` and `plain` presets;
  - `editCells` and `appendRows`, which edit files surgically;
  - `diff`;
  - `toCSV`, which escapes text that spreadsheets would run as formulas;
  - `fromCSV`, `toJSON`, `fromJSON`;
  - the Standard Schema types and the option and result types.
- **Surgical editing** (D-309 to D-312): only the edited sheets' `sheetData` and `dimension` change, plus new shared strings and style records, `calcPr` and the removal of the calculation chain. Every other part keeps its bytes. The XML parser now reports tag positions (`XmlSpan`).
- **Shared formulas** are translated to dependent cells when reading (D-313, closes D-107).
- **Errors:** 6 new error codes and 1 new warning (`XLV_W_FORMULAS_NOT_RECALCULATED`), documented in docs/errors.md.
- **Tests:**
  - unit tests for every function;
  - type-inference tests for Zod, Valibot and ArkType (`tests/types`);
  - security tests: prototype keys, CSV and formula injection, CSV limits, linear parsing;
  - property tests for CSV, editing and the formula shifter;
  - TSDoc examples, which all compile, and every public symbol has an `@example` (D-314);
  - LibreOffice also opens `writeTable`, `fromCSV` and `editCells` output. One case is a file that LibreOffice itself wrote, then edited by xlvane.
- **Final run** on Node.js 22.22.0 with LibreOffice 24.2.7.2: see the summary in the phase 3 commit message.
- **Package:** the size budget was raised to 100 kB gzip at the maintainer's decision (measured 84.5 kB). zod, valibot and arktype were added as development dependencies (D-301).
- **Facts:** docs/facts/csv.md (RFC 4180 facts PROVISIONAL; RFC not reachable) and docs/facts/formulas.md (the §18.17 grammar is mostly PROVISIONAL).
- **Not done in this phase:**
  - verification of the provisional facts (ECMA-376 §18.17 and §18.2.2, RFC 4180);
  - edits that change merges, column widths or other sheet properties;
  - recalculation;
  - streaming (phase 4).

## Phase 2: writing and styles (2026-10-30)

- **Public API** (`xlvane`): `createWorkbook({ dateSystem })`, `NewWorkbook` (`addSheet`, `getSheet`, `sheets`), `NewSheet` (`set`, `style`, `column`, `row`, `merge`, `freeze`, `autoFilter`), `writeWorkbook()`, `ERROR_VALUES` and the style types (`CellStyle`, `FontStyle`, `FillStyle`, `BorderStyle`, `AlignmentStyle`, `Color`, …). `xlvane/node`: `writeWorkbookFile(workbook, path, { overwrite })`, which never replaces a file unless asked (exclusive create).
- **Writer**: several sheets with visibility; numbers, strings (shared string table), booleans, errors, dates (serials with a date format, 1900 and 1904 systems), formula text with cached results; fonts, fills, borders, alignment and number formats, deduplicated; column widths, hidden columns, row heights, hidden rows, merges, frozen panes, autofilter. Deterministic output (D-202). XML escaping and the ST_Xstring `_xHHHH_` encoding are the exact inverses of the reader's rules.
- **Reader additions** (D-211): `cell.style`, `sheet.columns`, `sheet.rowInfo()`, `sheet.freeze`, `sheet.autoFilter`; formula text is decoded as ST_Xstring.
- **Validation**: 7 new error codes (`XLV_E_INVALID_VALUE`, `XLV_E_INVALID_STYLE`, `XLV_E_INVALID_SHEET_NAME`, `XLV_E_DUPLICATE_SHEET_NAME`, `XLV_E_OVERLAPPING_MERGE`, `XLV_E_FILE_EXISTS`, `XLV_E_FILE_WRITE`) in docs/errors.md; hints list allowed values and suggest close matches.
- **Tests**: writer unit tests, round trips (write, read, compare) for values, dates, formulas, styles and layout, property-based round trips, writer security tests (prototype keys in styles, getters, sheet names such as `__proto__`, markup in formula text), ZIP and XML writer tests, and a LibreOffice test (D-209) that converts written files to CSV and back to XLSX with `soffice --headless` and checks values, styles and layout. Final run on Node.js 22.22.0 with LibreOffice 24.2.7.2: `npm test` 521 passed (4 todo: corpus files not supplied yet), `npm run fuzz` 31 properties passed, `npm run typecheck` and `npm run check:pkg` passed.
- **Package**: size budget 70 kB gzip, measured 58.4 kB, enforced by `check:pkg` (D-208); CI workflow for Node.js 22/24/26 with LibreOffice, actions pinned to SHAs (D-210).
- **Facts**: docs/facts/sml-write.md. Many writer facts are PROVISIONAL because the ECMA-376 documents could not be downloaded in the cloud session (D-203); they must be checked when the maintainer supplies them.
- **Not done in this phase**: verification of the provisional facts against ECMA-376 and [MS-OI29500]; converting a read workbook into a writable one (phase 3, `editCells`); rich text, hyperlinks, comments, conditional formatting, column default styles; ZIP64 writing (phase 4); TSDoc example compilation test (phase 3).

## Phase 1: safe reading and `inspect()` (2026-10-30)

- **Public API** (`xlvane`): `readWorkbook(bytes, { limits })`, `inspect(input, options)`, `XlvError`/`isXlvError`, `ERROR_CODES`, `WARNING_CODES`, `DEFAULT_READ_LIMITS` and the model types. `xlvane/node`: `readWorkbookFile(path)`, `inspectFile(path)`.
- **ZIP** (D-102): own central-directory reader with ZIP64, fflate `Inflate` for Deflate. Limits on actually decompressed bytes (per entry, total, ratio), entry count and input size. Rejects overlapping entries, duplicate names (ignoring case), local/central mismatches, leading or trailing data, encryption, unsupported methods and multi-disk archives; verifies sizes and CRC-32. CFB files (legacy `.xls`, encrypted `.xlsx`) get their own error.
- **XML** (D-101): own iterative parser; DTDs always rejected; only the predefined entities and character references; Char production, line-end and attribute normalisation; namespaces; depth, attribute, name and text limits; UTF-8/UTF-16 only.
- **OPC**: content types (Default/Override), relationships with RFC 3986 target resolution, part-name grammar and ASCII case-insensitive equivalence (ECMA-376 Part 2).
- **SpreadsheetML**: workbook (sheets, states, `date1904`, Strict and Transitional), shared strings (rich text, phonetic runs excluded, `_xHHHH_` escapes), styles (number format per cell xf, as Office does), worksheets (all cell types, formulas with cached values, shared/array formulas, merged cells, dimension, positional cells without `r`), built-in and custom number formats with date detection, both date systems with the Strict/Transitional difference (D-106).
- **inspect()**: JSON report (docs/inspect-json.md) with sheets, visibility, used ranges, header detection, column types and samples, merged ranges, formulas and warnings.
- **Errors and warnings**: 16 error codes and 11 warning codes, documented in docs/errors.md (test-checked).
- **Tests**: unit, security (zip bombs with lying headers, overlapping entries, XXE, billion laughs, parameter entities, 100 000-level nesting, prototype pollution, ReDoS-style and amplification inputs), property-based fuzzing with fast-check, and a corpus harness for real-application files. Final run on Node.js 26.13.0: `npm test` 452 passed (4 todo: corpus files not supplied yet), `npm run fuzz` 30 properties × 20 000 runs passed, `npm run typecheck` and `npm run check:pkg` passed.
- **Security review**: an independent review pass found amplification inputs (quadratic format classification, long namespace names, merged-range sweep, warning keys) and two parsing bugs; all fixed with tests (D-115). Default `maxCells` lowered to 5 000 000 after measuring about 183 bytes per cell.
- **Process**: parts of the phase were implemented by parallel agents working on disjoint files against shared interfaces (`model.ts`, `sml/namespaces.ts`), each test-first.
- **Facts**: docs/facts/ (ECMA-376 Parts 1, 2, 4; APPNOTE; XML 1.0; Namespaces; Microsoft open specifications), decisions D-101 … D-115.
- **Not done in this phase**: real files from Excel, LibreOffice and Google Sheets (list in tests/corpus/REQUESTED.md); shared-formula translation (phase 3); TSDoc example compilation test (phase 3); streaming (phase 4).

## Phase 0: monorepo skeleton (2026-10-30)

- npm workspaces: `packages/xlvane` (`xlvane`), `packages/exceljs` (`@xlvane/exceljs`), `packages/mcp` (`@xlvane/mcp`), each ESM-only, MIT, `"private": true` until the release phase, with empty entry modules.
- Tooling (D-002 … D-005): tsdown 0.23.4 (build, declarations via oxc with `isolatedDeclarations`), Vitest 5.0.5 with vite 8.3.5, fast-check 4.10.2, publint 0.3.25, @arethetypeswrong/cli 0.18.5, TypeScript 7.0.2. No @types/node; a small own subset in `types/` (D-006).
- `npm run check:pkg` builds every package and runs `publint --strict`, `attw --pack --profile esm-only` and a `require()` load of each built file, which fails on top-level await (checked: Node.js 26 throws `ERR_REQUIRE_ASYNC_MODULE` for a module with top-level await).
- Policy tests (`tests/policy/packages.test.ts`): allowed runtime dependencies per package, ESM/MIT/LICENSE/`files`, no `require` export condition, no unapproved root development dependencies, no Node.js built-ins in core sources.
- Documents: CLAUDE.md (English translation of plan section 7), docs/PLAN.md (Polish original), docs/decisions.md, SECURITY.md, README files, tests/corpus/LICENSES.md.
- Not done: depgate policy file (depgate is unpublished on npm, D-008).
- Environment note: the system Node.js on the development machine is 25.8.1 (end-of-life, outside the tool engine ranges). This session used a portable Node.js 26.13.0 LTS (SHA-256 verified against nodejs.org SHASUMS256.txt) in a temporary directory.
