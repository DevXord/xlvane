# Development changelog

One entry per plan phase (docs/PLAN.md), newest first.

## Phase B: row and column operations (2026-10-30)

- **Reference engine** (D-B01, D-B02): `adjustFormula`, `adjustRange` and `shiftFormula` are public; rules R-01 … R-17 observed in LibreOffice 24.2.7.2 (Microsoft's documentation was not reachable from the session).
- **@xlvane/exceljs** (D-B03): `insertRow(s)`, `spliceRows`, `duplicateRow`, `spliceColumns`, `row.splice`, `addRow(values, 'i')`, moving formulas, merges, notes, links, autofilter and names like a spreadsheet application; `XLV_E_INVALID_ARGUMENT` for arguments exceljs corrupts data with.
- **Independent review**: 8 confirmed bugs fixed with regression tests (D-B01/D-B03 addenda).
- **Core** (D-B04): surgical `insertRows` / `deleteRows` on existing files; **CLI** (D-B05): `xlvane edit --insert-rows / --delete-rows`.
- **Research for phases C and D** done; third local review prompt prepared.

## Phase A2: hyperlinks, notes, views, app.xml (2026-10-30)

- **Sources:** second local review report (ECMA-376 Part 1 §18.3.1.47/48, §18.7, §12.3.3, §18.3.1.13/66/82/87/93, §22.2, §15.2.12; Part 4 §8.2, §19.4; [MS-OI29500]; [MS-XLSX]).
- **Core:** hyperlinks per the specification (D-A08), tab colour and sheet views (D-A09), `docProps/app.xml` (D-A10), `col@style` for new cells in `editCells` (D-A04 addendum), notes with comments and VML parts, per-row links and notes in the streaming writer (D-A12).
- **@xlvane/exceljs:** hyperlink values, tab colour, views, company and manager (D-A13); `cell.note` and streamed links (D-A14). No phase-A API throws `XLV_E_NOT_SUPPORTED` any more.
- **inspect, CLI, docs:** hyperlink and note counts with samples; README, SKILL.md, CHANGELOG.
- **Size budget** raised to 175 kB (D-A11); library at 127.7 kB.

## Phase A1: rich text, properties, defined names, column styles, format ids (2026-10-30)

- **Core** (D-A01 … D-A05): `Cell.richText` and `{ richText }` input; `docProps/core.xml` via `properties`; `Workbook.definedNames` and `defineName` (`XLV_E_INVALID_DEFINED_NAME`); `column(c, { style })` and `ColumnInfo.style`; `numFmt` as a built-in id.
- **@xlvane/exceljs** (D-A06): rich text, workbook properties, the `DefinedNames` API, column styles and date format 14 as in exceljs 4.4.0; ported exceljs specs for these features; cross-check sample extended.
- **inspect, CLI, docs** (D-A07): new bounded fields in `inspect()` and `xlvane inspect`; README, SKILL.md, CHANGELOG.
- **Roadmap**: docs/PLAN.md phases A–E and the "AI Friendly" stage; second local review report requested and received for phase A2.

## After phase 10: review with the supplied standards (2026-10-30)

- **Sources:** ECMA-376 Part 1, [MS-OI29500] rev. 25.0 and [MS-XLSX] rev. 30.0, read through a local review report (D-1104); RFC 4180 for CSV. Facts in docs/facts updated with section and page.
- **Writer limits** per [MS-OI29500] (D-1101): font size 1–409.55, font name ≤ 31 characters, number format code ≤ 254 characters, indent ≤ 255, styles-part record counts (`XLV_E_LIMIT_EXCEEDED`); existing styles in edited files are kept. Column width TSDoc corrected.
- **Shared formulas** (D-1102): `@`, spill `#` and references beyond the grid are not shifted; numbers per the union of the grammars.
- **Windows and Node.js 25** (D-1103): tests spawn CLIs without shims; `writeWorkbookStreamFile` closes its FileHandle explicitly.
- **Tests:** tests/unit/office-limits.test.ts, new cases in shared-formula and node tests.

## Phase 10: publication prepared (2026-10-30)

- **Packages:** `xlvane` and `@xlvane/exceljs` set to version 1.0.0, public, with provenance (D-1001); the "not published" texts are removed.
- **Release workflow** `release.yml` (D-1002 … D-1004):
  - tag `v*`;
  - full verification;
  - `npm stage publish` through trusted publishing (OIDC), with no secrets.
- **`scripts/check-release-version.mjs`** stops a release whose tag and versions differ.
- **docs/RELEASING.md** lists the maintainer's steps: account with 2FA, organization, trusted publisher with stage-only permission, approval with 2FA, disallowing tokens. **CHANGELOG.md** describes 1.0.0 for users.
- **`tests/policy/release.test.ts`** checks the package settings, the workflows (SHA pins, permissions, no secrets) and that no "not published" texts remain.
- **Not done here (the maintainer's steps):** the npm account and organization, a public repository, the tag, and approving the staged packages.

## Phase 9: corpus and environments (2026-10-30)

- **Corpus:**
  - 340 Apache POI test files (Apache-2.0, every file in LICENSES.md, LICENSE and NOTICE included), with the recorded outcome per file and deep checks of every readable file (D-901);
  - 3 files saved by LibreOffice 24.2 with value checks (D-902);
  - LibreOffice opens POI files after xlvane edited them.
- **Results:** 312 of 340 POI files are read (24 with warnings); 28 are refused with error codes (damaged, bombs, encrypted, beta namespace); no crash; streaming equals batch on all of them.
- **Tolerance** (D-903): rows with invalid numbers and cells with invalid references are skipped with the new warning `XLV_W_INVALID_CELL_REF`, as LibreOffice does; editing such sheets is refused.
- **Security** (D-904): new limit `maxSharedStringChars` against shared-string amplification, found through the POI file `poc-shared-strings.xlsx` (2 MB file, about 12 G characters when copied).
- **Environments** (D-905 … D-907): one smoke bundle runs in Node.js, Bun, Deno, Chromium and workerd (`npm run test:runtimes`); new CI job `runtimes`; dev dependencies playwright, workerd and deno.
- **Not done:**
  - files from Google Sheets and Numbers, and Excel files with the shared content, which are still requested from the maintainer;
  - the first CI run on GitHub, which happens on a push to `main` or a pull request.

## Phase 8: agent benchmark skipped, MCP server removed (2026-10-30)

- **Benchmark:** not built and not run, at the maintainer's decision: no paid API usage in the free version (D-801).
- **MCP server removed** (D-802):
  - `@xlvane/mcp`, the `xlvane mcp` command and its JSON schema;
  - the MCP tests;
  - the error codes `XLV_E_NOT_INSTALLED` and `XLV_E_PATH_NOT_ALLOWED`;
  - the dev dependency `@modelcontextprotocol/client`.
  - The CLI now has 8 commands; the welcome screen says "For AI agents and scripts: add --json to any command".
- **Kept from phase 7:**
  - the Agent Skill;
  - docs and `llms-full.txt` in the xlvane package;
  - the policy test of agent-facing texts;
  - tests building packages into their own directories.

## Phase 7: agent integration (2026-10-30)

- **`@xlvane/mcp`**, on the official MCP SDK v2 (`@modelcontextprotocol/server` 2.4.0, spec 2026-07-28) (D-701 … D-705):
  - tools: `inspect_workbook`, `read_table` (paging), `diff`, `validate`, and with `--allow-write` also `write_table`, `edit_cells`, `convert`;
  - JSON Schema inputs, annotations, structured results, errors with code and hint;
  - files only inside the `--root` directories, after resolving symbolic links (`XLV_E_PATH_NOT_ALLOWED`, new code);
  - no overwriting without `"overwrite": true`; size limit;
  - entry points `createServer`, `runCli` and the `xlvane-mcp` command; `xlvane mcp` starts the same server.
- **Agent Skill:** `skills/xlvane/SKILL.md` (Agent Skills format), installable with `npx skills add <path> --skill xlvane` (D-706).
- **Package `xlvane`:** ships `docs/`, `skills/` and `llms-full.txt` (D-707).
- **Tests:**
  - MCP server on an in-memory client: tool list and order, every tool, errors, schema validation;
  - security: absolute paths, `..`, symlinked directories and files (including a dangling link), reading and writing outside the root, overwrite protection, size limits, write tools hidden;
  - the built server over stdio, from `xlvane-mcp` and from `xlvane mcp`;
  - `runCli` arguments;
  - the skill's frontmatter, and its installation with the `skills` CLI;
  - TSDoc and SKILL.md code blocks compile;
  - a policy test of all texts read by agents (no persuasion, no instruction overrides, no hidden comments or invisible characters).
- **Test infrastructure:** tests build packages into their own directories instead of `packages/*/dist` (D-708).
- **Dev dependencies** (approved): `@modelcontextprotocol/client` 2.4.0, `skills` 1.7.4.
- **Docs:** packages/mcp/README.md, the MCP and skill sections of README.md, decisions D-701 … D-708, facts F-MCP-01, F-MCP-02, F-SKILL-01.
- **Not done in this phase:**
  - HTTP transport (stdio only);
  - `outputSchema` for the tools;
  - MCP resources or prompts.

## Phase 6: command line (2026-10-30)

- **`xlvane` command** (`packages/xlvane`, `bin` → `dist/cli.js`), on `util.parseArgs` and `util.styleText` without new dependencies (D-601 … D-603):
  - 9 commands: `inspect`, `read`, `convert`, `create`, `edit`, `diff`, `validate`, `migrate`, `mcp`;
  - a welcome screen, and `help <command>` / `--help` with 2–4 examples and related commands;
  - "Did you mean" for commands, options, sheet names and file names;
  - errors in three parts (Error / Where / Fix) with a ready-to-run `Try:` command;
  - overwrite protection with `--force`, atomic replacement, `--dry-run`, `--quiet`, `--verbose`;
  - colour only on terminals and never as the only carrier of information;
  - exit codes 0/1/2/3 (D-604).
- **`--json`:** one envelope per command, with JSON Schema files in docs/cli-json/ and docs/cli-json.md (D-605).
- **New error codes:** `XLV_E_USAGE`, `XLV_E_NOT_INSTALLED`, `XLV_E_INTERNAL` (docs/errors.md).
- **Library changes:**
  - Node.js file helpers moved to `src/util/node-files.ts` and shared with the CLI;
  - `inferValue` from csv.ts is reused by `edit --set` (D-606).
- **Tests** (`tests/cli/`):
  - snapshots of the welcome screen, every help screen and 27 error messages;
  - every example of every help screen and of the welcome screen runs against fixture files;
  - `--json` successes and errors of every command are validated against the schemas;
  - behaviour: colour, `--force`, `--dry-run`, CSV injection, `edit`, `validate`, `migrate`, `mcp`;
  - the built `dist/cli.js` is run with node;
  - a policy test checks that the core entry does not reach Node.js-only modules.
- **Docs:** a 5-minute CLI start in README.md and packages/xlvane/README.md; decisions D-601 … D-608; facts F-NODE-09 and F-NODE-10.
- **Package:** separate CLI size budget of 40 kB gzip (20.3 kB measured); library 104.1 kB of 125 kB (D-608).
- **Not done in this phase:**
  - progress for large files is a single status line (the reader reports no progress);
  - table widths count UTF-16 code units, so wide characters may misalign;
  - the MCP server itself (phase 7).

## Phase 5: exceljs compatibility (2026-10-30)

- **`@xlvane/exceljs`:** the commonly used exceljs 4.4.0 API on xlvane, for migration with one line (`"exceljs": "npm:@xlvane/exceljs@^1"`):
  - `Workbook` with `xlsx.load/readFile/read/writeBuffer/writeFile/write`, worksheets, rows, columns, cells, values, styles, merges, frozen panes and autofilter (D-503 … D-505);
  - `stream.xlsx.WorkbookWriter` and `WorkbookReader` (D-507);
  - `ValueType`, `FormulaType`, `ErrorValue` with the exceljs numbers;
  - default export, named exports and `require()` (Node.js 22.12+, `'module.exports'` export name, D-501);
  - everything outside the subset throws `XLV_E_NOT_SUPPORTED` (new error code, docs/errors.md, D-502).
- **`xlvane`:** `writeWorkbookStream` accepts `sheets` as an (async) iterable; the package parts are then written after the worksheets.
- **Tests:**
  - exceljs v4.4.0 specs ported to `tests/exceljs/` (96 passing in the ported and own files, 106 skipped with reasons: upstream binary fixtures or unsupported features; tests/exceljs/README.md);
  - alias test: a sample project written for exceljs runs unchanged from ESM (default and named import) and CommonJS against the built package; its output matches exceljs 4.4.0 (cross-check) and LibreOffice opens the files;
  - streaming writer layout rules; lazy sheets in `writeWorkbookStream`;
  - TSDoc examples test covers `packages/exceljs`;
  - in-process tests resolve `xlvane` to its sources (vitest alias), so the alias test can rebuild dist while they run.
- **Docs:** docs/exceljs-compat.md (support table with sources), packages/exceljs/README.md, decisions D-501 … D-507, F-NODE-08, F-EXCELJS-01.
- **Not done in this phase:** rich text and hyperlink writing, images, tables, conditional formatting, data validation, notes, defined names, protection, page setup, `insertRow`/`spliceRows`/`duplicateRow`, workbook CSV, document properties in the file.

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
