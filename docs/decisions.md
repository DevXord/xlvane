# Decisions and verified facts

This file records two things:

- **Decisions** (`D-NNN`): what we chose, why, and what it affects.
- **Facts** (`F-AREA-NN`): statements about file formats, tools and platforms that the code relies on, each with the source it was checked against and the date of the check. CLAUDE.md forbids writing such facts from memory.

Quotes are kept short (at most 15 words). Section numbers refer to the editions listed under "Sources".

## Sources

| Id | Document | Edition / version | Where |
|---|---|---|---|
| ECMA-376-1 | ECMA-376 Part 1, Fundamentals and Markup Language Reference | 5th edition, December 2016 | https://ecma-international.org/publications-and-standards/standards/ecma-376/ |
| ECMA-376-2 | ECMA-376 Part 2, Open Packaging Conventions | 5th edition, December 2021 | same page |
| ECMA-376-4 | ECMA-376 Part 4, Transitional Migration Features | 5th edition, December 2016 | same page |
| APPNOTE | PKWARE .ZIP File Format Specification (APPNOTE.TXT) | 6.3.10, revised Nov 01, 2022 | https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT |
| XML | W3C Extensible Markup Language (XML) 1.0 | Fifth Edition, W3C Recommendation 26 November 2008 | https://www.w3.org/TR/xml/ |
| XML-NS | W3C Namespaces in XML 1.0 | Third Edition, W3C Recommendation 8 December 2009 | https://www.w3.org/TR/xml-names/ |

---

## Decisions

### D-001 Monorepo with npm workspaces (2026-10-30)
Three workspaces as in docs/PLAN.md section 3: `packages/xlvane` (`xlvane`), `packages/exceljs` (`@xlvane/exceljs`), `packages/mcp` (`@xlvane/mcp`). All development tooling lives in the root `package.json`; packages have no `devDependencies` (checked by `tests/policy/packages.test.ts`).

### D-002 Build with tsdown, not tsup (2026-10-30)
The tsup README states that tsup is no longer actively maintained and recommends tsdown (F-TOOL-01). tsdown 0.23.4 is used with `format: 'esm'`, `platform: 'neutral'`, `target: 'es2022'` and `dts: true`. Note: tsup 9.0.0 was published to npm on 2026-10-14, but the README still carries the "not actively maintained" notice, so the recommendation stands.

### D-003 Vitest 5 and fast-check for tests (2026-10-30)
Vitest 5.0.5 runs all tests; fast-check 4.10.2 drives property-based tests. `npm test` runs the property tests with `XLV_FUZZ_RUNS=100`; `npm run fuzz` runs only `tests/fuzz` with `XLV_FUZZ_RUNS=20000`. Vitest 5 declares `vite` as a required peer dependency (F-TOOL-03), so `vite` is listed explicitly.

### D-004 Package checks (2026-10-30)
`npm run check:pkg` (`scripts/check-pkg.mjs`) builds all packages and then runs, for each package, `publint --strict` (warnings fail the check, F-TOOL-05) and `attw --pack <dir> --profile esm-only` (F-TOOL-06; CommonJS resolution is intentionally unsupported). It also loads every built `dist/*.js` with `require()`: Node.js rejects a module graph containing top-level await with `ERR_REQUIRE_ASYNC_MODULE` (F-NODE-01), which turns the "no top-level await" rule into a check.

### D-005 Approved development dependencies (2026-10-30)
Approved by the maintainer when the phase 0/1 plan was accepted: `tsdown`, `vitest`, `fast-check`, `publint`, `@arethetypeswrong/cli` (named in docs/PLAN.md), plus `typescript` (type checking; compiling TSDoc examples in phase 3) and `vite` (required peer of Vitest 5). The maintainer asked to add only dependencies that are really needed and to write the rest ourselves; therefore `@types/node` was not added (D-006). Exact versions are pinned. `tests/policy/packages.test.ts` fails if the root gains an unlisted development dependency.

### D-006 No @types/node (2026-10-30)
`types/node-modules.d.ts` declares the few members of `node:fs/promises`, `node:path`, `node:url` and `node:os` that the code uses, following the Node.js API documentation. `types/node-globals.d.ts` declares `process` for tests only; package sources cannot see it. `skipLibCheck` is on so that third-party declaration files that reference Node types do not fail the type check.

### D-007 Supported Node.js versions (2026-10-30)
Packages declare `engines.node: ">=22"`; the CI matrix in phase 9 covers Node.js 22, 24 and 26 (docs/PLAN.md). Development requires `^22.18.0 || ^24.11.0 || >=26.0.0`, the intersection of the tsdown and Vitest engine ranges (F-TOOL-02, F-TOOL-03). Node.js 25 is end-of-life (F-TOOL-07) and outside both ranges, so it is not supported for development.

### D-008 No depgate policy file (2026-10-30)
docs/PLAN.md phase 0 asks for a depgate policy file "if depgate is already available". The npm registry reports the `depgate` package as unpublished on 2026-10-29 (F-TOOL-08), so no policy file was added. Revisit when depgate is available.

### D-009 TypeScript configuration (2026-10-30)
TypeScript 7.0.2. `tsconfig.base.json`: `strict`, `noUncheckedIndexedAccess`, `exactOptionalPropertyTypes`, `verbatimModuleSyntax`, `moduleResolution: "bundler"`, `lib: ["ES2023", "DOM"]` (the DOM library provides `TextDecoder`, `URL` and Web Streams types for all target runtimes), `types: []`. Package configs add `isolatedDeclarations`, which lets tsdown generate declaration files with oxc-transform instead of the TypeScript compiler (F-TOOL-02). The configuration was validated by running `tsc` rather than by relying on release notes.

### D-010 Packages are `"private": true` until the release phase (2026-10-30)
Prevents an accidental `npm publish` before phase 10. Remove the flag as part of the release work.

### D-011 Language of the project files (2026-10-30)
CLAUDE.md is an English translation of section 7 of the plan (rule: documentation in English). docs/PLAN.md is the Polish original, moved unchanged from `plan-xlvane.md`; it takes precedence over the translation.

### D-012 License holder (2026-10-30) — to be confirmed by the maintainer
`LICENSE` files use the SPDX MIT text (F-TOOL-10) with "Copyright (c) 2026 DevXord" (the GitHub account that owns the repository). Change it if the copyright holder should be a person's legal name or an organisation.

### D-013 Security reports through GitHub private vulnerability reporting (2026-10-30)
SECURITY.md points reporters to GitHub private vulnerability reporting (F-TOOL-09). The repository owner must enable the feature in the repository settings; until then, reporters cannot use it.

---

## Facts: tools and platforms

### F-TOOL-01 tsup maintenance status
The tsup README says the project is no longer actively maintained and recommends tsdown. Quote: "This project is not actively maintained anymore." Source: https://github.com/egoist/tsup (README), checked 2026-10-30.

### F-TOOL-02 tsdown requirements
tsdown 0.23.4 requires Node.js `^22.18.0 || ^24.11.0 || >=26.0.0` (`npm view tsdown engines`; tsdown.dev/guide/getting-started says odd-numbered releases are not supported). Declaration files: with `isolatedDeclarations` enabled in tsconfig, tsdown "will use oxc-transform for .d.ts generation"; otherwise it uses the TypeScript compiler. `typescript` is a peer dependency (`^5.0.0 || ^6.0.0 || ^7.0.0`). Source: https://tsdown.dev/options/dts, https://tsdown.dev/guide/getting-started, npm registry, checked 2026-10-30.

### F-TOOL-03 Vitest 5 requirements
Vitest 5.0.5: `engines.node` is `^22.12.0 || ^24.0.0 || >=26.0.0`; `vite` (`^6.4.0 || ^7.0.0 || ^8.0.0`) is a peer dependency with `optional: false` in `peerDependenciesMeta`. Source: `npm view vitest@5.0.5 engines peerDependencies peerDependenciesMeta`, checked 2026-10-30.

### F-TOOL-04 Tool versions on 2026-10-30
npm `latest` versions: fflate 0.8.3, tsdown 0.23.4, vitest 5.0.5, vite 8.3.5, fast-check 4.10.2, publint 0.3.25, @arethetypeswrong/cli 0.18.5, typescript 7.0.2. Source: `npm view <name> version`.

### F-TOOL-05 publint CLI
`publint [path]` lints a directory or tarball; `--strict` reports warnings as errors; `--pack` chooses the package manager used to pack (default `auto`). Source: https://publint.dev/docs/cli, checked 2026-10-30.

### F-TOOL-06 @arethetypeswrong/cli
`attw --pack <dir>` packs a local directory and checks it; `--profile esm-only` "ignores CJS resolution failures". Source: README of `@arethetypeswrong/cli` (github.com/arethetypeswrong/arethetypeswrong.github.io, packages/cli/README.md), checked 2026-10-30.

### F-TOOL-07 Node.js release status
nodejs.org lists v22, v24 and v26 as LTS and v25 as EOL (last release March 31, 2026). Source: https://nodejs.org/en/about/previous-releases, checked 2026-10-30.

### F-TOOL-08 depgate on npm
`npm view depgate` returns `E404` with "Unpublished on 2026-10-29T21:30:32.875Z". Checked 2026-10-30.

### F-TOOL-09 GitHub private vulnerability reporting
Reporters open the repository's **Security and quality** tab and click **Report a vulnerability**; this works only when the owner has enabled private vulnerability reporting. Source: https://docs.github.com/en/code-security/security-advisories/guidance-on-reporting-and-writing-information-about-vulnerabilities/privately-reporting-a-security-vulnerability, checked 2026-10-30.

### F-TOOL-10 MIT license text
Copied from the SPDX license list data: https://raw.githubusercontent.com/spdx/license-list-data/main/text/MIT.txt, checked 2026-10-30.

---

## Decisions: phase 1 (reading and `inspect`)

### D-101 Own XML parser; DTDs are always rejected (2026-10-30)
`packages/xlvane/src/xml/parser.ts` is a non-validating, namespace-aware XML 1.0 parser with a SAX-style interface and an explicit element stack (no recursion). Any `<!DOCTYPE` fails with `XLV_E_XML_DTD_FORBIDDEN`, so external entities and entity expansion cannot happen. This is a deliberate deviation from XML 1.0 §5.1, which requires non-validating processors to read the internal DTD subset (fact X-14); ECMA-376 forbids DTDs in OPC's own parts (E-55) and SpreadsheetML parts do not use them. Only the five predefined entities and character references are recognised (X-06, X-07). Other rules implemented: the Char production on every character, so raw control characters are rejected (X-01, X-02); end-of-line normalisation before parsing (X-04); attribute-value normalisation as for CDATA attributes (X-05); Unique Att Spec and the namespaces rule Attributes Unique (X-08, N-07); reserved `xml`/`xmlns` prefixes, no prefix undeclaring, prefixed names must be declared (N-02 … N-04); comments, PIs (target `xml` only at the very start) and CDATA (X-10 … X-13). Only UTF-8 and UTF-16 with a BOM are decoded, and invalid byte sequences are fatal (X-15). Attributes are delivered in a `Map`, keyed by local name or `{namespace}local`. The XML declaration's pseudo-attributes are not validated beyond the encoding.

### D-102 Own ZIP central-directory reader, fflate only for Deflate (2026-10-30)
`zip/reader.ts` parses the end of central directory record, the central directory and the local headers itself (Z-01 … Z-16) and uses fflate's streaming `Inflate` for decompression, pushing 8 KiB of compressed data at a time. fflate calls `ondata` once per `push()` with everything that chunk produced (checked in `node_modules/fflate/esm/index.mjs`, `Inflate.prototype.c`), so the output can overshoot a limit by at most what 8 KiB of Deflate data expands to before the reader aborts. Limits apply to the bytes actually produced, never to declared sizes (W-07). Strict rules:
- the end of central directory record must end the file (no trailing data), and the first entry must start at offset 0 (no leading data);
- entries must not overlap each other or the central directory (Fifield's overlapping-files bomb, W-06, W-07);
- names must be unique ignoring case (OPC requires this too, E-58);
- the local header's name and method must match the central directory, stored entries need equal sizes, and the real size and CRC-32 must match the central directory;
- encryption (flag bits 0, 6, 13, method 99), methods other than 0 and 8, and multi-disk archives are rejected (Z-09, Z-10, Z-16; OPC forbids them as well, E-57).

ZIP64 is read (Z-11 … Z-13). CRC-32 uses the reflected polynomial 0xEDB88320 (RFC 1952 §8, W-11; APPNOTE does not name it, Z-17), and tests compare it with `node:zlib` `crc32` (F-NODE-02).
**Deviation from OPC:** ECMA-376 Part 2 Annex B says the UTF-8 name flag (bit 11) must be 0 and item names are ASCII with percent-encoding (E-57, E-58). xlvane still accepts UTF-8-flagged names for robustness. Names without the flag are decoded as IBM Code Page 437 (Z-18), using the mapping table from unicode.org.

### D-103 Default limits (2026-10-30) — to be reviewed against a real corpus
See `DEFAULT_READ_LIMITS` and the package README. The compression-ratio limit (200, checked only after 1 MiB of output) is a guess: real XLSX parts with very repetitive XML could exceed it. Revisit when the corpus from tests/corpus/REQUESTED.md exists.

### D-104 Error and warning model (2026-10-30)
`XlvError` carries `code`, `message`, `location` (part, sheet, cell, line, column, offset), `hint`, `docsUrl` and JSON-safe `details`, and `toJSON()` returns the same. Codes are listed in `ERROR_CODES` and `WARNING_CODES`, and a test checks that each has an entry in docs/errors.md. Warnings are aggregated per code and sheet (or part) with `count` and up to ten example cells, so a sheet with a million odd cells produces one warning. `docsUrl` points to `https://github.com/DevXord/xlvane/blob/main/docs/errors.md#<code>` because the `xlvane` GitHub organisation does not exist yet; change the base URL in `errors.ts` if the repository moves.

### D-105 Read model (2026-10-30)
`readWorkbook(bytes, options)` is synchronous and runtime-neutral. `readWorkbookFile` and `inspectFile` in `xlvane/node` add file access, with a size check before reading. Decisions in the model:
- Cells with neither a value nor a formula (for example formatting only) are not stored, which keeps memory proportional to the data.
- Numbers are parsed as finite `xsd:double`. Anything else, including `INF` and `NaN`, which have no JSON form, is returned as text with `XLV_W_INVALID_VALUE`.
- Numbers whose number format is a date or time format become `Date` objects whose UTC fields hold the spreadsheet's wall-clock value. Serials outside the date system stay numbers with a warning.
- `t="d"` cells (ISO 8601) are accepted in both conformance classes, although Transitional removes `d` from ST_CellType (E-21); a lenient reader loses nothing by accepting them.
- An `is` element is accepted even without `t="inlineStr"`, as in the spec's own example (E-22).
- Error cells have `type: "error"` and the error code as value. The eight codes of E-33 are not enforced.
- Rows and cells may appear in any order and may omit `r`, because the spec states no ordering rule and no rule for omitted references (E-24, E-25). An omitted reference is taken as the next row or next column, and `rows()` returns everything sorted.
- A cell or row reference outside A1:XFD1048576 (E-39, W-04) is an error (`XLV_E_INVALID_CELL_REF`), not a warning: such a file is damaged or crafted, and skipping the cell would silently change the data.
- Sheet lookup by name is exact (case-sensitive). The error lists the sheets and suggests close matches.

### D-106 Date conversion depends on the conformance class (2026-10-30)
The plan assumed a Strict `workbookPr@dateCompatibility` attribute; it does not exist in the 2016 edition (E-12). Instead, Strict serial dates count from 1899-12-30 without a phantom 29 February 1900 (E-13, E-14), while Transitional replaces that section with the 1900 system that Excel uses (Part 4 §15.3.4; E-15, E-16). `serialToDate` and `dateToSerial` therefore take the conformance class; see docs/facts/values.md.

### D-107 Shared formulas are not expanded yet (2026-10-30)
A cell that reuses a shared formula gets `formula.text = null` and `formula.sharedIndex`. Any text written in such a cell is ignored, as the spec requires (E-28), and `XLV_W_SHARED_FORMULA_NOT_EXPANDED` is reported. Translating the master formula to each cell needs a formula tokenizer and is left for phase 3.

### D-108 Macros, external links, encrypted and legacy files (2026-10-30)
Macros are detected through the VBA project relationship or content type, and through the Excel 4.0 macro sheet relationship or content types ([MS-OFFMACRO2], W-08 … W-10), because XLM macro sheets have no VBA part. External links are counted from `externalLink` relationships (E-05). Neither is ever executed or followed. Files that start with the Compound File Binary signature (W-01) are rejected with `XLV_E_UNSUPPORTED_CONTAINER`. Its hint covers both legacy `.xls` and password-protected `.xlsx`, because encrypted OOXML is stored in such a container (W-02, W-03).

### D-109 `inspect()` heuristics (2026-10-30)
- **Header row:** the first of the first 20 non-empty rows whose cells are all distinct non-blank strings covering at least half of the used width. Confidence is `high` when the next row has a non-string value.
- **`XLV_W_DIMENSION_MISMATCH`:** reported only when cells lie outside the declared dimension, because a larger dimension is normal for formatted empty cells.
- **`XLV_W_VALUES_IN_MERGED_CELLS`:** computed with a sweep over rows, so huge merged ranges cost nothing.
- **Output:** at most 200 columns are described by default, and dates are written as ISO 8601 without a time zone.

### D-110 Corpus from real applications is pending (2026-10-30)
The development machine has no Excel or LibreOffice, so the corpus test harness (`tests/corpus/corpus.test.ts`) currently has no real-application files. tests/corpus/REQUESTED.md lists the exact files needed from the maintainer. All other tests build their input from hand-written XML (`tests/helpers/xlsx-builder.ts`) or byte by byte (`tests/helpers/raw-zip.ts`).

### D-111 Portable Node.js for development in this session (2026-10-30)
The machine's system Node.js is 25.8.1, outside the tool engine ranges (D-007). With the maintainer's approval, the session used Node.js 26.13.0 LTS from https://nodejs.org/dist/v26.13.0/node-v26.13.0-win-x64.zip, verified against the published SHA-256 (`dcc60d59…150a`) and unpacked into a temporary directory. Nothing was installed system-wide.

### D-112 Time-only values in the Transitional 1900 system (2026-10-30)
Part 4 §15.3.4 puts the lower limit of the Transitional 1900 system at serial 1 (1900-01-01) and gives serial 0 as 1899-12-31 (docs/facts/values.md V-12). Taken literally, every time-only value such as 0.5 formatted `h:mm` would be outside the range. `serialToDate` therefore converts serials in [0, 1) onto 1899-12-31 (0.5 → 1899-12-31T12:00) and returns `null` only for negative serials and the nonexistent serial 60 (29 February 1900). `dateToSerial` accepts the same day back. Strict and 1904 serials follow Part 1 unchanged.

### D-113 Cell number formats follow Office, not the letter of ECMA-376 (2026-10-30)
ECMA-376 §18.8.9/§18.8.10 say both the cell xf and the cell style xf must be read to know a cell's formatting, and do not say how `applyNumberFormat` combines them (docs/facts/sst-styles.md). [MS-OI29500] "Part 1 Section 18.8.10, cellXfs (Cell Formats)", note b (https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/d0f88a45-c7c3-456b-9e21-2fd5577e9898, updated 2024-04-16) says that in Office only the cell xf record defines the formatting applied to a cell. xlvane does the same: a cell's number format is `cellXfs/xf[s]@numFmtId`, whatever `applyNumberFormat` and `cellStyleXfs` say.

### D-114 XML parser fixes found by fuzzing (2026-10-30)
The property tests found three deviations, now fixed with tests: `xmlns:` with an empty prefix is rejected (Namespaces in XML [2], N-01); character references may have any number of digits, including leading zeros (XML 1.0 [66], X-02); entity-name length is checked against `maxXmlNameLength` with a limit error, while character references are not limited by it. The XML declaration must now start with `version` (XML 1.0 [23], [24], X-13).

### D-115 Fixes from the security review: bounded work per byte (2026-10-30)
An independent review found inputs of tens of kilobytes that cost minutes of CPU or gigabytes of memory without hitting any limit. Fixed, each with a test in `tests/security/amplification.test.ts` or `tests/unit/read.test.ts`:
- Number formats are classified once per number format id per workbook (the cache was per style index and per sheet). Format codes longer than 1 024 characters are not classified as dates, with a warning. This is an implementation cap, not a documented Excel limit.
- Namespace names longer than `maxXmlNameLength` are rejected, because they are copied into every prefixed attribute key (`{namespace}local`).
- `maxXmlTreeElements` now counts attributes as well as elements, because both stay in memory.
- `WarningCollector` finds entries through nested maps instead of building a key that contains the sheet name.
- The `inspect()` check for values hidden under merged ranges drops duplicate ranges and stops after 5 000 000 steps per sheet; it only produces a warning.
- Malformed inline strings no longer leak text into the next cell, cells outside `sheetData` are ignored, and a missing styles or shared-strings part produces `XLV_W_PART_MISSING`.
- The default `maxCells` is lowered from 10 000 000 to 5 000 000. Measured with Node.js 26.13.0: 1 000 000 numeric cells (dense rows) retain about 183 MB of heap, about 183 bytes per cell, so 10 M cells would need about 1.8 GB.

## Decisions: phase 2 (writing and styles)

### D-201 Writing API: an in-memory model plus `writeWorkbook` (2026-10-30)
`createWorkbook()` returns a `NewWorkbook`; `addSheet()` returns a `NewSheet` with chainable `set`, `style`, `column`, `row`, `merge`, `freeze` and `autoFilter`. `writeWorkbook()` returns a `Uint8Array` and is runtime-neutral; `writeWorkbookFile()` in `xlvane/node` writes a file. The read model (`readWorkbook`) stays immutable and separate. Converting a read workbook into a `NewWorkbook` (needed by `editCells` and `appendRows`) is left for phase 3. Every value, style and option is validated when it is set, so errors point at the call that caused them. Option and style objects are read through `Object.keys` (own enumerable properties, each read once); unknown keys, including `__proto__`, are errors, and inherited properties are ignored.

### D-202 Own deterministic ZIP writer (2026-10-30)
`zip/writer.ts` writes local headers, the central directory and the end record itself (Z-02, Z-05, Z-07) and uses fflate's `deflateSync` (level 6) only to compress. An entry is stored instead when deflating does not make it smaller. Timestamps are fixed at 1980-01-01 00:00 (W2-40), flags are 0, names are ASCII, and there are no data descriptors and no extra fields, so the same input always gives the same bytes and every output passes our own strict reader. ZIP64 is not written: more than 65 534 parts or 4 GiB throws `XLV_E_LIMIT_EXCEEDED`. fflate's `zipSync` was not used because its timestamps default to the current time and its output layout is not ours to keep stable.

### D-203 Phase 2 facts are partly provisional (2026-10-30) — to be reviewed
The cloud session could not download ECMA-376 or Microsoft documentation (proxy refusal). The maintainer chose to supply the documents. Meanwhile the writer was built on a copy of the Transitional `sml.xsd` from the python-docx repository, the Open XML SDK part definitions and LibreOffice's behaviour, all pinned to commits in docs/facts/sml-write.md. Facts taken from them are marked PROVISIONAL and listed under "To verify". LibreOffice opening the files and keeping values, styles and layout (D-209) reduces the risk, but does not replace checking the specification.

### D-204 Number format ids (2026-10-30)
Format codes equal to a built-in, non-date "All Languages" format (ids 1–4, 9–13, 37–40, 48, 49; E-41) reuse that id. All other codes, including every date format, get declared `numFmt` records starting at id 164. Date formats are always declared because the built-in date ids are displayed differently per locale (E-41, E-42). 164 is the first id Excel uses for custom formats; ECMA-376 only says that ids outside the built-in list need a declared format (F-STY-04).

### D-205 Default font, reserved fills and the Normal style (2026-10-30) — convention
New workbooks use Calibri 11 as font 0, fills 0 and 1 are `none` and `gray125`, and there is one cell style "Normal" (W2-12, W2-14, W2-18). This matches the files Excel produces; it is not required by the parts of the schema checked so far. A cell's `style.font` is reported only for fonts other than font 0, and the writer completes partial fonts with the default name and size.

### D-206 Range styling cap (2026-10-30)
`sheet.style(range)` creates a cell for every position in the range, so ranges over 1 048 576 cells (one full column) throw `XLV_E_LIMIT_EXCEEDED`. Column and row default styles, which would style whole columns cheaply, are not modelled yet.

### D-207 Writer limits taken from Excel (2026-10-30)
Text up to 32 767 characters, column width up to 255 and row height up to 409 points (W-04, checked in phase 1). Sheet names (W2-20), formula length 8 192 (W2-21) and font size 409 (W2-19) are provisional. Font names are capped at 255 characters as an implementation limit. Strings with unpaired UTF-16 surrogates are rejected, because XML cannot hold them and the `_xHHHH_` escape of a lone surrogate is not decoded by readers that follow D-XS-3.

### D-208 Package size budget (2026-10-30)
70 kB gzip (level 9) for all `packages/xlvane/dist/*.js` files as published (unminified, comments kept). Measured 58.4 kB at the end of phase 2; fflate adds about 22 kB. `npm run check:pkg` enforces the budget; `tests/policy/size-budget.test.ts` checks that the README states the same number. Minifying the build would roughly halve the size but makes stack traces harder to read; not done.

### D-209 LibreOffice check of written files (2026-10-30)
`tests/libreoffice/libreoffice.test.ts` writes workbooks covering every feature of the phase, converts them with `soffice --headless` to CSV (all sheets, contents as shown; F-LO-01) and back to XLSX, and checks values in the CSV and merges, frozen panes, filter, widths, heights, states and styles in LibreOffice's own XLSX. The test is skipped when `soffice` is missing and fails instead when `XLV_REQUIRE_LIBREOFFICE=1` (set in CI). Each run uses its own LibreOffice profile directory.

### D-210 CI workflow (2026-10-30)
`.github/workflows/ci.yml` runs typecheck, tests (with LibreOffice from the Ubuntu 24.04 `libreoffice-calc` package) and `check:pkg` on Node.js 22, 24 and 26, with `contents: read` permissions and `persist-credentials: false`. Actions are pinned to commit SHAs (F-GH-01). Bun, Deno, browsers and Workers remain for phase 9.

### D-211 Reader additions for round trips (2026-10-30)
`Cell.style` (frozen, shared per style index; parts equal to the defaults left out), `Sheet.columns` (the `col` records as ranges, so a file cannot make the reader expand millions of entries), `Sheet.rowInfo(row)`, `Sheet.freeze` (first sheet view, `frozen` or `frozenSplit` panes only) and `Sheet.autoFilter`. Colours are returned as upper-case ARGB strings, or as `{ theme, tint }`, `{ indexed }`, `{ auto: true }`. Formula text is now decoded as ST_Xstring (W2-02); before, `_xHHHH_` escapes in formulas were returned undecoded. The `inspect()` JSON is unchanged.

## Facts: Node.js

### F-NODE-01 require() of ES modules and top-level await
`require()` loads ES modules without a flag from v20.19.0, v22.12.0 and v23.0.0, and throws `ERR_REQUIRE_ASYNC_MODULE` when the module or its imports use top-level await. Source: https://nodejs.org/api/modules.html#loading-ecmascript-modules-using-require (v26.13.0 docs), checked 2026-10-30 (W-05 in docs/facts/zip-xml-web.md). Also verified locally: Node.js 26.13.0 throws `ERR_REQUIRE_ASYNC_MODULE` for `await Promise.resolve(); export const x = 1;`.

### F-NODE-02 zlib.crc32
`zlib.crc32(data[, value])` was added in v22.2.0 and v20.15.0 and returns a 32-bit unsigned integer. Source: https://github.com/nodejs/node/blob/main/doc/api/zlib.md, checked 2026-10-30.

### F-NODE-03 fs.promises.stat and readFile
`fsPromises.stat(path)` resolves to a `Stats` object with `size` and `isFile()`; `fsPromises.readFile(path)` resolves to a Buffer (a `Uint8Array` subclass). Source: https://github.com/nodejs/node/blob/v26.x/doc/api/fs.md, checked 2026-10-30.

## Facts: file formats

Format facts live in separate files under docs/facts/ so that each one can carry its full source references (part, section, page, line in the extracted text). Code comments cite them by id.

| File | Ids | Sources |
|---|---|---|
| docs/facts/ecma-376.md | E-01 … E-59 | ECMA-376 Parts 1, 2 and 4 (5th editions) and their schemas |
| docs/facts/zip-xml-web.md | Z-*, X-*, N-*, W-* | APPNOTE.TXT 6.3.10, XML 1.0 5th ed., Namespaces in XML 1.0 3rd ed., [MS-CFB], [MS-OFFCRYPTO], [MS-OFFMACRO2], Excel specifications and limits, Node.js docs, Fifield's "A better zip bomb", RFC 1952 |
| docs/facts/values.md | see file | A1 references, serial dates, number formats |
| docs/facts/opc.md | see file | ECMA-376 Part 2 (packaging) |
| docs/facts/sst-styles.md | see file | shared strings, ST_Xstring, styles |
| docs/facts/sml-write.md | W2-*, F-LO-01, F-NODE-04, F-NODE-05, F-GH-01 | writer facts; partly PROVISIONAL (D-203): python-docx copy of sml.xsd, Open XML SDK, LibreOffice help, Node.js docs |
