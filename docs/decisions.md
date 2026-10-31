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

### D-012 License holder (2026-10-30) - to be confirmed by the maintainer
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

### D-103 Default limits (2026-10-30) - to be reviewed against a real corpus
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

### D-107 Shared formulas are not expanded yet (2026-10-30) - closed in phase 3 by D-313
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

### D-203 Phase 2 facts are partly provisional (2026-10-30) - to be reviewed
The cloud session could not download ECMA-376 or Microsoft documentation (proxy refusal). The maintainer chose to supply the documents. Meanwhile the writer was built on a copy of the Transitional `sml.xsd` from the python-docx repository, the Open XML SDK part definitions and LibreOffice's behaviour, all pinned to commits in docs/facts/sml-write.md. Facts taken from them are marked PROVISIONAL and listed under "To verify". LibreOffice opening the files and keeping values, styles and layout (D-209) reduces the risk, but does not replace checking the specification.

### D-204 Number format ids (2026-10-30)
Format codes equal to a built-in, non-date "All Languages" format (ids 1–4, 9–13, 37–40, 48, 49; E-41) reuse that id. All other codes, including every date format, get declared `numFmt` records starting at id 164. Date formats are always declared because the built-in date ids are displayed differently per locale (E-41, E-42). 164 is the first id Excel uses for custom formats; ECMA-376 only says that ids outside the built-in list need a declared format (F-STY-04).

### D-205 Default font, reserved fills and the Normal style (2026-10-30) - convention
New workbooks use Calibri 11 as font 0, fills 0 and 1 are `none` and `gray125`, and there is one cell style "Normal" (W2-12, W2-14, W2-18). This matches the files Excel produces; it is not required by the parts of the schema checked so far. A cell's `style.font` is reported only for fonts other than font 0, and the writer completes partial fonts with the default name and size.

### D-206 Range styling cap (2026-10-30)
`sheet.style(range)` creates a cell for every position in the range, so ranges over 1 048 576 cells (one full column) throw `XLV_E_LIMIT_EXCEEDED`. Column and row default styles, which would style whole columns cheaply, are not modelled yet.

### D-207 Writer limits taken from Excel (2026-10-30)
Text up to 32 767 characters, column width up to 255 and row height up to 409 points (W-04, checked in phase 1). Sheet names (W2-20), formula length 8 192 (W2-21) and font size 409 (W2-19) are provisional. Font names are capped at 255 characters as an implementation limit. Strings with unpaired UTF-16 surrogates are rejected, because XML cannot hold them and the `_xHHHH_` escape of a lone surrogate is not decoded by readers that follow D-XS-3.

### D-208 Package size budget (2026-10-30)
70 kB gzip (level 9) for all `packages/xlvane/dist/*.js` files as published (unminified, comments kept). Measured 58.4 kB at the end of phase 2; fflate adds about 22 kB. `npm run check:pkg` enforces the budget; `tests/policy/size-budget.test.ts` checks that the README states the same number. Minifying the build would roughly halve the size but makes stack traces harder to read; not done. **Phase 3:** the task API brought the size to 84.5 kB; the maintainer chose to raise the budget to 100 kB and keep the build unminified (a minified build measured about 51 kB). **Phase 4:** raised to 125 kB (D-408). **Phase A:** raised to 175 kB (D-A11). **Phase D:** raised to 250 kB (D-D01), then to 300 kB at the maintainer's request (D-D01 addendum).

### D-209 LibreOffice check of written files (2026-10-30)
`tests/libreoffice/libreoffice.test.ts` writes workbooks covering every feature of the phase, converts them with `soffice --headless` to CSV (all sheets, contents as shown; F-LO-01) and back to XLSX, and checks values in the CSV and merges, frozen panes, filter, widths, heights, states and styles in LibreOffice's own XLSX. The test is skipped when `soffice` is missing and fails instead when `XLV_REQUIRE_LIBREOFFICE=1` (set in CI). Each run uses its own LibreOffice profile directory.

### D-210 CI workflow (2026-10-30)
`.github/workflows/ci.yml` runs typecheck, tests (with LibreOffice from the Ubuntu 24.04 `libreoffice-calc` package) and `check:pkg` on Node.js 22, 24 and 26, with `contents: read` permissions and `persist-credentials: false`. Actions are pinned to commit SHAs (F-GH-01). Bun, Deno, browsers and Workers: see D-904 and D-905.

### D-211 Reader additions for round trips (2026-10-30)
`Cell.style` (frozen, shared per style index; parts equal to the defaults left out), `Sheet.columns` (the `col` records as ranges, so a file cannot make the reader expand millions of entries), `Sheet.rowInfo(row)`, `Sheet.freeze` (first sheet view, `frozen` or `frozenSplit` panes only) and `Sheet.autoFilter`. Colours are returned as upper-case ARGB strings, or as `{ theme, tint }`, `{ indexed }`, `{ auto: true }`. Formula text is now decoded as ST_Xstring (W2-02); before, `_xHHHH_` escapes in formulas were returned undecoded. The `inspect()` JSON is unchanged.

## Decisions: phase 3 (task-oriented API)

### D-301 Development dependencies for schema tests (2026-10-30)
Approved by the maintainer: zod 4.6.7, valibot 1.5.0 and arktype 2.2.9 as exact root devDependencies (SS-02), used only by tests (type inference, issue mapping) and by TSDoc examples. None of them is a dependency of `xlvane`; `tests/policy/packages.test.ts` lists them as approved.

### D-302 Standard Schema types copied, not imported (2026-10-30)
docs/PLAN.md asks to copy the interface types from the specification. `src/standard-schema.ts` holds the declarations of S-SS at a pinned commit with the MIT notice; `@standard-schema/spec` is not a dependency. Only version 1 is accepted (`XLV_E_SCHEMA_INVALID` otherwise).

### D-303 `readTable` row model (2026-10-30)
- Keys are the trimmed header texts; duplicates get `_2`, `_3`; empty headers use the column letter. Without a header (`header: false`, or `"auto"` finding none) keys are column letters.
- Every row has every key; empty cells are `null`; error cells give their code as a string; dates are `Date` (or ISO strings with `dates: "iso"`).
- Row objects are plain objects whose properties are defined with `Object.defineProperty`, so keys such as `__proto__` are ordinary own properties and no prototype changes.
- With a schema, `rows` holds only valid rows (output type of the schema), `rowNumbers` their sheet rows, and `issues` the problems, mapped to a cell through the first path key. `readTable` throws `XLV_E_ASYNC_SCHEMA` for asynchronous schemas instead of returning a Promise sometimes; `readTableAsync` handles them.
- Header detection is the one of `inspect()` (D-109), so both agree.

### D-304 `writeTable` report preset (2026-10-30)
Bold header with fill `D9E1F2` and a thin bottom border, frozen first row, autofilter over the table, column widths from the longest value (8 to 60 characters) unless given. Dates get the writer's date formats (D-204); numbers keep the General format unless a column sets `numFmt`, so values are not rounded on display. `"plain"` writes values only. Rows must all be objects or all arrays; values must be cell values (`XLV_E_INVALID_ROWS`, called `XLV_E_TABLE_INVALID` before 1.0.0, with row and key otherwise); inherited properties are ignored.

### D-305 `toJSON` and `fromJSON` (2026-10-30)
`toJSON` is `readTable` with ISO dates (without a time zone, as in CSV, since D-R02) and returns `{ sheet, headerRow, range, columns, rows }`. `fromJSON` accepts JSON text (parsed with `JSON.parse`, errors as `XLV_E_JSON_INVALID`) or an array, and writes it with `writeTable`. JSON keys such as `__proto__` are own properties after `JSON.parse` and become ordinary columns. ISO date strings stay strings: JSON has no date type, and guessing would change data. So dates survive a CSV round trip (`fromCSV` infers them) but come back from a JSON round trip as text; this is documented in the README and `xlvane help convert` (D-R02).

### D-306 `toCSV` (2026-10-30)
Output covers A1 to the last used row and column (as LibreOffice does), so positions match the sheet; `range` restricts it. Numbers use JavaScript's shortest round-trip form, booleans `TRUE`/`FALSE`, dates ISO 8601 (date only at midnight, otherwise `YYYY-MM-DDThh:mm:ss[.sss]` without a zone, because sheet dates have no zone), errors their code, formulas their cached value. Formula-injection escaping is on by default (CSV-10, CSV-11).

### D-307 `fromCSV` (2026-10-30)
Own linear parser (CSV-01 … CSV-03). Type inference (`types: "infer"`, default): plain decimal numbers without leading zeros and with at most 15 significant digits (so IDs, account and postal codes keep their digits), and since D-R02 also 16–17-digit numbers with a fraction or exponent written in JavaScript's shortest form (what `toCSV` writes), but not numbers a double cannot hold (`1e400`, `1e-400`, subnormals), `true`/`false` in any case, and ISO 8601 dates with a date part inside the date system; everything else is text. Nothing ever becomes a formula. Default preset `plain`. Limits: `maxChars` (default 256 Mi characters), `maxCells` (default 5 000 000 fields, the `maxCells` read limit; D-R02), 16 384 fields per line, 1 048 576 lines, 32 767 characters per field (the cell limit).

### D-308 `diff` (2026-10-30)
Sheets are matched by name; a renamed sheet shows as removed plus added. Cells are compared by type, value (dates by time) and formula text; styles only with `styles: true`. Cached formula results are compared as stored. At most `maxCells` changes (default 10 000) are listed per sheet, in row order; the summary counts all. Output is JSON-safe.

### D-309 Surgical editing (2026-10-30)
Chosen by the maintainer. `editCells` and `appendRows` rewrite only:
- the `sheetData` of edited sheets: untouched rows are copied as source text; in an edited row, untouched cells are copied as source text and new cells are inserted in column order; rows or cells without `r` get an explicit `r` so that positions cannot shift;
- the `dimension@ref` of edited sheets (grown to include new cells);
- appended `si` records in the shared string table (an existing equal string is reused; `count` and `uniqueCount` updated), or a new shared string part with its relationship and content type;
- appended `font`/`fill`/`border`/`numFmt`/`xf` records in the styles part (existing indexes never change), creating missing containers in schema order;
- `calcPr fullCalcOnLoad="1"` in the workbook part when it contains formulas, and removal of the calculation chain (part, relationship, content type) when formula cells were added, changed or removed (D-312).
Every other part keeps its decompressed bytes; the ZIP is rewritten with the D-202 writer. Tag positions come from new `XmlSpan` arguments of the SAX parser.

### D-310 Text of edited parts (2026-10-30)
Edited parts are produced from the parser's input after end-of-line normalisation (XML 1.0 §2.11), so CRLF in an edited part becomes LF. That does not change the XML information set. Parts declared as UTF-16 are written as UTF-8 and their declaration is corrected.

### D-311 Styles in edits (2026-10-30)
An edit's `style` is merged into the cell's current format like `NewSheet.style()`: given font, border and alignment properties replace the current ones, fill and number format replace. The new `xf` copies the current one's attributes and `protection`, and keeps unknown `alignment` attributes; its `extLst` and namespaced attributes are not copied. New fonts are completed with the workbook's font 0 (name and size). Dates written into cells without a date format get the writer's date format; cells that already have a date format keep it. When a container such as `fonts` is missing, it is created with the default records first, so that index 0 keeps its meaning.

### D-312 Formulas after edits (2026-10-30) - PROVISIONAL fact
xlvane does not calculate. When the workbook has formulas, `XLV_W_FORMULAS_NOT_RECALCULATED` is reported and `calcPr@fullCalcOnLoad="1"` asks applications to recalculate on load (attribute name from the schema copy S-XSD; its meaning is to be checked in ECMA-376 §18.2.2). The calculation chain lists formula cells, so it is removed when formula cells change; applications rebuild it. Editing a Strict workbook refuses dates before 1900-03-01, because the writer converts dates with the Transitional rules (D-106).

### D-313 Shared formula expansion (2026-10-30)
Closes D-107. A conservative linear tokenizer (`sml/formula-shift.ts`) translates the master formula to dependent cells inside the master's range by shifting relative A1 references (facts docs/facts/formulas.md, mostly PROVISIONAL). It refuses (keeps `text: null` with `XLV_W_SHARED_FORMULA_NOT_EXPANDED`) for R1C1-looking names, unbalanced quotes or brackets, unknown characters, references that would leave the grid (Excel would write `#REF!`), and ambiguous sheet-qualified or 3-D references. Each master is tokenized once and dependents render lazily, so memory stays proportional to the file. Moved references are written in upper case.

### D-314 TSDoc example check (2026-10-30)
`tests/policy/tsdoc-examples.test.ts` extracts every ```ts block from TSDoc in `packages/xlvane/src`, type-checks them in one `tsc` run (TypeScript 7 CLI; its JavaScript compiler API is not stable) with the project's strict options and `xlvane`/`xlvane/node` mapped to the sources, and checks that every symbol exported from `index.ts` and `node.ts` has an `@example`. The temporary project lives under `node_modules/.cache` so that examples can import zod.

## Decisions: phase 4 (streaming)

### D-401 Streaming read: structure first, worksheets row by row (2026-10-30)
`openWorkbookStream(source)` accepts a `Uint8Array`, `ArrayBuffer` or `Blob`. It reads the ZIP central directory, `[Content_Types].xml`, the relationship parts, the workbook part, the shared strings and the styles in full (all bounded by the limits). Worksheets are decompressed and parsed only while their rows are iterated. The batch and streaming readers share `analyzeWorkbook` (read.ts) and `createWorksheetParser` (sml/worksheet.ts). Streaming limits (`STREAM_LIMIT_DEFAULTS`): `maxCells` does not apply, because memory does not grow with cells; `maxEntryBytes` is 4 GiB and `maxTotalBytes` 8 GiB; the ratio, XML and other limits are unchanged. In Node.js, `openWorkbookStreamFile` reads through `fs.openAsBlob` (F-NODE-06), so the file is not held in memory either.

### D-402 Rows in file order; shared formulas only after their master (2026-10-30)
In streaming mode each row's cells are handed on at `</row>` and dropped. Rows therefore come in file order, which is ascending in every file we know of. Row heights travel with the row instead of being collected. A shared-formula dependent that comes before its master cannot be expanded (the stream cannot look ahead); it keeps `text: null` with `XLV_W_SHARED_FORMULA_NOT_EXPANDED`. Incremental XML parsing (`createXmlStreamParser`, `createXmlByteStreamParser`) gives the same events, spans and errors as whole-text parsing. Two differences: events from earlier chunks are already delivered when an error is found later, and in a document with several errors the one reported can differ (an invalid character is reported where it is met, not before everything else).

### D-403 Streaming write: pull-based, layout up front (2026-10-30)
`writeWorkbookStream({ sheets: [{ name, columns, freeze, merges, autoFilter, rows }] })` returns a `ReadableStream<Uint8Array>`. Rows are pulled from each sheet's (async) iterable only when the consumer reads, which gives backpressure and constant memory. Layout must be given before the rows. The autofilter range (`autoFilter: true`) and the styles part are computed after the rows, and the ZIP entries are produced lazily in this order: content types, package relationships, sheets, workbook, workbook relationships, styles, shared strings. No `dimension` element is written (it is optional, W2-01). Values and styles follow the batch writer's rules (`convertCellInput`), with errors naming sheet and cell. `writeWorkbookStreamFile` writes with the exclusive-create flag unless `overwrite`, and removes the file if the stream fails part-way.

### D-404 Inline strings by default when streaming (2026-10-30)
A shared string table needs every distinct string in memory until the end, so streamed text is written inline (`t="inlineStr"`, E-22/E-29) by default. `sharedStrings: true` writes a shared string table instead (smaller files when values repeat). Excel and LibreOffice read inline strings; the LibreOffice test covers them.

### D-405 Streaming ZIP (2026-10-30)
`zip/stream.ts`. Reading shares the central-directory and local-header validation of `openZip` (reader.ts was refactored into `checkSource`, `locateCentralDirectory`, `readCentralDirectory` and `validateLocalHeaders`). Entries are inflated incrementally, with the size, total and ratio limits checked on every output chunk and the CRC at the end. Writing deflates incrementally (level 6) and writes entries with flag bit 3 and a data descriptor carrying the 0x08074b50 signature (Z-15), with fixed timestamps. Entries are always deflated (storing would need the size in advance), and archives that would need ZIP64 are refused.

### D-406 Memory test (2026-10-30)
`tests/stream/memory.test.ts` measures heap retained after forced GC (Vitest `execArgv: ['--expose-gc']`) every 25 000 rows while writing and reading 50 000 and 400 000 rows × 5 columns. Measured in the cloud session: about 0.4 to 0.9 MB for every case. The test requires the difference between the two sizes to stay under 16 MB and the large read under 48 MB.

### D-407 Benchmark setup (2026-10-30)
At the maintainer's decision the benchmark libraries live in `bench/package.json` (exceljs 4.4.0, xlsx 0.18.5 from npm), installed only by `npm run bench`. Every variant runs in a fresh process; median time, peak RSS (`process.resourceUsage().maxRSS`, KiB per the Node.js docs) and peak heap are reported. The library APIs used are listed with file and line in bench/README.md. Results depend on the machine and are not a guarantee.

### D-408 Size budget raised to 125 kB (2026-10-30)
Streaming brought the bundle to 101.4 kB gzip. The maintainer chose to raise the budget to 125 kB and keep the build unminified.

## Decisions: phase 5 (exceljs compatibility)

### D-501 Module shape of `@xlvane/exceljs` (2026-10-30)
The package is ESM-only like `xlvane`. `src/index.ts` exports named classes and enums, a default export `ExcelJS` (a frozen object with `Workbook`, `ModelContainer`, `stream`, `ValueType`, `FormulaType`, `ErrorValue`, the members of the exceljs 4.4.0 CommonJS export `lib/exceljs.nodejs.js`), and the same object under the string export name `'module.exports'`, so that `require('exceljs')` returns it directly (F-NODE-08). tsdown 0.23.4 and TypeScript 7.0.2 accept the string export name; checked by building and by `tests/exceljs/alias.test.ts` (ESM default import, ESM named import, CommonJS `require`). `engines.node` is `>=22.12` at the maintainer's decision; `xlvane` stays at `>=22`.

### D-502 Unsupported API throws `XLV_E_NOT_SUPPORTED` (2026-10-30)
At the maintainer's decision every method, option or value outside the supported subset throws `XlvError('XLV_E_NOT_SUPPORTED')` naming the feature, with a hint linking to docs/exceljs-compat.md. Values (hyperlink, rich text, arbitrary objects) are rejected when assigned, not when written, so the stack points at the caller. Style parts that change how a cell looks (gradient fills, diagonal borders, `vertAlign`, `outline`, `shrinkToFit`, `readingOrder`, protection) throw; font `family`, `scheme` and `charset` only describe substitution and are ignored.

### D-503 In-memory model, converted on write (2026-10-30)
`Workbook`, `Worksheet`, `Row`, `Column` and `Cell` are an in-memory model that follows exceljs 4.4.0 behaviour (`lib/doc/*.js`): sparse 1-based `row.values`, `rowCount` as the last row number, column keys, merged cells that delegate to their master, styles inherited from rows and columns when a cell is created. Writing converts the model into an xlvane `NewWorkbook` (`convert.ts`); merged non-master cells keep their style but no value.

### D-504 Style mapping (2026-10-30)
exceljs `vertical: 'middle'` is SpreadsheetML `center`; `textRotation: 'vertical'` is 255. Fills read back as `{ type: 'pattern', pattern, fgColor }`, with `bgColor` only when the file has one, matching what exceljs 4.4.0 reads (checked by the cross-check sample).

### D-505 Reading into the model (2026-10-30)
`xlsx.load`/`readFile`/`read` read with xlvane (its limits for untrusted files apply) and fill the model: formulas as `{ formula, result }`, errors as `{ error }`, shared formulas translated to ordinary formulas, rich text as plain text, frozen panes as `views: [{ state: 'frozen', xSplit?, ySplit? }]`, the autofilter as an A1 string. A workbook with no worksheets cannot be written (xlvane needs one) and throws.

### D-506 Ported exceljs specs (2026-10-30)
Specs from exceljs tag v4.4.0 (commit d911329c, MIT) are ported to Vitest in `tests/exceljs/` with their attribution header; sources, HTTP status of each path and the reasons for skipped tests are in tests/exceljs/README.md. Tests that need upstream binary fixtures or unsupported features are skipped with the reason. A cross-check runs the same sample project against exceljs 4.4.0 (from `bench/node_modules`, when installed) and stores its output in `expected-exceljs-4.4.0.json`; our output must match it.

### D-507 Streaming writer and reader (2026-10-30)
`stream.xlsx.WorkbookWriter` feeds `writeWorkbookStream` through an async queue; committed rows are converted and dropped. To allow worksheets to be added while writing, `writeWorkbookStream` now accepts `sheets` as an (async) iterable and then writes the package parts ([Content_Types].xml, relationships, workbook, styles, shared strings) after the worksheets (allowed: ZIP entries may be in any order, Z-19; LibreOffice opens the result, checked in alias.test.ts). The worksheet layout is written before its rows, so it must be set before the first commit; a later change throws `XLV_E_NOT_SUPPORTED` (exceljs writes `autoFilter` and merges at the end). `stream.xlsx.WorkbookReader` iterates `openWorkbookStream`.

## Decisions: phase 6 (command line)

### D-601 CLI structure (2026-10-30)
The `xlvane` command lives in the `xlvane` package (docs/PLAN.md section 3) as `src/cli/`, built by tsdown into `dist/cli.js` and declared in `package.json` `bin`; it is not in `exports`, so it is not library API. `run(argv, io)` gets a `CliIO` (output functions, environment, working directory, TTY flags, module loader, version) and returns the exit code; only `src/cli/bin.ts` touches `process` (via `node:process`). Tests run the CLI in-process and once from a separate build (`tests/cli/bin.test.ts`). The CLI and `src/util/node-files.ts` are Node.js-only modules; `tests/policy/packages.test.ts` checks that `src/index.ts` does not reach them. No new dependencies: arguments use `util.parseArgs` (F-NODE-09), colour uses `util.styleText` (F-NODE-10).

### D-602 Colour (2026-10-30)
`util.styleText` respects `NO_COLOR`, `NODE_DISABLE_COLORS`, `FORCE_COLOR` and isTTY only from Node.js 22.8.0, and xlvane supports `>=22`. The CLI therefore applies the rules Node.js documents (doc/api/cli.md, v22.x) itself and calls `styleText` with `validateStream: false` only when colour is on. Colour only adds to text: every status is a word (`changed`, `added`, `warning`, `Error:`), and a test checks that output with colour removed equals output without colour.

### D-603 Arguments and typos (2026-10-30)
`parseArgs` runs with `strict: false, tokens: true`, and the CLI checks the tokens itself, so every mistake is an `XLV_E_USAGE` error (new code) with "Did you mean …?" from the existing `suggest()` (optimal string alignment distance, at most a third of the name). Commands must come first. `--help`/`-h` anywhere before `--` prints the command's help. `xlvane mcp` passes its arguments on unparsed.

### D-604 Writing and exit codes (2026-10-30)
An output file is created exclusively (`wx`); an existing file is replaced only with `--force`, through a temporary file in the same directory and a rename. `--dry-run` reports what would be written. `edit` never changes its input without `--force` (or writes to `--output`). Exit codes: 0 success, 1 a check found something (`diff`: files differ, `validate`: problems), 2 usage error, 3 input or file error. Unexpected exceptions become `XLV_E_INTERNAL` (new code; `--verbose` prints the stack).

### D-605 `--json` envelope and schemas (2026-10-30)
`--json` prints exactly one document: `{ xlvane: 1, command, ok: true, result }` or `{ xlvane: 1, command, ok: false, error, try }`, where `error` is `XlvError.toJSON()` and `try` the suggested command or `null`. The schemas are JSON Schema files in docs/cli-json/. At the maintainer's decision they are checked by a small validator in `tests/helpers/json-schema.ts` that supports only the keywords the schemas use and throws on any other keyword (no new dependency).

### D-606 Value typing in `edit --set` (2026-10-30)
`--set` values are typed like CSV fields with `types: "infer"` (the same `inferValue` as `fromCSV`): numbers, `true`/`false` and ISO 8601 dates; `=...` is a formula, a leading apostrophe keeps the text (as in spreadsheet programs), and an empty value clears the cell.

### D-607 `migrate` and `mcp` (2026-10-30)
At the maintainer's decision `migrate` reports and changes `package.json` only with `--force`, replacing just the `exceljs` value with `npm:@xlvane/exceljs@^1` so that formatting stays. It scans source files that mention `exceljs` (no `node_modules`, no dot folders, at most 10 000 files of 2 MB) for API listed as not or partly supported in docs/exceljs-compat.md; matching is textual. The report says that @xlvane/exceljs is not published yet; this line must be removed when it is. `mcp` loads `@xlvane/mcp` with a dynamic import and calls its `runCli(args): Promise<number>`; this is the contract for phase 7. Without the package it reports `XLV_E_NOT_INSTALLED` (new code) with `npm install @xlvane/mcp`.

### D-608 Size budgets (2026-10-30)
At the maintainer's decision the 125 kB budget covers the library files of `dist/` and `dist/cli.js` has its own budget of 40 kB gzip. Measured: library 104.1 kB (from 101.7 kB; adding the CLI entry changed how tsdown splits shared chunks), CLI 20.3 kB.

## Decisions: phase 7 (agent integration)

> D-701 … D-705 are superseded by D-802: the MCP server was removed. D-706 … D-708 still apply. The facts F-MCP-01 and F-MCP-02 are kept as research.

### D-701 MCP SDK and dependencies (2026-10-30)
At the maintainer's decision `@xlvane/mcp` uses the v2 SDK, `@modelcontextprotocol/server` 2.4.0, pinned exactly (F-MCP-01). Its runtime dependencies are `xlvane` and `@modelcontextprotocol/server` only (checked by `tests/policy/packages.test.ts`); the SDK brings `@modelcontextprotocol/core` and zod. Input schemas are plain JSON Schema passed through the SDK's `fromJsonSchema`, so the package imports no schema library. The maintainer approved two root development dependencies: `@modelcontextprotocol/client` 2.4.0 (tests on an in-memory and a stdio client) and `skills` 1.7.4 (the install test). Licenses: the SDK packages are Apache-2.0, zod, fflate and skills MIT.

### D-702 Tools (2026-10-30)
Seven tools as in docs/PLAN.md section 4, registered in a fixed order (the 2026-07-28 spec says tools/list SHOULD be deterministic, F-MCP-02): `inspect_workbook`, `read_table`, `diff`, `validate`, then the write tools `write_table`, `edit_cells`, `convert`. Each has a title, a factual description, an input schema with `additionalProperties: false` and annotations (`readOnlyHint` on the read tools; `destructiveHint: false` on the write tools, which never delete and replace only with `overwrite: true`). Results are `structuredContent` plus the same JSON as text; `XlvError`s become tool errors (`isError: true`) with code, location, hint and details. `read_table` pages with `offset`/`limit` (at most `--max-rows`, default 1000); `diff` lists at most `maxChanges` changes per sheet. Cell values in write calls accept `{ "date": "YYYY-MM-DD..." }` for dates, because JSON has no date type and strings must stay strings.

### D-703 Path policy (2026-10-30)
The 2026-07-28 spec deprecates Roots and suggests passing directories through server configuration (F-MCP-02), so the server takes `--root` directories at start; at the maintainer's decision at least one is required. Roots are resolved with `realpath`. A file is read only if its real path is inside a root; a file is written only if the real path of its directory is inside a root and the file itself is not a symbolic link. New files are created with `wx`; existing files are replaced only with `overwrite: true`, via a temporary file and a rename (a rename replaces a link instead of following it). Violations give `XLV_E_PATH_NOT_ALLOWED` (new code). Files above `--max-file-size` (default 64 MiB, below the library's 256 MiB default because tool calls are not reviewed by a person) are refused for reading and writing.

### D-704 Write tools only with --allow-write (2026-10-30)
At the maintainer's decision the write tools are not registered without `--allow-write`; the server instructions say so, so a client knows how writing is enabled. The instructions contain only the configuration (roots, writing, limits) and the error format.

### D-705 Entry points (2026-10-30)
`createServer(options): Promise<McpServer>` (asynchronous because the roots are resolved on the file system) and `runCli(args): Promise<number>`, the contract `xlvane mcp` uses (D-607). `runCli` serves stdio through the SDK's `serveStdio` and resolves when stdin ends. The package has a `bin`, `xlvane-mcp`, so clients can start it with `npx @xlvane/mcp --root <dir>`. stdout carries only the protocol; messages go to stderr.

### D-706 Agent Skill (2026-10-30)
`skills/xlvane/SKILL.md` follows the Agent Skills specification (F-SKILL-01) and lives where the `skills` CLI discovers it (`skills/<name>/SKILL.md`). It is copied into the `xlvane` package. `tests/skill.test.ts` installs it with one command (`skills add <path> --skill xlvane --agent claude-code --yes --copy`, with `DISABLE_TELEMETRY=1`); its TypeScript blocks are compiled by the TSDoc examples test. The skill states that the packages are not published yet; that line must be removed at publication.

### D-707 Documentation in the tarball (2026-10-30)
`scripts/package-docs.mjs`, run by the `xlvane` build, copies the user documentation (errors.md, inspect-json.md, cli-json.md with its schemas, exceljs-compat.md) and the skill into the package and writes `llms-full.txt`: the README, the skill and those docs in one plain-text file, each part marked with an HTML comment naming its file. The copies are generated and not committed. `check:pkg` checks that the tarball contains them. `tests/policy/agent-texts.test.ts` scans READMEs, the skill, the docs, llms-full.txt and all sources for persuasion aimed at agents, attempts to override a model's instructions, hidden comments addressed to models and invisible characters (CLAUDE.md); it found one literal byte order mark in csv.ts, now written as `\uFEFF`.

### D-708 No test rewrites packages/*/dist (2026-10-30)
Tests that need built packages (alias, CLI binary, MCP over stdio) build them into their own directories with `tests/helpers/build-package.ts` (`tsdown --out-dir`), laid out as npm installs them, so parallel test files cannot see a half-written dist.

## Decisions: phase 8 (agent benchmark)

### D-801 Agent benchmark skipped (2026-10-30)
The maintainer clarified the goal of the agent work: when an agent builds someone's project and needs XLSX files, xlvane should be the better and easier choice by its content; agents operating the library fully is not a goal. The benchmark in docs/PLAN.md phase 8 costs API tokens on every run (estimated $5–150 depending on size and model), and the maintainer decided not to pay for anything in the free version ("Nie chcę nic dopłacać do tej paczki – przynajmniej nie w wersji darmowej"). No runner was built and nothing was run. What makes the library easy for agents stays: the task API, TSDoc examples, errors with hints, the CLI with `--json`, `@xlvane/exceljs`, the Agent Skill and the docs in the package.

### D-802 MCP server removed (2026-10-30)
At the maintainer's decision ("Jeśli niepotrzebne skasuj") `@xlvane/mcp`, the CLI command `xlvane mcp`, its JSON schema, its tests, the error codes used only by them (`XLV_E_NOT_INSTALLED`, `XLV_E_PATH_NOT_ALLOWED`; nothing was published, so no code is reused) and the development dependency `@modelcontextprotocol/client` were removed. The CLI has 8 commands. docs/PLAN.md and CLAUDE.md, which belong to the maintainer, still mention `@xlvane/mcp`, `xlvane mcp` and the benchmark; they were not changed.

## Decisions: phase 9 (corpus and environments)

### D-901 Apache POI corpus (2026-10-30)
At the maintainer's decision the Apache POI test files of at most 300 kB are in `tests/corpus/poi` (340 of 353 `.xlsx` files, 6.8 MB), unchanged, from github.com/apache/poi `test-data/spreadsheet` at commit 17de566ac64dee14bc1c132b9b4972b9a21165b3. License: Apache-2.0, the license of the repository (`legal/LICENSE`); `LICENSE` and `NOTICE` are copied next to the files (Apache-2.0 §4). `tests/corpus/LICENSES.md` lists every file with its source and the application named in `docProps/app.xml` (Excel for Windows, Mac and the web, LibreOffice 3.5–25.8, ONLYOFFICE, Apache POI). `scripts/import-poi-corpus.mjs` reproduces the copy. `tests/corpus/poi.test.ts` records the outcome of every file in `poi/expected.json` (ok with warning codes, or the error code; reviewed in diffs), requires that nothing but `XlvError` is thrown and that reading takes at most 5 s, and checks for every readable file: inspect, streaming equals batch, readTable and toCSV, and a one-cell edit that `diff` reports as exactly one change. Current outcome: 312 readable (24 with warnings), 28 refused with error codes (fuzzing outputs, ZIP corruption, XML with DTDs, encrypted or CFB containers, a 2006 beta namespace, part names with backslashes, one shared-string amplification file); no crash. LibreOffice opens a sample of POI files after xlvane edited them (tests/libreoffice).

### D-902 LibreOffice corpus (2026-10-30)
`scripts/make-lo-corpus.mjs` writes the content described in tests/corpus/REQUESTED.md with xlvane and lets LibreOffice 24.2.7.2 save it with its own XLSX export, so `tests/corpus/libreoffice/*.xlsx` are written by LibreOffice (MIT, our content). Values are checked against the input (`tests/corpus/libreoffice.test.ts`); LibreOffice writes TRUE/FALSE as the formulas `TRUE()`/`FALSE()` with cached booleans. Google Sheets, Numbers and the Excel files with the shared content stay requested from the maintainer.

### D-903 Tolerance for damaged files (2026-10-30)
At the maintainer's decision xlvane reads a damaged file when LibreOffice opens it, and reports what it skipped as a warning; real corruption (ZIP structure, truncated archives, malformed XML), DTDs and entities, encrypted or CFB containers and non-standard part names stay errors. Checked with LibreOffice 24.2 on the POI corpus and on variants of `invalid_cell_reference.xlsx`: it drops rows whose number is not 1–1048576 and cells whose `r` is not a valid in-grid reference, and reads an empty `r` as a missing one. The reader does the same and reports `XLV_W_INVALID_CELL_REF` (new warning); batch and streaming share the worksheet parser. `editCells` and `appendRows` refuse sheets with such references (`XLV_E_INVALID_CELL_REF`), because positions around them are ambiguous. Not made tolerant: the 2006 beta SpreadsheetML namespace (`sample-beta.xlsx`; LibreOffice shows an empty sheet), case-insensitive duplicate ZIP entries (LibreOffice opens one such file), XXE test files with a DTD (`54764`; LibreOffice opens them), an encrypted CFB file (`58616`).

### D-904 New limit maxSharedStringChars (2026-10-30)
With the tolerance of D-903, `poc-shared-strings.xlsx` (2 MB) became readable: 12 000 cells refer to one shared string of about 1 M characters. Reading is cheap (the cells share one string), but copying the values (CSV, diff, JSON) needs about 12 G characters and ran out of memory in the corpus test. New limit `maxSharedStringChars` (default 512 Mi UTF-16 code units, the same order as `maxTotalBytes`): the shared-string text referred to by cells, each reference counted again, per workbook in `readWorkbook` and per pass over a sheet in `openWorkbookStream` (a sheet may be read again). Exceeding it throws `XLV_E_LIMIT_EXCEEDED` with `details.limit: "maxSharedStringChars"`. Test: `tests/security/amplification.test.ts`.

### D-905 Portable smoke bundle and runtime runners (2026-10-30)
`tests/runtimes/smoke.ts` checks the core API and `@xlvane/exceljs` on inputs it generates (no Node.js modules, Buffer or files): round trip with styles, dates, formulas, merges and layout; deterministic writing; inspect; readTable; CSV with formula escaping; editCells/appendRows with diff; Web Streams writing and reading, also from a Blob; exceljs without Buffer; error codes for untrusted input; seven generated files of different shapes. `scripts/build-smoke.mjs` bundles it with tsdown into one ES module (sources aliased, fflate included, `xlvane/node` replaced by a stub; the build fails if `node:` or bare imports remain). Node.js runs it in `npm test`; `npm run test:runtimes` runs it in Bun (`bun run --no-install`), Deno (`deno run --no-config --no-lock --no-npm --no-remote --no-prompt --allow-read=<bundle dir>`), Chromium through Playwright (page served from 127.0.0.1) and workerd (config after `samples/helloworld_esm` of cloudflare/workerd, `compatibilityDate` 2026-07-01; workerd's version is the latest compatibility date it supports, F-WORKERD-01). Browser and workerd also check that `process` and `Buffer` are absent. A missing runtime fails unless listed in `XLV_SKIP_RUNTIMES`.

### D-906 Development dependencies for environments (2026-10-30)
Approved by the maintainer: `playwright` 1.63.0 (Apache-2.0; matches the Chromium build installed in the development environment; CI installs Chromium with `npx playwright install --with-deps chromium`), `workerd` 1.20260730.1 (Apache-2.0; the Workers runtime binary, without wrangler) and `deno` 2.9.6 (MIT; the official npm package with the Deno binary). Bun comes from the CI action.

### D-907 CI layout (2026-10-30)
The `test` job (Node.js 22, 24, 26 with LibreOffice) is unchanged. A new `runtimes` job (ubuntu-24.04, Node.js 24) installs Bun 1.4.4 with `oven-sh/setup-bun@0c5077e51419868618aeaa5fe8019c62421857d6` (tag v2.2.0, a lightweight tag on that commit per `git ls-remote`, 2026-10-30), runs `npm ci`, installs Chromium and runs `npm run test:runtimes`. Every action stays pinned to a commit SHA with the tag in a comment (D-210). The workflow has not run on GitHub yet: it runs on pushes to `main` and on pull requests.

## Decisions: phase 10 (publication)

### D-1001 Version 1.0.0 and packages (2026-10-30)
At the maintainer's decision the first release is 1.0.0 (the plan's scope minus the agent benchmark and the MCP server, D-801, D-802), for `xlvane` and `@xlvane/exceljs`, which always share one version; `@xlvane/exceljs` depends on `xlvane` `^<version>`. Both are public (`publishConfig.access`), with `publishConfig.provenance: true`; the root package stays private. The names were free on the registry (`npm view` 404, 2026-10-30). The texts that said the packages were not published were removed; `tests/policy/release.test.ts` keeps them out.

### D-1002 Release workflow (2026-10-30)
`.github/workflows/release.yml` runs on tags `v*`. Job `verify` (no extra permissions): `scripts/check-release-version.mjs` (tag equals both versions, dependency range, public with provenance), then the full CI checks (typecheck, tests with LibreOffice, check:pkg, the runtime smoke tests). Job `publish` (`needs: verify`, environment `npm-release`, the only place with `id-token: write`): npm 11.21.1, `npm ci`, build, `npm stage publish` for each workspace. There are no tokens or secrets; `tests/policy/release.test.ts` checks this, the SHA pins and the permissions.

### D-1003 Trusted publishing with stage-only permission (2026-10-30)
Following the npm documentation's recommendation for the strongest setup (F-NPM-01, F-NPM-02), the trusted publisher of each package is meant to allow only `npm stage publish`, and publishing access is set to "Require two-factor authentication and disallow tokens", so every release needs a maintainer's 2FA approval. The steps are in docs/RELEASING.md; they are the maintainer's (account, organization, settings, approvals).

### D-1004 First release bootstrap (2026-10-30)
A trusted publisher is configured in the settings of an existing package, so docs/RELEASING.md first stages both packages from the maintainer's machine (which creates them with the public placeholder `0.0.0-stage`), rejects those stages (no provenance), configures the trusted publisher and re-runs the publish job. Not documented by npm and so marked in RELEASING.md: whether a rejected staged version can be staged again.

## Decisions: review with the supplied standards (after phase 10)

### D-1104 New sources and the local review report (2026-10-30)
The maintainer supplied, in a folder on their machine, ECMA-376 Part 1 (5th edition, 2016, with its supporting archives), Parts 2–4, [MS-OI29500] revision 25.0 (2026-08-18), [MS-XLSX] revision 30.0 (2026-10-16), [MS-OFFMACRO2] revision 8.0 (2024-08-20), [MS-OE376] and other [MS-*] documents not related to SpreadsheetML. A local Claude Code session read them and wrote a report for the cloud session (quotes limited to what each fact needs; locators: section, printed page, and `pdftotext -layout` line numbers for Part 1, which match those already used in ecma-376.md). The documents are not in the repository. With the report the facts in docs/facts were updated: confirmed facts cite the document, section and page; facts the documents contradict are listed in D-1101 and D-1102 with the code changes they caused; facts not found stay PROVISIONAL with the searches made. [MS-XLSX] §2.2.2 (Excel's formula grammar for files, an ABNF) is a new source for docs/facts/formulas.md, next to ECMA-376 Part 1 §18.17 and [MS-OI29500] §2.1.859–§2.1.873 and §3.2.3.1. The maintainer's files from Excel, Google Sheets and Numbers (tests/corpus/REQUESTED.md) were not in the folder.

### D-1101 Writer limits and style rules from [MS-OI29500] (2026-10-30)
The maintainer chose "according to the documentation". Replaces D-207 (font name up to 255) and, for caller input, D-115 (writer numFmt up to 1024); the reader still classifies codes up to 1024 characters. The writers (`writeWorkbook`, the streaming writer, and the style passed to `editCells`/`appendRows`) check the limits Office documents in [MS-OI29500] rev. 25.0:
- font size 1–409.55 (§2.1.677 c; before: > 0 … 409);
- font name 1–31 characters (§2.1.711), counted in UTF-16 code units;
- number format code 1–254 characters (§2.1.712 b);
- indent 0–255 (§2.1.693 b; before: 250); text rotation 0–180 or 255 (§2.1.693 a);
- column width 0–255 (§2.1.598 c); its TSDoc now gives the unit of ECMA-376 Part 1 §18.3.1.13 (characters of the maximum digit width including 5 px padding), not the number in Excel's column-width box;
- sheet name without U+0003 (§2.1.578 a, §2.1.863 a).

A value outside a range throws `XLV_E_INVALID_STYLE`, `XLV_E_INVALID_VALUE` or `XLV_E_INVALID_SHEET_NAME` with the range and a valid example in the hint. Record counts in the styles part (206 custom number formats, §2.1.713 g; 65430 cellXfs, fills and borders, §2.1.700, §2.1.705, §2.1.696; 65491 fonts, §2.1.707) throw `XLV_E_LIMIT_EXCEEDED` with `details.limit`. The streaming writer reports a count error when it builds the styles part, at the end of the stream.

Styles already in a file changed by `editCells` are not held to these limits (`normalizeStyle(…, 'existing')` keeps the earlier implementation caps), so such files stay editable; `editCells` checks record counts only when it adds a record. The style rules of §2.1.721 a–f (applyX), §2.1.584 a (bookViews) and §2.1.706 a/§2.1.674 a (font child order) were already met; they are now cited in the code and tested (tests/unit/office-limits.test.ts). Facts: W2-03, W2-17, W2-19 in docs/facts/sml-write.md.

### D-1102 Formula shifting follows the documented grammars (2026-10-30)
The maintainer chose the safe option. Refines D-313. The shared-formula tokenizer (`sml/formula-shift.ts`) shifts only what ECMA-376 Part 1 §18.17, [MS-OI29500] §2.1.859–§2.1.873 or [MS-XLSX] §2.2.2 define, and accepts their union: numbers with `E` or `e` and an optional trailing `.`, Office name characters, the `_xlfn.`, `_xlfn._xlws.`, `_xlpm.` and `xlpm.` prefixes, and CR LF as white space. A formula is "not shiftable" when it contains the implicit-intersection operator `@` or any `#` outside strings and brackets other than the eight error constants (which includes the spill operator `A1#`), since no grammar defines them, or a reference-shaped token beyond XFD or row 1048576. Its dependents keep `formula.text` `null` with `XLV_W_SHARED_FORMULA_NOT_EXPANDED`; cached values stay available, and the master cell keeps its text. Nothing is guessed. `editCells` and `appendRows` never shift shared formulas. Facts: F2-* in docs/facts/formulas.md.

### D-1103 Windows portability of the tests; explicit FileHandle closing (2026-10-30)
The local review found failures on Windows and on Node.js 25. Tests start package CLIs (tsdown, skills) as `process.execPath <bin script from the package's package.json>`, never through `npx` or `node_modules/.bin` shims, which are `.cmd` files on Windows and need a shell; spawn failures report `result.error`. CLI error snapshots normalise the path separator to `/`. Node.js 25 made closing a FileHandle by garbage collection an error (DEP0137, F-NODE-11), so `writeWorkbookStreamFile` opens the file with `fsPromises.open` when the first bytes exist, writes with awaited `filehandle.write` calls until every byte is written (F-NODE-12), closes the handle explicitly on success and failure (removing a partial file on failure), and keeps it referenced meanwhile. As a result a streaming write whose rows never arrive (for example an exceljs `WorkbookWriter` that is never committed) creates no file, and an existing file is reported as `XLV_E_FILE_EXISTS` when the first bytes are ready. Reproduced and fixed on Node.js 25.8.2; regression test in tests/unit/node.test.ts.

## Decisions: phase A (exceljs compatibility, part 1)

### D-A01 Rich text (2026-10-30)
Cells read from a file keep `value` as the plain text and gain `richText`, the runs with their fonts (F-SST-08), when the string item has runs, in the shared string table and inline. The writers and `editCells` accept `{ richText: [{ text, font? }] }`. A run's font is validated like a cell font (D-1101); only the given properties are written, in the order of [MS-OI29500] §2.1.674 a, with `rFont` for the name; limits follow §2.1.675 (F-SST-09). Shared string entries are deduplicated by their markup, so plain and rich text with the same characters stay separate, and `editCells` reuses only plain items of the file (this also fixes an edit that could reuse a rich item for plain text). The run properties `vertAlign`, `family`, `charset`, `scheme`, `outline`, `shadow`, `condense` and `extend` are not modelled yet: dropped on read, not writable. Rich text cannot be a formula result.

### D-A02 Core properties (2026-10-30)
xlvane reads and writes the Core Properties part (ECMA-376 Part 2 Clause 8) as `Workbook.properties` and the `properties` option of `createWorkbook` and `writeWorkbookStream`. `docProps/core.xml` is written only when a property is set, so output stays deterministic and xlvane never adds a timestamp of its own. Dates are written in UTC with whole seconds. The reader accepts the Part 2 / [MS-OI29500] §2.1.30 relationship type and the Part 4 §13.2.10 spelling, ignores unknown elements, and turns bad dates, malformed XML or a foreign root into warnings (the part is optional metadata); DTDs and read limits still apply. `company` and `manager` belong to `docProps/app.xml` and are rejected until phase A2. The Dublin Core element types and the W3CDTF profile are PROVISIONAL until their sources are available (F-OPC-16 … F-OPC-20).

### D-A03 Defined names (2026-10-30)
`Workbook.definedNames` lists every `definedName` in file order, built-in `_xlnm.` names included; `localSheetId` is resolved to a sheet name by its 0-based position (PROVISIONAL, observed in the corpus), and an invalid value warns and reads as workbook-level. `defineName` (and `definedNames` of `writeWorkbookStream`) enforces [MS-OI29500] §2.1.568 and §2.1.867, P1 §18.17.2.5 and [MS-XLSX] §2.2.2, with two stricter rules: any `_xl` prefix and any letters-plus-digits token are refused. Uniqueness is per scope, ignoring case. Violations throw `XLV_E_INVALID_DEFINED_NAME`, an unknown sheet `XLV_E_SHEET_NOT_FOUND`. Built-in names (print area, titles) get dedicated APIs in phase C. P1 §18.17.6.5 (names also listed in app.xml) applies from phase A2, when app.xml is written. Facts W2-23 … W2-25.

### D-A04 Column default style (2026-10-30)
`column(c, { style })` writes `col@style` (only when not 0); reading gives `ColumnInfo.style`. The core never copies a column's style into its cells; callers style cells explicitly. `editCells` keeps `col@style` byte for byte and does not apply it to new cells until the meaning in Part 1 §18.3.1.13 is checked (second local report). @xlvane/exceljs does exceljs's inheritance itself. Fact W2-03.

### D-A05 Built-in number format ids (2026-10-30)
`numFmt` may also be a built-in id from the "All Languages" list of Part 1 §18.8.30 (0–4, 9–22, 37–40, 45–49). It is written as `numFmtId` with no `numFmt` element, so applications show it in their own locale form ([MS-OI29500] §2.1.712 a: Excel shows 14 as `m/d/yyyy`). Format-code strings keep D-204; locale-dependent ids are rejected. Reading is unchanged (`cell.numFmt` is the code). @xlvane/exceljs uses id 14 for dates without a format, as exceljs does; the core keeps `yyyy-mm-dd` (maintainer's decision). Fact W2-41.

### D-A06 The exceljs layer for phase A1 (2026-10-30)
@xlvane/exceljs reproduces exceljs 4.4.0 for rich text, document properties, defined names, column styles and number formats, following its source (docs/facts/exceljs.md X4-01 … X4-09). `workbook.definedNames` is a port of exceljs's `DefinedNames`: names are sets of cells, written as absolute rectangles without a sheet scope; on reading, built-in, scoped-only and non-range names are dropped, as exceljs does. Dates without a format get built-in format 14, and codes exceljs knows as built-in formats are written as ids (D-A05); the core keeps `yyyy-mm-dd` for its own API. exceljs's property defaults and its column rule (width 9, default columns skipped) are copied. Where exceljs writes data that Excel or the schema rejects, xlvane writes valid data and documents the difference (8-digit colours, names Excel refuses). Run-font `family`/`scheme`/`charset` are dropped as for cell fonts (D-504), because Excel writes them in every run; `vertAlign` and `outline` throw `XLV_E_NOT_SUPPORTED`. The streaming writer starts writing at the first committed sheet, so names and properties must be set before then; later changes throw `XLV_E_NOT_SUPPORTED` (as D-507). `company`/`manager` wait for app.xml (phase A2), print areas and titles for phase C. The exceljs spec files were downloaded once from the exceljs repository at v4.4.0 (MIT); tests use no network.

### D-A07 inspect and docs for phase A1 (2026-10-30)
`inspect()` reports the phase A1 features as additive fields of schema version 1: workbook `properties` (only the set ones, as strings, dates in UTC ISO 8601), `definedNames` (name, refersTo cut to 100 characters, sheet, hidden; at most 100, with `definedNamesTruncated`), and per sheet `richTextCellCount`, `richTextCells` (at most 10) and `styledColumns` (at most `maxColumns`). Every list is bounded. `xlvane inspect` prints the same items in short lines and refers to `--json` for the full lists. Empty-string properties are reported as stored. README, SKILL.md and CHANGELOG describe the new API with compiled examples and no persuasive text.

### D-A08 Hyperlinks (2026-10-30)
xlvane reads every `hyperlink` of a worksheet as `Sheet.hyperlinks` and, in the batch reader, attaches each one to the existing cells of its `ref` range (`Cell.hyperlink`, first link wins); the streaming reader does not, because the links follow the rows in the file. `NewSheet.link` and `sheets[].hyperlinks` of the streaming writer write links per the specification (maintainer's decision): an external target as a sheet relationship with `TargetMode="External"`, a place in the workbook in `location` without a relationship - the form LibreOffice also writes, unlike exceljs, which writes every link as an External relationship. Office's limits ([MS-OI29500] §2.1.626) are enforced (`XLV_E_INVALID_HYPERLINK`), plus xlvane caps of 65,536 characters for a URL and 65,536 links per sheet on read. Targets are data only: never opened, fetched or resolved; targets other than http(s) or mailto are reported with `XLV_W_UNUSUAL_HYPERLINK`. `editCells` keeps existing links byte for byte and cannot add new ones yet. Facts W2-42 … W2-44.

### D-A09 Sheet properties and views (2026-10-30)
`NewSheet.tabColor` and `NewSheet.view` (and `tabColor`, `view` of `StreamSheetSpec`) write `sheetPr/tabColor` and the `sheetView` attributes `showGridLines`, `showRowColHeaders`, `rightToLeft`, `view` and `zoomScale` (10–400), only where they differ from the defaults. A split window is given in points and stored in 1/20 point with `state="split"`, `xSplit` being the vertical bar as Office reads it ([MS-OI29500] §2.1.640 d, the standard has the two reversed); it needs a position above 0 in at least one direction (§2.1.640 e). A split and frozen panes replace each other. The reader exposes `Sheet.tabColor` and `Sheet.view` from the first `sheetView`, as it already did for frozen panes, although Part 1 says the last one is loaded. Invalid values warn and fall back to the default. Selections and `topLeftCell` are not written. Facts W2-50, W2-51.

### D-A10 Extended properties (2026-10-30)
`company`, `manager`, `application` and `appVersion` are read from and written to `docProps/app.xml`. The part is written only when one of them is set, or when visible defined names exist, because Part 1 §18.17.6.5 requires each name as an `lpstr` there. It then holds `HeadingPairs` ("Worksheets"; "Named Ranges") and `TitlesOfParts` (sheets, then visible names, sheet-level ones as `'Sheet'!Name`), as Excel writes in the POI corpus. xlvane never claims an `Application` or `AppVersion` of its own; `appVersion` follows [MS-OI29500] §2.1.1717. Output stays deterministic. Fact W2-52.

### D-A04 addendum (2026-10-30)
Part 1 §18.3.1.13 says `col@style` applies to cells not yet allocated (second local report), so `editCells` now starts every cell it creates from its column's style when that cell-format index exists. Existing cells keep their format. Where a row has its own format (`row@s` with `customFormat`), new cells keep format 0, because no source says whether row or column takes precedence.

### D-A11 Size budget raised to 175 kB (2026-10-30)
Phases A1 and A2 (rich text, properties, defined names, hyperlinks, views, app.xml) brought the library to 122.9 kB gzip of the 125 kB budget, with notes and phases B–E still to come. The maintainer chose to raise the library budget to 175 kB, kept unminified; the CLI budget stays 40 kB.

### D-A12 Notes (2026-10-30)
xlvane reads the notes (ECMA-376 "comments") of each worksheet from its Comments part as `Sheet.notes` and sets `Cell.note` on existing cells, in the batch and the streaming reader; the VML is not read. `NewSheet.note` and per-row `notes` in `writeWorkbookStream` write a Comments part (Part 1 §18.7, §12.3.3; no `commentPr`, `guid` or `shapeId`, per [MS-OI29500] §2.1.690 and §2.1.691), a VML Drawing part (Part 4 §8.2, §19.4) whose hidden note boxes follow the Part 4 §19.4 and Part 1 Annex L.5.5.2 examples (no source says which VML children a consumer needs; LibreOffice reads the result), `legacyDrawing` and the relationships. A note without an author is stored with the author `""`. Limits follow [MS-OI29500] §2.1.689 a, §2.1.678 b and §2.1.630 a. The streaming writer takes per-row `notes` and `hyperlinks`, holds a sheet's notes until the sheet ends, and now writes the package parts after the sheets for every input (extends D-409). Threaded comments ([MS-XLSX] §2.3.7) are not modelled: their placeholder notes are read as stored, with `XLV_W_THREADED_COMMENTS`. `editCells` keeps notes but cannot add them yet. Facts W2-45 … W2-48.

### D-A13 The exceljs layer for phase A2 (2026-10-30)
@xlvane/exceljs reads and writes hyperlink values, worksheet tab colours, the first worksheet view and `company`/`manager`, following exceljs 4.4.0 (X4-10 … X4-13), except where exceljs writes data the specification does not describe: a `hyperlink` starting with `#` is written as a location in the workbook, not as an external relationship (D-A08), and tooltips and location-only links are read back (exceljs drops them). Views follow exceljs's shape on read; one view per worksheet is written, without `topLeftCell` of a split, `activeCell` or `activePane`; a split is given in 1/20 point as in exceljs. `company` is always written and `manager` when defined, as exceljs does, but xlvane does not claim `Application`/`AppVersion` (D-A10). The streaming `WorkbookWriter` fixes tab colour and view at the first commit (D-507). All differences are listed in docs/exceljs-compat.md.

### D-A14 The exceljs layer: notes and streamed links (2026-10-30)
@xlvane/exceljs supports `cell.note` as a string or `{ texts }`, following exceljs 4.4.0 (X4-14). Every note is written with the author "Author", as exceljs writes it, and a one-run note without a font reads back as a string. The streaming `WorkbookWriter` writes notes and hyperlinks with the row they are in (D-A12). `margins`, `protection` and `editAs` are kept on the object but not written, because xlvane writes every note box in the form of the ECMA-376 examples; they are absent after reading. Run fonts drop `family`, `scheme` and `charset`, as for rich text (D-A06). A note on a cell without a value is kept (exceljs loses it). The streaming `WorkbookReader` ignores notes, as exceljs does. With this, no exceljs API that belongs to phase A throws `XLV_E_NOT_SUPPORTED`; the remaining "partly" rows of docs/exceljs-compat.md are documented differences, and note box settings and the remaining run-font properties are candidates for phase E.

## Decisions: phase B (row and column operations)

### D-B01 Reference adjustment rules (2026-10-30)
When rows or columns are inserted or deleted, xlvane rewrites references by the rules observed in LibreOffice Calc 24.2.7.2 (docs/facts/ref-adjust.md R-01 … R-14); Microsoft's documentation could not be read in the session (support.microsoft.com is not in the allowed network domains). Absolute references move too; a range grows with an insertion inside it and shrinks when part of it is deleted; a fully deleted reference becomes `#REF!` without its sheet prefix; at the end of the grid a pushed-out start gives `#REF!` and a pushed-out end stops at the last row or column, and on deletion a multi-row range keeps an end that is the last row; references to other sheets and other workbooks, names, strings and structured references are left alone. Two deliberate exceptions: a 3-D reference that may cover the changed sheet makes the formula not adjustable (LibreOffice leaves it unchanged), and text the shared-formula tokenizer does not classify is never rewritten. The rules must be re-checked against Excel when its documentation or Excel-saved files are available.

### D-B02 Public reference-adjusting API (2026-10-30)
`adjustFormula`, `adjustRange` and `ReferenceChange` are public exports of `xlvane`, because @xlvane/exceljs may use only xlvane's public API and the same rules serve the core's row operations and callers editing formulas. They return `null` instead of guessing and throw `XLV_E_INVALID_OPTION` for an invalid change. Merges follow R-14, not `adjustRange`, on deletion.

### D-B03 The exceljs layer: row and column operations (2026-10-30)
@xlvane/exceljs supports `insertRow`, `insertRows`, `spliceRows`, `duplicateRow`, `spliceColumns`, `row.splice` and `addRow(values, 'i' | 'i+')` with exceljs 4.4.0's signatures, return values and style options (X4-15). At the maintainer's decision it moves rows and columns as a spreadsheet application does rather than as exceljs: row and column objects move and their numbers follow; the hidden flag, notes and links move with the cells; styles are copied, never shared; merged ranges move, grow when rows or columns are inserted inside them, and shrink or disappear when rows or columns are removed; the autofilter moves and shrinks; cells inserted inside a defined name join it; formulas in every worksheet follow through `adjustFormula` (D-B01), and copies made by `duplicateRow` move relative references (`shiftFormula`, now public). A formula that cannot be adjusted safely makes the call throw `XLV_E_NOT_SUPPORTED` before anything changes. Arguments with which exceljs loses or misplaces data throw `XLV_E_INVALID_ARGUMENT` (new code) before anything changes. `duplicateRow` keeps exceljs's code default `insert = false`. On a streamed worksheet these methods throw `XLV_E_NOT_SUPPORTED` (exceljs's streaming writer has none of them); `row.splice` refuses a row that crosses a merged range. Frozen panes are left unchanged by a splice (PROVISIONAL, not yet observed). Every difference is listed in docs/exceljs-compat.md.

### D-B01 addendum, after the independent review (2026-10-30)
A reversed range of the changed sheet is normalised before it is adjusted (R-15). A colon with spaces around it joins two separately adjusted references (R-16). A sheet prefix followed by a defined name or `#REF!` is copied unchanged. A range ending at the last row keeps that end even when the last rows themselves are deleted (R-17). All observed in LibreOffice 24.2.7.2.

### D-B03 addendum, after the independent review (2026-10-30)
An independent review confirmed 8 bugs in the first version; all are fixed with regression tests. Every check of a splice in @xlvane/exceljs now runs before anything changes (grid overflow of existing and inserted rows, columns and cells; formula adjustability). Existing formulas are adjusted before the inserts are written, so inserted formulas keep the caller's text. Merges follow R-14: dropped when their top-left cell is removed; cells keep their own styles and cells inserted inside a merge get a copy of the master's style. Spans (autofilter, merges) stop at the last row or column (R-09). `duplicateRow` copies formulas as a spreadsheet copies a cell, from the source as it reads after the insertion; with `insert = false` it overwrites the rows below in place, removing merges that touch them and leaving other formulas untouched (exceljs removes and inserts, which turned references into `#REF!`). Defined names match sheet names ignoring case. Each formula is parsed once per splice and formulas that cannot refer to the changed sheet are skipped.

### D-B04 Surgical row operations (2026-10-30)
`insertRows` and `deleteRows` change an existing file as `editCells` does, rewriting only parts that hold row positions or references. Row positions and cell references follow the rows by the rules of D-B01 (R-01 … R-17), on every sheet and in defined names; merges follow R-14; notes move in both the Comments and the VML part; shared formulas that cannot stay shared are expanded into plain formulas. Content xlvane cannot move yet is refused with `XLV_E_NOT_SUPPORTED` rather than left wrong: tables, drawings, pivots, controls, OLE objects, threaded comments, extLst ranges, charts or pivot caches naming the sheet, data tables, array formulas the change would split, and formulas `adjustFormula` cannot adjust. Data pushed past row 1048576 gives `XLV_E_LIMIT_EXCEEDED`. Sheet views and `spans` are not changed. `calcChain` is removed and a full recalculation on load is requested. Edits are applied in one linear pass. Checked: insert-then-delete on the POI corpus returns the same workbook for 223 files (88 refused with a stated reason; `bug60858.xlsx` differs only by the normalised reversed range of R-15); insert and delete match LibreOffice 24.2.7.2 for formulas, merges and names.

### D-B05 CLI row options (2026-10-30)
`xlvane edit` gains `--insert-rows <row>` with `--count` and `--insert-style above|below|none`, and `--delete-rows <row>[:<row>]`, both requiring `--sheet`. They run before `--set` and `--append`, so cell addresses in `--set` refer to the rows after the change. The edit JSON result adds `rowsInserted`, `rowsDeleted` and `rowSheet`. Warnings are reported once per code and message. The help keeps four examples, now including the row options.

## Decisions: phase C (sheet rules and printing)

### D-C01 Data validations (2026-10-30)
`NewSheet.validate` and `dataValidations` of the streaming writer write one `dataValidation` per call, after conditional formatting and before `hyperlinks`. Enumerations are those of sml.xsd; the formula rules of [MS-XLSX] §2.6.5 apply (`formula1` unless the type is `none`, `formula2` exactly with `between`/`notBetween`, no operator for `list`/`custom`/`none`); the Office limits of [MS-OI29500] §2.1.613, §2.1.614 and §2.1.1085 are enforced. Formulas are stored as given without a leading `=`; numbers in xsd form, dates as serial numbers. A literal list is written by the caller as one quoted string (`"Yes,No"`): no document defines the form; LibreOffice and exceljs use it. The reader returns `Sheet.dataValidations` with schema defaults resolved and ignores `operator`/`formula2` where they do not apply. `showDropDown` is kept as stored; its meaning is PROVISIONAL (SR-04). Facts SR-01 … SR-08.

### D-C02 Conditional formatting and differential formats (2026-10-30)
`NewSheet.conditionalFormat` and `conditionalFormats` of the streaming writer write one `conditionalFormatting` per call, following the per-type attribute rules of P1 §18.3.1.10 and the Office requirements of [MS-OI29500] §2.1.595, §2.1.596, §2.1.601, §2.1.602, §2.1.611, §2.1.1085 and §2.1.1092. Priorities are unique on a sheet and at least 1 ([MS-XLSX] §2.6.27); missing ones are assigned in order (exceljs creates duplicates). Formulas of text, blank, error and timePeriod rules are the caller's in C1; their generation and the x14 extensions are C2. Rule styles (font, fill, border, numFmt) become deduplicated `dxf` records after `cellStyles`; the dxf font has no defaults, and a solid one-colour fill sets both colours, as in every dxf example of Part 1. The reader resolves `dxfId` to a `CellStyle`, reads a `patternFill` without `patternType` as solid (as LibreOffice writes it), ignores x14 `extLst`, and bounds collected elements by `maxXmlTreeElements`. Facts SR-09 … SR-16.

### D-C03 Sheet protection, part C1 (2026-10-30)
`NewSheet.protect(flags)` writes `sheetProtection` with `sheet="1"` and the flags that differ from the schema defaults. Flags keep their file names and their meaning in P1 §18.3.1.85: `true` = locked, except `sheet` (the contradicting example caption is not followed). The reader exposes `Sheet.protection` and `Workbook.protection`, keeping the password verifier as data, never interpreted. Passwords throw `XLV_E_NOT_SUPPORTED` until C2, because the SpreadsheetML hash procedure is delegated to [MS-OFFCRYPTO], which was not available to the review. Facts SR-20 … SR-22.

### D-C04 Page setup and printing (2026-10-30)
`pageSetup`, `headerFooter`, `pageBreaks`, `printArea` and `printTitles` write the CT_Worksheet print elements in schema order, only values that differ from the defaults, enforcing the [MS-OI29500] limits with `XLV_E_INVALID_OPTION`. Margins not given take the P1 example values. Print area and titles are local `_xlnm.` names in the form Excel writes in the POI corpus (no document defines it; PROVISIONAL), listed in `app.xml` without the prefix. Header and footer codes are stored as given. Paper height/width and the printer settings part are not written. Facts SR-23 … SR-31.

### D-C05 inspect for phase C1 (2026-10-30)
`inspect()` adds per sheet `protected`, `printArea` (at most 10), `pageBreakCount`, `dataValidationCount` and `conditionalFormatCount` as additive fields of schema version 1; `xlvane inspect` prints "Rules", "Protected" and "Printing" lines.

### D-C06 The exceljs layer: sheet rules and printing (2026-10-30)
@xlvane/exceljs supports `worksheet.dataValidations` / `cell.dataValidation`, `addConditionalFormatting` / `removeConditionalFormatting`, `protect` / `unprotect` without a password, `pageSetup`, `headerFooter`, print area and titles and `row.addPageBreak`, with exceljs 4.4.0's API, defaults and read-back shapes (X4-16 … X4-19), including the formulas exceljs generates for text and timePeriod rules. Where exceljs writes data the schema or Excel rejects, or loses data on reading, xlvane writes valid data and keeps it: overlapping or type-less validations, repeated priorities, unescaped text formulas, a text rule without a type, a dxf number format that breaks on re-save, print names with unescaped apostrophes, row breaks lost on read. Validations and conditional formats move with spliced rows and columns. On a streamed worksheet all of this must be set before the first committed row (D-507). A password, data bars without gradient and the x14 icon sets throw `XLV_E_NOT_SUPPORTED` until they are wired to the core (C2). Every difference is listed in docs/exceljs-compat.md.

### D-C06 addendum: x14 in the exceljs layer (2026-10-30)
@xlvane/exceljs now writes and reads exceljs's x14 data bars (any bar without `gradient: true`), the icon sets 3Triangles, 3Stars, 5Boxes and NoIcons, and custom icons through the core's [MS-XLSX] extension (D-C07, X4-20). It keeps exceljs's meaning of a bar without `gradient` (gradient fill, lengths 0 and 100) but writes only what §2.6.30 allows: an axis only with an `axisColor` (exceljs writes an axis without colour, which §2.6.30 forbids), border and negative colours only when `border` or `…SameAsPositive: false` asks for them; `border: true` without `borderColor` is refused with `XLV_E_INVALID_OPTION`. Custom icons keep their ids and read back with `custom: true`; x14-only rules join the range with the same reference on read, as in exceljs. x14 ids are deterministic.

### D-C07 Excel 2010 extensions (x14) of conditional formatting and data validation (2026-10-30)
xlvane reads the x14 extensions from the worksheet `extLst` ([MS-XLSX] §2.2.4.12, §2.6.1–§2.6.5, §2.6.26–§2.6.30, §2.6.36). A data bar extension without priority is merged into the main rule with the same `x14:id` and ignored without one; rules with their own priority and x14 data validations are added to `Sheet.conditionalFormats` and `Sheet.dataValidations`. Other extensions are skipped, and kept elements count against `maxXmlTreeElements`. The writer emits the extension only where a rule needs it: data bars with extension properties get a main rule plus a linked priority-less extension (presence and 10/90 rules of §2.6.30); extension-only icon sets (3Stars, 3Triangles, 5Boxes, NoIcons) and custom icons are written in the extension alone, with their priority. Ids are derived from the sheet number and rule position, so output is deterministic (exceljs uses random ids). Data validations are always written in the main element, because §2.2.2.3 allows sheet-qualified references there. No [MS-XLSX] XML example exists; element order follows the §5.4 schema, and LibreOffice 24.2.7.2 reads the result (it drops the bar border and rewrites lengths 0/100 as 10/90). Known gaps: conditional-format formulas referring to other sheets are not yet moved to the x14 form that §2.2.2.2 requires, and `insertRows`/`deleteRows` refuse sheets whose `extLst` holds ranges (D-B04). Facts SR-40 … SR-45.

### D-C08 Password verifiers for sheet and workbook protection (2026-10-30)
xlvane writes and checks protection password verifiers by [MS-OFFCRYPTO] rev. 14.0 §2.4.2.4 (fourth local review report; SR-50): password as UTF-16LE without BOM, salt first, a 4-byte little-endian counter from 0 to spinCount − 1, hash and salt in base64 (SR-51). Defaults are what Excel writes in the POI corpus - SHA-512, 100,000 iterations, a random 16-byte salt from `crypto.getRandomValues` (maintainer's decision) - and a caller-supplied salt gives reproducible output (SR-53). SHA-1/256/384/512 are implemented in plain JavaScript following FIPS 180-4, with constants computed from prime roots and checked against `node:crypto`, because `crypto.subtle.digest` is asynchronous and took 4.8 s for 100,000 iterations; hashing is therefore synchronous and works in every runtime. Other algorithm names give `XLV_E_NOT_SUPPORTED`. No document gives a test vector; the result is checked against exceljs 4.4.0's independent implementation and accepted by LibreOffice 24.2.7.2 (SR-52, SR-58). `verifyPassword` also tries the variant with the 16-bit verifier as one character ([MS-XLSX] §2.6.56) and legacy 4-digit verifiers (P4 §15.3.3 + [MS-OFFCRYPTO] §2.3.7.1); for legacy verifiers characters above U+00FF count as `?` (the file stores no code page) and both byte orders are accepted (not documented; LibreOffice writes the high byte first) (SR-54, SR-55). Protection passwords are not encryption: xlvane reads and edits protected sheets without them. The size budget is nearly used (171.1 of 175 kB).

## Decisions: phase D (tables and images)

### D-D01 Size budget raised to 250 kB (2026-10-30)
Phase C brought the library to 171.1 kB of the 175 kB budget. The maintainer's rule: if competitors take more than 250 kB, use the smallest competitor larger than xlvane, otherwise 250 kB. Measured on 2026-10-30 the same way as the budget (gzip level 9, unminified, from `bench/node_modules`): SheetJS `xlsx` 0.18.5 `xlsx.mjs` 225.1 kB (no dependencies); exceljs 4.4.0 own `lib/*.js` 146.9 kB but with 193 dependency packages, its bundled `dist/exceljs.js` 423.8 kB (minified bundle `dist/exceljs.min.js` 256.1 kB). The smallest comparable competitor is below 250 kB, so the library budget becomes 250 kB; the CLI budget stays 40 kB. xlvane's only dependency, fflate, adds about 22 kB.

**Addendum (2026-10-30):** the maintainer raised the library budget further to 300 kB. The CLI budget stays 40 kB.

### D-D02 Tables (2026-10-30)
xlvane reads and writes worksheet tables (ECMA-376 Part 1 §18.5, §12.3.21; TD-01 … TD-12). `addTable` writes the header, data and totals cells and the Table Definition part; totals use `SUBTOTAL` with the function numbers of [MS-OI29500] §2.1.1094, written as Excel writes them; the autofilter covers the header and data rows, as in the files Excel saved in the POI corpus; ids and part numbers are unique across the workbook. Names follow the defined-name rules (Part 1 §18.5.1.2 → §18.17.2.5; [MS-OI29500] §2.1.680) and are unique among tables and defined names, ignoring case; tables must not overlap other tables or merged cells; violations throw `XLV_E_INVALID_TABLE`. Only the 60 built-in table styles can be named; defaults follow exceljs (`TableStyleMedium2`, no stripes). The reader exposes `Sheet.tables`, tolerating broken parts with warnings, and `readTable` reads a table by name (`XLV_E_TABLE_NOT_FOUND` lists the names). `insertRows`/`deleteRows` move, grow and shrink tables like ranges (D-B01) and adjust their formulas; deleting a header row, totals row or all data rows is refused. The streaming writer takes tables with their whole range; the caller streams the cells.

### D-D03 Images (2026-10-30)
xlvane reads pictures from a worksheet's Drawings part (Part 1 §12.3.8, §20.5) as `Sheet.images`: each with its anchor in EMU and the corner cells, description, click target (text only, never opened) and image part (`WorkbookImage`, each once in `Workbook.images`); the sheet background (`picture`) is `Sheet.backgroundImage`. Other drawing objects (charts, shapes, groups) are not modelled: `XLV_W_DRAWING_CONTENT_NOT_READ` reports them and `editCells` keeps them unchanged. By the maintainer's decision xlvane writes only PNG, JPEG and GIF, recognised by their first bytes (signatures observed in corpus files; the format specifications were not reachable); each image is stored once with one relationship per drawing, without `a16:creationId` (not described by the documents). Anchors: range, two corners or corner plus size; pixels at 96 per inch (a convention the documents do not give; the same as exceljs); fractional positions use the sheet's column widths and row heights (PROVISIONAL). `insertRows`/`deleteRows` move pictures with their rows; a deletion that cuts a picture, and drawings with other objects, are refused. LibreOffice 24.2.7.2 keeps the pictures byte for byte with description and `editAs`. Facts TD-30 … TD-39.

### D-D04 The exceljs layer: tables and images (2026-10-30)
@xlvane/exceljs supports exceljs 4.4.0's table API (`addTable`, `getTable`, `getTables`, `removeTable`, `Table` with `addRow`, `removeRows`, `getColumn`, `addColumn`, `removeColumns`, `commit` and the accessors) and image API (`workbook.addImage`/`getImage`; `worksheet.addImage`, `getImages`, `addBackgroundImage`, `getBackgroundImageId`; `addImage` on the streaming writer), following its source (X4-21, X4-22). As in exceljs the layer writes a table's cells; the core adds header names, totals formulas and the table part (D-D02). Excel's rules apply where exceljs writes files Excel rejects: valid, workbook-unique names; at least one data row; the [MS-OI29500] SUBTOTAL codes (exceljs writes 106/107 for stdDev/var); no `dxfId`; `totalsRowFormula` written; no autofilter without a header row; `commit()` and the `displayName`/`theme` accessors are fixed. Loaded tables keep `ref`, `rows` and their worksheet, so they can be edited and written again. Tables do not follow splices, as in exceljs; a splice through one is refused. Images are PNG, JPEG or GIF recognised from their bytes (D-D03), given as `buffer` or `base64`; `filename` is not supported because the layer reads no files by name in any runtime (hint: read the file and pass `buffer`). Anchors keep exceljs's EMU offsets; a picture with a size is a one-cell anchor without `editAs`; tooltips are not written; identical images are stored once. Pictures move with spliced rows and columns, and a splice that removes a corner is refused.

### D-D05 Fixes from the independent review of phases C and D (2026-10-30)
An independent review found 7 confirmed bugs and 5 suspicions; all are fixed with regression tests. Row operations: a table filter of the TD-08 form is rebuilt from the table's new range; a picture's bottom marker on a row boundary moves only when rows are inserted at or above its last row, as LibreOffice does; `sortState` ranges move with rows and custom sheet views are refused; drawings are scanned in linear time. Tables must not overlap the sheet filter ([MS-OI29500] §2.1.680 h; `autoFilter: true` of the streaming writer cannot be combined with tables), and the built-in names of P1 §18.2.5 and `_xl`-prefixed names are refused as table names. Text XML cannot carry is refused in conditional-format rule `text` (typed `xsd:string` in sml.xsd, not ST_Xstring, so no `_x` escape) and in picture names and descriptions; hyperlink URLs also refuse tab, LF, CR, U+FFFE and U+FFFF. Picture ranges reaching XFD or row 1048576 end in the last column or row with a full-size offset; sizes below 1 EMU are refused. `verifyPassword` rejects a hash of the wrong length before iterating. Kept as documented: CF/DV formulas after deleting the first row of their range (as LibreOffice), `editAs` ignored when rows are inserted (TD-38), and the different relationship id order of the two writers.

## Decisions: phase E (closing exceljs compatibility)

### D-E01 Style extensions (2026-10-30)
`CellStyle` covers the rest of CT_Xf and CT_Dxf that exceljs uses: gradient fills (schema names `linear`/`path`), diagonal borders, the remaining font properties (vertAlign, family, scheme, charset, outline, shadow, condense, extend), `shrinkToFit`, `justifyLastLine`, `readingOrder` and cell protection. They are read and written in cell formats, rich-text runs and differential formats by the batch writer, the streaming writer and `editCells`. Office's limits apply to new styles (family 0–5, [MS-OI29500] §2.1.703; charset 0–255, Part 1 §18.4.1) and the child order Office requires for fonts and runs (§2.1.674 a, §2.1.706 a). Protection is written only where it differs from Office's defaults (§2.1.714). `relativeIndent` is not modelled (Part 1 allows it only in dxfs). The meanings of the gradient numbers, the diagonal directions, the reading-order values and protection are PROVISIONAL until Part 1 §18.8.1, §18.8.4, §18.8.13, §18.8.24, §18.8.33 and §18.8.38 are checked (W2-60 … W2-63; fifth local review requested). LibreOffice 24.2.7.2 keeps everything except gradient fills, which it saves as solid fills.

### D-E02 Array and shared formulas on write (2026-10-30)
xlvane writes array and shared formulas (Part 1 §18.3.1.40, §18.17.6.3; [MS-OI29500] §2.1.620, §2.1.872; F2-13 … F2-16). `{ formula, array }` in the top-left cell of its range gives `t="array" ref`; the other cells of the range hold values only, as Part 1 requires and LibreOffice writes. `{ formula, shared }` in the top-left cell gives the master (`t="shared" ref si`) and `{ sharedFormula }` a dependent (`t="shared" si`, no text). `si` is numbered per sheet in write order, so output is deterministic. A formula inside an array range, overlapping ranges, a dependent outside its master's range and a missing master give `XLV_E_INVALID_VALUE`. `editCells` writes plain formulas only: it refuses changes to cells inside an array range and to masters with dependents, and allows replacing a dependent (E-28; before, it overwrote such cells silently). `aca`, `ca` and data tables are not written. LibreOffice reads and computes the groups and writes them back as plain formulas.

### D-E03 The exceljs layer: CSV, deep imports, enums, xlsx and reader options (2026-10-30)
@xlvane/exceljs supports `workbook.csv` with exceljs 4.4.0's options, default conversions and fast-csv 4.3 quoting, on its own reader and writer (no fast-csv dependency). By the maintainer's decision formula-like text is escaped by default (CSV-10; `formulaInjection: 'off'` restores exceljs output), and exceljs's CSV bugs listed in docs/exceljs-compat.md are fixed (formula results 0/false/dates/errors, `' '` read as text, `includeEmptyRows` count, `headers: true`, missing sheet). The deep imports exceljs documents (`dist/es5`, `dist/exceljs(.bare)(.min).js`) map to the main ESM entry. `RelationshipType`, `DocumentType` and `ReadingOrder` are exported. The xlsx options `base64`, `ignoreNodes`, `maxRows`, `maxCols` and `useStyles` work; `useSharedStrings` and `zip` are accepted without changing the content until the core writer takes such options. `WorkbookReader` checks its options, supports `styles: 'cache'`, `entries: 'emit'`, `read()`, `parse()` and events through an internal emitter (no `node:events`), and refuses `sharedStrings: 'emit'|'ignore'` and `hyperlinks: 'emit'` with `XLV_E_NOT_SUPPORTED`. JSON cell values are kept with type `String` and written as JSON text, as in exceljs.

### D-E04 Closing exceljs compatibility: styles, formulas, members, views, writer options (2026-10-30)
@xlvane/exceljs writes and reads every exceljs style part through the D-E01 core styles: gradient fills (`angle` → linear with `degree`; `path` with `center`, `right`/`bottom` defaulting to `left`/`top`), diagonal borders (`up`, `down`), all font properties in cells, rich-text runs and note runs (supersedes D-504 and D-A06, which dropped `family`, `scheme` and `charset`), `shrinkToFit`, `readingOrder` and exceljs's `textRotation` encoding, and `protection` of cells, rows, columns and styles (X4-27). Formulas follow exceljs (X4-29): `{ sharedFormula }`, `{ formula, shareType, ref }`, `fillFormula`, `formulaType`, the master's `ref` growing to its last dependent, and shared formulas read back as master plus `{ sharedFormula }` dependents with translated `cell.formula`; the streaming writer writes a `{ formula, ref }` master as a shared formula and other dependents as translated formulas. New members (X4-30): `workbook.model`, `removeWorksheetEx`, `clearThemes` (no-op: no theme is kept), `nextId`; `worksheet.model`, `findRows`, `lastColumn`, `mergeCellsWithoutStyle`, `destroy`; `row.getCellEx`, `outlineLevel`, `collapsed`, `protection`; column `outlineLevel`, `collapsed`, `protection`, `isDefault`, `headerCount`, multi-row `headers`, `toString`, `equivalentTo`; `cell.toCsvString`, `model`, `workbook`, `merge`/`unmerge`/`isMergedTo`. `model` objects are xlvane's own plain snapshots, not exceljs's xform models. The core gains row and column outline levels (0–7, [MS-OI29500]), `sheet.properties()` (`defaultRowHeight`, `outlineLevelRow`/`Col`, `summaryBelow`/`Right`; the levels are written as at least the highest used level, PROVISIONAL), `view()` options `topLeftCell`, `activeCell`, `zoomScaleNormal`, `showRuler`, `workbook.views()`, `workbook.calculation({ fullCalcOnLoad })` and `writeWorkbook(…, { sharedStrings, compressionLevel })` (also streaming and `xlvane/node`); the output stays deterministic (W2-65 … W2-68). The layer maps `workbook.views`, `calcProperties.fullCalcOnLoad`, `worksheet.properties`, `useSharedStrings` and the `zip` level options onto them. Not done, with reasons in docs/exceljs-compat.md: further sheet views and `workbookViewId` (one `sheetView` per sheet), split-view `topLeftCell`, `dyDescent` (x14ac extension, no checked source), `calcId` (xlvane does not calculate), styles of cells without a value (the core reader skips such cells), note box settings (REPORT-5).

### D-E05 Cells that hold only a style (2026-10-30)
The reader used to skip a cell element that has a style (`s` ≠ 0) but no value and no formula, so such styles were lost on reading (exceljs keeps them). xlvane now reads them as cells with `type: "empty"`, `value: null` and their style, kept apart from the cells with data: `Sheet.styledCells()` lists them, and each `StreamRow` of the streaming reader has `styledCells` (a row may have only such cells; its `cells` is then empty). `cells()`, `rows()`, `cell(ref)` and `cellCount` are unchanged, so `readTable`, `readTableStream`, `toCSV`, `toJSON`, `inspect` and `diff` see the same data as before (none of the corpus snapshots changed). Cell elements without style and without value are still nothing. A later definition with a value replaces a style-only one (`XLV_W_DUPLICATE_CELL`). Style-only cells count against `limits.maxCells`, so large styled empty areas stay bounded. `editCells` keeps a style-only cell's style when it gives that cell a value (before, the cell restarted from its column's style). @xlvane/exceljs gives these cells their style when reading, in the batch reader and with `styles: 'cache'` in `WorkbookReader`, as exceljs does; the upstream test "null cells retain style" is ported.

### D-E06 Fixes from the independent review of phase E (2026-10-31)
An independent review found 6 confirmed bugs, 4 smaller differences, a performance regression and several suspicions; all are fixed with regression tests. Splices in @xlvane/exceljs turn every shared formula group of the workbook into plain formulas, each adjusted for the splice (moving the master's text to its dependents gives wrong references when they cross the changed band); array formulas keep their group and their range moves. Styles from a file can be written again: the writers accept the schema's values for gradient numbers and stops (any finite `xsd:double`, any number of stops) and for `font.family` (0–14), superseding D-E01 on these points; the reader drops a `textRotation` outside ST_TextRotation with a warning; the core keeps an indent above 255 (D-1101) and @xlvane/exceljs leaves it out on load. CSV: dates are parsed as dayjs 1.11 customParseFormat parses them (literals skipped, tokens searched in the rest of the text, strict re-format check); `dateFormat: ''` uses the default; with `headers` an array row gives one field per header; header names are escaped against CSV injection; an unclosed quote is reported on its opening line; a quote or escape longer than one character is refused. The active cell of a view is the selection of the active pane, as exceljs reads it; the streaming `WorkbookWriter` writes workbook views. exceljs semantics matched: `textRotation` and `indent` follow `validInt`; `findRows` slices like exceljs; a JSON value's `toCsvString` is its JSON text; new worksheets have exceljs's default properties; outline levels outside 0–7 throw `XLV_E_INVALID_ARGUMENT` when set. A `$` in a shared master's `ref` is ignored on read; the last definition of a cell wins in both orders, also between a valued and a style-only definition. Style-only cells are stored as `[column, styleIndex]` pairs and built only when asked for; they still count toward `maxCells`. On a 2,000,000-cell styled sheet reading went from 2.8–3.1 s and 340 MB of heap to 2.1–2.2 s and 93 MB (the code before phase E: about 1.96 s and 61 MB).

### D-E07 Facts from the fifth review report; note box settings (2026-10-31)
The fifth local review report (R5, supplied by the maintainer) confirms the meanings xlvane had implemented from the schemas alone: Part 1 §18.8.24 and §18.8.38 with [MS-OI29500] §2.1.708 and §2.1.716 give gradient convergence values and stop positions as fractions 0–1 measured from the left and from the top; §18.8.4 the diagonal directions; §18.8.1 `readingOrder` 0/1/2 and the meanings of `shrinkToFit` and `justifyLastLine`; §18.8.33 that cell protection acts only on a protected sheet; §18.18.94, §18.4.1, §18.8.35 and §18.4.14 the font values. No code change was needed. The direction of a gradient's `degree` is stated in no document and stays an inference (W2-60 … W2-64). In array formulas only the top-left cell has `f` and every cell has `v`; no rule exists for `si` numbering or for the master's position, so both remain implementation decisions (F2-13, F2-15). Note boxes are now written and read (W2-69): `margins` in millimetres, always written with a unit (the unit of a unitless value is not documented); automatic or custom margins, with `o:insetmode` on the shape, where Excel reads it ([MS-OI29500] §2.1.1811 f/i); `locked` and `lockText` (omitted means locked, P4 §19.4.2.37/.38); `anchor` `absolute`, `oneCell` or `twoCell` using Excel's reversed `x:MoveWithCells`/`x:SizeWithCells` (§2.1.1881, §2.1.1886). @xlvane/exceljs maps `margins`, `protection` and `editAs` both ways with exceljs's defaults (X4-31). LibreOffice 24.2.7.2 keeps notes but replaces their box with its own.

## Decisions: pre-release review (before 1.0.0)

### D-R01 Pre-release fixes in the exceljs layer (2026-10-31)
Files loaded by @xlvane/exceljs must be writable again. Where a file holds something the core does not write in that form, the layer writes the nearest valid form and documents it in docs/exceljs-compat.md: zeros in `pageSetup` are left out ([MS-OI29500] §2.1.638); shared-formula groups that cannot be written as shared go out as plain formulas (F2-15); EMF and WMF pictures are accepted and written unchanged as `image/x-emf`/`image/x-wmf` (TD-31a, observed; an extension of D-D03 for pictures that come from files); row heights up to 409.5 (W-04a, observed in Excel-saved files); custom table styles from a file are written without their name; totals rows keep their cached results. Behaviour changes against exceljs: `errorStyle: 'error'` is written as `stop`; a 7-digit argb `tabColor` gets a trailing 0; solid conditional-format fills read back as `bgColor`; data-validation formulas that are not numbers stay text. Security: prototype-free maps (L1), `SPECIAL_VALUES` as a Map (M4), the formula check after skipping NUL (M3), a linear 3-D reference check (M5), CR/LF/delimiter refused with `quote: false` (L5). Bundling: `xlvane/node` is imported through a non-literal specifier, so browser and Workers bundles build and run (tested on the packed packages in Chromium and workerd). Types: the `ExcelJS` namespace and about 65 missing exceljs type names, top-level `types` and `main`. Every loadable file of the Apache POI corpus now also saves (test). For the maintainer's review: the wider picture formats for files, the 409.5 row-height limit, the trailing 0 for 7-digit `tabColor`.

### D-R02 Readable, safe console output; table and CSV fixes from the review (2026-10-31)
From the security, functional, AI-friendly and developer-friendly audits before 1.0.0.

**Terminal safety.** Text from files and arguments (cell text, sheet names, error messages and hints, warnings, notes) is written to the terminal with control characters escaped as `\xHH`: C0 except tab and line feed, DEL and C1 (which holds the 8-bit CSI U+009B). A final filter on everything printed for people keeps only the SGR sequences the CLI writes itself (bold, dim, red, green, yellow, cyan and their resets). A cell holding `_x001B_]0;title_x0007_` can therefore not change the window title, move the cursor or hide output. `--json` output is not changed: JSON escapes control characters itself. The progress line keeps its own carriage return and "erase in line".

**Colour.** TERM=dumb turns colour off, as `getColorDepth` does in Node.js lib/internal/tty.js (v22.x and v24.x, read 2026-10-31); FORCE_COLOR still wins. Colour marks errors (red), warnings (yellow), headings (bold), success lines (green), suggested commands (cyan) and secondary text (dim); it never carries information alone.

**Layout.** Width is the terminal's when it reports more than 0, else COLUMNS on a terminal, else 100. Errors: `Error:`, `Where:` (only what the message does not already say), `Fix:` wrapped with a hanging indent, `Try:`, and `Code:  XLV_E_… · docs/errors.md#anchor` (the full URL with `--verbose`). Warnings are grouped by code and message with their sheets, count and cells, the message wrapped, and a pointer to docs/errors.md. `validate` aligns the cells of problems and says "OK with N warnings … Use --strict to fail on warnings." in yellow. `read` right-aligns number columns and writes its summary line and the truncation note to stderr, so `xlvane read f.xlsx > out.txt` holds only the table. File sizes are shown as "812 bytes", "2.8 kB", "3.4 MB" (powers of 1000). JSON is indented only when stdout is a terminal; no `--pretty` option was added.

**Options and messages.** `read --limit all`; `read --json` includes the first 50 rows by default, like the text output (before: all rows), with `totalRows` and `truncated`. `read --range` is checked before the file is read (usage error, exit code 2). `-` as a file is a usage error ("Reading from standard input is not supported"). A missing file reads "File not found". `inspect --max-columns`, `--formula-samples`; built-in hidden names (`_xlnm.*`) only with `--verbose`; properties one per line. `read` and `convert` note warnings on stderr and list them under `warnings` in `--json`. `convert` notes a CSV whose first line has no comma but has `;` or tabs, with the command to run. `validate --columns` suggests the closest type and column name. `--force` gives the new file the permission bits of the file it replaces (owner and group are not changed).

**Library.** `readTable`, `readTableAsync`, `toJSON` and `toCSV` take `onWarning(warning)`, called once per workbook warning, because their results have no room for warnings. `XLV_E_TABLE_INVALID` is renamed `XLV_E_INVALID_ROWS` so that it is not confused with `XLV_E_INVALID_TABLE`. Dates in `readTable({ dates: "iso" })`, `toJSON`, `inspect` samples, table headers and the CLI use one form, ISO 8601 without a time zone as `toCSV` writes it (D-306): `2024-01-31` at midnight, otherwise `2024-01-31T13:45:00[.sss]`; before, `readTable`/`toJSON` wrote `toISOString()` with `Z` and `inspect` `…T00:00:00.000`. Document properties keep `toISOString()`, because they are instants. `inspect` lists at most 100 merged ranges per sheet, with `mergedRangeCount` and `mergedRangesTruncated`.

**Header detection** (`inspect`, `readTable`, `readTableStream`): formula cells in the candidate row are left out; whole numbers 1900–2100 count as names when they are consecutive and the next row has a value below them that is not such a year; repeated names are allowed with low confidence (docs/inspect-json.md).

**CSV numbers and limits.** `fromCSV` reads back every number `toCSV` writes (the shortest round-trip form, up to 17 digits) except whole numbers with more than 15 digits, which stay text because they are nearly always IDs that a spreadsheet would round; numbers that overflow, underflow to 0 or become subnormal stay text. `fromCSV` has `maxCells` (default `DEFAULT_READ_LIMITS.maxCells`, 5 000 000), counted while parsing, with `XLV_E_LIMIT_EXCEEDED`.

**Not done.** `read --limit` still reads the whole sheet: `readTableStream` detects columns from the leading rows only, so its columns can differ from `readTable`'s; using it would make `read` show different columns depending on `--limit`. Reading from standard input is not supported.

### D-R03 Security and core fixes before release (2026-10-31)
Linear complexity where the security audit found CPU denial of service: drawing anchors and VML shapes in `insertRows`/`deleteRows` are scanned with `indexOf` (8 KB drawing: 75 s → 6 ms; 40 000 VML shapes: 3.9 s → 20 ms), hyperlink relationships are looked up in a Map (60 000 links read: 60 s → 0.7 s), and `splice` applies edits in one pass (insertRows with 60 000 notes: 134 s → 0.7 s). The package editor unpacks parts only when they are needed and copies untouched entries compressed with their CRC and size (a file with 480 unused 1 MiB entries: 8 s and +588 MB → 23 ms); as a consequence, damaged unused entries pass through unchecked. Parts the streaming reader keeps in memory use the batch per-entry limit (256 MiB or `limits.maxEntryBytes`); sheets keep the streaming limit. The publish job installs with `npm ci --ignore-scripts`. Tables: `appendRows` extends a table that ends on the last used row (with its filter); a table with a totals row is refused with an `insertRows` hint. A header edit in `editCells` renames the table column ([MS-OI29500] §2.1.681 i: unique, non-empty text up to 255 characters; structured references are not rewritten). New checks: `XLV_W_MERGED_CELL_WRITE` (E-35), and `activeTab`/`firstSheet` must name an existing sheet, in the batch and the streaming writer. `location.cell` may hold a range or a whole row or column (`C:C`, `7:7`); one shared sheet-not-found error with suggestions and a case-sensitivity note. Each fix has a regression test (time budgets for the performance fixes).

### D-R04 Pre-release defaults approved by the maintainer (2026-10-31)
The maintainer approved these defaults for 1.0.0; each can be changed by the caller:
- Dates in `readTable({ dates: "iso" })` and `toJSON` are ISO 8601 without a time zone (D-R02); `dates: "date"` returns `Date` objects to format in any other way.
- EMF and WMF pictures loaded from files are kept and written unchanged (D-R01); a caller can remove them like any other image.
- Row heights are accepted up to 409.5 points (D-R01, W-04a); any lower height can be set.
- A 7-digit argb `tabColor` in @xlvane/exceljs gets a trailing 0 (D-R01); an 8-digit value is written as given.
- `fromCSV` stops at `maxCells` (default 5 000 000, D-R02), which the caller can raise or lower.

### D-R05 Licence review before release (2026-10-31)
- `xlvane` and `@xlvane/exceljs` are MIT. Neither imports or bundles exceljs: `@xlvane/exceljs` is a separate implementation of the exceljs 4.4.0 API on top of xlvane. The only runtime dependency of `xlvane` is `fflate` (MIT), imported and not bundled into dist, so its own licence ships with it.
- The API names and type declarations of `@xlvane/exceljs` follow exceljs's index.d.ts (cited by line in the sources). To leave no doubt, packages/exceljs/LICENSE also carries the exceljs MIT notice (Copyright (c) 2014-2019 Guyon Roche, copied from the npm package exceljs 4.4.0), and both READMEs say the project is not affiliated with or endorsed by the exceljs project. A policy test checks both.
- Repository-only material: tests ported from exceljs v4.4.0 and its image.png (MIT, tests/exceljs/LICENSE-exceljs.md); the Apache POI corpus (Apache-2.0 with its LICENSE and NOTICE, tests/corpus/poi/); every corpus file is listed in tests/corpus/LICENSES.md. None of it is in the npm tarballs. The benchmark libraries (exceljs, xlsx) are installed only by `npm run bench` and are not distributed.
- Standards: the ECMA-376 and Microsoft Open Specifications documents are not in the repository (spec/ and `docs/Nowy folder/` are git-ignored); docs/facts cite them by clause and page with short quotations.
- Open: the copyright holder in LICENSE is "DevXord" (D-012), to be confirmed by the maintainer. Whether the npm names `xlvane` and the scope `@xlvane` are free could not be checked from this environment.

### D-R06 The compatibility package is named `@xlvane/compat` (2026-10-31)
At the maintainer's decision the package published as `@xlvane/exceljs` until now is named `@xlvane/compat`, and its directory is packages/compat: the package name does not use the name of another project. Its README, description and documentation still say, descriptively, that it offers the exceljs 4.4.0 API, together with the statement that it is not affiliated with or endorsed by the exceljs project (D-R05). Users who want to keep `import ... from 'exceljs'` alias it in their own package.json: `"exceljs": "npm:@xlvane/compat@^1"`. The npm keyword is `exceljs-compatible`. Earlier entries of this file and of docs/CHANGELOG-dev.md keep the old name as history. No version was published under either name.

### D-R07 Node-only modules of @xlvane/compat through package imports (2026-10-31)
Found by testing the packed packages in outside projects: tsdown folds `const specifier = 'xlvane/node'; import(specifier)` into `import("xlvane/node")`, and esbuild ignores `/* @vite-ignore */` and `/* webpackIgnore */`, so `esbuild --bundle --platform=browser` (or `neutral`) of an app using @xlvane/compat failed with "Could not resolve node:fs". @xlvane/compat now loads `xlvane/node` and `node:fs/promises` through its own package imports `#xlvane-node` and `#node-fs` (package.json `imports`; Node.js 22, Bun 1.4.4 and Deno 2.9.6 were checked to resolve both mappings, see below): under the `node` condition, which Node.js, Bun and Deno use, they resolve to `xlvane/node` and to dist/node-fs.js, which re-exports `node:fs/promises` (Node.js 22 refuses `node:fs/promises` itself as an "imports" target with ERR_INVALID_PACKAGE_TARGET, observed); under any other conditions (browser, worker, workerd, neutral) to dist/no-node.js, which throws when loaded, so `readFile`/`writeFile` give `XLV_E_NOT_SUPPORTED` there. Checked with esbuild (browser and neutral: no `node:` import in the bundle, the bundle runs in Chromium and `xlsx.readFile` gives XLV_E_NOT_SUPPORTED), Node.js (file reads and writes work) and the repository's runtime tests; tests/policy/compat-bundling.test.ts keeps the mapping and the import sites.

### D-R08 Fixes from testing the packed packages in outside projects (2026-10-31)
**Library.** `appendRows` starts below every merged range, also one that holds no value (one pass over the merges); before, a new row could land inside a merge and be hidden. A split view (`view({ split })`) now writes the caller's `topLeftCell` (as `sheetView@topLeftCell`) and `activeCell` (as the selection of the active pane) in both writers instead of dropping them (W2-51). The hint for a conditional-format rule without `formulas` gives an example that fits the rule type. `insertRows`/`deleteRows` still refuse sheets with conditional formats in the [MS-XLSX] extension, also when xlvane wrote them (extended data bar settings, `autoMin`/`autoMax`, the icon sets `3Stars`, `3Triangles`, `5Boxes`, `NoIcons`); this is now stated in the README, the TSDoc and CHANGELOG.md.

**CLI.** `edit --dry-run` without `--output` now fails with `XLV_E_FILE_EXISTS` like the real run (before, the dry run reported a replacement the real run would refuse); with `--force` it reports the replacement. "Try:" commands fix the argument that was wrong: a sheet name that is not found is replaced where it was typed, in `--sheet` or in the `Sheet!` prefix of `--set` (quoted when the name needs it), and no command is suggested when the name is not in the arguments (before, `--sheet` was appended while `--set` kept the wrong name, which repeated the error). For a missing file, only the argument that was read is replaced (not an `--output` with the same name). The "Did you mean" candidates are limited to files the command can read (workbook extensions for XLSX inputs, the same extension for CSV/TSV/JSON) and leave out files another argument already names, so `diff a.xlsx typo.xlsx` no longer suggests comparing `a.xlsx` with itself. `create`/`convert` with a JSON object instead of an array names the property that holds an array and gives a `jq` command to extract it into `<input name>-<property>.json`, numbered so that it is never the input (which `jq … > input` would empty) or an existing file, instead of the library hint about passing a property. Human-readable output escapes the Unicode Bidi_Control characters (U+061C, U+200E, U+200F, U+202A–U+202E, U+2066–U+2069; the set equals `/\p{Bidi_Control}/u` in Node.js 22, Unicode 16.0, checked by a test) as `\u202e`, because they reorder the text displayed after them; `--json` is unchanged. Not changed: text tables pad by UTF-16 length, not display width, so CJK text misaligns columns; Node.js's own warning when both NO_COLOR and FORCE_COLOR are set.

**Tests.** vitest maps @xlvane/compat's package imports `#xlvane-node` and `#node-fs` (D-R07) to the sources, as it already did for `xlvane`; without this they resolved to packages/xlvane/dist, a second copy of xlvane.

### D-R09 @xlvane/compat fixes from the comparison with exceljs 4.4.0 in outside projects (2026-10-31)
Testing the packed package against exceljs 4.4.0 found the following; each fix cites the exceljs source and has a test in tests/exceljs/outside-project-fixes.test.ts.
- `worksheet.columns` now has a column for every column a cell was created in, as exceljs (row.js L48, worksheet.js L243–261). Default columns are not written, so the usual auto-fit loop works after `addRow` and on loaded files.
- The streaming `WorkbookWriter` is linear. Its queue no longer uses `Array#shift`, and the last-row lookup stops at the committed rows. 300 000 rows take 5.5 s against exceljs's 2.75 s (before: 75 s); the remaining factor is Deflate in JavaScript (fflate) against Node.js zlib, both at level 6.
- The declarations follow exceljs's index.d.ts for `getTable` (`Table`), `workbook.getImage` (`Image`), `getImages()[i].range` (required) and `fillFormula` results. At run time a missing table or image still gives `undefined`, as in exceljs.
- A single underline reads back as `true` (underline-xform.js).
- A split view's `activeCell` round-trips (core D-R08). Its exceljs `topLeftCell` (`pane@topLeftCell`, the bottom-right pane) is still not written.
- `getSheetValues()` gives `[]` for an existing empty row.
- `worksheet.orderNo` follows exceljs (max+1 for a new sheet, the 0-based position in the file for a loaded one) and orders `worksheets`.
- Tables read from a file keep their columns as read (no default `'none'` function or "Total" label).
- Documented, not changed: `sheetId` is not kept, because the core writes 1…n and does not report it on reading (a core option for it is left for a later release); note `protection` reads back as `'True'`/`'False'`; the module object is frozen.

### D-R10 The first release is published locally, without provenance (2026-10-31)
The first-release steps in docs/RELEASING.md staged both packages from the maintainer's machine and then rejected the stages, to let the workflow publish 1.0.0 with provenance. Rejecting the first stage of a new package removed the package itself: `npm view xlvane` gives "Unpublished on 2026-10-31T19:35:12.446Z", and npm's documentation (npm/documentation `unpublishing-packages-from-the-registry.mdx`, read 2026-10-31) says "If you unpublish an entire package, you may not publish any new versions of that package until 24 hours have passed" and that a published `name@version` cannot be reused (here only the placeholder `0.0.0-stage` was published; 1.0.0 never was). The staged-publishing page does not describe rejection. The maintainer chose to publish 1.0.0 from the maintainer's machine after the 24 hours, staged with `--provenance=false` and approved with 2FA, then to configure trusted publishing, so that releases from 1.0.1 on are staged by the workflow with provenance. RELEASING.md now describes this and warns against rejecting a first stage.

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
| docs/facts/sml-write.md | W2-*, F-LO-01, F-NODE-04, F-NODE-05, F-NODE-11 … F-NODE-13, F-GH-01 | writer facts, checked against ECMA-376 Part 1, [MS-OI29500] and [MS-XLSX] (D-1104); W2-21 and W2-32 PROVISIONAL |
| docs/facts/formulas.md | F2-* | formula tokenizer for shared formulas; checked against ECMA-376 Part 1 §18.17, [MS-OI29500] and [MS-XLSX] §2.2.2 (D-1102, D-1104) |
| docs/facts/ref-adjust.md | R-* | reference adjustment on insert/delete, observed in LibreOffice 24.2.7.2 (D-B01) |
| docs/facts/sheet-rules.md | SR-* | data validation, conditional formatting, dxfs, protection, printing (REPORT-3; D-C01 … D-C04) |
| docs/facts/tables-drawings.md | TD-* | tables, drawings and pictures (REPORT-3; D-D02, D-D03) |
| docs/facts/csv.md | CSV-*, SS-* | CSV (RFC 4180 PROVISIONAL), OWASP CSV Injection, Standard Schema |

### F-NODE-04 … F-NODE-05
See docs/facts/sml-write.md.

### F-NODE-06 fs.openAsBlob
`fs.openAsBlob(path[, options])` returns a Promise of a Blob backed by the file; it was added in v19.8.0 and marked stable in v22.17.0; the file must not be modified after the Blob is created. Source: https://github.com/nodejs/node/blob/v22.x/doc/api/fs.md, checked 2026-10-30.

### F-NODE-07 process.memoryUsage and resourceUsage
`process.memoryUsage()` returns `rss`, `heapTotal`, `heapUsed`, `external`, `arrayBuffers` (bytes); `process.resourceUsage().maxRSS` is in kibibytes. Source: https://github.com/nodejs/node/blob/v22.x/doc/api/process.md, checked 2026-10-30. Used by tests and the benchmark only.

### F-NODE-08 require(esm) and the `'module.exports'` export name
`require()` of an ES module returns its namespace object (the default export at `.default`); from v22.12.0 / v23.0.0 a module can export a value under the string name `'module.exports'`, which `require()` then returns directly. Source: Node.js `doc/api/modules.md` (v22.x), recorded as W-05 in docs/facts/zip-xml-web.md. Verified with the built `@xlvane/exceljs` on Node.js 26 (`tests/exceljs/alias.test.ts`, `check:pkg`).

### F-NODE-09 util.parseArgs
`util.parseArgs({ args, options, strict, allowPositionals, tokens })` was added in v18.3.0 / v16.17.0 and is no longer experimental from v20.0.0; `tokens: true` returns tokens of kind `option` (with `name`, `rawName`, `value`, `inlineValue`), `positional` and `option-terminator`; `allowNegative` was added in v22.4.0 (not used, because xlvane supports `>=22`). Source: https://github.com/nodejs/node/blob/v22.x/doc/api/util.md (`util.parseArgs([config])`), checked 2026-10-30.

### F-NODE-10 util.styleText and colour environment variables
`util.styleText(format, text[, options])` was added in v21.7.0 / v20.12.0; from v22.8.0 / v20.18.0 it respects isTTY and `NO_COLOR`, `NODE_DISABLE_COLORS`, `FORCE_COLOR`; options `validateStream` (default `true`) and `stream`; stable from v22.13.0. `FORCE_COLOR` values `1`, `true`, empty, `2`, `3` enable colour and then `NO_COLOR` and `NODE_DISABLE_COLORS` are ignored; any other value disables colour; `NO_COLOR` is an alias of `NODE_DISABLE_COLORS` and its value is arbitrary. Sources: https://github.com/nodejs/node/blob/v22.x/doc/api/util.md (`util.styleText`) and doc/api/cli.md (`FORCE_COLOR`, `NO_COLOR`), checked 2026-10-30. no-color.org could not be reached from the development environment (proxy refusal), so the Node.js definition is used.


### F-NODE-14 fsPromises.chmod and stats.mode
Node.js doc/api/fs.md (v22.x, read 2026-10-31): `fsPromises.chmod(path, mode)` (added v10.0.0) "Changes the permissions of a file"; `stats.mode` is "A bit-field describing the file type and mode". The CLI keeps the low 12 bits (`mode & 0o7777`) of the replaced file for `--force` (D-R02).
### F-EXCELJS-01 exceljs 4.4.0 behaviour
The exceljs behaviours copied by `@xlvane/exceljs` are taken from the exceljs 4.4.0 source installed in `bench/node_modules/exceljs` (`lib/doc/enums.js` for `ValueType`, `FormulaType`, `ErrorValue`; `lib/exceljs.nodejs.js` for the CommonJS export; `lib/doc/worksheet.js` L134–173 for sheet-name rules; `lib/doc/cell.js`, `row.js`, `column.js`; `lib/stream/xlsx/*`) and from its specs at tag v4.4.0. Each behaviour lists its source file and line in docs/exceljs-compat.md.

### F-NPM-01 Trusted publishing
From github.com/npm/documentation (main, read 2026-10-30) `content/packages-and-modules/securing-your-code/trusted-publishers.mdx`: requires npm CLI 11.5.1 or later and Node 22.14.0 or later; GitHub-hosted runners only; configured in the package settings on npmjs.com (organization or user, repository, workflow filename with `.yml`, optional environment, allowed actions); the workflow needs `permissions: id-token: write`; a new configuration expires if it is not used for a publish within 2 days; configurations created after 2026-09-03 allow `npm stage publish` by default; the documentation recommends stage-only permissions and "Require two-factor authentication and disallow tokens". Provenance is generated automatically with trusted publishing for a public repository and a public package; not for private repositories. `generating-provenance-statements.mdx`: `package.json` `repository` must match the repository (case-sensitive).

### F-NPM-02 Staged publishing
`staged-publishing.mdx` (same repository and date): requires npm CLI 11.15.0 or later and Node 22.14.0 or later; `npm stage publish` needs no 2FA; approval with `npm stage approve <stage-id>` or on npmjs.com needs 2FA; for a package that does not exist yet npm publishes a public placeholder `0.0.0-stage`. `npx npm@11.21.1 stage --help`: subcommands `publish`, `list`, `view`, `approve`, `reject`, `download`; `stage publish` accepts `-w`, `--access`, `--provenance`, `--dry-run`. A dry run of `npm stage publish -w packages/xlvane` staged `xlvane@1.0.0` (dry-run) with public access.

### F-NPM-03 npm and setup-node versions
`npm view npm`: 11.21.1 (2026-10-15) is the newest 11.x, engines `^20.17.0 || >=22.9.0`; 12.2.1 needs Node `^22.22.2 || ^24.15.0 || >=26.0.0`. `actions/setup-node` at 949feb24 (v7.1.0) `action.yml`: inputs `registry-url` and `package-manager-cache` (set to false to disable automatic caching).

### F-WORKERD-01 workerd configuration and compatibility date
The config format (`using Workerd = import "/workerd/workerd.capnp"`, `services`, `sockets` with `address`, `http`, `service`, worker `modules` with `esModule = embed`) follows `samples/helloworld_esm/config.capnp` in github.com/cloudflare/workerd (main, read 2026-10-30). `src/workerd/server/workerd.capnp`: "`compatibilityDate` must be specified". The README: "`workerd`'s version number is simply a date, corresponding to the maximum compatibility date supported by that version." CLI: `workerd serve <config-file> [<const-name>]` (`workerd serve --help`, 1.20260730.1).

### F-RUNTIMES-01 Deno, Bun and Playwright
Deno 2.9.6 `deno run --help`: `--no-config`, `--no-lock`, `--no-npm`, `--no-remote`, `--no-prompt`, `--allow-read=<PATH>`. Bun 1.4.4 `bun run --help`: `--no-install`. Playwright 1.63.0: `chromium.launch()`, `browser.newPage()`, `page.goto()`, `page.evaluate()` and the `pageerror` event, used as in its type declarations (`node_modules/playwright-core/types/types.d.ts`) and confirmed by running.

### F-MCP-01 MCP TypeScript SDK v2
`@modelcontextprotocol/server` 2.4.0 is npm `latest` (`npm view`, 2026-10-30), Apache-2.0, `engines.node >=20`, dependencies zod ^4.2.0 and `@modelcontextprotocol/core` 2.4.0. Its README: "This is v2 of the MCP TypeScript SDK, the stable release line, implementing the 2026-07-28 MCP spec. It replaces the single `@modelcontextprotocol/sdk` package from v1" (v1 is 1.32.2). API used, from `dist/*.d.mts` of the tarball: `McpServer(serverInfo, { capabilities, instructions })`, `registerTool(name, { title, description, inputSchema, annotations }, cb)` where `inputSchema` is a Standard Schema with JSON Schema, `fromJsonSchema(schema)`, `InMemoryTransport.createLinkedPair()`, and `serveStdio(factory)` from `@modelcontextprotocol/server/stdio`. Client side (`@modelcontextprotocol/client` 2.4.0): `Client`, `listTools`, `callTool`, `getInstructions`, `StdioClientTransport` from `/stdio`. Verified by a prototype on 2026-10-30 (input validation errors come back as `isError` results).

### F-MCP-02 MCP specification 2026-07-28
From `docs/specification/2026-07-28/changelog.mdx` in github.com/modelcontextprotocol/modelcontextprotocol (main, read 2026-10-30): the protocol is stateless (no `initialize` handshake, protocol version and client capabilities in each request's `_meta`); servers MUST implement `server/discover`; protocol-level sessions are removed; Roots, Sampling and Logging are deprecated ("pass directories or files via tool parameters, resource URIs, or server configuration instead of Roots"); servers SHOULD return tools from `tools/list` in a deterministic order; `inputSchema` and `outputSchema` may use any JSON Schema 2020-12 keywords.

### F-SKILL-01 Agent Skills
Specification (github.com/agentskills/agentskills `docs/specification.mdx`, main, read 2026-10-30): a skill is a directory with `SKILL.md`, YAML frontmatter then Markdown; `name` required, 1–64 characters, lowercase letters, digits and hyphens, not starting or ending with a hyphen, no consecutive hyphens, equal to the parent directory name; `description` required, 1–1024 characters; optional `license`, `compatibility` (at most 500 characters), `metadata`, `allowed-tools`. The `skills` CLI (vercel-labs/skills README and docs/creating-skills.md, npm 1.7.4): `npx skills add <repository | URL | local path>`, options `--skill`, `--agent`, `--yes`, `--copy`; discovers `skills/<name>/SKILL.md`; requires Node.js 22.20 or later; `DISABLE_TELEMETRY=1` or `DO_NOT_TRACK=1` turns off telemetry and security-audit requests. agentskills.io itself was not reachable (proxy refusal).

### F-TOOL-11 Vitest execArgv
Vitest 5 option `test.execArgv`: "Pass additional arguments to `node` process when spawning the worker" (default `[]`). Source: `node_modules/vitest/dist/chunks/plugin.d.C8mmycim.d.ts` (Vitest 5.0.5), checked 2026-10-30.
