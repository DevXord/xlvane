# Development changelog

One entry per plan phase (docs/PLAN.md), newest first.

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
