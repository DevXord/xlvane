# xlvane

[![npm version](https://img.shields.io/npm/v/xlvane)](https://www.npmjs.com/package/xlvane)
[![License: MIT](https://img.shields.io/npm/l/xlvane)](LICENSE)
[![CI](https://github.com/DevXord/xlvane/actions/workflows/ci.yml/badge.svg)](https://github.com/DevXord/xlvane/actions/workflows/ci.yml)

An MIT-licensed library for reading and writing XLSX files, designed for input from untrusted sources.

Releases are published from GitHub Actions with npm provenance: `xlvane` and `@xlvane/compat` from 1.0.1, `@xlvane/mcp` from 1.2.2 ([SECURITY.md](SECURITY.md)).

**Website:** [xlvane.github.io](https://xlvane.github.io) (guide, error codes, an in-browser demo). **Status:** published on npm ([xlvane](https://www.npmjs.com/package/xlvane), [@xlvane/compat](https://www.npmjs.com/package/@xlvane/compat), [@xlvane/mcp](https://www.npmjs.com/package/@xlvane/mcp)). Install with `npm install xlvane` (and `npm install @xlvane/compat` for the exceljs-compatible API). Changes for users are in [CHANGELOG.md](CHANGELOG.md), planned and deferred work in [ROADMAP.md](ROADMAP.md); how releases are made is in [docs/RELEASING.md](docs/RELEASING.md).

## Packages

| Package | Contents | Runtime dependencies |
|---|---|---|
| [`xlvane`](packages/xlvane) | Core library (safe reading, `inspect()`, writing with styles, task API, streaming) and the `xlvane` command line | `fflate` |
| [`@xlvane/compat`](packages/compat) | exceljs-compatible API ([docs/exceljs-compat.md](docs/exceljs-compat.md)) | `xlvane` |
| [`@xlvane/mcp`](packages/mcp) | MCP server for AI agents and other MCP clients | `xlvane`, `@modelcontextprotocol/server` |

## Command line in 5 minutes

```sh
npm install -g xlvane                    # or use npx xlvane in a project
xlvane                                   # overview and common commands
xlvane inspect report.xlsx               # sheets, columns, header rows, warnings
xlvane read report.xlsx --sheet Sales    # print a sheet as a table
xlvane convert report.xlsx sales.csv     # XLSX ⇄ CSV / JSON
xlvane create data.json report.xlsx      # formatted report from JSON or CSV
xlvane diff old.xlsx new.xlsx            # what changed
xlvane help <command>                    # usage and examples
```

Add `--json` to any command for machine-readable output ([docs/cli-json.md](docs/cli-json.md)). More in [packages/xlvane/README.md](packages/xlvane/README.md#command-line-in-5-minutes).

## MCP server, Agent Skill and offline docs

- `npx @xlvane/mcp --root ./data` (or `xlvane mcp --root ./data`) serves the tools `inspect_workbook`, `read_table`, `diff` and `validate` for files inside `./data`; `--allow-write` adds `write_table`, `edit_cells` and `convert`. See [packages/mcp/README.md](packages/mcp/README.md).
- [skills/xlvane/SKILL.md](skills/xlvane/SKILL.md) is an [Agent Skill](https://agentskills.io) with instructions and examples. Install it into a project with one command of the [`skills`](https://github.com/vercel-labs/skills) CLI: `npx skills add ./node_modules/xlvane/skills --skill xlvane` after `npm install xlvane` (works offline), or `npx skills add <path-to-a-clone-of-this-repository>/skills --skill xlvane`.
- The `xlvane` package also ships its documentation (`docs/`) and `llms-full.txt` (README, skill and docs in one file), so it is readable without network access.

## Moving from exceljs or SheetJS

- [docs/migrate-from-exceljs.md](docs/migrate-from-exceljs.md): switching a project from exceljs to `@xlvane/compat` with `npx xlvane migrate` and an npm alias, what to check afterwards and how to roll back.
- [How to fix the CVEs in the `xlsx` package from npm](docs/xlsx-cves.md): the two advisories of the `xlsx` (SheetJS) package on npm, the fixed SheetJS versions outside npm, and the equivalent xlvane calls.

## xlvane compared with other libraries

Every cell below is backed by a recorded source (registry document, README or source file at a pinned version or commit) in [docs/facts/comparison.md](docs/facts/comparison.md), checked on 2026-11-03. Performance is not compared here.

| | xlvane | ExcelJS | SheetJS Community Edition (npm `xlsx`) | excelize (Go) |
|---|---|---|---|---|
| Latest version (date) | 1.3.0 (2026-11-03) | 4.4.0 on npm (2023-10-19) | 0.18.5 on npm (2022-03-24); later versions are distributed by SheetJS outside npm | v2.11.1 Go module (2026-11-02) |
| License | MIT | MIT | Apache-2.0 | BSD-3-Clause |
| Runtime dependencies | 1 (`fflate`) | 9 | 7 | 9 direct module requirements in `go.mod` |
| Read / write XLSX | Yes / yes (XLSM, XLTX and XLTM are read too) | Yes / yes | Yes / yes | Yes / yes (also XLAM, XLSM, XLTM, XLTX) |
| Legacy `.xls` | No (rejected with an error) | Not documented | Read and write (BIFF2 to BIFF8) | Not listed |
| Other formats | CSV, JSON | CSV | XLSB, ODS, Numbers, CSV, HTML and others | None listed |
| Cell styles (fonts, fills, borders) when writing | Yes | Yes | Listed under SheetJS Pro in the README | Yes |
| Charts | From 1.3.0: write and read column, bar, line, pie, scatter and area charts. Earlier versions skip them when reading and keep them unchanged in `editCells` | No chart API | Listed under SheetJS Pro; chartsheets are read as their cached data | Write (`AddChart`, `AddChartSheet`) |
| Formula calculation | No | No | Not in the Community Edition (the README points to SheetJS Pro) | Partial (`CalcCellValue`, with documented gaps) |
| Streaming | Read and write XLSX | Read and write XLSX | Write CSV, HTML and JSON (`XLSX.stream`); for reading, the README buffers the stream first | Read and write worksheets |

What xlvane cannot do: it does not calculate formulas, does not open encrypted (password-protected) workbooks or legacy `.xls` files, does not model pivot tables, shapes, slicers or form controls, and reads and writes no format besides XLSX (plus XLSM, XLTX and XLTM on reading), CSV and JSON ([packages/xlvane/README.md](packages/xlvane/README.md#limitations), [SECURITY.md](SECURITY.md#out-of-scope)). `@xlvane/compat` is a replacement for the commonly used exceljs API; [docs/exceljs-compat.md](docs/exceljs-compat.md) lists what is supported, partly supported and not supported. xlvane is an independent project, not affiliated with or endorsed by the authors of ExcelJS, SheetJS or excelize.

## How xlvane is built and verified

- **Who writes the code.** The code, tests and documentation are written by AI coding agents (Claude Code), following the rules in [CLAUDE.md](CLAUDE.md) and the phases of [docs/PLAN.md](docs/PLAN.md). Each phase starts with a plan that the maintainer approves; the maintainer also approves every release (staged publishing with two-factor authentication, [docs/RELEASING.md](docs/RELEASING.md)). Decisions and checked facts, each with its source, are in [docs/decisions.md](docs/decisions.md); a summary of each phase is in [docs/CHANGELOG-dev.md](docs/CHANGELOG-dev.md).
- **Tests.** `npm test` ran 1982 tests in 181 files on 2026-11-03 (1927 passed, 52 skipped, 3 marked as to do; Node.js 22.23.4): unit tests, security tests ([tests/security](tests/security)), policy tests ([tests/policy](tests/policy): dependencies, release workflow, texts read by agents), command-line snapshots, the ported exceljs tests and the corpus tests.
- **Corpus.** 340 files of the Apache POI test corpus (Apache License 2.0; written by Excel for Windows, Mac and the web, LibreOffice, ONLYOFFICE and others, and damaged files) and files saved by LibreOffice; the origin and license of every file is in [tests/corpus/LICENSES.md](tests/corpus/LICENSES.md).
- **Fuzzing.** Property-based tests with [fast-check](https://fast-check.dev/) in [tests/fuzz](tests/fuzz). Each property runs `XLV_FUZZ_RUNS` times (100 in `npm test`, 20 000 by default in `npm run fuzz`; runs with `XLV_FUZZ_RUNS=100000` are recorded for 1.1.1, for 1.2.2 before the 1.2.3 fixes and for 1.3.0: docs/CHANGELOG-dev.md, docs/decisions.md D-K03 and D-K04), and the slower ones (CSV, sheet rules, edits, writing, streaming, streamed ZIP) a tenth to a fiftieth of that ([tests/fuzz/runs.ts](tests/fuzz/runs.ts) and the `RUNS` constant of each file).
- **LibreOffice.** [tests/libreoffice](tests/libreoffice) opens files written by xlvane in LibreOffice Calc (headless) and checks the values, styles and layout it sees; CI installs LibreOffice and requires these tests.
- **Runtimes.** CI runs the tests on Node.js 22, 24 and 26, and the same smoke bundle in Bun, Deno, Chromium and workerd (Cloudflare Workers) ([scripts/run-runtimes.mjs](scripts/run-runtimes.mjs), [.github/workflows/ci.yml](.github/workflows/ci.yml)).
- **Security reviews.** Independent reviews with proof-of-concept files were run before 1.1.1, on the published versions 1.0.0 to 1.2.2 before 1.2.3, and twice on the 1.3.0 release candidate. Each fixed finding has a test that failed before the fix; the points left unchanged are listed with their reasons ([docs/decisions.md](docs/decisions.md) D-H01, D-J01, D-K03 and D-K04).
- **Provenance.** Releases are built and published by [.github/workflows/release.yml](.github/workflows/release.yml) through npm trusted publishing, which adds provenance (`xlvane` and `@xlvane/compat` from 1.0.1, `@xlvane/mcp` from 1.2.2).

## Development

Requires Node.js 22.18+, 24.11+ or 26+ (the range supported by the build and test tools).

```sh
npm install
npm test            # unit, security, policy and corpus tests
npm run fuzz        # property-based tests with many runs
npm run typecheck   # TypeScript
npm run check:pkg   # build, publint, attw and package policy checks
```

How to report problems and propose changes is in [CONTRIBUTING.md](CONTRIBUTING.md). Project rules are in [CLAUDE.md](CLAUDE.md), decisions and verified format facts in [docs/decisions.md](docs/decisions.md), and the security policy in [SECURITY.md](SECURITY.md).

## License

MIT. `@xlvane/compat` is an independent project, not affiliated with or endorsed by the exceljs project or its authors; its LICENSE includes the exceljs MIT notice, because its API and type declarations follow exceljs 4.4.0.
