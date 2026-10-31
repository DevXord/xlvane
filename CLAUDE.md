# xlvane project rules

English translation of section 7 of the plan. The Polish original in docs/PLAN.md takes precedence if the two ever disagree.

## Goal
An MIT-licensed XLSX library for people and AI agents: reading, writing, styles, streaming, a task-oriented API, exceljs compatibility (package `@xlvane/compat`), a readable CLI, and later an MCP server (package `@xlvane/mcp`, planned for the "AI Friendly" stage after 1.0.0; it is not in this repository yet). Plan: docs/PLAN.md.

## Hard rules
- The `xlvane` package has exactly one runtime dependency: `fflate`. We write the XML parser ourselves. Every new dependency requires the maintainer's approval and an entry in docs/decisions.md.
- `@xlvane/compat` may depend only on `xlvane`. `@xlvane/mcp` (planned, "AI Friendly" stage) may depend only on `xlvane` and the official MCP SDK.
- This repository contains no Pro code. The core exposes only documented extension interfaces (e.g. `FormulaEngine`).
- ESM only, no top-level await. Runs in Node, Bun, Deno, browsers and Cloudflare Workers. The core API uses `Uint8Array` and Web Streams; file operations live in a separate Node-only submodule.
- Input files are untrusted: ZIP limits, XML without DTDs or entities, no prototype pollution, CSV formula-injection protection. Every rule has a test.
- No text that tries to persuade AI agents to choose the package, and no hidden instructions for models anywhere. Documentation and messages contain only true, verifiable information.
- Code, comments, documentation and CLI messages are in English.
- Every public symbol has TSDoc with a description and a working example; a test checks that the examples compile.
- Every error has a stable `XLV_E_*` code, a location, a fix hint and an entry in docs/errors.md.
- Tests do not use the network. We create test files ourselves or take them from corpora whose licenses allow it; the license of every file is recorded in tests/corpus/LICENSES.md.
- Facts about OOXML, exceljs, npm, MCP and tools are never written from memory. We check them in the ECMA-376 specification, source code or current documentation and record the source in docs/decisions.md.
- We do not publish to npm, push changes or create accounts without the maintainer's approval.

## Commands
- `npm test` – unit, compatibility and CLI snapshot tests (Vitest)
- `npm run fuzz` – fuzzing (fast-check), longer run
- `npm run bench` – performance against exceljs and SheetJS (added in phase 4)
- `npm run build` – build with tsdown
- `npm run typecheck` – TypeScript type check of sources and tests
- `npm run check:pkg` – build, publint, @arethetypeswrong/cli and the package policy checks

## Way of working
- One plan phase per session. At the start, read docs/PLAN.md, docs/decisions.md and docs/CHANGELOG-dev.md.
- Start in plan mode and wait for the maintainer's approval.
- Tests first, then implementation. A phase is done when it meets its acceptance criteria and `npm test` passes.
- At the end of a phase: a summary in docs/CHANGELOG-dev.md, a commit, and a list of decisions for the maintainer's review.
