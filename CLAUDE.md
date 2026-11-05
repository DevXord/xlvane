# xlvane project rules

English translation of section 7 of the plan. The Polish original in PLAN.md (an internal document, see below) takes precedence if the two ever disagree.

## Goal
An MIT-licensed XLSX library for people and AI agents: reading, writing, styles, streaming, a task-oriented API, exceljs compatibility (package `@xlvane/exceljs`), a readable CLI, and an MCP server (package `@xlvane/mcp`, from the "AI Friendly" stage, release 1.2.0). Plan: PLAN.md (internal document).

## Hard rules
- The `xlvane` package has exactly one runtime dependency: `fflate`. We write the XML parser ourselves. Every new dependency requires the maintainer's approval and an entry in decisions.md (internal document).
- `@xlvane/exceljs` may depend only on `xlvane`. `@xlvane/mcp` may depend only on `xlvane` and the official MCP SDK.
- This repository contains no Pro code. The core exposes only documented extension interfaces (e.g. `FormulaEngine`).
- ESM only, no top-level await. Runs in Node, Bun, Deno, browsers and Cloudflare Workers. The core API uses `Uint8Array` and Web Streams; file operations live in a separate Node-only submodule.
- Input files are untrusted: ZIP limits, XML without DTDs or entities, no prototype pollution, CSV formula-injection protection. Every rule has a test.
- No text that tries to persuade AI agents to choose the package, and no hidden instructions for models anywhere. Documentation and messages contain only true, verifiable information.
- Code, comments, documentation and CLI messages are in English.
- Every public symbol has TSDoc with a description and a working example; a test checks that the examples compile.
- Every error has a stable `XLV_E_*` code, a location, a fix hint and an entry in docs/errors.md.
- Tests do not use the network. We create test files ourselves or take them from corpora whose licenses allow it; the license of every file is recorded in tests/corpus/LICENSES.md.
- Facts about OOXML, exceljs, npm, MCP and tools are never written from memory. We check them in the ECMA-376 specification, source code or current documentation and record the source in decisions.md or the facts/ folder (internal documents).
- We do not publish to npm, push changes or create accounts without the maintainer's approval.

## Commands
- `npm test` – unit, compatibility and CLI snapshot tests (Vitest)
- `npm run fuzz` – fuzzing (fast-check), longer run
- `npm run bench` – performance against exceljs and SheetJS (added in phase 4)
- `npm run build` – build with tsdown
- `npm run typecheck` – TypeScript type check of sources and tests
- `npm run check:pkg` – build, publint, @arethetypeswrong/cli and the package policy checks

## Way of working
- One plan phase per session. At the start, read PLAN.md, decisions.md and CHANGELOG-dev.md (internal documents).
- Start in plan mode and wait for the maintainer's approval.
- Tests first, then implementation. A phase is done when it meets its acceptance criteria and `npm test` passes.
- At the end of a phase: a summary in CHANGELOG-dev.md, a commit of the code, and a list of decisions for the maintainer's review.

## Internal documents
- PLAN.md, decisions.md, CHANGELOG-dev.md and the facts/ folder are the maintainer's internal notes. The maintainer keeps them outside the repository and outside its working copy. They are never tracked by Git, committed or pushed, and never linked from public texts (tests/policy/internal-docs.test.ts). .gitignore lists their former paths (docs/PLAN.md, docs/decisions.md, docs/CHANGELOG-dev.md, docs/facts/) as a safeguard.
- In a cloud session the maintainer uploads them at the start; the session keeps them outside the repository (for example in its scratchpad). At the end the session gives the changed files back to the maintainer (for example as a .zip) and does not commit them.
- Without them (for example a session where they were not uploaded), ask the maintainer for them before work that needs a decision, a fact or the plan.
- Code comments may cite their identifiers (D-*, F-*, E-*) and their former paths (docs/decisions.md, docs/facts/*.md); public texts (READMEs, docs shipped in the package, CHANGELOG.md, SECURITY.md) do not.
