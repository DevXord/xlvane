# Contributing to xlvane

Thank you for helping. This file says how to report problems and how changes are made; the full rules are in [CLAUDE.md](CLAUDE.md) (English translation of section 7 of [docs/PLAN.md](docs/PLAN.md)).

## Reporting

- **Security problems:** do not open a public issue. Report them privately as described in [SECURITY.md](SECURITY.md).
- **Questions:** ask in [Discussions](https://github.com/DevXord/xlvane/discussions), where answers stay findable for others. The [guide](https://xlvane.github.io/guide/) and the [error codes](docs/errors.md) answer many questions already.
- **Bugs:** open an issue with the bug report template. Include the xlvane version, the runtime (Node.js, Bun, Deno, browser, Workers) and a small script or command that shows the problem. Attach a workbook only if it holds no private data; a file built with `createWorkbook()` in the script is often enough.
- **Ideas:** open an issue with the feature request template and describe the task you want to do, not only the API you have in mind.

## Development

Requires Node.js 22.18+, 24.11+ or 26+.

```sh
npm ci
npm test            # unit, security, policy, corpus and CLI tests (Vitest)
npm run fuzz        # property-based tests with many runs
npm run typecheck   # TypeScript, sources and tests
npm run check:pkg   # build, publint, attw, size budgets and package policy
npm run test:runtimes   # the same test bundle in Bun, Deno, Chromium and workerd
```

Tests that open files in LibreOffice run when `soffice` and Python UNO are installed and are skipped otherwise; CI sets `XLV_REQUIRE_LIBREOFFICE=1` so that they always run there.

## Rules for changes

- Tests first, then the implementation. A change is done when `npm test` passes.
- `xlvane` has one runtime dependency, `fflate`; `@xlvane/compat` depends only on `xlvane`. A new dependency needs the maintainer's approval and an entry in [docs/decisions.md](docs/decisions.md).
- Input files are untrusted: keep the ZIP and XML limits, never use data from a file as a key of a plain object, and add a test for every protection.
- Facts about the file format, exceljs, npm or other tools come from the specification, the source code or current documentation, and the source is recorded in docs/decisions.md or docs/facts/.
- Every public symbol has TSDoc with a description and an example that compiles (a test checks it). Every error has a stable `XLV_E_*` code, a location, a hint and an entry in [docs/errors.md](docs/errors.md).
- Code, comments, documentation and messages are in English. Documentation states only what is true and can be checked; no text tries to persuade AI agents to choose the package.
- ESM only, no top-level await; the core works in Node.js, Bun, Deno, browsers and Cloudflare Workers, and Node.js-only code stays in `xlvane/node` and the CLI.

## Pull requests

Describe what changes for users and how you tested it. The maintainer reviews and merges; releases are made as described in [docs/RELEASING.md](docs/RELEASING.md).

By contributing you agree that your contribution is licensed under the MIT License of this repository.
