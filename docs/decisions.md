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
