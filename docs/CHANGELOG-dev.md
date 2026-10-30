# Development changelog

One entry per plan phase (docs/PLAN.md), newest first.

## Phase 0: monorepo skeleton (2026-10-30)

- npm workspaces: `packages/xlvane` (`xlvane`), `packages/exceljs` (`@xlvane/exceljs`), `packages/mcp` (`@xlvane/mcp`), each ESM-only, MIT, `"private": true` until the release phase, with empty entry modules.
- Tooling (D-002 … D-005): tsdown 0.23.4 (build, declarations via oxc with `isolatedDeclarations`), Vitest 5.0.5 with vite 8.3.5, fast-check 4.10.2, publint 0.3.25, @arethetypeswrong/cli 0.18.5, TypeScript 7.0.2. No @types/node; a small own subset in `types/` (D-006).
- `npm run check:pkg` builds every package and runs `publint --strict`, `attw --pack --profile esm-only` and a `require()` load of each built file, which fails on top-level await (checked: Node.js 26 throws `ERR_REQUIRE_ASYNC_MODULE` for a module with top-level await).
- Policy tests (`tests/policy/packages.test.ts`): allowed runtime dependencies per package, ESM/MIT/LICENSE/`files`, no `require` export condition, no unapproved root development dependencies, no Node.js built-ins in core sources.
- Documents: CLAUDE.md (English translation of plan section 7), docs/PLAN.md (Polish original), docs/decisions.md, SECURITY.md, README files, tests/corpus/LICENSES.md.
- Not done: depgate policy file (depgate is unpublished on npm, D-008).
- Environment note: the system Node.js on the development machine is 25.8.1 (end-of-life, outside the tool engine ranges). This session used a portable Node.js 26.13.0 LTS (SHA-256 verified against nodejs.org SHASUMS256.txt) in a temporary directory.
