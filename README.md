# xlvane

An MIT-licensed library for reading and writing XLSX files, designed for input from untrusted sources.

**Status:** version 1.0.0. Install with `npm install xlvane` (and `npm install @xlvane/exceljs` for the exceljs-compatible API). Changes for users are in [CHANGELOG.md](CHANGELOG.md); how releases are made is in [docs/RELEASING.md](docs/RELEASING.md).

## Packages

| Package | Contents | Runtime dependencies |
|---|---|---|
| [`xlvane`](packages/xlvane) | Core library (safe reading, `inspect()`, writing with styles, task API, streaming) and the `xlvane` command line | `fflate` |
| [`@xlvane/exceljs`](packages/exceljs) | exceljs-compatible API ([docs/exceljs-compat.md](docs/exceljs-compat.md)) | `xlvane` |

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

## Agent Skill and offline docs

- [skills/xlvane/SKILL.md](skills/xlvane/SKILL.md) is an [Agent Skill](https://agentskills.io) with instructions and examples. Install it into a project with one command of the [`skills`](https://github.com/vercel-labs/skills) CLI: `npx skills add ./node_modules/xlvane/skills --skill xlvane` after `npm install xlvane` (works offline), or `npx skills add <path-to-a-clone-of-this-repository>/skills --skill xlvane`.
- The `xlvane` package also ships its documentation (`docs/`) and `llms-full.txt` (README, skill and docs in one file), so it is readable without network access.

## Development

Requires Node.js 22.18+, 24.11+ or 26+ (the range supported by the build and test tools).

```sh
npm install
npm test            # unit, security, policy and corpus tests
npm run fuzz        # property-based tests with many runs
npm run typecheck   # TypeScript
npm run check:pkg   # build, publint, attw and package policy checks
```

Project rules are in [CLAUDE.md](CLAUDE.md), decisions and verified format facts in [docs/decisions.md](docs/decisions.md), and the security policy in [SECURITY.md](SECURITY.md).

## License

MIT
