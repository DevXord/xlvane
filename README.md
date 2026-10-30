# xlvane

An MIT-licensed library for reading and writing XLSX files, designed for input from untrusted sources.

**Status:** early development, nothing is published to npm yet. The work follows [docs/PLAN.md](docs/PLAN.md) phase by phase; see [docs/CHANGELOG-dev.md](docs/CHANGELOG-dev.md) for what is done.

## Packages

| Package | Contents | Runtime dependencies |
|---|---|---|
| [`xlvane`](packages/xlvane) | Core library (safe reading, `inspect()`, writing with styles, task API, streaming) and the `xlvane` command line | `fflate` |
| [`@xlvane/exceljs`](packages/exceljs) | exceljs-compatible API ([docs/exceljs-compat.md](docs/exceljs-compat.md)) | `xlvane` |
| [`@xlvane/mcp`](packages/mcp) | MCP server (phase 7, not implemented) | `xlvane` |

## Command line in 5 minutes

```sh
npm install && npm run build
alias xlvane="node $PWD/packages/xlvane/dist/cli.js"   # until the package is published
xlvane                                   # overview and common commands
xlvane inspect report.xlsx               # sheets, columns, header rows, warnings
xlvane read report.xlsx --sheet Sales    # print a sheet as a table
xlvane convert report.xlsx sales.csv     # XLSX ⇄ CSV / JSON
xlvane create data.json report.xlsx      # formatted report from JSON or CSV
xlvane diff old.xlsx new.xlsx            # what changed
xlvane help <command>                    # usage and examples
```

Add `--json` to any command for machine-readable output ([docs/cli-json.md](docs/cli-json.md)). More in [packages/xlvane/README.md](packages/xlvane/README.md#command-line-in-5-minutes).

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
