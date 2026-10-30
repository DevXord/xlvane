# xlvane

An MIT-licensed library for reading and writing XLSX files, designed for input from untrusted sources.

**Status:** early development, nothing is published to npm yet. The work follows [docs/PLAN.md](docs/PLAN.md) phase by phase; see [docs/CHANGELOG-dev.md](docs/CHANGELOG-dev.md) for what is done.

## Packages

| Package | Contents | Runtime dependencies |
|---|---|---|
| [`xlvane`](packages/xlvane) | Core library: safe reading, `inspect()`, writing with styles | `fflate` |
| [`@xlvane/exceljs`](packages/exceljs) | exceljs-compatible API (phase 5, not implemented) | `xlvane` |
| [`@xlvane/mcp`](packages/mcp) | MCP server (phase 7, not implemented) | `xlvane` |

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
