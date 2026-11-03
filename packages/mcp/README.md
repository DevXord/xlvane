# @xlvane/mcp

An MCP (Model Context Protocol) server for XLSX files, built on [xlvane](https://www.npmjs.com/package/xlvane) and the official MCP TypeScript SDK (`@modelcontextprotocol/server` 2.4.0, protocol revision 2026-07-28). An MCP client, such as a coding agent, can inspect, read, compare and validate workbooks, and optionally write, edit and convert them, through tools instead of code.

Requires Node.js 22 or later. Dependencies: `xlvane` and `@modelcontextprotocol/server` (which brings `@modelcontextprotocol/core` and `zod`).

## Start

```sh
npx @xlvane/mcp --root ./data                 # read-only, files inside ./data
npx @xlvane/mcp --root ./data --allow-write   # also write_table, edit_cells, convert
xlvane mcp --root ./data                      # the same server, through the xlvane CLI
```

Many MCP clients take a configuration like this one; check your client's documentation for its exact format and file:

```json
{
  "mcpServers": {
    "xlvane": { "command": "npx", "args": ["@xlvane/mcp", "--root", "/absolute/path/to/data"] }
  }
}
```

| Option | Meaning |
|---|---|
| `--root <dir>` | A directory the tools may use; repeat for several. Required. Relative paths in tool calls are resolved against the first one. |
| `--allow-write` | Also offer the tools that write files. Without it they are not registered at all. |
| `--max-file-size <bytes>` | Largest file read or written. Default 64 MiB. |
| `--max-rows <n>` | Largest number of rows one `read_table` call returns. Default 1000. |
| `--max-response-chars <n>` | Largest tool result (also an error), in characters of its JSON text; the message carries the same data again as `structuredContent`. `read_table` returns the rows that fit (`truncated`, `nextOffset`); other tools return `XLV_E_LIMIT_EXCEEDED`. Default 1 Mi (1 048 576). |
| `--max-cells <n>` | Largest number of cells in a workbook a tool reads (the `maxCells` read limit). Default 5 000 000. Lower it when the server must answer quickly: `diff` holds both workbooks in memory. |

The server speaks MCP over stdin and stdout; messages for people go to stderr.

## Tools

| Tool | Writes | What it does |
|---|---|---|
| `inspect_workbook` | no | Sheets (also hidden ones), used ranges, header rows, column types with samples, merged cells, formulas, warnings. |
| `read_table` | no | Rows of a sheet as objects keyed by the header row; paging with `offset` and `limit`; dates as ISO 8601. |
| `diff` | no | Cell-by-cell differences between two workbooks. |
| `validate` | no | Read warnings, and column rules (`name`, `type`, `optional`) with the cell of each problem. |
| `write_table` | yes | Rows to a new XLSX file (report preset: bold header, frozen row, autofilter, widths). |
| `edit_cells` | yes | Set cells and append rows, keeping the rest of the file unchanged. |
| `convert` | yes | XLSX to CSV/TSV/JSON and back; CSV output escapes formula-like text. |

Every tool has a JSON Schema for its input and returns its result as `structuredContent` and as JSON text. Errors are tool results with `isError: true` and the `XlvError` data: a stable `code`, the `location`, a `hint` and `details` (docs/errors.md).

## Security

- **Directories:** files are read and written only inside the `--root` directories. Paths are checked after symbolic links are resolved, so `..`, absolute paths and links that lead outside are refused (`XLV_E_PATH_NOT_ALLOWED`). A file is checked again after it is opened (the path must still lead to the same file), so swapping it for a link meanwhile does not help. Writing through a symbolic link is refused. A path outside the roots gets the same answer whether or not something exists there, also when it is reached through a symbolic link inside a root. Only regular files are read. Messages on stderr show control characters as `\uXXXX`.
- **Read-only by default:** the write tools exist only with `--allow-write`. They create files exclusively and replace an existing file only when the call passes `"overwrite": true`, through a temporary file and a rename; the replacement keeps the file's permission bits (not setuid, setgid or sticky), and read-only files are not replaced. `edit_cells` writes formulas as given (for example `HYPERLINK(...)`); they are calculated when a spreadsheet application opens the file. Any file name inside a root may be written, also dot files such as `.gitignore`; choose the roots accordingly.
- **Limits:** files above `--max-file-size` are refused for reading and writing (`convert` to JSON stops building the text at that size), workbooks above `--max-cells` cells are refused, and results above `--max-response-chars` are cut (`read_table`) or refused. The official MCP SDK closes the connection after a message larger than 10 MiB. The xlvane reader applies its limits for untrusted files (ZIP bombs, XML without DTDs or entities, the budgets of xlvane 1.1.1). `validate` lists at most 1 000 problems (`problemsTotal` counts all).
- **File content is data:** values, sheet names and other text come from the files and may contain any text, also text meant to influence a model. The server returns it as data and does not interpret it.
- **No code execution:** macros are never run, external links are never followed, and formulas are not calculated.

## API

```ts
import { createServer, runCli } from '@xlvane/mcp';
```

`createServer({ roots, allowWrite, maxFileBytes, maxRows, maxResponseChars })` returns an `McpServer` to connect to any transport; `runCli(args)` is the command line.

## Links

- Changes in each version: [CHANGELOG.md](https://github.com/DevXord/xlvane/blob/main/CHANGELOG.md)
- Bugs and questions: [GitHub issues](https://github.com/DevXord/xlvane/issues)
- Reporting a vulnerability: [SECURITY.md](https://github.com/DevXord/xlvane/blob/main/SECURITY.md)

License: MIT. The MCP SDK is licensed under Apache-2.0.
