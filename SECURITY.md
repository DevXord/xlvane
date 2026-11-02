# Security policy

xlvane opens spreadsheet files from unknown sources, so every input file is treated as untrusted.

## Reporting a vulnerability

Please report vulnerabilities privately through GitHub private vulnerability reporting: on the repository page open the **Security and quality** tab and choose **Report a vulnerability**. Please do not open a public issue for a security problem.

A useful report contains the xlvane version (or commit), the runtime (Node.js, Bun, Deno, browser, Workers) and, if possible, a file or a script that reproduces the problem. The maintainer will acknowledge the report and agree on a disclosure timeline with you.

## Supported versions

Security fixes are released for the newest minor version of `xlvane` and `@xlvane/compat` (currently 1.1.x); please update to it. Releases are published from GitHub Actions with npm provenance (since 1.0.1), and the changes are listed in [CHANGELOG.md](CHANGELOG.md).

## What the library defends against

| Area | Protection |
|---|---|
| ZIP | Limits on the number of entries, on the decompressed size of each entry and of all entries read, and on the compression ratio. Sizes are measured on the bytes actually decompressed, not taken from ZIP headers. Overlapping entries, duplicate entry names, encrypted entries, unsupported compression methods and multi-disk archives are rejected. CRC-32 is verified. |
| XML | Own parser. Document type declarations (DTDs) are rejected, which rules out external entities (XXE) and entity expansion ("billion laughs"). Only the five predefined entities and character references are recognised. Limits on nesting depth, name length, attribute count and text length. The parser is iterative, so deep nesting cannot overflow the call stack. |
| Data | Inside the library, values read from a file never become keys of plain objects: lookups use `Map` or arrays, so names such as `__proto__`, `constructor` and `prototype` cannot change any object prototype. The row objects of `readTable`, `readTableStream` and `toJSON` are keyed by the header text of the file; their properties are defined with `Object.defineProperty`, so a header `__proto__` is an ordinary property, but code that copies such rows (for example `Object.assign`) or calls methods on them (such as `row.hasOwnProperty`) should treat the keys as untrusted. |
| Amplification | Work derived from a small file is bounded, not only the file itself: output positions of `toCSV`, `readTable`, `toJSON` and `readTableStream` (at most `maxCells` or 16 times the cells read), cells created by `@xlvane/compat` for merged ranges, tables and columns, hyperlinks resolved to cells, the expanded text of shared formulas, and a sheet part shared by several sheets (read once). Parsing, streaming from a `Blob`, row changes and reading style-only cells run in time linear in the input; tests check worst-case inputs (docs/decisions.md D-H01). |
| Regular expressions | Untrusted text is parsed by hand-written scanners instead of backtracking regular expressions; tests check worst-case inputs. |
| CSV export | `toCSV`, `xlvane convert` and `workbook.csv` of `@xlvane/compat` write text that starts with `=`, `+`, `-`, `@`, a tab, CR or LF, the full-width `＝`, `＋`, `－`, `＠`, or (as a precaution) another control character or the small `﹦`, `﹢`, `﹣`, `﹫`, with a leading `'`, so spreadsheet applications show it as text instead of running it as a formula (OWASP CSV Injection; docs/facts/csv.md CSV-10, CSV-12). `formulaInjection: "off"` writes the text unchanged. |
| Command line | `xlvane migrate` does not follow symbolic links and changes only the `exceljs` entry of the dependency sections. Human-readable output shows text from files with control and bidirectional characters escaped, so a cell cannot move the cursor, change the window title, add colours or reorder the output. `--json` prints the same data, with control and bidirectional characters as `\uXXXX` escapes. Files are never replaced without `--force`. |

| MCP server (`@xlvane/mcp`) | Tools read and write only inside the `--root` directories: real paths are checked after symbolic links are followed, a file is checked again after it is opened (same device and inode), writes create files exclusively or replace them through a temporary file only with `"overwrite": true`, and the write tools exist only with `--allow-write`. Files are limited in size (`--max-file-size`) and tool results in length (`--max-response-chars`). Cell text is returned as data and may contain any text, including text written to influence a model; the server does not interpret it. A local process that can change the root directories while the server runs can still race a write (the directory of a new file is checked before writing). |

All limits can be changed by the caller; the defaults are documented in the package README.

## Out of scope

- xlvane does not evaluate formulas. Cached formula results are returned as stored in the file.
- Macros (VBA projects) are never executed. `inspect()` reports their presence.
- External links are never followed. `inspect()` reports their presence.
- Password-protected (encrypted) workbooks and legacy `.xls` files are not supported and are rejected with an error.
