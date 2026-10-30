# Security policy

xlvane opens spreadsheet files from unknown sources, so every input file is treated as untrusted.

## Reporting a vulnerability

Please report vulnerabilities privately through GitHub private vulnerability reporting: on the repository page open the **Security and quality** tab and choose **Report a vulnerability**. Please do not open a public issue for a security problem.

A useful report contains the xlvane version (or commit), the runtime (Node.js, Bun, Deno, browser, Workers) and, if possible, a file or a script that reproduces the problem. The maintainer will acknowledge the report and agree on a disclosure timeline with you.

## Supported versions

No version has been released yet. Once 1.0 is released, this section will list the versions that receive security fixes.

## What the library defends against

| Area | Protection |
|---|---|
| ZIP | Limits on the number of entries, on the decompressed size of each entry and of all entries read, and on the compression ratio. Sizes are measured on the bytes actually decompressed, not taken from ZIP headers. Overlapping entries, duplicate entry names, encrypted entries, unsupported compression methods and multi-disk archives are rejected. CRC-32 is verified. |
| XML | Own parser. Document type declarations (DTDs) are rejected, which rules out external entities (XXE) and entity expansion ("billion laughs"). Only the five predefined entities and character references are recognised. Limits on nesting depth, name length, attribute count and text length. The parser is iterative, so deep nesting cannot overflow the call stack. |
| Data | Values read from a file never become keys of plain objects. Lookups use `Map` or arrays, so names such as `__proto__`, `constructor` and `prototype` cannot change any object prototype. |
| Regular expressions | Untrusted text is parsed by hand-written scanners instead of backtracking regular expressions; tests check worst-case inputs. |
| CSV export | Formula-injection protection (planned for phase 3 of docs/PLAN.md). |

All limits can be changed by the caller; the defaults are documented in the package README.

## Out of scope

- xlvane does not evaluate formulas. Cached formula results are returned as stored in the file.
- Macros (VBA projects) are never executed. `inspect()` reports their presence.
- External links are never followed. `inspect()` reports their presence.
- Password-protected (encrypted) workbooks and legacy `.xls` files are not supported and are rejected with an error.
