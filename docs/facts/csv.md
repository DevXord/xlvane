# Facts for CSV, Standard Schema and the task API (phase 3)

Ids are cited from `packages/xlvane/src/csv.ts`, `table.ts` and `standard-schema.ts`.

**Status (2026-10-30).** RFC 4180 could not be downloaded in the cloud session (rfc-editor.org and ietf.org were refused by the network proxy); the maintainer will supply it. Facts that were marked PROVISIONAL were checked against RFC 4180 when the maintainer supplied it (update below); none is open.

**Update (2026-10-30).** The maintainer supplied the full text of RFC 4180 (Y. Shafranovich, "Common Format and MIME Type for Comma-Separated Values (CSV) Files", October 2005, Informational) in the session; CSV-01 … CSV-04 are checked against it (S-RFC). The ECMA-376 archives (Parts 2, 3 and 4) do not cover CSV.

| Id | Source | Version |
|---|---|---|
| S-OWASP | OWASP "CSV Injection", `pages/attacks/CSV_Injection.md` in github.com/OWASP/www-community | branch `master` at 4ce6d55dc4a476504d430a2e74238c388401a9c9 |
| S-SS | Standard Schema spec, `packages/spec/src/index.ts` in github.com/standard-schema/standard-schema (MIT, Copyright (c) 2024 Colin McDonnell) | branch `main` at 86764d849573fc9cbb7cc2f2195d224f22b812c7; npm `@standard-schema/spec` 1.1.0 |
| S-LO | LibreOffice help "CSV Filter Options" (docs/facts/sml-write.md F-LO-01) | see there |
| S-NPM | npm registry metadata | read 2026-10-30 |
| S-RFC | RFC 4180, text supplied by the maintainer | October 2005 |

## CSV format

### CSV-01 - Records and line breaks - VERIFIED
Records are separated by CRLF; the last record may end with a line break. `toCSV` writes CRLF by default (option `"\n"`). `fromCSV` accepts CRLF, LF and lone CR as record ends, and a final line break does not start another record. S-RFC §2 rule 1 (p.2): "Each record is located on a separate line, delimited by a line break (CRLF)"; rule 2: "The last record in the file may or may not have an ending line break". Accepting LF and CR follows §3 (p.4–5): "implementors should be aware that some implementations may use other values" and "be liberal in what you accept from others".

### CSV-02 - Fields and quoting - VERIFIED
Fields are separated by the delimiter (RFC 4180: comma). A field that contains the delimiter, a double quote, CR or LF is enclosed in double quotes, and a double quote inside it is written as two double quotes. Spaces belong to the field. `toCSV` quotes exactly these fields (plus escaped formula-like text, CSV-10). S-RFC §2 rule 4 (p.3): fields "separated by commas", "Spaces are considered part of a field and should not be ignored", "The last field in the record must not be followed by a comma"; rule 6: "Fields containing line breaks (CRLF), double quotes, and commas should be enclosed in double-quotes"; rule 7: a double quote inside is "escaped by preceding it with another double quote". Quoting a field with a lone CR or LF goes beyond rule 6 (which names CRLF); it keeps such fields intact. Other delimiters (option `delimiter`) are outside RFC 4180, which defines only the comma.

### CSV-03 - Malformed input - VERIFIED (no rule) / implementation choice
RFC 4180 has no rule for a quoted field that is never closed or for text after a closing quote; `fromCSV` rejects both with `XLV_E_CSV_MALFORMED` (line and column). A double quote inside an unquoted field is kept as text (lenient, as common readers do). S-RFC defines no error handling; its ABNF (§2, p.3–4) requires `escaped = DQUOTE *(TEXTDATA / COMMA / CR / LF / 2DQUOTE) DQUOTE` and allows nothing between a closing quote and the next COMMA or CRLF, so both rejected cases do not match the grammar. Rule 5 says that in unquoted fields "double quotes may not appear"; keeping them is a deliberate leniency (§3: "be liberal in what you accept").

### CSV-04 - Header line - VERIFIED / implementation choice
RFC 4180 describes an optional header line with the same format as records. `fromCSV` treats the first line as header by default (`header: true`). S-RFC §2 rule 3 (p.2): "There maybe an optional header line appearing as the first line of the file with the same format as normal record lines"; §3 (p.4): the MIME parameter `header` is "present" or "absent", and "Implementors choosing not to use this parameter must make their own decisions as to whether the header line is present or absent". The default is such a decision.

### CSV-05 - Character set
`fromCSV` decodes bytes as UTF-8 with `fatal: true` (WHATWG Encoding, `TextDecoder`), strips a leading U+FEFF, and rejects invalid sequences. `toCSV` can start with U+FEFF (`bom: true`); LibreOffice's CSV filter has a "byte-order mark" export token (S-LO token 14). S-RFC §3 (p.4): "Common usage of CSV is US-ASCII, but other character sets defined by IANA for the \"text\" tree may be used in conjunction with the \"charset\" parameter"; its ABNF `TEXTDATA` covers only %x20-7E. xlvane uses UTF-8 (charset=utf-8), beyond the ASCII grammar, as spreadsheet data needs.

## Formula injection

### CSV-10 - Characters that start a formula
S-OWASP lists cells beginning with `=`, `+`, `-`, `@`, tab (0x09), carriage return (0x0D), line feed (0x0A), and full-width variants such as `＝`, `＋`, `－`, `＠` (Japanese locales), and recommends as sanitization: wrap the field in double quotes, prepend a single quote, and escape double quotes by doubling them; example `=1+2";=1+2` → `"'=1+2"";=1+2"`. Quote (anchor): "Prepend each cell field with a single quote".

### CSV-11 - Limits of the mitigation
S-OWASP warns that Microsoft Excel may drop the quote when a CSV is saved and reopened, so the escape is not reliable after such a round trip, and mentions a tab-prefix variant observed to work in Excel. xlvane uses the single-quote form (CSV-10), documents the caveat and offers `formulaInjection: "off"`. Escaping applies only to text cells; numbers such as `-5` are written unchanged.

## Standard Schema

### SS-01 - Interface
A schema exposes `~standard` with `version: 1`, `vendor: string`, `validate(value, options?)` returning `Result | Promise<Result>`, and optional `types: { input, output }`. A result is `{ value }` on success or `{ issues }` on failure; an issue has `message` and an optional `path` of `PropertyKey | { key: PropertyKey }`. `InferOutput<S>` is `NonNullable<S['~standard']['types']>['output']`. Source: S-SS L1–80. xlvane copies these declarations (flattening `StandardTypedV1.Props` into `Props`; structurally identical).

### SS-02 - Library versions used in tests
zod 4.6.7, valibot 1.5.0, arktype 2.2.9 (MIT; arktype depends on `@ark/util`, `@ark/schema`, `arkregex`). Source: S-NPM `npm view <name>@latest`. Each implements `~standard` (verified by the tests in tests/unit/table.test.ts and tests/types/schemas.test.ts, which use the libraries' own schemas).
