# Facts for CSV, Standard Schema and the task API (phase 3)

Ids are cited from `packages/xlvane/src/csv.ts`, `table.ts` and `standard-schema.ts`.

**Status (2026-10-30).** RFC 4180 could not be downloaded in the cloud session (rfc-editor.org and ietf.org were refused by the network proxy); the maintainer will supply it. Facts marked **PROVISIONAL** describe what the code assumes and must be checked against RFC 4180 when it is available.

| Id | Source | Version |
|---|---|---|
| S-OWASP | OWASP "CSV Injection", `pages/attacks/CSV_Injection.md` in github.com/OWASP/www-community | branch `master` at 4ce6d55dc4a476504d430a2e74238c388401a9c9 |
| S-SS | Standard Schema spec, `packages/spec/src/index.ts` in github.com/standard-schema/standard-schema (MIT, Copyright (c) 2024 Colin McDonnell) | branch `main` at 86764d849573fc9cbb7cc2f2195d224f22b812c7; npm `@standard-schema/spec` 1.1.0 |
| S-LO | LibreOffice help "CSV Filter Options" (docs/facts/sml-write.md F-LO-01) | see there |
| S-NPM | npm registry metadata | read 2026-10-30 |

## CSV format

### CSV-01 — Records and line breaks — PROVISIONAL
Records are separated by CRLF; the last record may end with a line break. `toCSV` writes CRLF by default (option `"\n"`). `fromCSV` accepts CRLF, LF and lone CR as record ends, and a final line break does not start another record.

### CSV-02 — Fields and quoting — PROVISIONAL
Fields are separated by the delimiter (RFC 4180: comma). A field that contains the delimiter, a double quote, CR or LF is enclosed in double quotes, and a double quote inside it is written as two double quotes. Spaces belong to the field. `toCSV` quotes exactly these fields (plus escaped formula-like text, CSV-10).

### CSV-03 — Malformed input — PROVISIONAL / implementation choice
RFC 4180 has no rule for a quoted field that is never closed or for text after a closing quote; `fromCSV` rejects both with `XLV_E_CSV_MALFORMED` (line and column). A double quote inside an unquoted field is kept as text (lenient, as common readers do).

### CSV-04 — Header line — PROVISIONAL
RFC 4180 describes an optional header line with the same format as records. `fromCSV` treats the first line as header by default (`header: true`).

### CSV-05 — Character set
`fromCSV` decodes bytes as UTF-8 with `fatal: true` (WHATWG Encoding, `TextDecoder`), strips a leading U+FEFF, and rejects invalid sequences. `toCSV` can start with U+FEFF (`bom: true`); LibreOffice's CSV filter has a "byte-order mark" export token (S-LO token 14).

## Formula injection

### CSV-10 — Characters that start a formula
S-OWASP lists cells beginning with `=`, `+`, `-`, `@`, tab (0x09), carriage return (0x0D), line feed (0x0A), and full-width variants such as `＝`, `＋`, `－`, `＠` (Japanese locales), and recommends as sanitization: wrap the field in double quotes, prepend a single quote, and escape double quotes by doubling them; example `=1+2";=1+2` → `"'=1+2"";=1+2"`. Quote (anchor): "Prepend each cell field with a single quote".

### CSV-11 — Limits of the mitigation
S-OWASP warns that Microsoft Excel may drop the quote when a CSV is saved and reopened, so the escape is not reliable after such a round trip, and mentions a tab-prefix variant observed to work in Excel. xlvane uses the single-quote form (CSV-10), documents the caveat and offers `formulaInjection: "off"`. Escaping applies only to text cells; numbers such as `-5` are written unchanged.

## Standard Schema

### SS-01 — Interface
A schema exposes `~standard` with `version: 1`, `vendor: string`, `validate(value, options?)` returning `Result | Promise<Result>`, and optional `types: { input, output }`. A result is `{ value }` on success or `{ issues }` on failure; an issue has `message` and an optional `path` of `PropertyKey | { key: PropertyKey }`. `InferOutput<S>` is `NonNullable<S['~standard']['types']>['output']`. Source: S-SS L1–80. xlvane copies these declarations (flattening `StandardTypedV1.Props` into `Props`; structurally identical).

### SS-02 — Library versions used in tests
zod 4.6.7, valibot 1.5.0, arktype 2.2.9 (MIT; arktype depends on `@ark/util`, `@ark/schema`, `arkregex`). Source: S-NPM `npm view <name>@latest`. Each implements `~standard` (verified by the tests in tests/unit/table.test.ts and tests/types/schemas.test.ts, which use the libraries' own schemas).
