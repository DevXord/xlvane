# `xlvane --json`, version 1

Every `xlvane` command accepts `--json`. The command then prints exactly one JSON document to stdout, without colour. Progress lines and notes, which normally go to stderr, are not printed. The exit codes are the same as without `--json`. `--help` and `--version` always print text.

The schemas are JSON Schema (draft 2020-12) files in [docs/cli-json/](cli-json/). `tests/cli/json.test.ts` checks the output of every command against them. They use only this subset of JSON Schema:
- `type`, `properties`, `required`, `additionalProperties`, `items`;
- `enum`, `const`, `oneOf`, `anyOf`, `minimum`, `pattern`;
- `$ref` to `$defs`.

The field `xlvane` changes only when a field is removed or changes meaning; new fields may be added within a version.

## Envelope

On success:

```json
{ "xlvane": 1, "command": "read", "ok": true, "result": { } }
```

On failure:

```json
{
  "xlvane": 1,
  "command": "read",
  "ok": false,
  "error": {
    "name": "XlvError",
    "code": "XLV_E_SHEET_NOT_FOUND",
    "message": "The workbook has no sheet named \"Salse\".",
    "location": { "sheet": "Salse" },
    "hint": "Did you mean \"Sales\"? Available sheets: \"Sales\", \"Costs\". Sheet names are case-sensitive.",
    "docsUrl": "https://github.com/DevXord/xlvane/blob/main/docs/errors.md#xlv_e_sheet_not_found",
    "details": { "available": ["Sales", "Costs"], "path": "report.xlsx" }
  },
  "try": "xlvane read report.xlsx --sheet Sales --json"
}
```

Field notes:
- **`error`:** the same data as `XlvError.toJSON()` in the library. Codes are listed in [errors.md](errors.md). When the error comes from a file, `details.path` names it.
- **`try`:** a ready-to-run command that fixes the error when there is an obvious one (for example the closest sheet name, or `--force`), otherwise `null`. It is quoted for POSIX shells.
- **`command`:** `null` when the command itself is unknown.

The error envelope is defined in [cli-json/common.schema.json](cli-json/common.schema.json) (`$defs/errorEnvelope`).

## Exit codes

| Code | Meaning |
|---|---|
| 0 | Success. |
| 1 | The command ran, and a check found something: `diff` found differences, or `validate` found problems. |
| 2 | Usage error (`XLV_E_USAGE`): unknown command or option, missing or extra argument, invalid option value. |
| 3 | Input or file error: the file cannot be read, is not a valid workbook, the sheet does not exist, the output exists without `--force`, and so on. |

## Commands

| Command | Schema | `result` |
|---|---|---|
| `inspect` | [cli-json/inspect.schema.json](cli-json/inspect.schema.json) | The `inspect()` report ([inspect-json.md](inspect-json.md)). With `--sheet`, only that sheet and its warnings. |
| `read` | [cli-json/read.schema.json](cli-json/read.schema.json) | `sheet`, `headerRow`, `range`, `columns` (`key`, `header`, `letter`), `rows` (objects keyed by column key; dates as ISO 8601 strings in UTC), `rowNumbers`, `totalRows`, `truncated`. Without `--limit` every row is included. |
| `convert` | [cli-json/convert.schema.json](cli-json/convert.schema.json) | `input`, `output`, `from`, `to`, `bytes`, `written` (`false` with `--dry-run`), `replaced`. |
| `create` | [cli-json/create.schema.json](cli-json/create.schema.json) | Same as `convert`. |
| `edit` | [cli-json/edit.schema.json](cli-json/edit.schema.json) | `input`, `output`, `rowsInserted`, `rowsDeleted`, `rowSheet` (the sheet of the row options, or `null`), `cellsSet`, `rowsAppended`, `appendSheet`, `warnings`, `bytes`, `written`, `replaced`. |
| `diff` | [cli-json/diff.schema.json](cli-json/diff.schema.json) | The `diff()` result: `equal`, `sheetsAdded`, `sheetsRemoved`, `sheets` (cell changes with `before` and `after`), `summary`. Every change is included. |
| `validate` | [cli-json/validate.schema.json](cli-json/validate.schema.json) | `file`, `ok`, `sheet`, `rowsChecked`, `problems` (`kind`, `column`, `cell`, `row`, `message`), `warnings`. |
| `migrate` | [cli-json/migrate.schema.json](cli-json/migrate.schema.json) | `directory`, `packageJson`, `changes` (`section`, `from`, `to`), `alreadyMigrated`, `enginesNode`, `filesScanned`, `scanLimited`, `findings` (`file`, `line`, `column`, `api`, `status`), `written`. |
