# Formula facts for shared-formula expansion (phase 3, checked 2026-10-30)

Ids `F2-*` are cited from `packages/xlvane/src/sml/formula-shift.ts`. Phase 3 wrote these without the formula grammar (D-203). On 2026-10-30 they were checked against the three grammars (D-1104):
- **P1:** ECMA-376 Part 1, 5th edition (2016), §18.17, printed p.2039–2071;
- **OI:** [MS-OI29500] rev 25.0, §2.1.859–§2.1.873, p.385–396;
- **XLSX:** [MS-XLSX] rev 30.0, §2.2.2 Formulas, p.38–64.

xlvane accepts the union of what the three grammars define, and treats anything outside them as "not shiftable" (D-1102).

### F2-01 - What a shared formula is - VERIFIED
The master cell holds the text with `t="shared"`, `ref` and `si`; dependent cells carry `t="shared"` and `si` and get the master formula "derived from their position relative to the master"; text written in a dependent is ignored; a cell outside the master's `ref` is implementation-defined. Source: docs/facts/ecma-376.md E-28 (Part 1 §18.3.1.40). xlvane expands only dependents inside the master's range.

### F2-02 - A1 references and the grid - VERIFIED
`A1-column A1-row`, `A1-column ":" A1-column` and `A1-row ":" A1-row`; a column is letters, a row a non-zero digit sequence, each optionally prefixed by `$`; column letters are not case-sensitive; references are always stored in A1 form in the file, whatever the display mode. Sources: E-37 (P1 §18.17.2.3.1, §18.17.6.1); XLSX §2.2.2 p.39–40 (`A1-relative-column` A…XFD, `row-digit-sequence` 1…1048576); P1 §18.17.5.1 p.2065. Grid: A1 to XFD1048576 (E-39). Implementation: a reference-shaped token beyond the grid (`XFE1`, `A1048577`, `A0`) makes the formula not shiftable instead of being read as a defined name.

### F2-03 - Moving a formula - PARTLY
P1 §18.17.2.3 p.2051: a relative reference is based on the relative position of the formula cell and the referenced cell and changes along with the formula cell; an absolute reference keeps referring to the same location; a mixed reference has one absolute and one relative part. P1 §18.17.2.3.1 p.2054: `$` marks the absolute part (`$A$2`, `$A2`, `B$34`). The arithmetic (adding the row and column offsets to the relative parts) is implied, not stated. OI §2.1.863 c p.392 says Office does not relate cut-and-paste *moves* to relative references; that concerns the application, not shared formulas. Implementation: a reference that would leave the grid makes xlvane refuse the expansion instead of writing `#REF!`.

### F2-04 - String literals - VERIFIED
`"…"` with `""` for a quote character. Sources: P1 §18.17.2.3 p.2051, §18.17.2.3.1 p.2054; OI §2.1.863 c p.392. Implementation: string contents are copied unchanged and are never references.

### F2-05 - Names and functions - VERIFIED against the Office grammars; P1 alone is narrower
Sources: P1 §18.17.2 p.2040, §18.17.2.4 p.2055–2056, §18.17.2.5 p.2057, §18.17.5.1 p.2065; OI §2.1.867 p.394–395; XLSX p.39–40, p.43, p.45.
- Name characters are letters, `_`, `\`, `.`, `?`, digits and anything above U+007F.
- A call is a name directly followed by `(`; XLSX `cell-function-call` is why `LOG10(` is a call while `LOG10` alone is a reference.
- The `_xlfn.`, `_xlfn._xlws.`, `_xlpm.` and `xlpm.` prefixes are part of the identifier (XLSX only).
- Other identifiers that are not references are defined names or TRUE/FALSE and are copied.

Implementation: xlvane accepts the Office character set, a superset of P1's.

### F2-06 - Bracketed text - VERIFIED
`[...]`: structured references such as `Table1[[#Headers],[Col]]` and external workbook indexes such as `[1]`. Sources: P1 §18.17.2.3 p.2052 (external `[1]`); OI §2.1.859 a p.385, §3.2.3.1.1 p.846–847; XLSX p.40–41. In structured references `'` escapes only `'`, `[`, `]` and `#`. Implementation: brackets nest, `'` escapes the next character, and the contents are copied unchanged.

### F2-07 - Operators and separators - PARTLY
- The operators `+ - * / ^ & = < > <= >= <> % , :` and space (intersection), `;` in array constants, `{}` and `()` are confirmed: P1 §18.17.2.2 p.2047; OI §2.1.862 a p.391, §2.1.860 b p.386; XLSX p.38–39.
- White space: P1 allows only U+0020; XLSX `whitespace = space / %x0D %x0A` also allows CR LF.
- `@` is in **no** grammar (NOT FOUND). Implementation: a formula containing `@` outside strings and brackets is not shiftable.
- Implementation: a lone LF or CR is accepted because XML end-of-line handling turns CR LF into LF (to check against XML 1.0 §2.11). Tab is accepted although no grammar defines it, because copying it cannot change a formula's meaning.

### F2-08 - Errors and the spill operator - VERIFIED; spill `#` NOT FOUND
Error constants `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, `#GETTING_DATA`: P1 §18.17.2.1 p.2045, §18.17.3 p.2060–2061; OI §2.1.861; XLSX p.38 `error-constant` (E-33). A spill `#` after an operand (`A1#`) is in no grammar. XLSX `#SPILL!` (§2.3.6.1.3.7 p.93) is a value, and the XLDAPR metadata (§2.2.4.4 p.73) is markup; neither is formula text. Implementation: any `#` outside strings and brackets that does not start an error constant makes the formula not shiftable.

### F2-09 - R1C1-looking names - VERIFIED
Identifiers of the form `R`, `C`, `Rn`, `Cn`, `RnCn` (any case) make the formula not shiftable. Sources: P1 §18.17.2.3.2 p.2054–2055 and §18.17.2.5 (a name shall not have the form of a cell reference, which includes R1C1); XLSX p.40 (a name must not be an `R1C1-cell-reference`).

### F2-10 - Numbers - PARTLY (the grammars differ)
- P1 §18.17.2.1 p.2045, §18.17.5.3 p.2066: exponent `E` or `e`; a trailing `.` is allowed (`210.`).
- OI §2.1.861 a: `E` or `e`, no trailing `.`.
- XLSX: `E` only.

Implementation: xlvane accepts the union (`1e5`, `1.`, `1.E5`, `.5`, `1E+10`) and copies numbers unchanged. `E10` is a reference. A number directly followed by a letter or `:` other than a row range is rejected.

### F2-11 - Sheet-qualified and 3-D references - VERIFIED
Sources: P1 §18.17.2.3 p.2049–2051 (`work-sheet-prefix`, `work-sheet-prefix-special`); OI §2.1.861 a p.390–391, §2.1.863 a p.392; XLSX p.38–39.
- `'quoted name'!ref` with `''` for an apostrophe; the name must be followed by `!`.
- `Name!ref` and 3-D `Sheet1:Sheet3!ref`.
- A quoted 3-D prefix uses one pair of apostrophes: `'Sheet 1:Sheet 3'!A1`.
- U+0003 is excluded from quoted names (OI, XLSX).

Implementation: an unquoted sheet name that reads as a cell or R1C1 reference, and an unquoted column-like 3-D range such as `Jan:Feb!A1`, make the formula not shiftable; the quoted form shifts.

### F2-12 - Conservative failure - implementation choice
Anything else (unknown characters, unbalanced quotes or brackets, `$` inside a name, `(` after white space following a reference, `@`, spill `#`, references beyond the grid) makes the formula not shiftable; its dependents keep `formula.text = null` with `XLV_W_SHARED_FORMULA_NOT_EXPANDED`, and their cached values stay available.

## Writing array and shared formulas (phase E, D-E02)

### F2-13 - Array formulas - VERIFIED
The `f` element of an array formula has `t="array"` and the range in `ref` ([MS-OI29500] §2.1.620 e, §2.1.872: Office uses `ref`, where the P1 examples write `r`); for a multi-cell array the top-left cell's `f` designates the range and **the other cells of the range shall not have an `f` element but shall each have a `v` element** (P1 §18.17.6.3, p.2068–2069; report 1 §4.8). `aca` is ignored unless `t="array"` (E-26). Implementation: `{ formula, result?, array: "B1:B3" }` is accepted only on the top-left cell of its range; the writer refuses a formula in another cell of the range and overlapping ranges (`XLV_E_INVALID_VALUE`); the other cells take their cached results as plain values. `aca` is not written.
**Checked 2026-10-31 (R5 Item 7.5, 7.8):** P1 §18.17.6.3 (p.2068–2069) confirmed: only the top-left cell has `f`, every cell of the range has `v`; the examples write `r`, which neither schema has - Office uses `ref` ([MS-OI29500] §2.1.872 a).

### F2-14 - LibreOffice 24.2.7.2 writes array formulas the same way - OBSERVED
A sheet where `B1:B3` got `setArrayFormula("A1:A3*2")` through the UNO API and was saved with the "Calc Office Open XML" filter has `<c r="B1" …><f aca="false" t="array" ref="B1:B3">A1:A3*2</f><v>2</v></c>` and `<c r="B2" …><v>4</v></c>`, `<c r="B3" …><v>6</v></c>` (observed 2026-10-30). LibreOffice reads xlvane's array formula back and computes it (tests/libreoffice/formulas.test.ts).

### F2-15 - Shared formula groups on write - VERIFIED (rules) / implementation (numbering, master position)
Rules: E-28 (P1 §18.3.1.40): the master holds the text with `t="shared"`, `ref` and `si`; dependents carry `t="shared"` and `si` and need no text; `si` is required with `t="shared"` ([MS-OI29500] §2.1.620 h); master ranges on a sheet shall not overlap; a cell outside the master's `ref` is implementation-defined. Implementation: `{ formula, result?, shared: "C1:C10" }` on the top-left cell of the range writes the master; `{ sharedFormula: "C1", result? }` writes a dependent with an empty `f`; `si` numbers the groups of a sheet from 0 in the order their masters are written (row, then column), so output is deterministic; a dependent outside its master's range, a dependent whose master is missing and overlapping array or shared ranges are refused (`XLV_E_INVALID_VALUE`). Requiring the master in the top-left cell matches the definition of the master cell for data tables in [MS-OI29500] §2.1.620 d and guarantees that the master is written before its dependents in both writers.
**Checked 2026-10-31 (R5 Item 7.4):** P1 §18.3.1.40 `si` row (raw L57929–L57956): the first formula of a group is the master; dependents need no text and their text is ignored; a cell in the range without `si`/`t` keeps its own formula; master ranges on a sheet shall not overlap. **NOT FOUND** in P1, [MS-OI29500] or [MS-OE376]: any rule on `si` numbering (start, consecutiveness, scope) and any rule that the master is the top-left cell of `ref`; every example (Annex L.2.2.4, L.2.3.7) has the master at the top-left and `si="0"`. Both remain implementation decisions.

### F2-16 - LibreOffice 24.2.7.2 and shared formulas - OBSERVED
LibreOffice reads xlvane's shared group (master and dependents without text), computes every cell, and writes the group back as one ordinary formula per cell (`A1+10`, `A2+10`, `A3+10`; observed 2026-10-30, tests/libreoffice/formulas.test.ts).
