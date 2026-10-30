# Formula facts for shared-formula expansion (phase 3)

Ids `F2-*` are cited from `packages/xlvane/src/sml/formula-shift.ts`. The formula grammar of ECMA-376 Part 1 §18.17 was **not available** in the session that wrote this (the ECMA-376 documents could not be downloaded, docs/decisions.md D-203). Facts that rest only on phase 1 facts are marked VERIFIED; all others are **PROVISIONAL** and must be checked against §18.17 when the documents are supplied.

### F2-01 — What a shared formula is — VERIFIED
The master cell holds the text with `t="shared"`, `ref` and `si`; dependent cells carry `t="shared"` and `si` and get the master formula "derived from their position relative to the master"; text written in a dependent is ignored; a cell outside the master's `ref` is implementation-defined. Source: docs/facts/ecma-376.md E-28 (Part 1 §18.3.1.40). xlvane expands only dependents inside the master's range.

### F2-02 — A1 references — VERIFIED
`A1-column A1-row`, `A1-column ":" A1-column` and `A1-row ":" A1-row`; a column is letters, a row a non-zero digit sequence, each optionally prefixed by `$`; column letters are not case-sensitive; references are always stored in A1 form in the file, whatever the display mode. Source: E-37 (§18.17.2.3.1, §18.17.6.1). Grid: A1 to XFD1048576 (E-39).

### F2-03 — Moving a formula — PROVISIONAL
Moving a formula by (r, c) adds r to every row number and c to every column number that is not marked with `$`; `$`-marked parts stay. This is the meaning of "relative" and "absolute" references as used in E-28 ("derived from their position relative to the master") and E-37 (`$` prefix); the exact rule is in §18.17.2.3 and was not read. A reference that would leave the grid makes xlvane refuse the expansion instead of writing `#REF!` (implementation choice).

### F2-04 — String literals — PROVISIONAL
`"…"` with `""` for a quote character. Text inside strings is never a reference.

### F2-05 — Names and functions — PROVISIONAL
A name directly followed by `(` is a function call (including names such as `LOG10` that would otherwise read as a cell reference; future-function prefixes such as `_xlfn.` are part of the name). Other identifiers that are not A1 references are defined names or the constants TRUE/FALSE and are copied. Identifier characters assumed: letters, digits, `_`, `.`, `\`, `?`, and non-ASCII characters.

### F2-06 — Bracketed text — PROVISIONAL
`[...]` (structured references such as `Table1[[#Headers],[Col]]`, external workbook indexes such as `[1]`) is copied unchanged; brackets nest and `'` escapes the next character inside them.

### F2-07 — Operators and separators — PROVISIONAL
`+ - * / ^ & = < > % , ; ( ) { } : @`, space (intersection) and line breaks are copied. `{…}` array constants contain only constants.

### F2-08 — Errors and the spill operator — VERIFIED / PROVISIONAL
Error constants `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, `#GETTING_DATA` are VERIFIED (E-33). `#` after an operand (`A1#`, dynamic-array spill reference) is PROVISIONAL (not in the 2016 edition as far as checked; it appears in files from current Excel).

### F2-09 — R1C1-looking names — implementation choice
Identifiers of the form `R`, `C`, `Rn`, `Cn`, `RnCn` (any case) make the formula "not shiftable". They cannot be ordinary names in A1 mode in Excel, and guessing would risk a wrong formula.

### F2-10 — Numbers — PROVISIONAL
`digits[.digits][E[+|-]digits]` and `.digits`. `1E10` is a number, `E10` a reference. A number directly followed by a letter or `:` other than a row range is rejected.

### F2-11 — Sheet-qualified references — PROVISIONAL
`'quoted name'!ref` (with `''` for an apostrophe) and `Name!ref`; 3-D ranges `Sheet1:Sheet3!ref`. An unquoted sheet name that reads as a cell reference or R1C1 reference, and a column-like 3-D range such as `Jan:Feb!A1`, make the formula "not shiftable" (they are ambiguous with references).

### F2-12 — Conservative failure
Anything else (unknown characters, unbalanced quotes or brackets, `$` inside a name, `(` after white space following a reference) makes the formula "not shiftable"; its dependents keep `formula.text = null` with `XLV_W_SHARED_FORMULA_NOT_EXPANDED`.

## To verify
Checked on 2026-10-30: the supplied ECMA-376 Part 4 (5th edition, December 2016) does not define the formula grammar; its §15.9 (printed p.221–222) covers only an attribute synonym of `c` and date representations, and refers to Part 1 for the rest. So F2-03 … F2-08 (spill operator part), F2-10 and F2-11 stay PROVISIONAL until Part 1 is supplied. To check them against ECMA-376 Part 1 §18.17.2 (syntax), §18.17.2.3 (references), §18.17.2.4/§18.17.2.5 (names, functions) and Part 4 (Transitional additions such as `_xlfn.` prefixes).
