# Facts: how references change when rows or columns are inserted or deleted (phase B)

Ids `R-*` are cited from `packages/xlvane/src/sml/ref-adjust.ts` and docs/decisions.md D-B01.

**Sources.**
- ECMA-376 does not define this. It is application behaviour: Part 1 §18.17 defines the formula grammar (docs/facts/formulas.md), not what an application does to formulas when the grid changes.
- **Microsoft's documentation could not be read.** support.microsoft.com is outside this cloud session's allowed network domains: WebFetch failed with `ENOTFOUND` on 2026-10-30, for "How to correct a #REF! error" and "Insert or delete rows and columns". No Microsoft wording is quoted below. Checking these facts against Excel (documentation or files saved by Excel) is still open.
- **Observation (S-LO-REF).** LibreOffice Calc 24.2.7.2 (Ubuntu 24.04 package, the same as F-LO-01) was run headless in this container on 2026-10-30. The steps:
  1. Workbooks written by xlvane were loaded through LibreOffice's UNO API.
  2. One operation was applied: `XTableRows`/`XTableColumns` `insertByIndex`/`removeByIndex` at row 5 or column E, two rows or columns.
  3. The result was saved with the "Calc Office Open XML" filter and read back with xlvane.

  To reproduce it:
  - script: `tests/libreoffice/ref-adjust-facts.py`;
  - test: `tests/libreoffice/ref-adjust-facts.test.ts`, which re-runs the observation whenever soffice and Python UNO are installed;
  - recorded results: `tests/libreoffice/ref-adjust-facts.cases.ts`, with 166 formula cases plus the merges and defined names asserted in the test.

  Every rule below is an observation of LibreOffice, not of Excel.

In the examples, two rows are inserted before row 5 ("insert"), or rows 5–6 are deleted ("delete"), on sheet Main. Columns behave the same way at column E.

### R-01 — Every reference at or after the change moves
On insert, a reference to row ≥ 5 moves down by 2: `B9` → `B11`. On delete, a reference to a row after the deleted band moves up by 2: `B9` → `B7`. References before the change stay.

### R-02 — Absolute references move too
`$B$9*2` → `$B$11*2`. `SUM(B$5:$B8)` → `SUM(B$7:$B10)`. `$` marks only what stays fixed when a formula is copied; it does not stop the reference from following its cells.

### R-03 — A range that the insertion falls into grows
- `SUM(B2:B10)` → `SUM(B2:B12)`.
- `SUM(B1:B5)` → `SUM(B1:B7)`: the insertion happens before row 5, which is inside the range.
- `SUM(B4:B5)` → `SUM(B4:B7)`.
- Whole rows: `SUM(3:9)` → `SUM(3:11)`.
- Columns: `SUM(B2:J2)` → `SUM(B2:L2)` and `SUM(C:I)` → `SUM(C:K)`.

### R-04 — A range that starts at the insertion point moves; one that ends before it stays
`SUM(B5:B8)` → `SUM(B7:B10)`. `SUM(B2:B4)` stays `SUM(B2:B4)`. The range does not grow at its bottom edge.

### R-05 — A deleted single reference becomes `#REF!`
`B5` → `#REF!`. `E2` with columns E–F deleted → `#REF!`.

### R-06 — `#REF!` replaces the whole reference, including its sheet prefix
- `'Main'!B5` on another sheet → `#REF!`.
- `SUM(Main!B5:B6)` → `SUM(#REF!)`.
- The defined name `Main!$B$5` → `#REF!`.

### R-07 — A range entirely inside the deleted band becomes `#REF!`
- `SUM(B5:B6)` → `SUM(#REF!)`.
- `SUM(5:6)` → `SUM(#REF!)`.
- Defined name `Main!$B$5:$B$6` → `#REF!`.
- Columns: `SUM(E2:F2)` → `SUM(#REF!)`.

### R-08 — A range partly inside the deleted band shrinks
- End deleted: `SUM(B4:B6)` → `SUM(B4:B4)`. The range form is kept even for one cell. Also `SUM(4:6)` → `SUM(4:4)`.
- Start deleted: `SUM(B6:B10)` → `SUM(B5:B8)`. The new start is the first row after the band, renumbered.
- Band inside the range: `SUM(B2:B10)` → `SUM(B2:B8)`.

### R-09 — Insertion at the end of the grid
- A single reference pushed beyond row 1048576 (or column XFD) becomes `#REF!`: `B1048576` → `#REF!`, `XFD2` → `#REF!`.
- A range whose start is pushed beyond the grid becomes `#REF!`: `SUM(B1048575:B1048576)` → `SUM(#REF!)`, `SUM(1048576:1048576)` → `SUM(#REF!)`, `SUM(XFC2:XFD2)` → `SUM(#REF!)`.
- A range whose end is pushed beyond the grid stops at the last row or column:
  - `SUM(B1048574:B1048575)` → `SUM(B1048576:B1048576)`;
  - `SUM(B10:B1048576)` → `SUM(B12:B1048576)`;
  - `SUM(B2:B1048576)` stays;
  - `SUM(5:1048576)` → `SUM(7:1048576)`;
  - `SUM(E:XFD)` → `SUM(G:XFD)`.

### R-10 — Deletion keeps a range end that is the last row or column
A range of more than one row or column whose end is row 1048576 (or column XFD) keeps that end:
- `SUM(B10:B1048576)` → `SUM(B8:B1048576)`;
- `SUM(B2:B1048576)` stays;
- `SUM(1048575:1048576)` → `SUM(1048573:1048576)`;
- `SUM(B5:B1048576)` and `SUM(B6:B1048576)` → `SUM(B5:B1048576)`;
- `SUM(E:XFD)` stays;
- `SUM(XFC2:XFD2)` → `SUM(XFA2:XFD2)`.

A single row, column or cell moves as usual: `SUM(1048576:1048576)` → `SUM(1048574:1048574)`, `B1048576` → `B1048574`, `SUM(XFC2:XFC2)` → `SUM(XFA2:XFA2)`.

### R-11 — Defined names follow the same rules
- Insert: `Main!$B$2:$B$10` → `Main!$B$2:$B$12`; `Main!$B$5:$B$6` → `Main!$B$7:$B$8`; `Main!$B$5` → `Main!$B$7`.
- Delete: `Main!$B$2:$B$8`, `#REF!`, `#REF!`, and `Main!$B$4:$B$6` → `Main!$B$4:$B$4`.
- Formulas that use a name (`SUM(RangeAll)`) are not changed.

### R-12 — Only references to the changed sheet change
- On sheet Other, `B9` (its own sheet) stays when Main changes.
- `Main!B9` → `Main!B11`.
- On Main, `Other!B9` stays.
- When rows are inserted on Other, `Other!B9` on Main → `Other!B11`.
- Text in strings (`INDIRECT("B9")`), names and whole columns (on a row change) or whole rows (on a column change) are not changed: `SUM(A:A)` and `SUM(1:1)` stay.
- LibreOffice writes sheet names without the quotes they do not need (`'Main'!B5` → `Main!B7`). That is formatting only.

### R-13 — 3-D references are not adjusted by LibreOffice
`SUM(Main:Other!B9)` stays (saved as `SUM(Main:Other!B9:B9)`) after rows are inserted into or deleted from Main, and after rows are inserted into Other. Excel's behaviour is not known (Microsoft's documentation was unreachable).

xlvane does not copy this (D-B01). A 3-D reference that covers the changed sheet, or whose sheets are unknown, makes the formula not adjustable (`null`).

### R-14 — Merged ranges
- **Insert:** merges change like ranges (R-03, R-04): `G4:G5` → `G4:G7`, `H4:H8` → `H4:H10`, `D5:E8` → `D7:E10`, `I6:I7` → `I8:I9`. For columns: `E4:H5` → `G4:J5`, `D7:E7` → `D7:G7`.
- **Delete:**
  - A merge whose top-left cell is deleted is dropped, even though part of it remains: `D5:E8` with rows 5–6 deleted, and `E4:H5` with columns E–F deleted.
  - A merge that would shrink to one cell is dropped: `G4:G5`, `I6:I7`, `D7:E7`.
  - Otherwise it shrinks: `H4:H8` → `H4:H6`.

  This is a rule for merges only. Formula ranges whose start is deleted shrink instead (R-08). @xlvane/compat follows this rule for merged ranges too (D-B03 addendum).

### R-15 — A reversed range is written in order
Observed 2026-10-30 with `tests/libreoffice/ref-adjust-facts.py` on a separate workbook (formulas in Main!C1:C7). LibreOffice writes a reversed range in order, whether or not the change touches it: `SUM(B9:B2)` → `SUM(B2:B7)` (delete), `SUM(B2:B11)` (insert), `SUM(B2:B9)` (a change that does not touch it); `SUM(B6:B4)` → `SUM(B4:B4)` (delete); `SUM(Other!B9:B2)` → `SUM(Other!B2:B9)` after an insert on Main. xlvane normalises a reversed range of the changed sheet before adjusting it; a reversed range on another sheet keeps its text.

### R-16 — A colon with spaces around it joins two separate references
Same observation. `SUM(B4 : B6)` → `SUM(B4 : #REF!)` and `SUM(B5 : B10)` → `SUM(#REF! : B8)` (delete), `SUM(B4 : B8)` and `SUM(B7 : B12)` (insert): each side is adjusted as a single reference, not as a range.

### R-17 — Deleting the last rows
Same observation, rows 1048575–1048576 deleted: `SUM(B1048570:B1048576)` stays (R-10 also holds when the deleted rows include the last one), `SUM(B1048575:B1048576)` → `SUM(#REF!)`, `B1048575` → `#REF!`.

### Not observed
- Conditional-formatting and data-validation `sqref`, hyperlink `ref` and comments. xlvane cannot write the first two yet. xlvane applies the formula rules to these by analogy (`adjustRange`), which is not an observation.
- Excel itself (see Sources).
