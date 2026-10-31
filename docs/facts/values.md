# Facts fragment: cell references, serial dates, ISO date cells, number formats

> **Provenance.** Compiled on 2026-10-30. Line numbers (`Lnnn`) refer to text extracted with `pdftotext -layout` from the ECMA-376 PDFs (see docs/facts/ecma-376.md); `scratchpad/specs/…` paths are the working copies used during the check and are not stored in this repository.

Facts relied on by `packages/xlvane/src/sml/cell-ref.ts`, `dates.ts` and `numfmt.ts`.

## Conventions

- **Sources.** ECMA-376 texts in `scratchpad/specs/`:
  - `part1.txt` (Part 1, 5th ed., Dec 2016, Strict);
  - `part4.txt` (Part 4, 5th ed., Dec 2016, Transitional);
  - `xsd-strict/sml.xsd`.
  - Web sources are given by URL; all were read on 2026-10-30.
- **Locators.**
  - "p." is the printed page number from the page footer.
  - The PDF page is printed page + 10 in Part 1 and + 16 in Part 4. I checked this for Part 1 p.1776 = PDF 1786.
  - `Lnnn` is the line in the `.txt` file; run `sed -n 'NNN,+15p' part1.txt` to check.
- **Quoting.** Verbatim prose quotes are kept to a minimum (one, in V-12, at most 15 words); other entries paraphrase closely with exact locators.
  - All other prose is **paraphrased closely**. Each entry has exact line locators so the wording can be checked.
  - Grammar productions, format codes, attribute and type names, and numbers are technical tokens. They are given **verbatim** (marked "Tokens").
- **Labels.** "Derived" marks my own arithmetic or inference. "Decision" marks a behaviour I chose where the spec is silent or ambiguous.
- **Cross-reference.** Several facts repeat entries of `facts-ecma.md` (E-xx). I re-read each passage myself in the spec text.

---

## A. Cell references (`cell-ref.ts`)

### V-01 - Grid size
- **Fact:** Cell references run from A1–A1048576 to XFD1–XFD1048576. An implementation may extend this range.
  - Derived: XFD = 24·26² + 6·26 + 4 = 16 384 columns, so the grid is 1 048 576 rows × 16 384 columns.
  - This matches Microsoft's Excel limits (`facts-other.md` W-04).
- **Source:** Part 1 §18.17.5.1 Limits, p.2065 (PDF 2075), L97529–97535.
- **Tokens:** `A1`, `A1048576`, `XFD1`, `XFD1048576`.
- **Code:** `MAX_ROWS = 1_048_576`, `MAX_COLUMNS = 16_384`.

### V-02 - A1 grammar, column naming, case
- **Fact:**
  - A column is one or more letters a–z or A–Z, optionally preceded by `$`.
  - A row is a digit sequence whose value shall be non-zero, optionally preceded by `$`.
  - Columns are named A–Z, then AA–AZ, BA–BZ, …, ZA–ZZ, then AAA–AAZ, ABA–ABZ, and so on (bijective base 26).
  - Column letters are not case-sensitive.
  - A cell reference is the column letters followed immediately by the row number.
- **Source:**
  - Part 1 §18.17.2.3.1 A1-Style Cell References, p.2053–2054 (PDF 2063–2064), L96908–96955.
  - The case rule is at L96948. The `digit-sequence` production is at L96471.
- **Tokens:**
  - `A1-relative-column= letter, {letter} ;`
  - `A1-absolute-column= "$", A1-relative-column ;`
  - `A1-relative-row= digit-sequence ; (* shall be non-zero; see 3.17.2.3.1 *)`
  - `digit-sequence= decimal-digit, {decimal-digit} ;`

### V-03 - Types of the stored references
- **Fact:**
  - `c@r` (cell) has type `ST_CellRef` and is described as an A1-style reference to the cell's location.
  - `ST_CellRef` represents a single cell reference.
  - `ST_Ref` is a cell or a range of cells.
  - `ST_RefA` is a separate type: a single cell reference that may be absolute.
  - In the schema all three are `xsd:string` restrictions with **no pattern**.
  - `mergeCell@ref` is `ST_Ref` (required).
  - `dimension@ref` is `ST_Ref`. It is the range that would contain all `c` elements, and it does **not** support whole-column or whole-row notation.
- **Source:**
  - Part 1 §18.3.1.4 c (Cell), attribute `r`, p.1594, L75059–75062.
  - §18.18.7 ST_CellRef, p.2440, L116110–116113.
  - §18.18.62 ST_Ref, p.2480, L118030.
  - §18.18.63 ST_RefA, p.2480, L118044–118046.
  - §18.3.1.54 mergeCell, p.1646, L77645.
  - §18.3.1.35 dimension, p.1622, L76458–76476.
  - `xsd-strict/sml.xsd` L201–209.

### V-04 - Decision: what `parseCellRef` / `parseRange` accept
- **Lowercase letters: accepted.** Column letters are not case-sensitive (V-02). `c@r` is an "A1 style reference" (V-03), so the A1 rules apply.
- **`$` markers: rejected.**
  - The A1 grammar allows `$`, but that grammar is the formula grammar.
  - For stored attributes, the spec gives a separate type, `ST_RefA`, for references that may be absolute (V-03).
  - `c@r`, `mergeCell@ref` and `dimension@ref` use `ST_CellRef` and `ST_Ref`, not `ST_RefA`.
  - Derived inference: `$` is not part of `ST_CellRef` or `ST_Ref`.
  - **Maintainer review:** if real files turn out to contain `$A$1` in these attributes, a lenient mode could strip the markers.
- **Leading zeros in the row: accepted** (`A01` → row 1). A `digit-sequence` may have leading zeros, and only its value must be non-zero (V-02). The row is capped at 7 digits and the column at 3 letters.
- **Whole-row and whole-column references (`A:C`, `1:3`): rejected** (`dimension` does not support them, V-03).
- **Single cell: accepted** as a one-cell range. `ST_Ref` can be a single cell.
- **Range corners:** normalised so that start is top-left and end is bottom-right. The spec does not say the corners must be in that order.

---

## B. Serial dates (`dates.ts`)

### V-05 - Serial date-time definition (Strict)
- **Fact:**
  - A serial date-time is a **signed** number of days relative to the base date of the selected date system. It goes up by 1 for each following day and down by 1 for each earlier day.
  - The fractional part is a fraction of one day.
  - Base dates and serials represent **local** date and time, so no time-zone conversion applies.
- **Source:** Part 1 §18.17.4.1 Date Conversion for Serial Date-Times, p.2063 (PDF 2073), L97438–97447.

### V-06 - Strict 1900 date system
- **Fact:**
  - Lower limit: 0001-01-01 00:00:00 = serial −693593.
  - Upper limit: 9999-12-31 23:59:59.999 = serial 2,958,465.9999884.
  - Base date: 1899-12-30 00:00:00 = serial 0.
  - Spec examples: 1.5 = 1899-12-31T12:00; −4.25 = 1899-12-25T18:00.
- **Source:** Part 1 §18.17.4.1, p.2063, L97440–97456.
- **Tokens:** `-693593`, `2,958,465.9999884`.
- **Derived:**
  - JavaScript `Date` uses the proleptic Gregorian calendar, and in it 1899-12-30 − 693593 days = 0001-01-01. Tested.
  - Serial 60 = 1900-02-28 and 61 = 1900-03-01. There is no 29 February 1900 (V-11).

### V-07 - 1904 date system (Strict and Transitional)
- **Fact:**
  - Lower limit: 0001-01-01 00:00:00 = −695055.
  - Upper limit: 9999-12-31 23:59:59.999 = 2,957,003.9999884.
  - Base date: 1904-01-01 00:00:00 = serial 0.
- **Source:**
  - Part 1 §18.17.4.1, p.2063–2064, L97458–97463.
  - Part 4 §15.3.4 restates the same values for Transitional documents (p.215–216, L9675–9680).
- **Tokens:** `-695055`, `2,957,003.9999884`.

### V-08 - Range and selection of the system
- **Fact:**
  - A serial outside the temporal range of the selected date system is invalid.
  - The `date1904` attribute of `workbookPr` selects the system. It is `xsd:boolean` with default `false` (1900 system).
- **Source:**
  - Part 1 §18.17.4.1, p.2064, L97463–97470.
  - §18.2.28 workbookPr, attribute `date1904`, p.1577, L74188.
- **Code:** `serialToDate` returns `null` outside the range. `dateToSerial` throws `RangeError`.

### V-09 - Time of day
- **Fact:**
  - The time of day is a serial in [0, 1): 0 is 00:00:00 and 0.99999999 is 23:59:59.
  - The day of any serial is the greatest integer not exceeding it (floor). The time is the serial minus that day.
  - The time component grows by 1/86 400 per second.
  - Spec examples: −2.00000001 is day −3 with time 0.99999999; 4.66666667 is day 4 with time 0.66666667; 0.4207639… is 10:05:54.
- **Source:** Part 1 §18.17.4.2 Time Conversion for Serial Date-Times, p.2064, L97472–97501.

### V-10 - Combined worked examples (used as tests)
- **Fact (1900 system):**
  - −2337.999989… = 1893-08-05T00:00:01
  - 3687.4207639… = 1910-02-03T10:05:54
  - 2.5 = 1900-01-01T12:00:00
  - 2958465.9999884… = 9999-12-31T23:59:59
  - −1.25 = 1899-12-28T18:00
- **Fact (1904 system):**
  - −3799.999989… = 1893-08-05T00:00:01
  - 2225.4207639… = 1910-02-03T10:05:54
  - 0.5 = 1904-01-01T12:00:00
  - 2957003.9999884… = 9999-12-31T23:59:59
- **Also:** §18.3.1.96 gives 1976-11-22T08:30 = 28086.3541666667 (1900 system).
- **Source:**
  - Part 1 §18.17.4.3, p.2064–2065, L97503–97526.
  - §18.3.1.96 v (Cell Value), p.1707–1708, L80660–80682.
- **Note:** The spec prints these serials cut short ("…"). Tests use exact fractions (e.g. `3687 + 36354/86400`). The printed values are tested to within one unit of their last printed decimal.

### V-11 - Proleptic Gregorian calendar
- **Fact:** SpreadsheetML relates all dates to the proleptic Gregorian calendar of ISO 8601. Calendar-reform gaps and leap seconds are ignored.
- **Source:** Part 1 §18.17.4, note, p.2063, L97425–97429.
- **Use:**
  - In the Strict 1900 system there is no 1900-02-29.
  - `parseIsoDateTime` validates days per proleptic Gregorian leap years, so `1900-02-29` → `null` and `2000-02-29` is valid.

### V-12 - Transitional 1900 date system (Part 4 replaces Part 1 §18.17.4.1)
- **Fact:**
  - For Transitional documents, Part 4 replaces the text of Part 1 §18.17.4.1.
  - 1900 system:
    - lower limit: 1900-01-01 00:00:00 = serial 1;
    - upper limit: 9999-12-31 23:59:59 = 2,958,465.9999884;
    - base date: 1899-12-31 = serial 0.
  - The 1904 system is unchanged (V-07).
  - Consequence: serials below 1 (zero, negatives and time-only fractions 0.x) are **outside** the Transitional 1900 range. **Implementation decision D-112 (docs/decisions.md):** time-only values in [0, 1) are still converted, onto 1899-12-31 (the day the section gives for serial 0); negative serials return `null`.
- **Source:** Part 4 §15.3.4 Modified content for Date Conversion for Serial Date-Times, p.215–216 (PDF 231–232), L9655–9691. The lower limit is at L9671.
- **Verbatim (the one prose quote):** "January 1, 1900, 00:00:00, which has a serial date-time of 1" (Part 4 L9671, p.215).
- **Spec inconsistency:** The replacement text keeps the old example, which says the 1900 system's base date is 30 December 1899 and that 1.5 = 1899-12-31T12:00 (L9663–9666). That contradicts the stated lower limit. The code follows the normative limits.

### V-13 - 29 February 1900 (Transitional serial 60)
- **Spec:** Silent. No mention of 1900-02-29 or serial 60 in Parts 1 or 4. The searches are listed in `facts-ecma.md` E-16.
- **Derived:**
  - Transitional puts 1900-01-01 at 1 and 9999-12-31 at 2958465. That is 2 958 464 days apart.
  - The proleptic Gregorian calendar has 2 958 463 days between them (Strict: 2 and 2958465).
  - So the Transitional 1900 system counts one extra day.
- **Corroboration:** The DATE function examples for the "1900 backward compatibility date-base" give DATE(1900,1,1) = 1 and DATE(9999,12,31) = 2958465 (Part 1 §18.17.7.74, p.2164, L102718–102724).
- **Microsoft:** Excel treats 1900 as a leap year, a choice inherited from Lotus 1-2-3 for serial-date compatibility. Only dates before 1 March 1900 are affected (the WEEKDAY function is wrong for them).
  - https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/wrongly-assumes-1900-is-leap-year
- **Decision (per coordinator):** In the Transitional 1900 system:
  - serials in [1, 60) map to 1899-12-31 + n days;
  - [60, 61) is the nonexistent 1900-02-29, so `serialToDate` returns `null`;
  - serials ≥ 61 map to 1899-12-30 + n days.
  - `dateToSerial` never produces values in [60, 61).

### V-14 - Upper limit precision
- **Spec:**
  - Part 1 gives the upper instant as 23:59:59.999 but its serial as 2,958,465.9999884. That serial is about 23:59:58.998.
  - Part 4 gives 23:59:59 for the 1900 system.
- **Decision (derived):** In every system the last valid instant is 9999-12-31T23:59:59.999. The range check is on the instant after rounding to milliseconds (V-15), so all serials of day 2958465 (1904: 2957003) are valid.

### V-15 - Rounding
- **Spec:** No rule on conversion precision. §18.17.5.2 (precision of numbers) does not cover it.
- **Decision (from the brief):**
  - The serial is converted to whole milliseconds with round-half-up (`Math.round(serial × 86 400 000)`) **before** the range check and before the serial-60 check.
  - So 0.9999999999 (Transitional) is 1900-01-01T00:00:00.000, and 59.9999999999 rounds to 60 and gives `null`.

### V-16 - Choice of variant: conformance class, not `dateCompatibility`
- **Fact (5th edition):** `workbookPr@dateCompatibility` does not exist in Part 1, Part 4 or either `sml.xsd` (`facts-ecma.md` E-12). The 1900 variant depends on the conformance class (V-06 vs V-12).
- **Fact:** If `workbook@conformance` is omitted, its default is `transitional` (Part 1 §18.2.27, `facts-ecma.md` E-04). So `serialToDate` and `dateToSerial` default to `'transitional'`.
- **Microsoft (ISO/IEC 29500-1:2008 text, as reproduced in the Open XML SDK reference):**
  - `dateCompatibility` existed in the 1st edition, default true.
  - true: the 1900 or 1904 backward-compatibility date base, as chosen by `date1904`.
  - false: the 1900 date base with the full ISO 8601 range, and `date1904` is ignored.
  - https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.workbookproperties?view=openxml-3.0.1 (Remarks, attribute table).
- **Microsoft implementer notes ([MS-OI29500] §18.2.28):**
  - Office 2010 ignores `dateCompatibility` and always uses a compatibility date base (note g).
  - Office 2013 ignores it too: with `workbook@conformance="strict"` it uses the 1900 date system, otherwise the 1900 compatibility date system (note j).
  - Excel does not support negative serial numbers (note j).
  - https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/a251e7c6-7b8a-4e2f-b284-025f9f09ba3e
- **Decision:** The third parameter is `conformance: Conformance = 'transitional'`. This matches the 5th edition and Office 2013+ behaviour.
  - A legacy 2008-edition file with `dateCompatibility="0"` is not handled specially. The reader would need to map it to system `'1900'` with `'strict'`. **Maintainer review.**
**Checked 2026-10-30 (D-1104):** [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.585 g, j (P1 §18.2.28 workbookPr), p.245–246 - confirmed from the PDF: `dateCompatibility` ignored; `conformance="strict"` → 1900 date system, otherwise 1900 compatibility date system; Excel does not support negative serial numbers.

### V-17 - Negative serials
- **Fact:** Negative serials are valid in the Strict 1900 system (down to −693593) and the 1904 system (down to −695055), in both conformance classes. In the Transitional 1900 system they are invalid (V-12).
- **Microsoft:** Excel does not support negative serials (V-16, note j). The code follows the spec, not Excel.
- **Source:** V-06, V-07, V-12.

### V-18 - ISO 8601 lexical forms for `t="d"` cells
- **Fact:**
  - Dates and times in cells are stored as ISO 8601 strings.
  - Earliest 0001-01-01 00:00; latest 9999-12-31 23:59:59.999.
  - Midnight uses hour 0, never 24. Leap seconds are not permitted.
  - Date only: Complete, Extended Format Calendar Date (`1975-10-05`).
  - Time only: Complete, Extended Format Time Of Day (`08:30`, `08:30:00`, `08:30:00.000`).
  - Date and time: Complete, Extended Format Calendar Date and Time Of Day (`1976-11-22T08:30`, `…T08:30:00`, `…T08:30:00.000`).
  - Only seconds may carry a decimal separator, and it shall be a full stop. Fractional seconds *should* have at most three decimals.
- **Source:** Part 1 §18.17.4 Dates and Times, p.2062–2063 (PDF 2072–2073), L97363–97418.
- **Decisions:**
  - 1 to 9 fraction digits are accepted (the 3-digit limit is a "should"), rounded to the nearest millisecond.
  - Total length is capped at 30 characters. The parser is a linear scanner.
  - Seconds 60 are rejected (leap seconds not permitted).
  - The same sentence also says the most seconds expressed in a minute "shall be 60". I read that as a count of seconds, not as allowing a 60th second.

### V-19 - Time-only ISO values
- **Spec:** Silent on which date a time-only value has.
- **Derived:** A time of day is a serial in [0, 1) (V-09). In the Strict 1900 system, where `t="d"` exists (V-21), serial 0 is 1899-12-30.
- **Decision:** `parseIsoDateTime('08:30')` → 1899-12-30T08:30:00.000.

### V-20 - Trailing `Z`
- **Spec:**
  - The normative forms (V-18) have no time-zone designator.
  - Informative Annex L.2.16.9 says an instant has a date, time and time-zone component, and its examples end in `Z` (e.g. `1893-08-05T00:00:01Z`).
- **Source:** Part 1 Annex L.2.16.9 (p.4770, L254936–254942); L.2.16.9.3 (p.4772, L255011–255030).
- **Decision:**
  - A single trailing `Z` after a time is accepted and ignored. Under our model the UTC fields are the wall-clock fields.
  - Numeric offsets (`+02:00`) and `Z` after a date-only value give `null`.

### V-21 - `t="d"` is Strict only
- **Fact:**
  - Part 4 removes the `d` value of `ST_CellType` for Transitional documents.
  - In Transitional documents each instant is stored as an ISO 8601 string or as a serial value.
- **Source:** Part 4 §15.8.3 (p.221, L9981–9984) and §15.9.2 (p.222, L10006–10009).
- **Use:** `parseIsoDateTime` takes no conformance argument and works for both classes, as the coordinator asked.

---

## C. Number formats (`numfmt.ts`)

### V-22 - Built-in ("implied") formats, All Languages
- **Fact:**
  - For these ids the `xf` record writes `numFmtId` and no `numFmt` element is written.
  - Some ids can be interpreted differently depending on the UI language.
  - The table, verbatim (raw extraction of PDF p.1786–1787 confirms the spaces before `;` in 37 and 38):

| ID | Code |
|---|---|
| 0 | `General` |
| 1 | `0` |
| 2 | `0.00` |
| 3 | `#,##0` |
| 4 | `#,##0.00` |
| 9 | `0%` |
| 10 | `0.00%` |
| 11 | `0.00E+00` |
| 12 | `# ?/?` |
| 13 | `# ??/??` |
| 14 | `mm-dd-yy` |
| 15 | `d-mmm-yy` |
| 16 | `d-mmm` |
| 17 | `mmm-yy` |
| 18 | `h:mm AM/PM` |
| 19 | `h:mm:ss AM/PM` |
| 20 | `h:mm` |
| 21 | `h:mm:ss` |
| 22 | `m/d/yy h:mm` |
| 37 | `#,##0 ;(#,##0)` |
| 38 | `#,##0 ;[Red](#,##0)` |
| 39 | `#,##0.00;(#,##0.00)` |
| 40 | `#,##0.00;[Red](#,##0.00)` |
| 45 | `mm:ss` |
| 46 | `[h]:mm:ss` |
| 47 | `mmss.0` |
| 48 | `##0.0E+0` |
| 49 | `@` |

- **Source:** Part 1 §18.8.30 numFmt (Number Format), p.1776–1777 (PDF 1786–1787), L83957–84010.
- **Microsoft:** Excel's own codes for 14, 22, 37–40, 47 and Korean 55 differ, e.g. 14 = `m/d/yyyy` and 47 = `mm:ss.0` ([MS-OI29500] §18.8.30 note a). `BUILTIN_NUMBER_FORMATS` keeps the spec text. The date classification is the same for both.
  - https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/17d11129-219b-4e2c-88db-45844d21e528
**Checked 2026-10-30 (D-1104):** [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.712 a, p.299 - confirmed from the PDF (built-in codes 14, 22, 37–40, 47, KOR 55).

### V-23 - Ids outside the lists
- **Fact:**
  - When values outside the lists are used, the behaviour is implementation-defined.
  - Ids not in the listing, such as 5, 6, 7 and 8, shall follow the `formatCode` attribute, so they have no implied format.
- **Source:** Part 1 §18.8.30, p.1776, L83971–83974.
- **Code:** `BUILTIN_NUMBER_FORMATS` has no 5–8, 23–36, 41–44 or ≥ 50.

### V-24 - Locale-specific implied ids
- **Fact:** The spec lists language-dependent codes for ids 27–36 and 50–58 in four tables:
  - zh-tw / zh-cn: p.1778–1780, L84058–84158;
  - ja-jp / ko-kr: p.1780–1781, L84160–84245;
  - each also has a "with unicode values" version.
  - Every code listed for these ids is built only from date/time codes (`y`, `m`, `d`, `h`, `s`, `e`, `g`, AM/PM glyphs) and quoted CJK literals. Examples: zh-tw 27 `[$-404]e/m/d`; ja-jp 28 `[$-411]ggge"5E74"m"6708"d"65E5"`; ko-kr 34 `yyyy-mm-dd`; zh-cn 32 `h"65F6"mm"5206"`.
  - th-th (p.1781–1783, L84247–84301):
    - 59–62 and 67–70 are numeric: `t0`, `t0.00`, `t#,##0`, `t#,##0.00`, `t0%`, `t0.00%`, `t# ?/?`, `t# ??/??`;
    - 71–80 use Thai letters, e.g. 75 `0E0A:0E190E19`, 79 `[0E0A]:0E190E19:0E170E17`, 80 `0E190E19:0E170E17.0`;
    - 81 is `d/m/bb`.
- **Code:** These codes differ per language, so they are **not** in `BUILTIN_NUMBER_FORMATS`. Their ids are in `isDateFormatId`.

### V-25 - Thai ids 71–80 classified as dates (derived)
- **Spec:** §18.8.31 does not define the Thai letters used in 71–80 as format codes.
- **Derived:** They are classified as date/time because they follow the shape of the All-Languages date and time ids:
  - 75/76 `x:yy` / `x:yy:zz` mirror 20/21 `h:mm` / `h:mm:ss`;
  - 79 `[x]:yy:zz` mirrors 46 `[h]:mm:ss`;
  - 80 `yy:zz.0` mirrors 47;
  - 71–74 are three-part `/` and `-` patterns like 14–17.
- **Maintainer review:** this is inference, not spec text.

### V-26 - Sections
- **Fact:**
  - A code has up to four sections separated by `;`: positive, negative, zero, text.
  - Two sections: positive and zero, then negative.
  - One section: all numbers.
  - A skipped section still has its terminating `;`.
  - A text section, if present, shall be the last section.
- **Source:** Part 1 §18.8.31 numFmts, "Number Format Codes", p.1783–1784, L84332–84345; text section p.1786, L84456–84460.
- **Microsoft:** Office reads a three-section code as positive; negative; zero ([MS-OI29500] §18.8.31 note a).
  - https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/c6046ec6-f96d-4609-98e4-0dd549292955
- **Decision:** `isDateFormat` is true if **any** section contains a date or time code.
  - Microsoft's ABNF (V-32) allows only `@` and AM/PM in a text section, so date codes there are not legal anyway.
**Checked 2026-10-30 (D-1104):** [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.713 a, p.310 - confirmed from the PDF. Also g: Office limits custom number formats to 206.

### V-27 - Literal text, escapes, spacing and fill
- **Fact:**
  - Characters inside double quotes, or right after a backslash, are never interpreted as format codes. They are always literal.
  - Single quotes shall not be used to denote text.
  - `\` shows the next character.
  - `_` skips the width of the next character.
  - `*` repeats the next character to fill the cell.
- **Source:** Part 1 §18.8.31: symbol table p.1784–1785, L84354–84428 (`\` L84405, `*` L84409, `_` L84416); "Text and spacing" p.1785, L84429–84437.
- **Code:**
  - `"…"` is skipped to the closing quote.
  - The character after `\`, `_` or `*` is skipped.
  - An unterminated quote runs to the end of the code (decision).

### V-28 - Exponent codes vs the era-year code `e`
- **Fact:**
  - `E-`, `E+`, `e-`, `e+` are the scientific-format codes.
  - Separately, `e` is the era-year code. It becomes `yyyy` in "OTHER" locales, per the international-considerations table.
- **Source:**
  - Part 1 §18.8.31, symbol table p.1784, L84370, and "Display scientific notations" p.1788, L84565–84574.
  - Era codes p.1790–1791, L84656–84702.
- **Microsoft:**
  - Office accepts only `E+` and `E-` for scientific notation (note e).
  - Office maps `e`/`ee` to `yyyy`, not `yy`, outside JPN/CHT (note f).
  - [MS-OI29500] §18.8.31, URL in V-26.
- **Code:**
  - `e` followed by `+` or `-` is an exponent. Any other `e` is an era year, hence a date.
  - Uppercase `E` is never a date code.
**Checked 2026-10-30 (D-1104):** [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.713 e, f, p.310 - confirmed from the PDF.

### V-29 - Colour and condition brackets
- **Fact:**
  - Colours: `[Black]`, `[Blue]`, `[Cyan]`, `[Green]`, `[Magenta]`, `[Red]`, `[White]`, `[Yellow]`, and `[ColorN]` with N from 1 to 56.
  - Conditions: `[` operator value `]` with `=`, `>`, `<`, `>=`, `<=`, `<>`.
- **Source:** Part 1 §18.8.31, p.1787–1788, L84529–84558.
- **Code:** Skipped. They never make a format a date.

### V-30 - Locale/currency brackets and the system date/time values
- **Fact:**
  - Syntax: `[$<Currency String>-<language info>]`.
  - The language info is a 32-bit hex value: bytes 0–1 are the language ID, byte 2 the calendar type, byte 3 the number system.
  - Special values: `0xf800` is the system long date format and `0xf400` the system time format.
- **Source:** Part 1 §18.8.31, "International Considerations", p.1791–1792 (PDF 1801–1802), L84703–84723.
- **Code:**
  - `[$…]` brackets are skipped.
  - If the hex digits after the last `-` (1–8 digits) have low 16 bits equal to `0xF800` or `0xF400`, the code is a date/time format. **Derived:** comparing only the low 16 bits (the language-ID bytes) is my reading.

### V-31 - Date and time codes
- **Fact:**
  - Months: `m`, `mm`, `mmm`, `mmmm`, `mmmmm`.
  - Days: `d`, `dd`, `ddd`, `dddd`.
  - Years: `yy`, `yyyy`.
  - Hours: `h`, `hh`. Minutes: `m`, `mm`. Seconds: `s`, `ss`.
  - Fractional seconds: `h:mm:ss.00`.
  - 12-hour markers: `AM/PM`, `A/P`.
  - Elapsed time: `[h]:mm`, `[mm]:ss`, `[ss].00`.
  - `m`/`mm` mean minutes immediately after `h`/`hh` or immediately before `ss`, otherwise months.
  - International codes: `r`, `rr` (ja-jp/zh-tw), `g`, `gg`, `ggg` (era), `e`, `ee` (era year), `b1`, `b2` (calendar).
- **Source:** Part 1 §18.8.31 "Dates and times", p.1789–1790, L84576–84655; "International Considerations", p.1790–1791, L84656–84702.
- **Code:**
  - `y`, `m`, `d`, `h`, `s`, `g`, `r` and `e` (V-28) are date/time codes. Month/minute disambiguation is not needed: either way the format displays a date or time.
  - Elapsed brackets are a run of one letter `h`, `m` or `s`.
  - `b1`/`b2` only select a calendar, so they are ignored.
  - `AM/PM` and `A/P` alone do not make a date (V-32).

### V-32 - Case of date codes; `General`; AM/PM in number formats
- **Spec:**
  - All date/time codes are printed in lowercase.
  - The spec states case-insensitivity explicitly for A1 column letters (V-02) but not for format codes.
- **Microsoft:**
  - [MS-OI29500] §18.8.30 note b gives Office's ABNF for format codes.
  - `NFPartYear`, `NFPartMonth`, `NFPartDay`, `NFPartHour`, `NFPartMinute`, `NFPartSecond` and the bracketed absolute-time tokens are built from ASCII **small** letters.
  - `INTL-NUMFMT-GENERAL` is `General` with a capital G.
  - `NFNumber` may end with `*INTL-AMPM`, so AM/PM can occur in a number format.
  - URL in V-22.
- **Decisions:**
  - Date codes are matched case-sensitively, so `YYYY-MM-DD` is **not** a date. **Maintainer review:** some third-party writers emit uppercase codes. A lenient mode could accept them, but no source documents that as Excel behaviour.
  - `General` is matched in any case. This only prevents false positives from its letters `e` and `r`.
  - `AM/PM` alone is not a date.
**Checked 2026-10-30 (D-1104):** [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.712 b, p.299–303 - confirmed from the PDF; `formatCode` shorter than 255 characters.

### V-33 - General
- **Fact:** `General` is id 0. The spec notes that dates do not follow General but are converted to date formatting automatically.
- **Source:** Part 1 §18.8.30, "General" Format, p.1777–1778, L84013–84027.
- **Code:** `isDateFormat('General')` is `false`. Detecting date input under General is the writer's concern, not the reader's.

---

## D. Summary of items the spec does not settle

| Item | What I did | Entry |
|---|---|---|
| Transitional serial 60 (29 Feb 1900) | `null`; Microsoft's leap-year article cited | V-13 |
| Conversion precision | round half up to 1 ms before the range checks | V-15 |
| `dateCompatibility` (absent in 5th ed.) | use conformance class; default `transitional` | V-16 |
| Excel rejects negative serials | follow spec (negatives valid where the spec says) | V-17 |
| Date of a time-only ISO value | 1899-12-30 | V-19 |
| `Z` suffix | accepted after a time; offsets rejected | V-20 |
| `$` in `c@r` / `ST_Ref` | rejected (ST_RefA is the absolute type) | V-04 |
| Thai ids 71–80 | classed as date/time by analogy | V-25 |
| Uppercase date codes | not dates | V-32 |
| `[$-xF800]` language info | low 16 bits compared | V-30 |
