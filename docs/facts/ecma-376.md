# ECMA-376 verified facts for an XLSX reader

> **Provenance.** Compiled on 2026-10-30 from the official documents listed below. Line numbers (`Lnnn`) refer to text extracted with `pdftotext -layout` (poppler) from the PDFs in the ECMA-376 download archives (https://ecma-international.org/publications-and-standards/standards/ecma-376/): Part 1 `ECMA-376-1_5th_edition_december_2016.zip`, Part 2 `ECMA-376-2_5th_edition_december_2021.zip`, Part 4 `ECMA-376-4_5th_edition_december_2016.zip`. Schema line numbers refer to the XSD files inside the same archives. Paths such as `scratchpad/specs/…` below are the working copies used during the check; the specs are not stored in this repository.


## Conventions

- **Sources.** All files are in `scratchpad/specs/`:
  - `part1.txt`: ECMA-376 Part 1, 5th ed. (Dec 2016), Strict.
  - `part2.txt`: Part 2, 5th ed. (Dec 2021), OPC.
  - `part4.txt`: Part 4, 5th ed. (Dec 2016), Transitional.
  - The schemas: `xsd-strict/`, `xsd-trans/` and `opc-xsd/`.
- **Locators.** Each fact gives:
  - the Part, the clause (§) and its title;
  - the printed page number ("p."), taken from the page footer. The PDF page index is the printed page + 10 for Parts 1 and 2, and + 16 for Part 4;
  - `Lnnn`, the line number in the matching `.txt` file. To check a fact, run `sed -n 'NNN,+20p' part1.txt`.
- **Schema facts.** These give the xsd file, the type name and the line number.
- **How text was extracted.** In `pdftotext -layout` output, attribute tables are often shifted. For example, the type sentence of one attribute can print next to the name of the following attribute. Every attribute table cited here was cross-checked with `pdftotext -raw` on the same PDF page, or against the xsd.
- **Quoting.** Verbatim prose quotes are kept to a minimum (at most 15 words each). That single quote is in E-13. Everywhere else:
  - prose is **paraphrased closely**, with exact locators so you can check the wording yourself;
  - technical tokens are given **verbatim**: URIs, enumeration values, type, attribute and default names, format codes and numbers.
- **"Derived"** marks my own arithmetic or inference, not spec text.

---

## A. Namespaces, relationship types, content types (facts 1–2)

### E-01 - Main SpreadsheetML namespace
- **Fact:** The SpreadsheetML main namespace is different in each conformance class.
  - Strict: `http://purl.oclc.org/ooxml/spreadsheetml/main`
  - Transitional: `http://schemas.openxmlformats.org/spreadsheetml/2006/main`
- **Source (Strict):**
  - Part 1 §12.3.23 Workbook Part, "Root Namespace", p.98, L4737. The same URI appears for every §12.3.x SpreadsheetML part.
  - `xsd-strict/sml.xsd` L7 (`targetNamespace`).
- **Source (Transitional):**
  - Part 4 §10.2.23 Workbook Part, p.21, L1492.
  - Part 4 §15.8 Simple Types, p.221, L9957.
  - `xsd-trans/sml.xsd` L7.
- **Strict/Transitional:** Different URIs, as listed above.

### E-02 - Namespace of `r:id` (and other explicit-relationship attributes)
- **Fact:** The `r:` namespace is also different in each class.
  - Strict: `http://purl.oclc.org/ooxml/officeDocument/relationships`
  - Transitional: `http://schemas.openxmlformats.org/officeDocument/2006/relationships`
  - The attribute type is `ST_RelationshipId` (an `xsd:string` restriction).
- **Source (Strict):**
  - Part 1 §18.2.19 sheet (Sheet Information), attribute `id`, "Namespace:" cell, p.1568, L73724–73728.
  - Part 1 §22.8.2.1 ST_RelationshipId, p.3788, L177413.
  - `xsd-strict/shared-relationshipReference.xsd` (`targetNamespace`).
- **Source (Transitional):**
  - Part 4 §15.10.3 Changed attribute for sheet element. It shows the namespace as `.../officeDocument/2006/relationships`. p.222, L10037–10045.
  - `xsd-trans/shared-relationshipReference.xsd` (`targetNamespace`).
- **Strict/Transitional:** Different, as listed above.

### E-03 - OPC namespaces (`.rels` and `[Content_Types].xml`)
- **Fact:** The OPC namespaces are the same in both conformance classes. Part 2 defines no purl.oclc.org variants. A search of `part2.txt` for "purl" finds only the Dublin Core URIs.
  - Relationships part XML: `http://schemas.openxmlformats.org/package/2006/relationships`
  - Media Types stream: `http://schemas.openxmlformats.org/package/2006/content-types`
  - Relationships part media type: `application/vnd.openxmlformats-package.relationships+xml`
- **Source:**
  - Part 2 §6.5.2.1 Relationships part, p.20, L1079–1081.
  - Part 2 Annex E, Tables E.1 and E.2, p.76–77, L3612–3660.
  - `opc-xsd/opc-relationships.xsd` and `opc-xsd/opc-contentTypes.xsd` (`targetNamespace`).
- **Strict/Transitional:** Same.

### E-04 - `workbook@conformance`
- **Fact:** The `workbook` element has an optional `conformance` attribute.
  - Type: `ST_ConformanceClass`. Values: `strict`, `transitional`.
  - Part 1 says that when the attribute is omitted, its default value is `transitional`.
  - The xsd declares no default.
- **Source:**
  - Part 1 §18.2.27 workbook, p.1575, L74085–74100.
  - Part 1 §22.9.2.2 ST_ConformanceClass, p.3792, L177590.
  - `xsd-strict/sml.xsd` CT_Workbook L4068 (attribute `conformance`, no default).
  - `xsd-strict/shared-commonSimpleTypes.xsd` L89.
- **Strict/Transitional:** `CT_Workbook` is identical in both xsds (checked with diff).

### E-05 - Relationship type URIs
- **Fact:** The base URIs are:
  - Strict: `http://purl.oclc.org/ooxml/officeDocument/relationships/<name>`
  - Transitional: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/<name>`

| Target part | `<name>` | Strict source (Part 1) | Transitional source (Part 4) |
|---|---|---|---|
| Workbook (from package `.rels`) | `officeDocument` | §12.3.23, p.98, L4739 | §10.2.23, p.21, L1493 |
| Worksheet | `worksheet` | §12.3.24, p.100, L4824 | §10.2.24, p.21, L1501 |
| Chartsheet | `chartsheet` | §12.3.2, p.69, L3285 | §10.2.2, p.18, L1314 |
| Dialogsheet | `dialogsheet` | §12.3.7, p.76, L3614 | §10.2.7, p.19, L1357 |
| Shared String Table | `sharedStrings` | §12.3.15, p.89, L4257 | §10.2.15, p.20, L1425 |
| Styles | `styles` | §12.3.20, p.94, L4545 | §10.2.20, p.21, L1469 |
| External Workbook References | `externalLink` | §12.3.9, p.79, L3771 | §10.2.9, p.19, L1373 |
| Theme | `theme` | §14.2.7, p.135, L6623 | §12.2.7, p.26, L1713 |

- **Additional:**
  - The `externalLinkPath` type is used for the externalLink part's link to the external file. Source: §12.4, p.~101, L4925; Part 4 §10.3, L1508.
  - Spec inconsistency: the worked example in Part 1 §18.17.2.3 (p.2052, L96860–96866) uses `externalLinkPath` even for the Workbook→externalLink-part relationship, which contradicts §12.3.9.
- **Strict/Transitional:** Only the base URI differs. The `<name>` suffixes are the same.

### E-06 - Macrosheet
- **NOT FOUND:** No macrosheet part, relationship type or content type exists in Part 1 or Part 4.
  - Searched for: "macrosheet" (case-insensitive), "relationships/macrosheet", "xlMacrosheet".
  - The only hits are `CT_Macrosheet` in the schemas: `xsd-strict/sml.xsd` L2106 and `xsd-trans/sml.xsd` L2118.
  - No global `macrosheet` element is declared. (`dialogsheet` and `worksheet` are declared as elements, L2105/L2103.)

### E-07 - Workbook main-part content types
- **Fact:** ECMA-376 defines exactly two content types for the Workbook part:
  - `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml` (workbook)
  - `application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml` (template)
- **Macro-enabled workbook and macro-enabled template content types: NOT defined by ECMA-376.**
  - Searched Parts 1, 2 and 4 for: "macroEnabled", "ms-excel", "vbaProject", ".xlsm", ".xltm", "macro-enabled", "macro enabled". All had zero hits.
  - Part 1 §9.1.7 adds a rule: macro and scripting mechanisms must use correct content types for their executable parts, and must not reuse any content type already defined in ECMA-376. Source: p.22, L1115–1117.
- **Source:** Part 1 §12.3.23 Workbook Part, "Content Type(s)", p.98, L4735–4736.
- **Strict/Transitional:** Part 4 §10.2.23 changes only the root namespace and the relationship type, not the content types.

### E-08 - Part cardinality (relevant for a reader)
- **Fact:**
  - **Workbook:** exactly one per package, targeted from the package relationship item. §12.3.23, p.98, L4746.
  - **Worksheet:** exactly one part per worksheet. Each is the target of an explicit relationship from the Workbook, and the `sheet` element's `id` attribute references it. §12.3.24, p.100, L4832–4834.
  - **Chartsheet:** zero or more, each targeted explicitly. §12.3.2, p.69.
  - **Dialogsheet:** "one or more" (sic), targeted explicitly. §12.3.7, p.76, L3622.
  - **Styles:** no more than one, implicit relationship from the Workbook. §12.3.20, p.94, L4553.
  - **Theme:** zero or one. §14.2.7, p.135.
  - **Shared String Table:** **exactly one**, implicit relationship from the Workbook. §12.3.15, p.89, L4264.
- **Spec inconsistency:** The "minimal conformant SpreadsheetML package" example in §12.2 (p.65–66, L3070–3130) contains no Shared String Table part, which contradicts the "exactly one" rule. A reader must tolerate a missing SST.
- **Strict/Transitional:** No difference found.

---

## B. Workbook structure (facts 3–4)

### E-09 - `sheets/sheet` attributes
- **Fact:** `sheets` holds 1..n `sheet` elements (CT_Sheets, `minOccurs="1"`). Each `sheet` has these attributes:

| Attribute | Type | Required | Default | Rule |
|---|---|---|---|---|
| `name` | `ST_Xstring` | yes | none | must be unique |
| `sheetId` | `xsd:unsignedInt` | yes | none | must be unique |
| `state` | `ST_SheetState` | no | `visible` | |
| `r:id` | `ST_RelationshipId` | yes | none | identifies the sheet part |

- **Not stated:** No case rule or length limit for sheet names was found. Searched "31 characters" and "sheet name".
- **Source:**
  - Part 1 §18.2.19 sheet (Sheet Information), p.1568–1569, L73718–73757 (checked with raw extraction of PDF p.1578–1579).
  - `xsd-strict/sml.xsd` CT_Sheet L4187, CT_Sheets L4182.
- **Strict/Transitional:** Same attributes. `xsd-trans/sml.xsd` CT_Sheet L4216 is identical. Only the `r:` namespace differs (E-02).

### E-10 - ST_SheetState
- **Fact:** The values are `visible`, `hidden` and `veryHidden`. The default (on `sheet@state`) is `visible`.
  - `hidden`: the sheet is hidden but can be shown again through the UI. (The spec text says "workbook window", an editorial slip.)
  - `veryHidden`: hidden, cannot be shown through the UI, only programmatically.
  - `visible`: visible.
- **Source:**
  - Part 1 §18.18.68 ST_SheetState, p.2482–2483, L118146.
  - `xsd-strict/sml.xsd` L4193.
- **Strict/Transitional:** Identical (`xsd-trans/sml.xsd` L4222).

### E-11 - `workbookPr@date1904`
- **Fact:** `date1904` is `xsd:boolean`, optional, default `false`.
  - `1`/`true`: the workbook uses the 1904 date system.
  - `0`/`false`: the workbook uses the 1900 date system.
  - The spec text points to §18.17.4.1 for the definitions.
- **Source:**
  - Part 1 §18.2.28 workbookPr, p.1577, L74188–74200 (raw extraction confirms the type is boolean).
  - `xsd-strict/sml.xsd` CT_WorkbookPr L4201.
- **Strict/Transitional:** CT_WorkbookPr is identical in both xsds (diff). The *meaning* of the 1900 system differs; see E-15.

### E-12 - `workbookPr@dateCompatibility`
- **NOT FOUND** anywhere in Part 1, Part 4, `xsd-strict/sml.xsd` or `xsd-trans/sml.xsd`.
  - Searched (case-insensitive): "dateCompatibility", "dateCompat", "date compat".
  - In this edition, CT_WorkbookPr has only `date1904` for date-system selection. Its attributes, in order: `date1904`, `showObjects`, `showBorderUnselectedTables`, `filterPrivacy`, `promptedSolutions`, `showInkAnnotation`, `backupFile`, `saveExternalLinkValues`, `updateLinks`, `codeName`, `hidePivotFieldList`, `showPivotChartFilter`, `allowRefreshQuery`, `publishItems`, `checkCompatibility`, `autoCompressPictures`, `refreshAllConnections`, `defaultThemeVersion`.
- **Leftover from an older model:** The DATE function example and Annex L still mention a "1900 backward compatibility" date base (E-19). No attribute selects it.

---

## C. Dates and times (fact 5)

### E-13 - Strict serial date-time: base dates, limits, negative values (Part 1 §18.17.4.1)
- **Fact:** A serial date-time is a **signed** number of days relative to the base date of the selected date system. The fractional part is the fraction of a day. Base dates and serials represent **local** date and time.
  - **1900 system:**
    - base: 1899-12-30T00:00:00 = serial 0;
    - lower limit: 0001-01-01T00:00:00 = −693593;
    - upper limit: 9999-12-31T23:59:59.999 = 2,958,465.9999884.
  - **1904 system:**
    - base: 1904-01-01T00:00:00 = serial 0;
    - lower limit: 0001-01-01T00:00:00 = −695055;
    - upper limit: 9999-12-31T23:59:59.999 = 2,957,003.9999884.
  - A serial outside the range of the selected system is invalid.
  - **So negative serials are allowed in both Strict systems.**
- **Source:** Part 1 §18.17.4.1 Date Conversion for Serial Date-Times, p.2063–2064, L97438–97470.
- **Verbatim (the one prose quote):** "The base date for this system is 00:00:00 on December 30th, 1899" (L97451, p.2063).
- **Strict/Transitional:** The 1900 system differs in Transitional (E-15).

### E-14 - Strict: what serial 0, 1 (and 60) mean; worked examples
- **Fact (spec examples, 1900 system):**

| Serial | Date-time |
|---|---|
| 1.5 | 1899-12-31T12:00 |
| −4.25 | 1899-12-25T18:00 |
| −1.25 | 1899-12-28T18:00 |
| 2.5 | 1900-01-01T12:00:00 |
| 3687.4207639 | 1910-02-03T10:05:54 |
| −2337.999989 | 1893-08-05T00:00:01 |
| 2958465.9999884 | 9999-12-31T23:59:59 |

- **Fact (spec examples, 1904 system):**

| Serial | Date-time |
|---|---|
| 0.5 | 1904-01-01T12:00:00 |
| −3799.999989 | 1893-08-05T00:00:01 |
| 2225.4207639 | 1910-02-03T10:05:54 |

- **Derived:** In Strict 1900, serial 0 = 1899-12-30, 1 = 1899-12-31, 2 = 1900-01-01 and **60 = 1900-02-28** (61 = 1900-03-01).
  - There is no 1900-02-29, because §18.17.4 relates all dates to the proleptic Gregorian calendar (p.2063, L97425–97429).
  - Check: the §18.3.1.96 example puts 1976-11-22T08:30 at serial 28086.3541666667 in the 1900 system (p.1707–1708, L80663–80680). I computed 1899-12-30 → 1976-11-22 = 28086 days, which agrees.
- **Source:** Part 1 §18.17.4.1–§18.17.4.3, p.2063–2065, L97440–97526.

### E-15 - Transitional replaces §18.17.4.1 (Part 4 §15.3.4)
- **Fact:** For transitional documents, Part 4 replaces the text of Part 1 §18.17.4.1.
  - **1900 system:**
    - lower limit: 1900-01-01T00:00:00 = serial **1**;
    - upper limit: 9999-12-31T23:59:59 = 2,958,465.9999884;
    - base date: **1899-12-31** = serial 0.
  - **1904 system:** unchanged from Strict (lower −695055, base 1904-01-01 = 0).
- **Consequences:**
  - In Transitional 1900, serials below 1 are outside the valid range, so negative serials (and 0) are invalid.
  - In Transitional 1904, negative serials remain valid.
- **Spec inconsistency:** In the same replacement text, the example still says the 1900 system has a base date of 30 December 1899 and that 1.5 = 1899-12-31T12:00. That contradicts the stated lower limit of serial 1 = 1900-01-01 (L9664–9667 vs L9671–9673).
- **Source:** Part 4 §15.3.4 Modified content for Date Conversion for Serial Date-Times (Part 1, §18.17.4.1), p.215–216, L9655–9691.
- **Strict/Transitional:** Different, as described.

### E-16 - 29 February 1900 / serial 60
- **NOT FOUND:** Neither Part 1 nor Part 4 mentions 1900-02-29, a 1900 leap-year bug, or special handling of serial 60.
  - Searched (case-insensitive): "leap year", "1900-02-29", "February 29", "29 February" (all hits are in day-count-basis notes of financial functions), "1900 was", "1900 is", "February 1900", "serial date-time of 60", "serial value 60". Same searches in Part 4.
- **Derived (my arithmetic, not spec text):**
  - Proleptic Gregorian has 2,958,464 days from 1899-12-31 to 9999-12-31.
  - Transitional states 1900-01-01 = 1 and 9999-12-31 = 2,958,465.
  - The extra day is consistent with a phantom 1900-02-29 being counted somewhere in the Transitional 1900 system, but the spec never says so.
  - Strict has no such day (E-14).
**Checked 2026-10-30 (D-1104):** still not found: ECMA-376 Part 1, 5th ed. (2016), Part 4, [MS-OI29500] rev. 25.0 (2026-08-18), [MS-OE376] and [MS-XLSX] rev. 30.0 (2026-10-16) were searched for `1900` and `leap`.

### E-17 - Time fraction (Part 1 §18.17.4.2)
- **Fact:** Time of day is a serial in [0, 1): 0 is 00:00:00 and 0.99999999 is 23:59:59.
  - The time component grows by 1/86,400 per second, so 12:00 = 0.5.
  - For any serial, the day is the greatest integer not exceeding it (floor), and the time of day is the serial minus that day.
  - Spec examples: −2.00000001 is day −3 with time 0.99999999; 4.66666667 is day 4 with time 0.66666667.
  - More examples: 0.0000115 = 00:00:01; 0.4207639 = 10:05:54; 0.9999884 = 23:59:59.
- **Source:** Part 1 §18.17.4.2 Time Conversion for Serial Date-Times, p.2064, L97472–97501.
- **Strict/Transitional:** Part 4 does not modify §18.17.4.2.

### E-18 - ISO 8601 lexical format for date cells (`t="d"`) (Part 1 §18.17.4, §18.17.6.7)
- **Fact (Strict):** Dates and times in cells are stored as strings in ISO 8601 lexical formats.
  - **Range:** earliest 0001-01-01 00:00; latest 9999-12-31 23:59:59.999.
  - **Midnight:** must use hour 0, never hour 24.
  - **Leap seconds:** not permitted. (The same sentence oddly says the maximum seconds in a minute shall be 60.)
  - **Date only:** ISO 8601 Complete, Extended Format Calendar Date (§B.1.1/§B.2.1), e.g. `1975-10-05`.
  - **Time only:** Complete, Extended Format Time of Day (§B.1.2/§B.2.2), e.g. `08:30`, `08:30:00`, `08:30:00.000`.
  - **Date and time:** Complete, Extended Format Calendar Date and Time of Day (§B.1.3/§B.2.3), e.g. `1976-11-22T08:30`, `1976-11-22T08:30:00`, `1976-11-22T08:30:00.000`.
  - **Decimal separator:** a full stop, allowed only in seconds; fractional seconds *should* have at most 3 decimal places.
  - **Calendar:** all dates use the proleptic Gregorian calendar.
  - **Example cell:** `<c t="d"><f>DATE(1582,10,15)+0.5</f><v>1582-10-15T12:00</v></c>`
- **Source:**
  - Part 1 §18.17.4 Dates and Times, p.2062–2063, L97363–97429.
  - §18.17.6.7, p.2070, L97813–97828.
- **Strict/Transitional:**
  - Part 4 §15.9.2 (p.222, L10006–10009): in transitional documents each instant shall be stored as an ISO 8601 string **or** as a serial value.
  - The Transitional `ST_CellType` has **no `d` value** (E-21).

### E-19 - Other (conflicting) date material a reader may meet
- **Fact (DATE function):** Part 1 §18.17.7.74 DATE (p.2164, L102718–102735) gives examples "for the 1900 backward compatibility date-base date system". It gives DATE(1900,1,1) = 1, DATE(0,1,1) = 1 and DATE(9999,12,31) = 2958465. This matches the Transitional 1900 system, not Strict §18.17.4.1, where 1900-01-01 = 2.
- **Fact (Annex L, informative):** Annex L.2.16.9.1 (p.4771, L254961–254975) describes **three** date bases:
  - "1900 date base": base 1899-12-30, lower limit January 1, −9999;
  - "1900 backward compatibility": 1900-01-01 = 1, base 1899-12-31;
  - "1904 backward compatibility".
- **Annex L contradictions:**
  - It says a serial is UTC. §18.17.4.1 says local.
  - It gives 1900-01-01 = 2.0 but also 1.5 = 1900-01-01T12:00:00Z (L254981 vs L255021, p.4771–4772).
  - Annex L is informative (L244847).
- **Use:** Treat §18.17.4.1 (Strict) and Part 4 §15.3.4 (Transitional) as normative.

---

## D. Cells, rows, values, formulas (facts 6–9, 12)

### E-20 - `c` (Cell) attributes
- **Fact:** CT_Cell has the child sequence `f?`, `v?`, `is?`, `extLst?`. Its attributes:

| Attr | Type | Required | Default | Meaning |
|---|---|---|---|---|
| `r` | `ST_CellRef` | **no** | none | A1-style reference to the cell's location |
| `s` | `xsd:unsignedInt` | no | `0` | index of the cell's style; style records are in the Styles part (see E-47) |
| `t` | `ST_CellType` | no | `n` | the cell's data type |
| `cm` | `xsd:unsignedInt` | no | `0` | zero-based index of a cell metadata record (Metadata part) |
| `vm` | `xsd:unsignedInt` | no | `0` | zero-based index of a value metadata record (Metadata part) |
| `ph` | `xsd:boolean` | no | `false` | whether to show phonetic information (East Asian) |

- **Note:** Defaults come only from the xsd. The prose does not state them.
- **Source:**
  - Part 1 §18.3.1.4 c (Cell), p.1593–1594, L75005–75085 (raw-checked, PDF p.1603–1604).
  - `xsd-strict/sml.xsd` CT_Cell L2256–2269.
- **Strict/Transitional:** Identical (`xsd-trans/sml.xsd` L2274).
**Checked 2026-10-30 (D-1104):** [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.590 a, b, f, g (P1 §18.3.1.4 c), p.251: Office treats `cm` and `vm` as one-based indexes (P1 says zero-based); `s` indexes cellXfs, 0 … 65490.

### E-21 - ST_CellType values and meaning
- **Fact (Strict):**

| Value | Meaning |
|---|---|
| `b` | the cell contains a boolean |
| `d` | the cell contains a date in ISO 8601 format |
| `e` | the cell contains an error |
| `inlineStr` | the cell contains an inline rich string (not in the shared string table); the value is in `is`, not `v` |
| `n` | the cell contains a number (this is the default) |
| `s` | the cell contains a shared string |
| `str` | the cell contains a formula string |

- **Source:**
  - Part 1 §18.18.11 ST_CellType, p.2441, L116149–116176.
  - `xsd-strict/sml.xsd` L2270–2280.
- **Strict/Transitional:** **`d` is removed in Transitional.**
  - Part 4 §15.8.3 "Removed enumeration values for ST_CellType" (p.221, L9981–9984).
  - `xsd-trans/sml.xsd` ST_CellType L2288–2297 and the Part 4 schema listing (L54300–54331) list only `b n e s str inlineStr`.

### E-22 - `inlineStr` and the `is` child
- **Fact:** When a cell's `t` is `inlineStr`, the only child element allowed is `is`. (The xsd sequence would still allow `f`/`v` too.)
- **Spec inconsistency:** The §18.3.1.53 `is` example writes `<c r="A1"><is>…` **without** `t="inlineStr"` (p.1646, L77634–77638). A lenient reader should accept `is` even when `t` is absent.
- **Source:** Part 1 §18.3.1.4, p.1593, L75021–75023; §18.3.1.53 is (Rich Text Inline), p.1646, L77628–77643.

### E-23 - `v` element; how values are written (including booleans)
- **Fact:** `v` holds the cell's value. Its type is `ST_Xstring`.
  - For a string cell, `v` holds the index into the shared string table. Otherwise it holds the value directly.
  - For a formula cell, `v` holds the last calculated result.
- **Fact (value representation for formula results):**

| Result type | Text written in `v` |
|---|---|
| array | the text form of the array's value |
| error | the text form of the error value |
| logical | `0` for FALSE, `1` for TRUE |
| number | the unformatted text form of the number, as accurately as possible |
| text | all characters of the text |

- **Not found:** No separate rule for non-formula `t="b"` cells. The `0`/`1` rule above is the only boolean serialisation rule found.
- **Source:**
  - Part 1 §18.3.1.96 v (Cell Value), p.1707–1708, L80654–80700.
  - §18.17.6.6 Value Representation, p.2070, L97802–97811.
  - `xsd-strict/sml.xsd` CT_Cell L2259.
**Checked 2026-10-30 (D-1104):** [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.660 a, b (P1 §18.3.1.96 v), p.286: `v` at most 32767 characters; Boolean `1`/`0`; error cells hold one of `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`.

### E-24 - `row` element: is `r` required?
- **Fact:** `row@r` is `xsd:unsignedInt` and **optional**; it is the row index. Other attributes:
  - `spans` (`ST_CellSpans`) is an optional optimisation and need not be written;
  - `s` (default `0`) applies only if `customFormat` = 1.
  - Annex L (informative) says each row carries an index (`r`) so that empty rows need not be written.
- **Source:**
  - Part 1 §18.3.1.73 row (Row), p.1675–1680, L79048–79300 (`r` at L79173–79176, p.1678; `spans` at L79187–79200, p.1679).
  - `xsd-strict/sml.xsd` CT_Row L2238–2255.
  - Annex L.2.2.8, p.4646, L249438–249440.
- **Strict/Transitional:** Identical (`xsd-trans/sml.xsd` L2256).

### E-25 - Ordering of rows/cells; omitted `c@r`
- **Fact:** `c@r` is optional in the schema (E-20).
- **NOT FOUND:**
  - No requirement that `row` elements or `c` elements appear in ascending order.
  - No rule for deriving a cell's position when `c@r` (or `row@r`) is omitted.
  - Searched Part 1 L74800–81000 and Annex L for: "ascending", "sorted order", "increasing order", "order of rows", "rows shall", "cells shall be", "omitted", "not present", "row-major".
  - The examples always show ascending order with explicit `r`, but that is never stated normatively.
**Checked 2026-10-30 (D-1104):** [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.590 c, p.251: a cell without `r` is in the column one greater than the previous cell of the row, the first such cell in column 1. §2.1.644 a, p.279: `row@r` 1 … 1048576, non-decreasing.

### E-26 - `f` (Formula) attributes
- **Fact:** The content type is `ST_Formula` (a restriction of `ST_Xstring`). CT_CellFormula attributes:

| Attr | Type | Default | Meaning |
|---|---|---|---|
| `t` | `ST_CellFormulaType` | `normal` | formula type |
| `aca` | boolean | `false` | array formulas only: calculate the whole array; ignored unless `t=array` |
| `ref` | `ST_Ref` | none | range the formula applies to; required only for shared, array or data-table formulas; written only on the master formula |
| `dt2D` | boolean | `false` | the data table is two-dimensional; master cell only |
| `dtr` | boolean | `false` | a one-dimensional data table is a row (else a column); master only |
| `del1` / `del2` | boolean | `false` | input cell 1 / 2 of the data table deleted; master only |
| `r1` / `r2` | `ST_CellRef` | none | first / second input cell of the data table (`r2` when `dt2D` = 1); master only |
| `ca` | boolean | `false` | recalculate on the next calculation (set for volatile functions and circular references) |
| `si` | `xsd:unsignedInt` | none | shared group index |
| `bx` | boolean | `false` | the formula assigns a value to a name |

- `xml:space` is also allowed.
- **Source:**
  - Part 1 §18.3.1.40 f (Formula), p.1630–1633, L76864–77038 (raw-checked, PDF p.1641–1643).
  - `xsd-strict/sml.xsd` CT_CellFormula L2730–2747; ST_Formula L213.
- **Strict/Transitional:** Identical (`xsd-trans/sml.xsd` L2751).
**Checked 2026-10-30 (D-1104):** [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.620 b, e, h (P1 §18.3.1.40 f), p.260–261: `bx` must be false; `t="array"` marks array formulas and `aca` is ignored otherwise; `si` is required with `t="shared"` and ignored otherwise.

### E-27 - ST_CellFormulaType
- **Fact:** The values are:
  - `normal`: a regular cell formula (default);
  - `array`: an array formula, one formula applied across a range that can return multiple results;
  - `dataTable`: a data-table formula;
  - `shared`: part of a shared formula.
- **Source:**
  - Part 1 §18.18.6, p.2440, L116089–116106.
  - §18.3.1.40, p.1630–1631.
  - `xsd-strict/sml.xsd` L2281.
- **Strict/Transitional:** Identical.

### E-28 - Shared formulas
- **Fact:**
  - The first formula in an `si` group is the **master**. Its `f` holds the formula text plus `t="shared"`, `ref` and `si`.
  - Later cells need not write formula text. They carry `t="shared"` and `si`, and their formula is derived from their position relative to the master.
  - A cell is shared only when **both** `si` is present and `t` = `shared`.
  - For a non-master shared cell, any formula text written in it shall be **ignored**; the master overrides.
  - If a cell inside the master's `ref` range does not carry `si`/`t="shared"`, its own formula overrides. Cells before and after it keep sharing.
  - Handling an `si`/`shared` cell outside the master's `ref` range is implementation-defined.
  - Master ranges on the same sheet shall not overlap.
  - Two formulas count as "the same" when their R1C1 representations are identical.
- **Annex L example (informative):** The dependent cell H8 does include formula text, `SUM(E8:G8)`. It also uses an attribute `ce` that is not in the schema. A reader should still use the master formula.
- **Source:**
  - Part 1 §18.3.1.40, p.1631–1633, L76910–77021.
  - Annex L.2.2.9.2.1, p.4647, L249478–249508.

### E-29 - `is` (inline rich string) structure
- **Fact:** `is` has type `CT_Rst`, the same type as shared-string `si`. Its sequence:
  - `t?` (ST_Xstring);
  - `r*` (rich-text run, CT_RElt: `rPr?` then `t`);
  - `rPh*` (phonetic run, CT_PhoneticRun: `t`; required attributes `sb` and `eb`, both unsignedInt);
  - `phoneticPr?` (required `fontId`; `type` default `fullwidthKatakana`; `alignment` default `left`).
- **Source:**
  - `xsd-strict/sml.xsd` CT_Rst L1833, CT_RElt L1808, CT_PhoneticRun L1801, CT_PhoneticPr (after CT_Rst).
  - Part 1 §18.3.1.53, p.1646.
- **Strict/Transitional:** Identical (diff of the CT_Sst…CT_PhoneticPr block).

### E-30 - Shared strings part: `sst`, `si`, `t`, `r`
- **Fact:** The root is `sst` (CT_Sst): `si*` then `extLst?`, with optional attributes `count` and `uniqueCount` (unsignedInt).
  - Each `si` (String Item) is one string.
  - A simple string *should* be a single `t`.
  - A string with character-level formatting *shall* be made of multiple rich-text runs `r` that together express the string.
  - `r`: a run of text sharing one set of properties. Properties are in `rPr`; the displayed text is in `t`.
  - `t`: the text content, typed `ST_Xstring`, with `xml:space` allowed.
  - Cells reference `si` by **zero-based index**: the `v` of a `t="s"` cell (Annex L.2.2.9.1, p.4646).
- **Source:**
  - Part 1 §18.4 Shared String Table, p.1717–1719, L81158–81245.
  - §18.4.4 r, p.1722, L81397; §18.4.8 si, p.1724, L81475; §18.4.9 sst, p.1724, L81486; §18.4.12 t, p.1725, L81552.
  - `xsd-strict/sml.xsd` L1776–1784.
- **Strict/Transitional:** Identical.
**Checked 2026-10-30 (D-1104):** [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.675 a–c, §2.1.678 b, p.289–290: `t` at most 32767 characters; at most 32767 `r` and `rPh`.

### E-31 - `rPh` (phonetic run): is its text part of the value?
- **Fact:** `rPh` is a run of text that shows a phonetic hint for East Asian pronunciation, displayed across the top of the cell.
  - `sb` and `eb` are zero-based offsets into the *base text*. Each must be between 0 and the base text length, and **sb < eb** is required.
  - Non-overlapping consecutive runs are recommended.
- **NOT FOUND:** No explicit statement on whether `rPh/t` text is part of the cell's string value.
- **Derived:** The `si` definition (§18.4.8) builds the string from `t` or the rich-text runs `r` only, and `rPh` is described as a hint over the base text. Excluding `rPh/t` from the value is the reading consistent with the spec.
- **Source:** Part 1 §18.4.6 rPh, p.1723, L81426–81462; §18.4.3 phoneticPr, p.1721–1722, L81342–81395.

### E-32 - `sst@count` / `@uniqueCount`
- **Fact:**
  - `count` is the total number of strings in the workbook. It counts text strings only, not numbers.
  - `uniqueCount` is the number of unique entries in the table. A string with different character-level formatting counts as unique.
  - Each attribute is optional unless the other is present, in which case it is required. Write both or neither.
- **Spec inconsistency:** The §12.3.15 example uses `totalCount` and `<sstItem>`, which match neither the schema nor §18.4 (p.89, L4280–4292).
- **Source:** Part 1 §18.4.9 sst, p.1724, L81486–81518.

### E-33 - Cell error values
- **Fact:** The complete list of error values is: `#DIV/0!`, `#N/A`, `#NAME?`, `#NULL!`, `#NUM!`, `#REF!`, `#VALUE!`, `#GETTING_DATA`.
  - `#GETTING_DATA` means a value not yet retrieved or calculated, for example from an OLAP cube.
  - The same set is the `error-constant` production of the formula grammar.
  - ERROR.TYPE maps them to: `#NULL!` 1, `#DIV/0!` 2, `#VALUE!` 3, `#REF!` 4, `#NAME?` 5, `#NUM!` 6, `#N/A` 7, `#GETTING_DATA` 8.
- **Source:**
  - Part 1 §18.17.3 Error values, p.2060–2061, L97294–97355.
  - §18.17.2.1 Constants (`error-constant`), p.2045, L96442–96444.
  - ERROR.TYPE, p.2199–2200, L104466–104484.

---

## E. Strings, escaping, ranges (facts 10–11, 13, 16)

### E-34 - ST_Xstring and the `_xHHHH_` escape
- **Fact (§22.9.2.19 ST_Xstring "Escaped String"):** Characters that cannot be represented in XML 1.0 are escaped as `_xHHHH_`, where H is a hexadecimal digit of the character's value. Example: U+0008 → `_x0008_`. The base type is `xsd:string`.
- **Literal `_x` sequences:** §22.9.2.19 does **not** say how to escape them.
  - The rule appears only for the `vt:bstr` variant type (§22.4.2.4 bstr, p.3736, L174882–174889).
  - Rule: to store an escape sequence literally, escape its initial underscore as `_x005F_`. Example: the literal string `_x0008_` is stored as `_x005F_x0008_`.
  - §22.4.2.17 lpstr / §22.4.2.18 lpwstr and §22.2.2.21 Properties restate only the `_xHHHH_` rule.
- **Source:**
  - Part 1 §22.9.2.19, p.3801, L178057–178068.
  - `xsd-strict/shared-commonSimpleTypes.xsd` L67; `xsd-trans/shared-commonSimpleTypes.xsd` L114 (plain `xsd:string`).
- **Real-data examples in the spec:** `_x000a_` (L79381) and `_x000d__x000a_` (L94292).
- **Strict/Transitional:** No difference. No Part 4 text on ST_Xstring.

### E-35 - `mergeCells` / `mergeCell@ref`
- **Fact:**
  - `mergeCells` holds 1..n `mergeCell` and an optional `count` (unsignedInt).
  - `mergeCell@ref` is `ST_Ref` and **required**.
  - Formatting and content of a merged range are always stored in its **top-left** cell (from the example text).
- **Source:**
  - Part 1 §18.3.1.54 mergeCell and §18.3.1.55 mergeCells, p.1646–1647, L77645–77678.
  - `xsd-strict/sml.xsd` CT_MergeCells L2452, CT_MergeCell L2458.
- **Strict/Transitional:** Identical.

### E-36 - `dimension@ref`
- **Fact:** `dimension` gives the used range of the worksheet, and the element is optional.
  - `ref` (`ST_Ref`, required) is the range that would contain all `c` elements under `sheetData`.
  - Whole-column and whole-row notation is not supported.
  - Used cells include cells with formulas, text or formatting.
- **Source:**
  - Part 1 §18.3.1.35 dimension, p.1622, L76458–76476.
  - `xsd-strict/sml.xsd` CT_SheetDimension L2306.
- **Reader note:** It is an optimisation only. Do not trust it for bounds.

### E-37 - A1 reference syntax
- **Fact (grammar):**
  - `A1-reference = A1-column ":" A1-column | A1-row ":" A1-row | A1-column A1-row`
  - A column is one or more letters a–z/A–Z, optionally prefixed by `$`.
  - A row is a digit sequence, which **shall be non-zero**, optionally prefixed by `$`.
  - Columns run A–Z, AA–AZ, …, ZA–ZZ, AAA…
  - **Column letters are not case-sensitive.**
  - Areas are joined with `:`; the comma is the union operator.
- **Fact (storage):** Regardless of `calcPr@refMode`, cell references are always stored in XML in A1 form. R1C1 is a runtime display style only.
- **Source:**
  - Part 1 §18.17.2.3.1 A1-Style Cell References, p.2053–2054, L96906–96955.
  - §18.17.6.1 Cell Reference Style, p.2067, L97612–97628.

### E-38 - ST_CellRef, ST_Ref, ST_RefA, ST_Sqref
- **Fact:**
  - `ST_CellRef`: a single cell reference.
  - `ST_Ref`: a cell or range reference.
  - `ST_RefA`: a single cell reference that may be absolute.
  - `ST_Sqref`: a space-delimited list of `ST_Ref`.
  - In both xsds, ST_CellRef, ST_Ref and ST_RefA are plain `xsd:string` restrictions **with no pattern**, and ST_Sqref is `xsd:list` of ST_Ref. The schema gives no syntactic validation.
- **Source:**
  - Part 1 §18.18.7 (p.2440, L116110), §18.18.62 (p.2480, L118030), §18.18.63 (p.2480, L118044), §18.18.76 (p.2487, L118381).
  - `xsd-strict/sml.xsd` L201–212; `xsd-trans/sml.xsd` L204–215.

### E-39 - Maximum rows and columns
- **Fact:**
  - Cell references range from A1–A1048576 to XFD1–XFD1048576, i.e. 1,048,576 rows × 16,384 columns. The column count is derived from XFD.
  - An implementation **can extend** this range.
  - A defined name of the form used by cells in A1–XFD1048576 shall be treated as an error.
  - Guidance in §18.17.2.3 encourages implementations not to limit rows and columns unnecessarily.
- **Source:**
  - Part 1 §18.17.5.1 Limits, p.2065, L97529–97545.
  - §18.17.2.3, p.2049, L96684–96685.

---

## F. Number formats and styles (facts 14–15)

### E-40 - `numFmts` / `numFmt`
- **Fact:**
  - `numFmts` holds `numFmt*` and an optional `count`.
  - `numFmt@numFmtId` has type `ST_NumFmtId` (`xsd:unsignedInt`) and is required.
  - `numFmt@formatCode` has type `ST_Xstring` and is required.
  - `numFmtId` is a real ID, not an index. `xf` records reference a `numFmt` by matching `numFmtId` (Annex L.2.7.3.1, informative, p.4677).
- **Source:**
  - Part 1 §18.8.30 numFmt (p.1776–1783, attribute table at L84303–84311), §18.8.31 numFmts (p.1783–1792), §18.18.47 ST_NumFmtId (L117452).
  - `xsd-strict/sml.xsd` CT_NumFmts L3547, CT_NumFmt L3553, ST_NumFmtId L3623.
- **Strict/Transitional:** Identical (diff).

### E-41 - Built-in (implied) number formats, "All Languages"
- **Fact:** These formats are implied: an `xf` may use the `numFmtId` without a `numFmt` element. Some IDs may be interpreted differently depending on the UI language.
  - Using IDs outside the lists is implementation-defined.
  - **IDs not in the listing (e.g. 5, 6, 7, 8) shall use the `formatCode` attribute.** That is, they need an explicit `numFmt`.
  - Codes are verbatim from the spec (raw extraction, PDF p.1786–1787). Date/time marking is mine, based on the tokens.

| ID | formatCode | Date/time? |
|---|---|---|
| 0 | `General` | |
| 1 | `0` | |
| 2 | `0.00` | |
| 3 | `#,##0` | |
| 4 | `#,##0.00` | |
| 9 | `0%` | |
| 10 | `0.00%` | |
| 11 | `0.00E+00` | |
| 12 | `# ?/?` | |
| 13 | `# ??/??` | |
| 14 | `mm-dd-yy` | **date** |
| 15 | `d-mmm-yy` | **date** |
| 16 | `d-mmm` | **date** |
| 17 | `mmm-yy` | **date** |
| 18 | `h:mm AM/PM` | **time** |
| 19 | `h:mm:ss AM/PM` | **time** |
| 20 | `h:mm` | **time** |
| 21 | `h:mm:ss` | **time** |
| 22 | `m/d/yy h:mm` | **date+time** |
| 37 | `#,##0 ;(#,##0)` (a space before `;` in the spec text) | |
| 38 | `#,##0 ;[Red](#,##0)` (a space before `;`) | |
| 39 | `#,##0.00;(#,##0.00)` | |
| 40 | `#,##0.00;[Red](#,##0.00)` | |
| 45 | `mm:ss` | **time** |
| 46 | `[h]:mm:ss` | **elapsed time** |
| 47 | `mmss.0` | **time** |
| 48 | `##0.0E+0` | |
| 49 | `@` | (text) |

- **Not in the list:** 5–8, 23–26, 41–44.
- **Source:** Part 1 §18.8.30 numFmt, p.1776–1777, L83957–84010.
- **Also stated ("General" format):** dates do not follow "General" and are converted to date formatting automatically. p.1778, L84027.

### E-42 - Locale-specific implied IDs (27–36, 50–58 CJK; 59–62, 67–81 Thai)
- **Fact:** The spec lists locale-dependent built-ins.
  - In the glyph tables, `pdftotext` lost the CJK and Thai glyphs. The spec also gives "with unicode values" tables, where a glyph is shown as a hex code point (e.g. `5E74` = U+5E74). Those versions are transcribed below.
  - Code points written bare (not in quotes) stand for literal glyphs. For example, `4E0A5348/4E0B5348` stands for U+4E0A U+5348 / U+4E0B U+5348, the AM/PM markers.
  - All of these IDs are date or time formats.
- **zh-tw / zh-cn:**

| ID | zh-tw | zh-cn |
|---|---|---|
| 27 | `[$-404]e/m/d` | `yyyy"5E74"m"6708"` |
| 28 | `[$-404]e"5E74"m"6708"d"65E5"` | `m"6708"d"65E5"` |
| 29 | `[$-404]e"5E74"m"6708"d"65E5"` | `m"6708"d"65E5"` |
| 30 | `m/d/yy` | `m-d-yy` |
| 31 | `yyyy"5E74"m"6708"d"65E5"` | `yyyy"5E74"m"6708"d"65E5"` |
| 32 | `hh"6642"mm"5206"` | `h"65F6"mm"5206"` |
| 33 | `hh"6642"mm"5206"ss"79D2"` | `h"65F6"mm"5206"ss"79D2"` |
| 34 | `4E0A5348/4E0B5348hh"6642"mm"5206"` | `4E0A5348/4E0B5348h"65F6"mm"5206"` |
| 35 | `4E0A5348/4E0B5348hh"6642"mm"5206"ss"79D2"` | `4E0A5348/4E0B5348h"65F6"mm"5206"ss"79D2"` |
| 36 | `[$-404]e/m/d` | `yyyy"5E74"m"6708"` |
| 50 | `[$-404]e/m/d` | `yyyy"5E74"m"6708"` |
| 51 | `[$-404]e"5E74"m"6708"d"65E5"` | `m"6708"d"65E5"` |
| 52 | `4E0A5348/4E0B5348hh"6642"mm"5206"` | `yyyy"5E74"m"6708"` |
| 53 | `4E0A5348/4E0B5348hh"6642"mm"5206"ss"79D2"` | `m"6708"d"65E5"` |
| 54 | `[$-404]e"5E74"m"6708"d"65E5"` | `m"6708"d"65E5"` |
| 55 | `4E0A5348/4E0B5348hh"6642"mm"5206"` | `4E0A5348/4E0B5348h"65F6"mm"5206"` |
| 56 | `4E0A5348/4E0B5348hh"6642"mm"5206"ss"79D2"` | `4E0A5348/4E0B5348h"65F6"mm"5206"ss"79D2"` |
| 57 | `[$-404]e/m/d` | `yyyy"5E74"m"6708"` |
| 58 | `[$-404]e"5E74"m"6708"d"65E5"` | `m"6708"d"65E5"` |

- **ja-jp / ko-kr:**

| ID | ja-jp | ko-kr |
|---|---|---|
| 27 | `[$-411]ge.m.d` | `yyyy"5E74" mm"6708" dd"65E5"` |
| 28 | `[$-411]ggge"5E74"m"6708"d"65E5"` | `mm-dd` |
| 29 | `[$-411]ggge"5E74"m"6708"d"65E5"` | `mm-dd` |
| 30 | `m/d/yy` | `mm-dd-yy` |
| 31 | `yyyy"5E74"m"6708"d"65E5"` | `yyyy"B144" mm"C6D4" dd"C77C"` |
| 32 | `h"6642"mm"5206"` | `h"C2DC" mm"BD84"` |
| 33 | `h"6642"mm"5206"ss"79D2"` | `h"C2DC" mm"BD84" ss"CD08"` |
| 34 | `yyyy"5E74"m"6708"` | `yyyy-mm-dd` |
| 35 | `m"6708"d"65E5"` | `yyyy-mm-dd` |
| 36 | `[$-411]ge.m.d` | `yyyy"5E74" mm"6708" dd"65E5"` |
| 50 | `[$-411]ge.m.d` | `yyyy"5E74" mm"6708" dd"65E5"` |
| 51 | `[$-411]ggge"5E74"m"6708"d"65E5"` | `mm-dd` |
| 52 | `yyyy"5E74"m"6708"` | `yyyy-mm-dd` |
| 53 | `m"6708"d"65E5"` | `yyyy-mm-dd` |
| 54 | `[$-411]ggge"5E74"m"6708"d"65E5"` | `mm-dd` |
| 55 | `yyyy"5E74"m"6708"` | `yyyy-mm-dd` |
| 56 | `m"6708"d"65E5"` | `yyyy-mm-dd` |
| 57 | `[$-411]ge.m.d` | `yyyy"5E74" mm"6708" dd"65E5"` |
| 58 | `[$-411]ggge"5E74"m"6708"d"65E5"` | `mm-dd` |

- **th-th:**
  - Numeric formats (`t` prefix): 59 `t0`, 60 `t0.00`, 61 `t#,##0`, 62 `t#,##0.00`, 67 `t0%`, 68 `t0.00%`, 69 `t# ?/?`, 70 `t# ??/??`.
  - Date/time formats:
    - 71 `0E27/0E14/0E1B0E1B0E1B0E1B`
    - 72 `0E27-0E140E140E14-0E1B0E1B`
    - 73 `0E27-0E140E140E14`
    - 74 `0E140E140E14-0E1B0E1B`
    - 75 `0E0A:0E190E19`
    - 76 `0E0A:0E190E19:0E170E17`
    - 77 `0E27/0E14/0E1B0E1B0E1B0E1B 0E0A:0E190E19`
    - 78 `0E190E19:0E170E17`
    - 79 `[0E0A]:0E190E19:0E170E17`
    - 80 `0E190E19:0E170E17.0`
    - 81 `d/m/bb`
- **Observation:** The token `bb` (in 81) is not defined anywhere in the grammar. Only `b1`/`b2` are defined (E-46).
- **Source:** Part 1 §18.8.30, p.1778–1783, L84060–84294 (raw-checked, PDF p.1789–1793).

### E-43 - Format code sections
- **Fact:**
  - A format code has up to four sections separated by `;`: positive; negative; zero; text.
  - With two sections, the first is for positive numbers and zero, and the second for negatives.
  - With one section, it applies to all numbers.
  - To skip a section, write its terminating `;`.
  - A text section, if present, shall be the last section. `@` marks where the cell text goes; without `@`, the text is not displayed.
  - If there is no text section, text is not affected by the format.
- **Source:** Part 1 §18.8.31 numFmts, "Number Format Codes", p.1783–1786, L84333–84462.

### E-44 - Format symbols, literals, `\`, `"…"`, `_`, `*`
- **Fact:**
  - **Digit placeholders:**
    - `0` shows insignificant zeros;
    - `#` shows significant digits only;
    - `?` pads insignificant zeros with spaces so decimal points align; it is also used in fractions.
  - **Decimal point:** `.`
  - **Percent:** `%` multiplies by 100 and appends `%`.
  - **Comma:** `,` is a thousands separator when it is surrounded by `#` or `0`. A comma that follows a placeholder scales the number by 1000 per comma (e.g. `0.0,,`).
  - **Exponent:** `E-`, `E+`, `e-`, `e+` give scientific format. `E+` and `e+` also write `+` for positive exponents.
  - **Displayed as-is:** `$ - + ( ) : space`.
  - **Slash:** `/` is a fraction symbol when preceded and followed by `0`, `#` or `?`; otherwise it is a literal slash.
  - **Backslash:** `\` shows the next character literally; the backslash itself is not shown.
  - **Asterisk:** `*` repeats the next character to fill the column width. At most one per section; if there are more, all but the last are ignored.
  - **Underscore:** `_` skips the width of the next character, e.g. `_)`.
  - **Quotes:** `"text"` shows the quoted text.
  - **At sign:** `@` is the text placeholder.
  - **Literal text:** Characters inside double quotes or right after a backslash are **never** interpreted as format tokens. Single quotes **shall not** be used to denote text.
  - **Characters displayed without quoting:** `$ - + / ( ) : ! ^ & ' ~ { } < > =` and space.
- **Source:**
  - Part 1 §18.8.31, symbol table p.1784–1785, L84354–84429 (raw-checked, PDF p.1794–1795).
  - "Text and spacing", p.1785–1786, L84430–84480.

### E-45 - Bracketed sections: colors, conditions, locale/currency
- **Fact:**
  - **Color names:** `[Black]`, `[Blue]`, `[Cyan]`, `[Green]`, `[Magenta]`, `[Red]`, `[White]`, `[Yellow]`. The color must be the first item in its section.
  - **Indexed colors:** `[ColorN]` with N in 1–56 indexes the legacy palette. `[Color1]` maps to `indexed="8"` … `[Color56]` to `indexed="63"`.
  - **Conditions:** `[op value]` with op one of `=`, `>`, `<`, `>=`, `<=`, `<>`, e.g. `[Red][<=100];[Blue][>100]`. If no condition matches, `#` is shown across the cell.
  - **Locale/currency:** `[$<Currency String>-<language info>]`, e.g. `[$USD-409]`. The language info is a 32-bit hex value:
    - bytes 0–1: LCID;
    - byte 2: calendar type (high bit: parse input with this calendar);
    - byte 3: number system type (high bit: parse input with this number system).
  - **Special language-info values:** `0xf800` = system long date format; `0xf400` = system time format.
- **Source:**
  - Part 1 §18.8.31: colors p.1787–1788, L84531–84546; conditions p.1788, L84548–84558.
  - `[$…]`: p.1791–1792, L84700–84722 (raw-checked, PDF p.1801).

### E-46 - Date/time tokens
- **Fact:**
  - **Months:** `m` 1–12, `mm` 01–12, `mmm` Jan–Dec, `mmmm` January–December, `mmmmm` J–D.
  - **Days:** `d` 1–31, `dd` 01–31, `ddd` Sun–Sat, `dddd` Sunday–Saturday.
  - **Years:** `yy` 00–99, `yyyy` date-base minimum value–9999.
  - **Hours:** `h` 0–23, `hh` 00–23.
  - **Minutes:** `m` 0–59, `mm` 00–59.
  - **Seconds:** `s` 0–59, `ss` 00–59.
  - **Fractional seconds:** e.g. `h:mm:ss.00`.
  - **12-hour clock:** `AM/PM` or `A/P`. If the format contains AM or PM, the hour uses the 12-hour clock; otherwise 24-hour.
  - **Elapsed time:** `[h]:mm`, `[mm]:ss`, `[ss].00`.
  - **Minutes vs month:** `m`/`mm` mean **minutes** only when they come immediately after `h`/`hh` or immediately before `ss`; otherwise they mean month.
  - **Illegal values:** date/time-formatted cells with values outside the requirements show `#` across the cell.
  - **International codes:**
    - `r` (ja-jp/zh-tw only): becomes `ee` in ja-jp and `e` in zh-tw.
    - `rr`: becomes `gggee` in ja-jp and `e` in zh-tw.
    - `g`: single Roman-character era (ja-jp); same as `gg` in zh-tw.
    - `gg`: single Kanji era (ja-jp); last era short name since 1911 (zh-tw).
    - `ggg`: two-Kanji era (ja-jp); long name (zh-tw).
    - `e`: era year (ja-jp); era year since 1912 (zh-tw); **in OTHER locales becomes `yyyy`**.
    - `ee`: era year with leading zero; **in OTHER locales becomes `yy`**.
    - `b2`: Hijri calendar.
    - `b1`: Gregorian calendar.
- **Source:**
  - Part 1 §18.8.31 "Dates and times", p.1789–1790, L84576–84655 (raw-checked, PDF p.1799–1800).
  - "International Considerations", p.1790–1791, L84656–84699. The `b1`/`b2` pairing was confirmed with raw extraction of PDF p.1801; in layout mode the description column is shifted one row.

### E-47 - How `c@s` reaches a number format
- **Fact:**
  - **`cellXfs`:** holds the master formatting records (`xf`) for cells. Cells in the Sheet part reference these `xf` records by **zero-based index**, i.e. `c@s` indexes `cellXfs/xf`; default `s=0`.
  - **Both layers are read:** a cell can have direct formatting and a cell style, so both the cell-style `xf` records (`cellStyleXfs`) and the cell `xf` records shall be read to get the full formatting.
  - **`xf@xfId`:** for an `xf` in `cellXfs`, this is the zero-based index of the `cellStyleXfs` `xf` for the cell's style. It is absent on `cellStyleXfs` records.
  - **`xf@numFmtId`:** the ID of the `numFmt` record, looked up either among the built-ins (E-41/E-42) or in `numFmts` by ID.
  - **`applyNumberFormat`:** a boolean saying whether this `xf`'s number formatting should be applied.
- **Schema:**
  - All `xf` attributes are optional.
  - `applyNumberFormat` (and the other `apply*` attributes) have **no default** in the xsd.
  - `quotePrefix` and `pivotButton` default to `false`.
- **NOT FOUND:** No further normative rule for resolving `applyNumberFormat` absent/0 against the style `xf`.
- **Source:**
  - Part 1 §18.8.10 cellXfs, p.1763, L83343–83360.
  - §18.8.9 cellStyleXfs, p.1762, L83301–83341.
  - §18.8.45 xf (Format), p.1799–1801, L85084 (raw-checked, PDF p.1809–1811).
  - Annex L.2.7.3.6 (informative, p.4679, L250890–250897): an `xf` in `cellXfs` is referenced by zero-based index (`s`) from a `c`.
  - `xsd-strict/sml.xsd` CT_CellXfs L3563, CT_Xf L3569–3588.
- **Strict/Transitional:** Identical (`xsd-trans/sml.xsd` CT_Xf L3598).

---

## G. Packaging, Part 2 (facts 17–21)

### E-48 - `[Content_Types].xml`: Default and Override semantics and precedence
- **Fact:**
  - **Coverage:** For every part **other than Relationships parts**, the Media Types stream shall specify one matching `Default`, one matching `Override`, or both. When both match, **`Override` takes precedence**.
  - **Uniqueness:** At most one `Default` per extension and at most one `Override` per part name.
  - **Order** of elements is not significant. Unused `Default`s are allowed.
  - **Default matching:** a `Default` matches any part whose name ends with `.` followed by its `Extension`.
  - **Override matching:** an `Override` matches the part whose name equals its `PartName`.
  - **Lookup procedure (§7.2.3.5):**
    1. Compare the part name with Override `PartName`s using **ASCII case-insensitive** matching. On a match, return its media type.
    2. Otherwise take the extension: the substring right of the rightmost `.` in the last segment.
    3. Compare it with Default `Extension`s using **ASCII case-insensitive** matching.
  - **Comparing media types** when adding parts: case-insensitive, every character included (§7.2.3.4).
  - **Stream rules:** The stream is not a part and is not URI-addressable. Its root is `Types` (no attributes). It shall not use the ECMA-376-3 MCE mechanisms.
- **Schema:** `opc-contentTypes.xsd` CT_Types is a choice of `Default` and `Override`, minOccurs 0. The prose says "one or more".
- **Source:** Part 2 §7.2.3.1–§7.2.3.5, p.27–31, L1444–1632 (raw-checked, PDF p.38–39).

### E-49 - ZIP name of the Media Types stream
- **Fact:** In ZIP the stream is stored as the item `[Content_Types].xml`. When interleaved, it is the complete sequence `[Content_Types].xml/[0].piece` … `[n].last.piece`. The brackets were chosen because they violate the part-name grammar.
- **Source:** Part 2 §7.3.7, p.35, L1838–1850.

### E-50 - Part name grammar
- **Fact:**
  - A part name is a Unicode string matching `part_name = 1*( "/" isegment-nz )` (RFC 3987 `isegment-nz`), with these constraints:
    - no segment contains a percent-encoded `/` or `\`;
    - no segment contains percent-encoded characters that match RFC 3987 `iunreserved`;
    - no segment ends with `.`.
  - Support for non-ASCII part names is recommended but not required.
  - No part name may be derivable from another: `/a/b` cannot coexist with `/a`.
- **Source:** Part 2 §6.2.2.2 Syntax, p.10–11, L581–612; §6.2.2.3, p.11, L626–640.

### E-51 - Part name equivalence (case-insensitive)
- **Fact:**
  - Part-name equivalence is **ASCII case-insensitive matching**: code points 0x41–0x5A (A–Z) are compared as 0x61–0x7A (a–z). Non-ASCII characters are compared exactly.
  - Two different parts shall not have equivalent names (`/a` vs `/A`).
  - NFC/NFD forms of two part names *should* not be equivalent, and names *should* not be "weakly derivable" from one another.
- **Source:**
  - Part 2 §6.2.2.3 Part name equivalence and integrity in an abstract package, p.11–12, L616–674.
  - Definition 3.3.11, p.6, L493–496.

### E-52 - Relationships part naming; package relationships
- **Fact:**
  - **Package relationships:** the package Relationships part is named `/_rels/.rels`.
  - **Part relationships:** the Relationships part for a source part is built by adding `.rels` to the source's last segment and inserting a `_rels` segment before it. Examples: `/foo` → `/_rels/foo.rels`; `/foo/bar.xml` → `/foo/_rels/bar.xml.rels`.
  - **Reserved names:** `/_rels/.rels`, and any name whose second-to-last segment is equivalent to `_rels` and whose last segment ends with `.rels`.
  - **No chaining:** there shall be no relationships from or to a Relationships part.
  - **XML validity:** Relationships parts are XML, schema-valid against `opc-relationships.xsd` after MCE processing, and must not contain `xml:base`.
- **Source:** Part 2 §6.2.2.2 (p.11, L600–602), §6.5.2.1–§6.5.2.3 (p.20–21, L1077–1145), §6.5.3.1 (p.21, L1149–1157).

### E-53 - `Relationship` element attributes; Id uniqueness; Type comparison
- **Fact:**

| Attribute | Required | Type | Rule |
|---|---|---|---|
| `Id` | yes | `xsd:ID` | unique within the Relationships part |
| `Type` | yes | `xsd:anyURI` | compared as plain strings: case-sensitive, character for character, no escaping done or undone |
| `Target` | yes | `xsd:anyURI` | see TargetMode |
| `TargetMode` | no | `Internal` / `External` | default `Internal` |

  - With `TargetMode=Internal`, `Target` shall be a relative reference to a part.
  - With `External`, it may be a relative reference or an absolute IRI.
  - The `Relationships` root has no attributes and holds 0..n `Relationship` elements.
- **Source:**
  - Part 2 §6.5.3.3–§6.5.3.4, p.22–23, L1167–1231 (raw-checked).
  - `opc-xsd/opc-relationships.xsd` CT_Relationship.

### E-54 - Resolving relative Targets
- **Fact:** Relative references are resolved per **RFC 3986 §5 (Reference Resolution)**, extended by RFC 3987 §6.5. The base IRI depends on where the reference sits:
  - **Internal, part Relationships part:** the pack IRI of the package plus the **source part name**. `bar.xml` from `/a/b/_rels/foo.xml.rels` resolves to `/a/b/bar.xml`, and `../bar.xml` to `/a/bar.xml`.
  - **Internal, package Relationships part:** the pack IRI of the package root `/`.
  - **External (either kind):** the absolute IRI of the package itself.
  - **Leading `/`:** a reference starting with `/` ignores the base path.
  - **Above the root:** `/../` collapses to `/` (RFC 3986 remove_dot_segments).
  - **Not a relationships part:** a relative reference inside a non-relationships part uses that part's pack IRI.
- **Source:**
  - Part 2 §6.4.1–§6.4.3 (p.16–19, L900–1036).
  - §6.5.2.2 (p.20, L1094–1112) and §6.5.2.3 (p.21, L1125–1145).

### E-55 - XML encoding and DTD prohibition, and its scope
- **Fact:** For XML content in the parts and streams **defined in Part 2**, namely the Media Types stream, the Core Properties part, Digital Signature XML Signature parts and Relationships parts:
  - **(a) Encoding:** UTF-8 or UTF-16. Any encoding declaration shall name only UTF-8 or UTF-16.
  - **(b) DTDs:** shall not be used (stated as a denial-of-service mitigation).
  - **(c) Namespaces:** the XML shall conform to XML Namespaces.
  - **(d) Schema validity:** the XML shall be schema-valid.
  - **(e) Reserved namespaces:** no `xml`/`xsi` items unless allowed.
- **Scope:** The scope is explicitly limited to those OPC-defined parts.
- **NOT FOUND:**
  - **Requirement IDs:** No `[M1.17]` or any `[Mx.y]` requirement ID exists in this 2021 edition. Searched `\[M[0-9]`. The introduction says the conformance-guidelines annex of earlier editions was dropped.
  - **Part 1:** No equivalent encoding or DTD rule for SpreadsheetML parts. Searched Part 1 for "DTD", "document type definition", "UTF-8 or", "UTF-16", and "encod" in clauses 1–9. Hits concern only password hashing and DOCTYPE examples.
- **Source:** Part 2 §6.2.5 XML usage, p.12–13, L693–720; Introduction, L236–240.

### E-56 - ECMA-376's own packaging constraints (Part 1 §9.1)
- **Fact:**
  - **ZIP:** every Office Open XML document is a ZIP archive (§9.1.3).
  - **Unknown parts:** parts other than relationship parts that are not the target of any relationship are unknown parts. They shall be ignored on reading (§9.1.4).
  - **Trash items:** they use `[trash]/HHHH.dat`, are not OPC part names, and have no content type (§9.1.5).
  - **Invalid items:** other ZIP items that are not valid part names or lack a content type are not allowed (§9.1.6).
  - **Unknown relationships:** consumers shall not fail on them (§9.1.7).
  - **Interleaving:** shall not be used (§9.1.8).
- **Source:** Part 1 §9.1.3–§9.1.8, p.21–22, L1071–1125.

### E-57 - ZIP features allowed or prohibited (Part 2 §7.3.6 and Annex B, normative)
- **Compression:**
  - §7.3.6: "shall not use compression algorithms except DEFLATE".
  - Table B.4 (p.69–70): method **0 stored** = Yes/Yes/Yes; method **8 deflated** = Yes/Yes/Yes; methods 1–7 and 9–11 (Shrunk, Reduced 1–4, Imploded, Tokenizing, Deflate64, PKWARE DCL Implode, Reserved) = No.
- **Encryption:**
  - §7.3.6: shall not be included.
  - Table B.3: all encryption features (traditional PKWARE, DES, 3DES, RC2, RC4, AES, non-OAEP, central-directory encryption) = No.
  - Table B.5: bit 0 (encrypted) = No; bits 6 and 13 = No.
  - Table B.1: archive decryption header and archive extra data record = No.
- **Multi-disk:** not supported. Disk-number fields are "partial - no multi disk archives"; production writes "always 1 disk" (Table B.2, p.66–68).
- **ZIP64:**
  - Table B.3: version 4.5 "File uses ZIP64 format extensions" = Yes/Yes/Yes.
  - Table B.1: Zip64 EOCD record **V1** and locator: consumption Yes (partially); production "used only when needed".
  - Zip64 EOCD **V2** (spec 6.2) = No/No/No.
  - Table B.6: extra field 0x0001 (ZIP64 extended info) = Yes/Yes/Optional.
  - All 64-bit sizes and offsets must have the high-order bit 0. "Number of entries" fields must not exceed 2,147,483,647 (p.73).
- **Data descriptors:**
  - Table B.1: consumption Yes; production Optional.
  - Table B.5 bit 3 (CRC and sizes zero in the local header, real values in the data descriptor) = Yes/Yes/Yes.
  - B.3: the descriptor may begin with signature `0x08074b50`; readers should handle either form.
  - B.2: values in the central-directory file header, local header and data descriptor must be equal, except as allowed for bit 3.
- **Other general-purpose bits:** bits 1–2 (deflate options) = Yes. Bits 4, 5 and 7–15 = consumption No, production "Bits set to 0". Bit 11 (the UTF-8 name flag in later APPNOTEs) is listed only as "Currently unused".
- **Volume labels and folders:** entries with version 1.1 (volume label) and 2.0 (folder/directory) are consumption "Yes (do not interpret as a part)" and production "No (rewrite/remove)".
- **Extra fields:** most PKWARE and third-party extra fields are consumption No, pass-through Yes. Padding `0xA220` (the growth hint) is consumption No, production Optional.
- **Limits:** a central-directory file header shall not exceed 65,535 bytes.
- **Source:** Part 2 §7.3.6 (p.35, L1816–1836); Annex B §B.1–B.4 and Tables B.1–B.7 (p.63–73, L2960–end), raw-checked against PDF p.73–83.

### E-58 - ZIP item names ↔ part names; case; duplicates
- **Fact:**
  - **Item names:** ZIP item names shall conform to the ZIP Appnote, **shall be ASCII only**, and **shall be unique** within the ZIP file (§7.3.3).
  - **Part → ZIP:** (a) remove the leading `/`; (b) **percent-encode every non-ASCII character** (§7.3.4).
  - **ZIP → part:** (a) **un-percent-encode every non-ASCII character**; (b) add a leading `/` (§7.3.5). Only items that represent files are mapped.
  - **Equivalent names:** logical item names are compared with **ASCII case-insensitive** matching. A physical package shall not contain equivalent logical item names (§7.2.5.2). So `xl/Workbook.xml` and `xl/workbook.xml` together are non-conformant.
  - **Unmappable names:** a logical item whose name does not follow the part-name grammar is not mapped to a part, e.g. `/[trash]/0000.dat` (§7.2.5.5).
  - **Informative preprocessing (Annex A):** examples include `/%41/%61.xml` → `/A/a.xml`, `\a.xml` → `/a.xml`, `/A B.xml` → `/A%20B.xml`.
- **Source:** Part 2 §7.2.5.2–§7.2.5.5 (p.32–33, L1684–1739), §7.3.3–§7.3.5 (p.34–35, L1778–1814), Annex A (p.61–62, L2903–2945).

### E-59 - Interleaving (for completeness)
- **Fact:** OPC defines interleaved pieces named `<name>/[0].piece` … `<name>/[n].last.piece` (complete, gap-free sequences). Part 1 §9.1.8 forbids interleaving in Office Open XML documents, so a strict XLSX reader can reject piece names.
- **Source:** Part 2 §7.2.4 (p.31), §7.2.5.2 (p.32, L1686–1710); Part 1 §9.1.8 (p.22).

---

## H. Summary of NOT FOUND items

| Item | What was searched |
|---|---|
| Strict `workbookPr@dateCompatibility` (E-12) | "dateCompatibility", "dateCompat", "date compat" in part1/part4 text and both sml.xsd |
| Explicit treatment of 29 Feb 1900 / serial 60 (E-16) | "leap year", "1900-02-29", "February 29", "29 February", "February 1900", "serial … 60" |
| Macrosheet part / relationship / content type (E-06) | "macrosheet", "relationships/macrosheet", "xlMacrosheet" |
| Macro-enabled workbook / template content types (E-07) | "macroEnabled", "ms-excel", "vbaProject", ".xlsm", ".xltm", "macro-enabled" |
| Row/cell ordering requirement; meaning of omitted `c@r` / `row@r` (E-25) | "ascending", "sorted order", "increasing order", "order of rows", "rows shall", "cells shall be", "omitted", "not present", "row-major" |
| Whether `rPh` text is part of the string value (E-31) | §18.4.3, §18.4.6, §18.4.8, Annex L.2.3 |
| `_x005F_` escape rule inside ST_Xstring itself (E-34) | Exists only for `vt:bstr`, §22.4.2.4 |
| `[M1.17]`-style requirement IDs (E-55) | `\[M[0-9]` in part2.txt; none in the 2021 edition |
| Part 1 encoding / DTD rule for non-OPC parts (E-55) | "DTD", "document type definition", "UTF-8 or", "UTF-16", "encod" |
| A normative boolean rule for non-formula `t="b"` cells beyond §18.17.6.6 (E-23) | - |
| A rule for resolving `applyNumberFormat` between cell and style `xf` (E-47) | - |
| Sheet-name length or case rules (E-09) | "31 characters", "sheet name" |

## I. Spec inconsistencies worth coding defensively against

1. **SST cardinality:** §12.3.15 requires exactly one Shared String Table part, but the §12.2 minimal package has none (E-08).
2. **SST example names:** §12.3.15 uses `totalCount`/`sstItem`, which are not schema names (E-32).
3. **Inline string example:** §18.3.1.53 `is` example has no `t="inlineStr"` (E-22).
4. **Shared formula example:** the Annex L example writes formula text plus a non-schema `ce` attribute on dependent shared-formula cells (E-28).
5. **Transitional 1900 base date:** Part 4 gives base 1899-12-31 = 0 and lower limit 1900-01-01 = 1, while keeping an example with base 1899-12-30 (E-15).
6. **Leftover date-system wording:** the DATE() examples and Annex L still describe a "1900 backward compatibility" date base, and Annex L has three date bases and UTC serials (E-19).
7. **externalLink relationship type:** the §18.17.2.3 example uses the `externalLinkPath` type for the Workbook→externalLink-part relationship, but §12.3.9 says `externalLink` (E-05).
8. **Element order in examples:** the §12.3.23 and §18.2.27 workbook examples place `workbookPr` after `sheets`, which violates the CT_Workbook sequence (`fileVersion`, `fileSharing`, `workbookPr`, `workbookProtection`, `bookViews`, `sheets`, …; `xsd-strict/sml.xsd` L4068–4076).
