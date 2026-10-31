# Facts: shared strings, ST_Xstring, styles (number formats)

> **Provenance.** Compiled on 2026-10-30. Line numbers (`Lnnn`) refer to text extracted with `pdftotext -layout` from the ECMA-376 PDFs (see docs/facts/ecma-376.md); `scratchpad/specs/…` paths are the working copies used during the check and are not stored in this repository.

Fragment for `packages/xlvane/src/sml/xstring.ts`, `shared-strings.ts`, `styles.ts`.

## Conventions

- Sources are in `scratchpad/specs/`. `part1.txt` is ECMA-376 Part 1, 5th ed. (2016), Strict; `part4.txt` is Part 4 (Transitional); `xsd-strict/` and `xsd-trans/` are the schemas; `../text/xml10.txt` is XML 1.0 (5th ed.).
- Locators: Part, clause (§) and title, printed page ("p.", from the page footer), and `Lnnn`, the line in the `.txt` file (`sed -n 'NNN,+15p' part1.txt` to check).
- Quoting: verbatim prose quotes are kept to a minimum (one, in F-STY-06, at most 15 words). Everywhere else prose is paraphrased closely, with exact line locators; technical tokens (element, attribute and type names, `_xHHHH_`, `_x005F_`, code-point ranges) are verbatim.
- "Decision" marks a choice made where the spec is silent or ambiguous. These are not spec rules; they are listed for the maintainer's review.
- Every fact was checked against the spec text by me; it agrees with `facts-ecma.md` E-29 … E-34, E-40, E-47 where they overlap.

## Shared String Table

### F-SST-01 - Root element of the Shared String Table part
- **Fact:** The root element of the part is `sst`. Strict namespace `http://purl.oclc.org/ooxml/spreadsheetml/main`; Transitional `http://schemas.openxmlformats.org/spreadsheetml/2006/main`.
- **Source:** Part 1 §12.3.15 Shared String Table Part, p.89, L4278 (root element sentence). Part 4 §10.2.15 Shared Strings Table Part, p.20, L1422 (Transitional root namespace).
- **Code:** `parseSharedStrings` throws `XLV_E_PART_INVALID` for any other root name or namespace.

### F-SST-02 - Content model (schema)
- **Fact:** `CT_Sst` = sequence `si*` (type `CT_Rst`), `extLst?`; optional attributes `count`, `uniqueCount` (`xsd:unsignedInt`). `CT_Rst` = sequence `t?` (`ST_Xstring`), `r*` (`CT_RElt`), `rPh*` (`CT_PhoneticRun`), `phoneticPr?`. `CT_RElt` = `rPr?`, `t` (`ST_Xstring`, exactly one). `CT_PhoneticRun` = `t` (`ST_Xstring`), required `sb`, `eb`.
- **Source:** `xsd-strict/sml.xsd` L1776–1840 (CT_Sst L1777, CT_PhoneticRun L1801, CT_RElt L1808, CT_Rst L1833). `xsd-trans/sml.xsd` L1788–1852: byte-identical block (checked with `diff`).
- **Note:** `t` is declared with a simple type, so the schema itself does not allow child elements or attributes on it (the prose in §18.4.12 still lists `xml:space`, F-SST-05).

### F-SST-03 - `si` (String Item)
- **Fact:** An `si` is one string of the table. A string with only cell-level formatting should be a single text element; a string with character-level formatting shall be made of several rich text runs that together express the string.
- **Source:** Part 1 §18.4.8 si (String Item), p.1724, L81475–81482.
- **Code:** the value of an `si` is its own `t` followed by the `t` of each `r`, concatenated in document order. Index `i` in the returned array is `si` number `i` (zero-based, see `facts-ecma.md` E-30).

### F-SST-04 - `r` (Rich Text Run)
- **Fact:** A run is a region of text sharing one set of properties; the properties are in `rPr` and the text shown to the user is in the run's `t`.
- **Source:** Part 1 §18.4.4 r (Rich Text Run), p.1722, L81397–81401.
- **Code:** text inside `rPr` and any other element of a run is ignored.

### F-SST-05 - `t` (Text): ST_Xstring and `xml:space`
- **Fact:** `t` is the text content shown as part of a string; its values are `ST_Xstring` (§22.9.2.19). Its attribute `xml:space` (namespace `http://www.w3.org/XML/1998/namespace`) says how white space is handled, with values as defined in XML 1.0 §2.10.
- **Source:** Part 1 §18.4.12 t (Text), p.1725, L81552–81566.
- **Code:** `decodeXstring` is applied to the complete character data of each `t` separately (after XML references, CDATA and comments are resolved by the XML parser). An escape split across two runs is therefore not decoded.

### F-SST-06 - `rPh` (Phonetic Run)
- **Fact:** `rPh` is a run of text that gives a phonetic (pronunciation) hint for East Asian languages for its String Item; the hints are displayed across the top of the cell. `sb`/`eb` are zero-based offsets into the *base text* (the value must lie between 0 and the base text length; `sb < eb`).
- **Source:** Part 1 §18.4.6 rPh (Phonetic Run), p.1723, L81426–81462; §18.4.3 phoneticPr, p.1721–1722, L81342–81395.
- **Not found:** no sentence says whether the `rPh/t` text belongs to the string's value (also searched in `facts-ecma.md` E-31).
- **Decision D-SST-2:** `rPh/t` is excluded from the value. Reasons: §18.4.8 builds the string from the text element or the rich text runs only; §18.4.6 describes `rPh` as a hint *about* a base text, and its `sb`/`eb` offsets point into that base text, which only makes sense if the hint is not part of it.

### F-SST-07 - `sst@count`, `sst@uniqueCount`
- **Fact:** `count` is the total number of text strings in the workbook (numbers are not counted); `uniqueCount` is the number of unique entries in the table. Each is optional unless the other is present.
- **Source:** Part 1 §18.4.9 sst (Shared String Table), p.1724, L81486–81518.
- **Decision D-SST-3:** both are informational only. They are never used to size or pre-allocate anything (untrusted input), and a mismatch with the real number of `si` elements is not reported (it would be noise; the number of `si` elements is authoritative for cell indexes).

### F-SST-08 - `r` and `rPr` content model (rich text runs)
- **Fact:** CT_RElt is the sequence `rPr` (CT_RPrElt, 0..1) then `t` (ST_Xstring, exactly 1). CT_RPrElt is an unbounded `xsd:choice` of `rFont` (CT_FontName), `charset`, `family` (CT_IntProperty), `b`, `i`, `strike`, `outline`, `shadow`, `condense`, `extend` (CT_BooleanProperty), `color` (CT_Color), `sz` (CT_FontSize), `u` (CT_UnderlineProperty), `vertAlign` (CT_VerticalAlignFontProperty), `scheme` (CT_FontScheme), each 0..1. CT_Rst is the sequence `t` 0..1, `r` 0..n, `rPh` 0..n, `phoneticPr` 0..1. The font name of a run is `rFont`, where a `font` record uses `name`.
- **Source:** Transitional `spec/ECMA-376-4_5th_edition_december_2016/xsd/sml.xsd` CT_RElt L1820–1825, CT_RPrElt L1826–1844, CT_Rst L1845–1852; Strict `spec/ECMA-376-1_supporting/OfficeOpenXML-XMLSchema-Strict/sml.xsd` CT_RElt L1808, CT_RPrElt L1814, CT_Rst L1833 (same content). The schema allows the children of `rPr` in any order. [MS-OI29500] rev. 25.0 §2.1.674 a, p.289 (local review report, D-1104): Excel requires the order `b, i, strike, condense, extend, outline, shadow, u, vertAlign, sz, color, rFont, family, charset, scheme`.
- **Code:** the reader reads `rPr` with the font parser of the styles part (`parseRunFont`, CT_BooleanProperty `val` default true as in W2-11); a leading `t` of an item with runs is returned as a run without font; the writer writes only the properties given, in the Office order. `vertAlign`, `family`, `charset`, `scheme`, `outline`, `shadow`, `condense` and `extend` are not modelled (left out on read, not writable).
- **Update (D-E01):** `vertAlign`, `family`, `charset`, `scheme`, `outline`, `shadow`, `condense` and `extend` are now read and written in runs too, with the values and limits of W2-64 (the run's `family` is xsd:int; xlvane applies the Office range 0 to 5 to new runs and reads 0 to 14).

### F-SST-09 - Limits of a string item
- **Fact:** Office limits `t` to 32767 characters and a string item to at most 32767 `r` and 32767 `rPh`.
- **Source:** [MS-OI29500] rev. 25.0 §2.1.675 a–c (§18.4.8 si), p.289–290; §2.1.678 b (§18.4.12 t), p.290 (local review report, D-1104; E-30).
- **Code:** the writer rejects rich text with more than 32767 runs or more than 32767 characters in all (`XLV_E_INVALID_VALUE`), the same limit as for plain text.

### Decision D-SST-1 - White space in `t`
- **Fact (XML 1.0 §2.10 White Space Handling, `text/xml10.txt` L1029–1068):** an XML processor must pass all non-markup characters to the application; `xml:space="preserve"` signals the intent that white space be preserved, `xml:space="default"` signals that the application's default handling is acceptable.
- **Decision:** white space in `t` is kept exactly as written regardless of `xml:space` (absent, `preserve` or `default`). Preserving is always within what the attribute allows, and it is the only choice that does not change cell values. (The XML parser already normalises line ends per XML 1.0 §2.11, so CR LF in a file becomes LF; a real CR is written as `_x000D_` or `&#13;`.)

### Decision D-SST-4 - Elements the reader ignores
- Elements of other namespaces, unknown elements, `phoneticPr`, `extLst`, `si` elements nested where the schema does not allow them, and character data that is not directly inside a `t` (for example indentation between elements) are ignored. Child elements inside `t` (not allowed by the schema) are skipped, but the character data around them is kept.

## ST_Xstring

### F-XS-01 - The `_xHHHH_` escape
- **Fact:** `ST_Xstring` ("Escaped String") is a restriction of `xsd:string` in which every character that XML 1.0 cannot represent is escaped in the form `_xHHHH_`, each H being a hexadecimal digit of the character's value. The spec's example: U+0008 is not allowed in XML 1.0, so it is written `_x0008_`.
- **Source:** Part 1 §22.9.2.19 ST_Xstring (Escaped String), p.3801, L178057–178066. Schema: `xsd-strict/shared-commonSimpleTypes.xsd` L67, `xsd-trans/shared-commonSimpleTypes.xsd` L114 (both plain `xsd:string`, no pattern).
- **Not found:** §22.9.2.19 says nothing about how a literal `_xHHHH_` text is written; Part 4 has no text on ST_Xstring.

### F-XS-02 - Literal escape sequences (`_x005F_`), stated for `vt:bstr`
- **Fact:** For the `bstr` variant type the same `_xHHHH_` escape is used, and to store the literal form of an escape sequence its initial underscore is itself escaped, i.e. stored as `_x005F_`. Example given: the literal text `_x0008_` is stored as `_x005F_x0008_`.
- **Source:** Part 1 §22.4.2.4 bstr (Basic String), p.3736, L174882–174889.
- **Scope:** this sentence is in the `bstr` clause, not in §22.9.2.19.
- **Code:** no special rule is needed. A single left-to-right pass decodes `_x005F_` to `_` and continues *after* the escape, so `_x005F_x0008_` yields the text `_x0008_`. This is the same result whether or not the bstr sentence applies to ST_Xstring, because `_x005F_` is also an ordinary `_xHHHH_` escape of U+005F.

### F-XS-03 - Hex digit case in the spec's own examples
- **Fact:** Spec examples of ST_Xstring-typed values use lower-case hex: `_x000a_` in `scenario@comment` and `_x000d__x000a_` in `dbPr@command`, while §22.9.2.19 uses `_x0008_`.
- **Source:** Part 1 §18.3.1.76 scenarios, p.1682, L79381; §18.13.3 dbPr, p.2001, L94291–94292.
- **Code:** hex digits are accepted in upper and lower case. The `x` must be lower case (`_X0041_` is text), and exactly four hex digits followed by `_` are required.

### F-XML-CHAR - XML 1.0 `Char`
- **Fact:** `Char ::= #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]`; the production's comment says it is any Unicode character except the surrogate blocks, FFFE and FFFF.
- **Source:** XML 1.0 (5th ed.) §2.2 Characters, production [2], `text/xml10.txt` L446–465.

### Decisions for ST_Xstring decoding
- **D-XS-1 (every `_xHHHH_` is decoded):** an escape is decoded even when the character could have been written directly (for example `_x0041_` → `A`, `_x000a_` → LF). Reason: the spec's own examples escape representable characters (F-XS-03), and `_x005F_` (F-XS-02) is an escape of a representable character.
- **D-XS-2 (characters outside `Char` are produced):** `_x0000_`, `_x0001_` … `_x001F_` (except tab, LF, CR, which are allowed anyway), `_xFFFE_` and `_xFFFF_` decode to those code points; this is exactly what the escape exists for (F-XS-01, F-XML-CHAR).
- **D-XS-3 (lone surrogates stay text):** `_xD800_` … `_xDFFF_` on their own are *not* decoded and remain as the seven characters of text. Reason: the escape encodes "a character", and D800–DFFF are not characters (they are excluded from `Char`, F-XML-CHAR, and are not Unicode scalar values). Leaving them as text is lossless and keeps every string well-formed UTF-16, so downstream code such as `encodeURIComponent` (which throws on lone surrogates) or `TextEncoder` (which replaces them) behaves predictably. The alternatives were U+FFFD (loses information) or a lone surrogate (ill-formed string).
- **D-XS-4 (escaped surrogate pairs are decoded):** an escaped high surrogate immediately followed by an escaped low surrogate (`_xD83D__xDE00_`) decodes to the one supplementary character they encode. The spec does not mention this (supplementary characters are valid XML and need no escape); it is accepted because the result is a valid character and the input is unambiguous. **Review:** if the maintainer prefers strictness, drop this and treat both halves as text (D-XS-3).
- **D-XS-5 (no overlap after a non-escape):** when a candidate `_x…` is not an escape, scanning resumes at the next character, so an underscore that ends a non-escape can start a real one (`_xZZZZ_x0041_` → `_xZZZZA`). After a real escape, scanning resumes after its closing underscore (`_x005F_x0041_` → `_x0041_`).
- **Complexity:** one pass with `indexOf`; each candidate inspects at most 14 characters, so the decoder is linear.

## Styles part (number formats only)

### F-STY-01 - Root element of the Styles part
- **Fact:** The root element of the Styles part is `styleSheet`; a package has at most one Styles part.
- **Source:** Part 1 §12.3.20 Styles Part, p.94–95, L4553 (cardinality), L4570 (root element sentence).
- **Code:** `parseStyles` throws `XLV_E_PART_INVALID` for another root name or a namespace other than the Strict or Transitional SpreadsheetML namespace. Workbooks without a Styles part use `EMPTY_STYLES`.

### F-STY-02 - `CT_Stylesheet` sequence
- **Fact:** `CT_Stylesheet` is a sequence of optional elements, each `maxOccurs="1"`: `numFmts`, `fonts`, `fills`, `borders`, `cellStyleXfs`, `cellXfs`, `cellStyles`, `dxfs`, `tableStyles`, `colors`, `extLst`.
- **Source:** `xsd-strict/sml.xsd` L3359–3373; `xsd-trans/sml.xsd` L3386–3400 (same sequence).
- **Decision D-STY-5:** element order is not enforced; if `numFmts` or `cellXfs` occurs more than once, the first is used and `XLV_W_INVALID_VALUE` is reported (`cellStyleXfs` is not read at all, D-STY-2). Elements in other namespaces are ignored.

### F-STY-03 - `numFmt` attributes
- **Fact:** `numFmt@numFmtId` (`ST_NumFmtId`) is the id by which the master style records (`xf`) reference the number format; `numFmt@formatCode` (`ST_Xstring`) is the format code. Both are `use="required"`. `numFmts` holds `numFmt*` and an optional `count`.
- **Source:** Part 1 §18.8.30 numFmt (Number Format), attribute table p.1783, L84300–84310; `xsd-strict/sml.xsd` CT_NumFmts L3547, CT_NumFmt L3553–3556. Transitional identical (`diff` of `xsd-trans/sml.xsd` L3574–3619).
- **Code:** ids are matched by value, not by position, and kept in a `Map<number, string>`. `formatCode` is passed through `decodeXstring`. An empty `formatCode` is kept as `""`.
- **Decision D-STY-3:** a `numFmt` with a missing or invalid `numFmtId`, or a missing `formatCode`, is ignored with `XLV_W_INVALID_VALUE`. If an id is declared twice, the first declaration is kept and `XLV_W_INVALID_VALUE` is reported (the spec calls it an Id but states no uniqueness rule for this case).

### F-STY-04 - Implied (built-in) formats
- **Fact:** For the listed built-in ids, the format code is implied: a `numFmtId` is written on the `xf` without a matching `numFmt` element. Use of ids outside the lists is implementation-defined; ids not in the listing (for example 5–8) follow the `formatCode` of a declared `numFmt`.
- **Source:** Part 1 §18.8.30 numFmt, p.1776, L83957–83975.
- **Code:** `customFormatCode(id)` returns only what `numFmts` declares; the built-in table and date detection are in `sml/numfmt.ts` (another agent).

### F-STY-05 - `cellXfs` and `cellStyleXfs`
- **Fact:** `cellXfs` holds the master formatting records for cells; they are the starting point for a cell's formatting, and cells in the Sheet part reference them by zero-based index. A cell can have both direct formatting and a cell style, so both the cell-style `xf` records (`cellStyleXfs`) and the cell `xf` records shall be read to know the full formatting. Both elements have an optional `count` attribute giving the number of `xf` elements (§18.8.10 L83354; §18.8.9 L83332).
- **Source:** Part 1 §18.8.10 cellXfs (Cell Formats), p.1763, L83343–83360; §18.8.9 cellStyleXfs (Formatting Records), p.1762, L83301–83311. Informative: Annex L.2.7.3.6, p.4679, L250890–250897 (an `xf` in `cellXfs` is referenced by zero-based index, the `s` attribute of `c`).
- **Code:** `cellFormatCount` is the number of `xf` elements found, not the `count` attribute. `numFmtId(i)` returns 0 (General) for any `i` that is not an integer in `[0, cellFormatCount)`. Only the cell xf records are read; the cell style records are not, following Office (F-STY-09, D-STY-2).

### F-STY-06 - `xf` attributes that decide the number format
- **Fact (§18.8.45 xf (Format)):**
  - `numFmtId`: the id of the number format (`numFmt`) record used by this cell format (`ST_NumFmtId`). p.1800, L85138–85146.
  - `applyNumberFormat`: a boolean indicating "whether the number formatting specified for this xf should be applied." p.1800, L85114–85117. (This is the one verbatim quote of this fragment.)
  - `xfId`: for `xf` records in `cellXfs`, the zero-based index of the `cellStyleXfs` record of the cell style applied to the cell; not present on records in `cellStyleXfs`. p.1801, L85160–85169.
- **Schema:** every `CT_Xf` attribute is optional; `applyNumberFormat` (`xsd:boolean`) has **no default**, nor do `numFmtId` and `xfId`. `xsd-strict/sml.xsd` CT_Xf L3569–3588 (`numFmtId` L3575, `applyNumberFormat` L3582); `xsd-trans/sml.xsd` CT_Xf L3598 (identical, checked with `diff`).
- **Part 4:** no Transitional text about `applyNumberFormat` beyond the schema (searched `part4.txt`; only the xsd listing lines 57017 and 93819 match).

### F-STY-07 - `applyNumberFormat="0"` on a cell style record
- **Fact:** The §18.8.9 example explains that a `cellStyleXfs` record carrying `applyNumberFormat="0"` (and the other apply attributes set to 0) means the number format, alignment and protection formatting are not applied to the cell, even when related formatting records specify them.
- **Source:** Part 1 §18.8.9 cellStyleXfs, example, p.1762, L83312–83327.

### F-STY-08 - `ST_NumFmtId`
- **Fact:** `ST_NumFmtId` identifies a number format entry in `CT_NumFmts`; it is a restriction of `xsd:unsignedInt` with no further facets.
- **Source:** Part 1 §18.18.47 ST_NumFmtId (Number Format Id), p.2467, L117452–117459; `xsd-strict/sml.xsd` ST_NumFmtId L3623–3625.
- **Not verified here:** the XML Schema Part 2 lexical rules of `xsd:unsignedInt` and `xsd:boolean` (range 0…4294967295, digits, white-space collapsing, `true/false/1/0`) are not in the local spec set; they are used as commonly specified and should be checked against XML Schema Part 2 (§3.3.22 unsignedInt, §3.2.2 boolean) when that document is added.

### F-STY-09 - Office reads only the cell xf record ([MS-OI29500])
- **Fact:** [MS-OI29500] (Microsoft's notes on how Office implements ECMA-376) has a page titled "Part 1 Section 18.8.10, cellXfs (Cell Formats)" (the coordinator gives its clause number as §2.1.700; the page I fetched shows only the title). In **note b**, Microsoft restates the standard's requirement that both the cell style xf records and the cell xf records be read to get a cell's full formatting, then says that in Office only the cell xf record defines the formatting applied to the cell. (On the same page, notes a and c give Excel's limit of 65430 `xf` records.)
- **Source:** https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/d0f88a45-c7c3-456b-9e21-2fd5577e9898, note b; page updated 2024-04-16; fetched and checked by me on 2026-10-30.
- **Spec/Office difference:** ECMA-376 §18.8.10 / §18.8.9 (F-STY-05) say both layers *shall* be read, and §18.8.45 (F-STY-06) defines `applyNumberFormat`. Office does not combine the two layers: the cell xf record alone decides.
**Checked 2026-10-30 (D-1104):** [MS-OI29500] rev. 25.0 (2026-08-18) §2.1.700 b (P1 §18.8.10 cellXfs), p.298, and §2.1.699 b, p.297: only the cell xf record defines the formatting applied to a cell; a, c: at most 65430 xf records, cellXfs plus non-built-in cellStyleXfs ≤ 65430. Confirmed from the PDF (replaces the web-page citation).

### Decision D-STY-1 - Validating ids
- `numFmt@numFmtId` and `cellXfs/xf@numFmtId` are parsed as `xsd:unsignedInt`: optional surrounding white space, ASCII digits (leading zeros allowed), value at most 4294967295. Anything else is reported as `XLV_W_INVALID_VALUE`: negative, fractional, exponent, hex, `+`, empty, names such as `__proto__`, values above the range. On a `numFmt` the entry is ignored; on a cell xf the result is General (0). Parsing is linear and never throws. A leading `+` is rejected (see the F-STY-08 note: to be checked against XML Schema Part 2).
- `xfId`, `applyNumberFormat` and `cellStyleXfs` are not read (D-STY-2), so their values are not validated and never warn.

### Decision D-STY-2 - Follow Office: the cell xf record only
- **Rule:** `numFmtId(s)` is the `numFmtId` of `cellXfs/xf[s]`. It is 0 (General) when that attribute is absent or invalid, or when `s` is out of range. `applyNumberFormat`, `xfId` and `cellStyleXfs` are ignored.
- **Reasoning:** ECMA-376 gives no default for `applyNumberFormat` and no rule for combining the cell and style records (F-STY-05, F-STY-06, F-STY-07; also NOT FOUND in `facts-ecma.md` E-47). Office documents that it uses only the cell xf record (F-STY-09), so following Office is the interoperable choice. For example, `numFmtId="14" applyNumberFormat="0"` gives 14 (a date format). That is what note b implies; I have not checked it in Excel itself.
- **Difference from the standard:** a strictly ECMA-376 reader would also read `cellStyleXfs[xfId]` and honour `applyNumberFormat`. xlvane does not, on purpose.
- **History:** the first version of this module used a layered rule (cell xf unless `applyNumberFormat` is explicitly false, then the cell style record). It was replaced on the coordinator's request after F-STY-09 was found.
- The result is computed once per cell format at parse time, so `numFmtId(i)` is O(1).

### Decision D-STY-4 - (removed)
- No longer applies: `xfId` is not read on any record (D-STY-2).

### Decision D-STY-6 - Out-of-range style index
- `numFmtId(i)` returns 0 for an index that is negative, fractional, NaN, infinite or not less than `cellFormatCount`; it does not warn (it has no context). Reporting a cell whose `s` is out of range is left to the worksheet parser, which knows the cell reference.

## Open items for other modules

- `sml/worksheet.ts` decodes inline strings (`is`, also `CT_Rst`, F-SST-02) per text event. For consistency with this fragment, it should decode per `t` element, concatenate `t` + `r/t` and skip `rPh/t` (D-SST-2). Not my file; reported to the coordinator.
