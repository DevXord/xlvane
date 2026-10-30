# Facts used by `opc/package.ts` and `opc/part-name.ts` (F-OPC-*)

> **Provenance.** Compiled on 2026-10-30. Line numbers (`Lnnn`) refer to text extracted with `pdftotext -layout` from the ECMA-376 PDFs (see docs/facts/ecma-376.md); `scratchpad/specs/…` paths are the working copies used during the check and are not stored in this repository.

Sources:
- `scratchpad/specs/part2.txt`: ECMA-376 Part 2, 5th ed. (Dec 2021), OPC. `Lnnn` is the line number in that file.
- `scratchpad/specs/part1.txt`: ECMA-376 Part 1, 5th ed. (Dec 2016).
- `scratchpad/specs/opc-xsd/opc-contentTypes.xsd` and `opc-relationships.xsd`.
- RFC 3986 and RFC 3987, read on 2026-10-30 from https://www.rfc-editor.org/rfc/rfc3986.txt and https://www.rfc-editor.org/rfc/rfc3987.txt. Part 2 §6.2.2.2 and §6.4.1 delegate to them.
- Cross-checked against `scratchpad/facts-ecma.md` E-48 … E-59. Each fact below was re-read in `part2.txt` at the given lines.

Requirement ids: the 2021 edition of Part 2 has no `[Mx.y]` ids. A search of `part2.txt` for `\[M[0-9]` finds nothing (see also facts-ecma.md E-55). Facts are cited by clause number.

Quotes: verbatim prose quotes are kept to a minimum (one, in F-OPC-07, at most 15 words). Everything else is a close paraphrase with exact line numbers. Technical tokens (names, URIs, enumeration values, grammar rules) are verbatim.

## Part names

### F-OPC-01: Part name grammar
- **Fact:** a part name is a Unicode string matching `part_name = 1*( "/" isegment-nz )`, with `isegment-nz` from RFC 3987 §2.2. Extra constraints:
  - (a) no segment contains a percent-encoded `/` or `\`;
  - (b) no segment contains a percent-encoded character that matches `iunreserved`;
  - (c) no segment ends with `.`.
- **Source:** Part 2 §6.2.2.2 Syntax, L583–595.
- **RFC 3987 §2.2:**
  - `isegment-nz = 1*ipchar`;
  - `ipchar = iunreserved / pct-encoded / sub-delims / ":" / "@"`;
  - `iunreserved = ALPHA / DIGIT / "-" / "." / "_" / "~" / ucschar`;
  - `ucschar = %xA0-D7FF / %xF900-FDCF / %xFDF0-FFEF / %x10000-1FFFD / … / %xD0000-DFFFD / %xE1000-EFFFD` (each plane N from 1 to D contributes `%xN0000-NFFFD`).
- **RFC 3986:**
  - §2.1: `pct-encoded = "%" HEXDIG HEXDIG`;
  - §2.2: `sub-delims = "!" / "$" / "&" / "'" / "(" / ")" / "*" / "+" / "," / ";" / "="`;
  - §2.3: `unreserved = ALPHA / DIGIT / "-" / "." / "_" / "~"`.
- **Code:** `isPartName()` in `opc/part-name.ts`.

### F-OPC-02: Reserved Relationships part names
- **Fact:** `/_rels/.rels` is reserved. So is any part name whose second-to-last segment is equivalent to `_rels` and whose last segment ends with `.rels`.
- **Source:** Part 2 §6.2.2.2, L600–602.
- **Code:** `isRelationshipsPart()`.

### F-OPC-03: Part name equivalence
- **Fact:** equivalence uses ASCII case-insensitive matching: only A–Z (0x41–0x5A) are folded to a–z. The names of two different parts shall not be equivalent.
- **Source:** Part 2 §6.2.2.3, L618–622; definition 3.3.11, L494.
- **Code:** `partKey()` folds A–Z only. Non-ASCII letters are compared exactly.
- **Derivability:** a part name must not be derivable from another one (`/a` and `/a/b`), L633. This is not checked; see D-OPC-07.

## XML in OPC parts

### F-OPC-04: XML usage in OPC's own parts
- **Fact:** these rules apply to the Media Types stream, Core Properties, Digital Signature XML and Relationships parts:
  - (a) UTF-8 or UTF-16 only, and any encoding declaration names only one of them (L702);
  - (b) DTD declarations shall not be used, as a denial-of-service mitigation (L706–709);
  - (c) the XML conforms to XML Namespaces;
  - (d) the XML is schema-valid against Annex C;
  - (e) no `xml`/`xsi` items unless allowed.
- **Source:** Part 2 §6.2.5 XML usage, L693–720.
- **Code:** (a) and (b) are enforced by the XML layer: `decodeXml` throws `XLV_E_XML_ENCODING` and `parseXml` throws `XLV_E_XML_DTD_FORBIDDEN`. `opc/package.ts` only routes both parts through those functions. (d) is applied partially; see D-OPC-04.

## Pack IRIs and reference resolution

### F-OPC-05: Query and fragment of a pack IRI
- **Fact:** a query component in a pack IRI is not used when resolving it to a part. A fragment applies to whatever resource the pack IRI identifies.
- **Source:** Part 2 §6.3.2, L777–781.
- **Part 1 §9.1.2 (L1066–1068):** fragments are supported in all external targets and in some internal ones.
- **Code:** `resolveTarget` drops `?…` and `#…` before resolving.

### F-OPC-06: Resolving relative references
- **Fact:** relative references are resolved by RFC 3986 §5 (Reference Resolution), as extended by RFC 3987 §6.5. Part 2 changes nothing in that procedure.
- **Source:** Part 2 §6.4.1, L903–907.
- **Base for a non-Relationships part:** the pack IRI of the package plus the part's own name (§6.4.2, L922–923).
- **Worked examples (§6.4.3, L955–1036).** From base `/a/b/foo.xml`:
  - `/b/bar.xml` → `/b/bar.xml`;
  - `bar.xml` → `/a/b/bar.xml`;
  - `./bar.xml` → `/a/b/bar.xml`;
  - `../bar.xml` → `/a/bar.xml`.
- **From base `/`:** `bar.xml`, `./bar.xml` and `../bar.xml` all give `/bar.xml`. remove_dot_segments replaces `/../` with `/` (L1024).
- **RFC 3986 §5.2.2–§5.2.4:**
  - A reference with a scheme is used as is.
  - A reference with an authority (`//…`) keeps its own path.
  - A path that starts with `/` ignores the base path.
  - Any other path is merged: the base path up to and including its last `/`, then the reference.
  - remove_dot_segments then applies steps A–E. Step C removes the last output segment for `/../`, and does nothing extra at the root.
  - An empty reference path takes the base path.
- **RFC 3986 §4.2:** a relative-path reference cannot have a `:` in its first segment, because that would be read as a scheme. The ABNF is `scheme = ALPHA *( ALPHA / DIGIT / "+" / "-" / "." )` (§3.1).
- **Code:** `resolveTarget()`.

## Relationships

### F-OPC-07: Relationships parts; target modes
- **Media type:** `application/vnd.openxmlformats-package.relationships+xml`. **Root namespace:** `http://schemas.openxmlformats.org/package/2006/relationships` (§6.5.2.1, L1077–1081).
- **No chaining:** there shall be no relationships from or to a Relationships part (L1083–1084).
- **Package relationships:** the package Relationships part is named `/_rels/.rels` (§6.5.2.2, L1095). For an internal target, the base is the package root `/` (L1097–1099).
- **Part relationships:** the name is built by adding `.rels` to the last segment of the source part and inserting a `_rels` segment before it. Examples: `/foo` → `/_rels/foo.rels` and `/foo/bar.xml` → `/foo/_rels/bar.xml.rels` (§6.5.2.3, L1120–1127). For an internal target, the base is the source part name (L1130–1133).
- **External targets:** the base is the absolute IRI of the package itself (L1097, L1130). A reader that has no package IRI cannot resolve them, and must not try to.
- **Code:** `relationshipsPartName()`, `relationships()`, `RELATIONSHIPS_CONTENT_TYPE`.

### F-OPC-08: Relationship markup
- **Root:** `Relationships`. It has no attributes and contains zero or more `Relationship` elements (§6.5.3.3, L1168–1172). The schema CT_Relationships is a sequence of `Relationship`, minOccurs 0.
- **Attributes of `Relationship`** (§6.5.3.4, L1174–1228; `opc-relationships.xsd` CT_Relationship):
  - `TargetMode` is optional and defaults to `Internal`. ST_TargetMode is an enumeration: `External`, `Internal`.
  - `Target` is required, `xsd:anyURI`. Internal: shall be a relative reference to a part. External: a relative reference or an absolute IRI.
  - `Type` is required, `xsd:anyURI`. Types compare as strings: case-sensitive, with no escaping done or undone.
  - `Id` is required, `xsd:ID`.
- **The one verbatim quote (Part 2 §6.5.3.4, L1219–1220):** "The value of the Id attribute shall be unique within the Relationships part."
- **MCE and `xml:base`:** Relationships parts may use MCE (ECMA-376-3) extensions, and the result of MCE processing must not contain `xml:base` (§6.5.3.1–§6.5.3.2, L1149–1162).
- **Code:**
  - A duplicate `Id` → `XLV_E_PART_INVALID`.
  - A missing required attribute or an unknown `TargetMode` → `XLV_E_PART_INVALID`.
  - `Type` is kept verbatim.
  - MCE: see D-OPC-04.

## Media Types stream (`[Content_Types].xml`)

### F-OPC-09: The stream is mandatory in ZIP packages and is not a part
- **Fact:** a physical format without a native media-type mechanism shall include an XML stream, the Media Types stream. ZIP has no such mechanism, and Table 1 maps part media types to that stream. The stream shall not represent a part and is not URI-addressable.
- **Source:** Part 2 §7.2.3.1, L1446–1455; §7.3.1 Table 1, L1758–1772.
- **ZIP item name:** `[Content_Types].xml` (§7.3.7, L1840–1846). The brackets were chosen because they break the part-name grammar.
- **Annex H (L3923–3925)** shows the item next to items that are parts.
- **Code:** if no ZIP item `[Content_Types].xml` exists (ASCII case-insensitive, see F-OPC-13), `openPackage` throws `XLV_E_PART_MISSING`. The name never becomes a part, because `[` and `]` are not `ipchar`.

### F-OPC-10: Media Types stream markup
- **Root:** the root element is `Types` (§7.2.3.2.1, L1463–1465; §7.2.3.2.3, L1486–1488). It has no attributes.
- **Namespace:** `http://schemas.openxmlformats.org/package/2006/content-types` (the `opc-contentTypes.xsd` targetNamespace).
- **Coverage:** every part except Relationships parts shall have a matching `Default`, a matching `Override`, or both. When both match, `Override` wins (L1468–1474).
- **Uniqueness:** at most one `Default` per extension, and at most one `Override` per part name (L1476–1477).
- **Order:** element order is not significant. Unused `Default`s are allowed (L1478–1479).
- **No MCE:** the stream shall not use MCE (§7.2.3.2.2, L1483).
- **`Default`:** `Extension` and `ContentType` are both required. It matches any part whose name ends with `.` followed by `Extension` (§7.2.3.2.4, L1500–1515).
  - The ST_Extension pattern has no `.`, so an extension never contains a dot.
- **`Override`:** `ContentType` and `PartName` (`xs:anyURI`) are both required. It matches the part whose name equals `PartName` (§7.2.3.2.5, L1527–1546).
- **Schema:** CT_Types is a choice of `Default` and `Override`, minOccurs 0. The prose of §7.2.3.2.1 says "one or more"; the code accepts zero.
- **Code:** `readContentTypes()`.
  - A wrong root → `XLV_E_PART_INVALID`.
  - A missing attribute → `XLV_E_PART_INVALID`.
  - A duplicate extension or part name, compared ASCII case-insensitively → `XLV_E_PART_INVALID`.

### F-OPC-11: Looking up a part's media type
- **Fact (§7.2.3.5, L1610–1632):**
  - (a) Compare the part name with `Override@PartName` using ASCII case-insensitive matching. On a match, return its `ContentType`.
  - (c) Otherwise, the extension is the substring to the right of the rightmost `.` in the rightmost segment. Compare it with `Default@Extension`, again ASCII case-insensitively.
  - (d) On a match, return its `ContentType`.
  - A part name without an extension needs an `Override` (§7.2.3.4 b, L1585–1586).
- **Code:** `contentType()`. For a Relationships part with no entry in the stream, it returns the fixed media type from F-OPC-07, because the stream is not required to cover those parts (F-OPC-10).

## Mapping ZIP items to parts

### F-OPC-12: ZIP item names to part names
- **ZIP item names** shall follow the ZIP Appnote, shall be ASCII only, and shall be unique (§7.3.3, L1779–1781).
- **Part → ZIP:** (a) remove the leading `/`; (b) percent-encode every non-ASCII character (§7.3.4, L1795–1800).
- **ZIP → part:** all ZIP items are mapped to logical item names, except items that do not represent files. The steps are (a) un-percent-encode every non-ASCII character, then (b) add a `/` (§7.3.5, L1806–1814).
- **Folders:** Annex B Table B.3 lists version 2.0 "File is a folder (directory)" as consumption Yes, with the instruction not to interpret it as a part (L3228–3236).
- **Code:** `openPackage` skips names that end with `/`. It decodes percent-encoded UTF-8 sequences that form non-ASCII characters, then adds `/`.

### F-OPC-13: Logical item names
- **Equivalence:** logical item names compare with ASCII case-insensitive matching. A physical package shall not contain equivalent logical item names (§7.2.5.2, L1700–1702).
- **Mapping to parts:** a logical item name is mapped to a part only if it follows the part-name grammar. Other items, for example `/[trash]/0000.dat`, are not parts (§7.2.5.5, L1726–1739).
- **Interleaving:** piece names (`…/[0].piece`) contain brackets, so they are never part names.
- **Part 1:**
  - §9.1.8 (L1123–1125): Office Open XML documents shall not use interleaving.
  - §9.1.5–§9.1.6: trash items `[trash]/HHHH.dat` are allowed. Other items that are not valid part names are not allowed in a document.
  - §9.1.4: unknown parts shall be ignored on reading.
- **Code:**
  - Items whose logical name is not a part name are left out of `partNames` silently.
  - Two items with equivalent logical names → `XLV_E_PART_INVALID`.
  - When only interleaved pieces of `[Content_Types].xml` exist, the missing-stream error says so.

### F-OPC-14: Annex A preprocessing (informative)
- **Fact:** some implementations preprocess near-references before resolving them. The listed steps include:
  - un-percent-encode each percent-encoded unreserved character;
  - percent-encode spaces and brackets;
  - convert `\` to `/`.
- **Examples:** `/%41/%61.xml` → `/A/a.xml`; `/%2e/%2e/a.xml` → `/a.xml`; `/%D1%86.xml` stays as written. The last example shows that "unreserved" in Annex A means the ASCII `unreserved` of RFC 3986: Annex A produces URI references, not IRIs.
- **Source:** Part 2 Annex A (informative), L2900–2945.
- **Code:** only the first step is applied, to `Target` and `Override@PartName` (D-OPC-02). Backslashes and spaces are not repaired.

### F-OPC-15: Percent-encoding of non-ASCII characters is UTF-8
- **Fact:** RFC 3987 §3.1 step 2 converts each non-ASCII character to UTF-8 octets, and writes each octet as `%HH`. §3.2 (URI → IRI) converts percent-encodings back to characters, except `%`, characters in `reserved`, and US-ASCII characters that URIs do not allow.
- **Why it matters:** Part 2 §7.3.4 and §7.3.5 percent-encode and un-percent-encode "every non-ASCII character" without naming an encoding. Part 2 bases non-ASCII part names on RFC 3987 (Introduction, L210–212).
- **Code:** `percentAt()` decodes a sequence of `%HH` octets as one UTF-8 character. Invalid sequences stay encoded.

## Decisions for review (D-OPC-*)

- **D-OPC-01 — UTF-8 ZIP names.**
  - **Background:** the ZIP layer accepts names flagged as UTF-8 (bit 11) for robustness, although §7.3.3 requires ASCII names and Annex B lists bit 11 as unused.
  - **Mapping:** a raw non-ASCII name such as `xl/ц.xml` and its conformant spelling `xl/%D1%86.xml` map to the same decoded part name, `/xl/ц.xml` (F-OPC-12 step a).
  - **Comparison:** names are compared in decoded form, folding A–Z only.
  - **Duplicates:** if both spellings exist, they are equivalent logical items → `XLV_E_PART_INVALID`.
  - **`zip.find`:** never used for part lookups, because it folds case with `toLowerCase`, which also folds non-ASCII letters. `openPackage` keeps its own `Map` of folded name → entry.

- **D-OPC-02 — Canonical form of references.**
  - **Rule:** `resolveTarget` and `Override@PartName` decode percent-encoded `iunreserved` characters: ASCII unreserved characters, and UTF-8 sequences that decode to `ucschar`. This happens before dot-segment removal.
  - **Effect:** `%D1%86.xml` finds the part stored as `xl/%D1%86.xml`.
  - **Basis:** a name that still contains such an encoding is not a valid part name (F-OPC-01 b), so it could never match a part.
    - The reference is converted to its IRI form (RFC 3987 §3.2, F-OPC-15), restricted to `iunreserved` characters.
    - Annex A (informative) suggests the same step for ASCII unreserved characters only.
  - **Not repaired:** other invalid targets (backslash, space, `[`, empty segment, a segment ending in `.`). `resolveTarget` returns `null` for them.

- **D-OPC-03 — Climbing above the root.** `../../../x` is not an error. RFC 3986 remove_dot_segments collapses it at `/`, as Part 2 §6.4.3.5 case 2 shows, so it resolves to `/x`. The result is only ever used as a key into the archive's own entry map, so it cannot leave the package.

- **D-OPC-04 — Partial schema validation.**
  - **Checked:** root namespace and local name; the required attributes; the `TargetMode` enumeration; uniqueness of `Id`, `Extension` and `PartName`.
  - **Not checked:**
    - unknown elements and attributes are ignored. §6.2.5 d forbids foreign namespaces and §7.2.3.2.2 forbids MCE in the Media Types stream;
    - no MCE processing is done for `.rels`. `mc:AlternateContent` and its children are ignored, and `mc:MustUnderstand` is not honoured;
    - the syntax of `xsd:ID`, `xsd:anyURI` and `ST_ContentType` is not checked;
    - the absence of `xml:base` is not checked;
    - no whitespace collapsing is applied to attribute values. They are compared exactly as the XML layer delivers them.

- **D-OPC-05 — Internal targets that cannot resolve.** An internal `Target` that is not a relative reference to a part makes `relationships()` throw `XLV_E_PART_INVALID`. Examples: an absolute IRI, `//host/…`, a backslash, a name ending in `.`. The relationship is not dropped silently. An alternative is a warning, but no warning code exists for this.

- **D-OPC-06 — Required attributes and duplicates fail the whole part.** Missing attributes, a duplicate `Default` or `Override`, or a duplicate `Id` throw `XLV_E_PART_INVALID`; the reader does not pick one of the duplicates. Duplicates are ambiguous, and two readers could resolve them differently.

- **D-OPC-07 — Not enforced.**
  - Derivable part names (`/a` together with `/a/b`).
  - That every part has a media type (§7.2.3.2.1, Part 1 §9.1.6). `contentType()` just returns `undefined`.
  - The bans on relationships from or to Relationships parts.
  - The NFC/NFD "should" rules.
