# Verified facts — ZIP, XML 1.0, Namespaces, and supporting web sources

> **Provenance.** Compiled on 2026-10-30 from the official documents listed below. Line numbers (`Lnnn`) refer to text extracted with `pdftotext -layout` (poppler) from local copies of the sources listed in the table below (APPNOTE.TXT, W3C HTML pages, Microsoft Learn pages). Paths such as `scratchpad/specs/…` below are the working copies used during the check; the specs are not stored in this repository.


Compiled 2026-10-30 for a security-focused XLSX reader (own ZIP central-directory reader, own XML parser with no DTD support).

## Sources actually read

| Key | Document | Version / date (as stated in the document) | Location |
|---|---|---|---|
| APPNOTE | PKWARE "APPNOTE.TXT - .ZIP File Format Specification" | **Version 6.3.10**, Status FINAL (replaces 6.3.9), **Revised Nov 01, 2022** (header lines 1–5) | `scratchpad/specs/APPNOTE.TXT` |
| XML | W3C "Extensible Markup Language (XML) 1.0" | **Fifth Edition**, W3C Recommendation **26 November 2008** (modified in place 7 Feb 2013 to fix RFC4646/4647 links) | `scratchpad/specs/xml10.html` (= http://www.w3.org/TR/2008/REC-xml-20081126/) |
| NS | W3C "Namespaces in XML 1.0" | **Third Edition**, W3C Recommendation **8 December 2009** | `scratchpad/specs/xml-names.html` (= http://www.w3.org/TR/2009/REC-xml-names-20091208/) |
| MS-CFB | [MS-CFB] §2.2 Compound File Header | page "Last updated" 2021-06-24 | https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/05060311-bfce-4b12-874d-71fd4ce63aea |
| MS-OFFCRYPTO | [MS-OFFCRYPTO] §1.3.3.4, §1.3.8, §2.3.4, §2.3.4.4, §2.3.4.5, §2.3.4.10 | per-page dates given below | learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/… (URLs below) |
| MS-OFFMACRO2 | [MS-OFFMACRO2] §2.2.1.1–§2.2.1.4 | per-page dates given below | learn.microsoft.com/en-us/openspecs/office_standards/ms-offmacro2/… (URLs below) |
| EXCEL | Microsoft Support "Excel specifications and limits" | page `ms.date` 04/13/2026; "Applies To" Excel for Microsoft 365, 2024, 2021, 2019, 2016 | https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3 |
| NODE | Node.js docs "Modules: CommonJS modules" → "Loading ECMAScript modules using require()" | rendered as **Node.js v26.13.0** documentation | https://nodejs.org/api/modules.html#loading-ecmascript-modules-using-require |
| FIFIELD | David Fifield, "A better zip bomb" | dated 2019-07-02, last update listed 2023-05-18; a version appeared at USENIX WOOT 2019 | https://www.bamsoftware.com/hacks/zipbomb/ |
| RFC1952 | RFC 1952 "GZIP file format specification version 4.3" (P. Deutsch, May 1996) — *extra source, only used to fill a gap in APPNOTE (CRC-32)* | May 1996 | https://www.rfc-editor.org/rfc/rfc1952.txt |

Raw copies of the web pages are in `scratchpad/web/`, and plain-text conversions of every source are in `scratchpad/text/`.

### How quotes are handled (please read)
- Verbatim prose is kept to a minimum. APPNOTE §1.4.2 also prohibits reproducing it "in whole or in part" without PKWARE's permission. So each fact has a **verbatim anchor** instead of a full sentence. The anchor is the exact identifier, heading, constant, field name, grammar production, or a very short phrase as it appears in the source, so you can grep for it. The **Fact** column paraphrases the normative sentence as closely as I could.
- Field-layout tables are built from the field names and sizes in the spec. **Offsets are my own calculation** from the listed sizes (APPNOTE gives sizes, not offsets), and are marked *derived*.
- For local files, **APPNOTE line numbers** are given (`L123`). For XML and NS, the HTML anchor `#id` is given, so the URL is `…/REC-xml-20081126/#id` or `…/REC-xml-names-20091208/#id`.
- "Derived" means a logical consequence of cited text that the source does not state itself. "NOT FOUND" means I searched and the source does not say it.

---

## ZIP (APPNOTE 6.3.10)

### Z-01 — General encoding rules
- **Fact:** All values are stored little-endian unless a specific element says otherwise (§4.3.3). All fields are unsigned unless noted (§4.4.1.1). String fields are not null-terminated because their length is given explicitly (§4.4.1.2). Every record type is identified by a header signature, and signature values begin with the two-byte marker 0x4b50 ("PK") (§4.2.1).
- **Source:** APPNOTE §4.2.1 (L326–330), §4.3.3 (L348–351), §4.4.1.1–4.4.1.2 (L654–658).
- **Anchor:** "little-endian byte order" · "Intel low-byte:high-byte" · "0x4b50"
- **Notes:** Because signatures are stored little-endian, 0x04034b50 is the byte sequence `50 4B 03 04` on disk (*derived*).

### Z-02 — End of central directory record (EOCD): signature and layout
- **Fact:** The EOCD starts with the signature 0x06054b50 and has these fields in this order:

| Offset (*derived*) | Size | Field (APPNOTE name) |
|---|---|---|
| 0 | 4 | end of central dir signature (0x06054b50) |
| 4 | 2 | number of this disk |
| 6 | 2 | number of the disk with the start of the central directory |
| 8 | 2 | total number of entries in the central directory on this disk |
| 10 | 2 | total number of entries in the central directory |
| 12 | 4 | size of the central directory |
| 16 | 4 | offset of start of central directory with respect to the starting disk number |
| 20 | 2 | .ZIP file comment length |
| 22 | n | .ZIP file comment (variable size) |

  The fixed part is 22 bytes (*derived*: 4+2+2+2+2+4+4+2).
- **Source:** APPNOTE §4.3.16 "End of central directory record" (L632–647). Field semantics are in §4.4.19–4.4.26 (L1014–1075).
- **Anchor:** "end of central dir signature    4 bytes  (0x06054b50)"
- **Notes:** §4.3.1 (L336–341): a ZIP file MUST contain an EOCD and MUST have only one. An archive with only an EOCD is an empty ZIP file. §4.3.4 (L353–356): the EOCD MUST NOT be compressed. §4.4.26 (L1070–1075): the comment is stored unsecured, with no encryption or authentication.

### Z-03 — Maximum comment length, and so how far back to search for the EOCD
- **Fact:** The .ZIP file comment length is a 2-byte unsigned field (§4.4.25 with §4.4.1.1), so the comment is at most 65,535 bytes. The EOCD therefore starts within the last 22 + 65,535 = **65,557 bytes** of the file (*derived*).
- **Source:** APPNOTE §4.3.16 (L646–647), §4.4.25 (L1066–1068), §4.4.1.1 (L654–655).
- **Anchor:** ".ZIP file comment length        2 bytes"
- **Notes:** APPNOTE gives **no algorithm** for locating the EOCD (NOT FOUND: searched §4.3, §4.4 and the whole file for "search", "scan", "locate"). FIFIELD (W-10) points out there is no single well-defined algorithm for locating the central directory, so two parsers may disagree about an archive's contents. A defensive reader should scan backwards and require that `comment length == bytes remaining after the EOCD` (*recommendation, not in APPNOTE*).

### Z-04 — Sentinel values in the EOCD mean "look in ZIP64"
- **Fact:** If an EOCD field is too small for the value, the field SHOULD be set to −1 (0xFFFF or 0xFFFFFFFF) and the ZIP64 record SHOULD be created (§4.4.1.4). Each EOCD field description says the same: if the archive is ZIP64 and the field is 0xFFFF (2-byte fields) or 0xFFFFFFFF (4-byte fields), the real value is in the corresponding zip64 end of central directory field.
- **Source:** APPNOTE §4.4.1.4 (L663–666), §4.4.19–4.4.24 (L1014–1064).
- **Anchor:** "set to -1 (0xFFFF or 0xFFFFFFFF)"
- **Notes:** §4.4.1.5 (L668–670): the EOCD and the Zip64 EOCD locator MUST be on the same disk when an archive is split or spanned.

### Z-05 — Central directory file header (CDFH): signature and layout
- **Fact:** Each central directory file header starts with 0x02014b50:

| Offset (*derived*) | Size | Field |
|---|---|---|
| 0 | 4 | central file header signature (0x02014b50) |
| 4 | 2 | version made by |
| 6 | 2 | version needed to extract |
| 8 | 2 | general purpose bit flag |
| 10 | 2 | compression method |
| 12 | 2 | last mod file time |
| 14 | 2 | last mod file date |
| 16 | 4 | crc-32 |
| 20 | 4 | compressed size |
| 24 | 4 | uncompressed size |
| 28 | 2 | file name length |
| 30 | 2 | extra field length |
| 32 | 2 | file comment length |
| 34 | 2 | disk number start |
| 36 | 2 | internal file attributes |
| 38 | 4 | external file attributes |
| 42 | 4 | relative offset of local header |
| 46 | n | file name, then extra field, then file comment (each variable size) |

  The fixed part is 46 bytes (*derived*). FIFIELD independently says a CDH is "46 bytes long".
- **Source:** APPNOTE §4.3.12 "Central directory structure" → "File header" (L511–542).
- **Anchor:** "central file header signature   4 bytes  (0x02014b50)"
- **Notes:** The central directory is `[CDH 1]…[CDH n][digital signature]`. The digital signature record has signature 0x05054b50 (§4.3.13, L544–548). Entries in the central directory need not be in the same order as the files in the archive (§4.4.1.3, L660–661). Individual CD records MUST NOT be compressed, but the CD as a whole MAY be, under central-directory encryption (§4.3.4).

### Z-06 — CDFH field semantics relevant to a reader
- **Fact:**
  - **version made by:** the upper byte is the host system (0 = MS-DOS/OS/2 FAT, 3 = UNIX, 10 = Windows NTFS, 19 = OS X, 20–255 unused). The lower byte is the spec version, where value/10 = major and value mod 10 = minor (§4.4.2).
  - **version needed to extract:** minimum feature version, in the same encoding. Values include 2.0 (deflate, directory, traditional PKWARE encryption), 4.5 (ZIP64), 5.1 (AES), 6.2 (central directory encryption) and 6.3 (LZMA, PPMd+, Blowfish, Twofish) (§4.4.3.2).
  - **name/extra/comment lengths:** the combined length of a directory record and these three fields SHOULD NOT generally exceed 65,535 bytes (§4.4.10–4.4.12).
  - **disk number start:** the disk on which the file begins. 0xFFFF means the value is in the 4-byte zip64 extra field (§4.4.13).
  - **internal attributes:** the lowest bit means "apparently text". Bit 0x0002 means a 4-byte record-length control field precedes each logical record (§4.4.14).
  - **external attributes:** host-dependent. For MS-DOS, the low byte is the DOS directory attribute byte (§4.4.15).
  - **relative offset of local header:** offset from the start of the first disk on which the file appears to where the local header SHOULD be found. 0xFFFFFFFF means the value is in the 8-byte zip64 extra field (§4.4.16).
- **Source:** APPNOTE §4.4.2 (L672–701), §4.4.3 (L703–764), §4.4.10–4.4.16 (L937–987).
- **Anchor:** "4.5 - File uses ZIP64 format extensions" · "relative offset of local header: (4 bytes)"
- **Notes:** The wording "where the local header SHOULD be found" means the spec does not guarantee the pointer is valid. Validate it: in range, signature present, and no overlap (see W-09/W-10).

### Z-07 — Local file header (LFH): signature and layout
- **Fact:**

| Offset (*derived*) | Size | Field |
|---|---|---|
| 0 | 4 | local file header signature (0x04034b50) |
| 4 | 2 | version needed to extract |
| 6 | 2 | general purpose bit flag |
| 8 | 2 | compression method |
| 10 | 2 | last mod file time |
| 12 | 2 | last mod file date |
| 14 | 4 | crc-32 |
| 18 | 4 | compressed size |
| 22 | 4 | uncompressed size |
| 26 | 2 | file name length |
| 28 | 2 | extra field length |
| 30 | n | file name, then extra field (variable size) |

  The fixed part is 30 bytes (*derived*). FIFIELD independently says an LFH is "30 bytes long".
- **Source:** APPNOTE §4.3.7 "Local file header" (L387–402).
- **Anchor:** "local file header signature     4 bytes  (0x04034b50)"
- **Notes:**
  - Each file MUST be preceded by an LFH, and each LFH MUST have a matching CD header (§4.3.2, L343–346).
  - File data SHOULD immediately follow the local header. An encryption header, if present, sits between them (§4.3.8, L404–412).
  - Zero-byte files and directories MUST NOT include file data (§4.3.8, L414–415).
  - The LFH has **no** comment length, and its name/extra lengths may differ from the CDH's. Data starts at `LFH offset + 30 + LFH.nameLen + LFH.extraLen` (*derived*).

### Z-08 — File name path rules
- **Fact:** The file name is a name with an optional relative path. It MUST NOT contain a drive or device letter or a leading slash, and all slashes MUST be forward slashes '/' (§4.4.17.1). If the archive came from standard input, there is no file name field.
- **Source:** APPNOTE §4.4.17.1 (L989–997).
- **Anchor:** "MUST NOT contain a drive or device letter, or a leading slash"
- **Notes:** **NOT FOUND:** APPNOTE says nothing about `..` path segments or path traversal (searched for `..`, "parent director", "traversal"). Rejecting them is a reader policy, not a spec rule. Under CD encryption with bit 13 set, the LFH name is a masked sequential hex value, not the real name (§4.4.17.2, §7.3.2).

### Z-09 — General purpose bit flags
- **Fact (APPNOTE §4.4.4, L767–852):**
  - **Bit 0:** set means the file is encrypted.
  - **Bits 1–2:** meaning depends on method. Method 6 (implode): dictionary/tree options. Methods 8/9 (deflate): compression option (normal/max/fast/superfast). Method 14 (LZMA), bit 1: EOS marker used. Undefined for other methods.
  - **Bit 3:** set means the crc-32, compressed size and uncompressed size fields **are set to zero in the local header**, and the correct values are in the data descriptor immediately after the compressed data. (PKZIP 2.04g honoured this only for method 8; newer versions honour it for any method.)
  - **Bit 4:** reserved for enhanced deflating. **Bit 5:** compressed patched data.
  - **Bit 6:** strong encryption. When set, version needed MUST be ≥ 50 **and bit 0 MUST also be set**. With AES, version needed MUST be ≥ 51.
  - **Bits 7–10:** currently unused.
  - **Bit 11:** language encoding flag (EFS). When set, the filename and comment MUST be encoded in UTF-8 (see Appendix D).
  - **Bit 12:** reserved by PKWARE for enhanced compression.
  - **Bit 13:** set when the Central Directory is encrypted, meaning selected Local Header values are **masked** to hide their real values.
  - **Bit 14:** reserved for alternate streams. **Bit 15:** reserved by PKWARE.
- **Source:** APPNOTE §4.4.4 (bit 0 L769; bit 3 L802–809; bit 6 L818–826; bit 11 L836–838; bit 13 L842–848).
- **Anchor:** "Bit 3: If this bit is set, the fields crc-32, compressed" · "Bit 11: Language encoding flag (EFS)"
- **Notes:** The masked LFH fields under bit 13 are compression method, time, date, crc-32, compressed size and uncompressed size (mask value 0), plus the file name (sequential base-16 value). ZIP64 sentinel values SHOULD NOT be masked (§7.3.2, L2990–3024). For an XLSX reader, bits 0, 6 and 13 can be treated as "encrypted, reject" (*recommendation*).

### Z-10 — Compression method codes
- **Fact:** Method **0** means "The file is stored (no compression)". Method **8** means "The file is Deflated". Other codes: 1 Shrunk, 2–5 Reduced, 6 Imploded, 7 reserved (Tokenizing), 9 Deflate64, 10 PKWARE DCL Imploding, 12 BZIP2, 14 LZMA, 16 IBM z/OS CMPSC, 18 IBM TERSE (new), 19 IBM LZ77 z, 20 deprecated (use 93), 93 Zstandard, 94 MP3, 95 XZ, 96 JPEG variant, 97 WavPack, 98 PPMd I rev 1, **99 AE-x encryption marker**. Codes 11, 13, 15 and 17 are reserved.
- **Source:** APPNOTE §4.4.5 (L854–886). Methods 1–6 are legacy and no longer recommended (§4.4.5.1). Most programs use method 8 by default (§4.1.3, L274–275). "Stored" means copied uncompressed (§4.1.7, L295–297).
- **Anchor:** "0 - The file is stored (no compression)" · "8 - The file is Deflated"
- **Notes:** Appendix E (L3768–3797): AE-x (WinZip AES) sets bit 0, uses method 99, and keeps the real method in extra field **0x9901**. **AE-2 stores 0 in the CRC-32 field** (E.6.2), so a CRC check would fail on AE-2 entries. That is another reason to reject method 99.

### Z-11 — Zip64 end of central directory record
- **Fact:**

| Offset (*derived*) | Size | Field |
|---|---|---|
| 0 | 4 | zip64 end of central dir signature (0x06064b50) |
| 4 | 8 | size of zip64 end of central directory record |
| 12 | 2 | version made by |
| 14 | 2 | version needed to extract |
| 16 | 4 | number of this disk |
| 20 | 4 | number of the disk with the start of the central directory |
| 24 | 8 | total number of entries in the central directory on this disk |
| 32 | 8 | total number of entries in the central directory |
| 40 | 8 | size of the central directory |
| 48 | 8 | offset of start of central directory with respect to the starting disk number |
| 56 | n | zip64 extensible data sector (variable size) |

  The "size of zip64 end of central directory record" SHOULD be the size of the remaining record and SHOULD NOT include the leading 12 bytes: Size = SizeOfFixedFields + SizeOfVariableData − 12 (§4.3.14.1). With no extensible data the fixed part is 56 bytes, so the size field is 44 (*derived*).
- **Source:** APPNOTE §4.3.14 (L560–619).
- **Anchor:** "(0x06064b50)" · "SHOULD NOT include the leading 12 bytes"
- **Notes:** This layout is Version 1. Version 2 belongs to Central Directory Encryption (§4.3.14.2). The extensible data sector holds blocks of Header ID (2) + Data Size (**4**) (§4.3.14.3, L601–614). It is "currently reserved for use by PKWARE" (§4.4.27, L1077–1079). Note the 4-byte Data Size here, compared with the 2-byte size in ordinary extra fields.

### Z-12 — Zip64 end of central directory locator
- **Fact:**

| Offset (*derived*) | Size | Field |
|---|---|---|
| 0 | 4 | zip64 end of central dir locator signature (0x07064b50) |
| 4 | 4 | number of the disk with the start of the zip64 end of central directory |
| 8 | 8 | relative offset of the zip64 end of central directory record |
| 16 | 4 | total number of disks |

  Total: 20 bytes (*derived*).
- **Source:** APPNOTE §4.3.15 (L621–630). The overall-format order puts `[zip64 EOCD record][zip64 EOCD locator][EOCD]` at the end of the archive (§4.3.6, L382–384).
- **Anchor:** "(0x07064b50)" · "total number of disks           4 bytes"
- **Notes:** The locator is listed immediately before the EOCD, so readers look for it at `EOCD offset − 20` (*derived*). APPNOTE never explicitly says "no gap". Confirm the record at the locator's offset has signature 0x06064b50.

### Z-13 — Zip64 Extended Information Extra Field (header ID 0x0001)
- **Fact:** Layout: Tag 0x0001 (2 bytes), Size of this extra block (2), then **Original (uncompressed) Size** (8), **Compressed Size** (8), **Relative Header Offset** (8), **Disk Start Number** (4). "The order of the fields … is fixed, but the fields MUST only appear if the corresponding Local or Central directory record field is set to 0xFFFF or 0xFFFFFFFF." In the Local header this entry MUST include BOTH the original and compressed sizes.
- **Source:** APPNOTE §4.5.3 (L1148–1180). §4.4.8/4.4.9 (L929–932): a size field of 0xFFFFFFFF means the 8-byte value is in this extra field. §4.4.13 (L950–953): disk start 0xFFFF means the 4-byte value is here. §4.4.16 (L985–987): local-header offset 0xFFFFFFFF means the value is here.
- **Anchor:** "fields MUST only appear if the corresponding Local or Central"
- **Notes:** **Gotcha:** the extra field puts uncompressed size **before** compressed size, the reverse of the header order. Parse it as a cursor: for each of (uncompressed, compressed, offset, disk), read the value only if the CDH field holds the sentinel (*derived*). Reading files that use ZIP64 needs version 4.5 (§4.4.3.2).

### Z-14 — Extra field framing
- **Fact:** The extra field is a sequence of `header + data` blocks. Each header is Header ID (2 bytes) + Data Size (2 bytes), little-endian. IDs 0–31 are reserved for PKWARE, and the rest may be used by third parties.
- **Source:** APPNOTE §4.5.1 (L1097–1116). PKWARE IDs are listed in §4.5.2 (L1118–1145).
- **Anchor:** "Header IDs of 0 thru 31 are reserved for use by PKWARE."
- **Notes:** Some IDs: 0x0001 Zip64, 0x000a NTFS, 0x000d UNIX, 0x0017 Strong Encryption Header. Third-party IDs: 0x6375/0x7075 Info-ZIP Unicode (§4.6.8–4.6.9), 0xa220 Microsoft Open Packaging Growth Hint (§4.6.10, L1820–1828), 0x9901 AE-x (App. E). FIFIELD's "extra-field quoting" bomb abuses this length–value structure (W-09). A reader must check that each block's Data Size fits inside the declared extra-field length (*recommendation*).

### Z-15 — Data descriptor
- **Fact:** Layout: crc-32 (4), compressed size (4), uncompressed size (4). It MUST exist if bit 3 is set. It is byte-aligned and immediately follows the last byte of compressed data. In ZIP64 archives the two sizes are 8 bytes each (§4.3.9.1). When extracting, if the zip64 extra field is present, the sizes are 8-byte values (§4.3.9.2). The value **0x08074b50** has commonly been adopted as an optional signature. Readers SHOULD handle descriptors with or without it (§4.3.9.3), and writers SHOULD include it (§4.3.9.4).
- **Source:** APPNOTE §4.3.9 (L417–464). Bit 3 is in §4.4.4 (L802–809). CRC and sizes go "in the data descriptor and in the central directory" (§4.4.7 L911–914, §4.4.8–9 L926–929).
- **Anchor:** "0x08074b50 has commonly been adopted as a signature value"
- **Notes:** The possible descriptor sizes are 12/16 bytes (32-bit, without/with signature) and 20/24 bytes (ZIP64) (*derived*). §8.5.5 (L3431–3436): 0x08074b50 is **also** the spanning signature at the start of the first segment, and position tells the two uses apart. A central-directory-driven reader can take CRC and sizes from the CDH and never needs to parse the descriptor (*derived* from §4.4.7–4.4.9). Under CD encryption a present descriptor MUST be all zeros (§4.3.9.6).

### Z-16 — Multi-disk (spanned/split) fields
- **Fact:**
  - The EOCD has "number of this disk" (the disk that holds the EOCD, §4.4.19), "number of the disk with the start of the central directory" (§4.4.20), and "total number of entries in the central dir on this disk" (§4.4.21) as well as the overall total (§4.4.22).
  - Each CDH has "disk number start" (the disk on which the file begins, §4.4.13).
  - The ZIP64 EOCD has 4-byte versions of the two disk numbers (§4.3.14).
  - The ZIP64 locator has "number of the disk with the start of the zip64 end of central directory" and "total number of disks" (§4.3.15).
  - Section 8: split archives are named `filename.z01 … filename.z(n-1)`, with `filename.zip` last (§8.3.3).
  - Spanned/split archives made by PKZIP ≥ 2.50 start with the spanning signature 0x08074b50 (§8.5.3). A one-segment "temporary spanning marker" uses 0x30304b50 (§8.5.4).
  - Capacities: at most 4,294,967,295 − 1 segments, and a minimum segment size of 64K (§8.5.1).
- **Source:** APPNOTE §4.4.13, §4.4.19–4.4.22 (L948–1046), §4.3.14–4.3.15, §8.0–8.5 (L3341–3436).
- **Anchor:** "number of the disk with the start of the central directory"
- **Notes:** **NOT FOUND:** APPNOTE never states that disk numbering starts at 0 (searched "first disk", "zero", "numbered"). Treating `this disk == 0 && CD-start disk == 0 && entries-on-this-disk == total entries` (and locator total disks == 1) as "single-volume" is a de-facto convention, not a quoted rule. A reader that rejects multi-disk archives can also reject a leading 0x08074b50 or 0x30304b50 (*recommendation*).

### Z-17 — CRC-32
- **Fact:** APPNOTE says the CRC-32 algorithm came from David Schwaderer, and its "magic number" is **0xdebb20e3**. The register is pre-conditioned to all ones (0xffffffff), and the result is post-conditioned by taking the one's complement of the residual. If bit 3 is set, the LFH crc field is zero and the real value is in the data descriptor and the central directory. Data integrity MUST be provided for each file using CRC32 (§4.1.5).
- **Source:** APPNOTE §4.4.7 (L900–917), §4.1.5 (L288).
- **Anchor:** "The 'magic number' for the CRC is 0xdebb20e3."
- **Notes:**
  - **NOT FOUND in APPNOTE:** (a) the generator polynomial, and (b) **which bytes** the CRC covers (uncompressed or compressed). I checked every "CRC" occurrence (L288–3797). The same CRC-32 is reused in traditional PKWARE encryption key updates (§6.1.3, L2600–2612).
  - **Supplement (RFC1952 §2.3.1 "CRC32"):** gzip's CRC32 is computed over the **uncompressed data** using the CRC-32 of ISO 3309 / ITU-T V.42. The RFC1952 §8 sample code builds its table with `0xedb88320L` and inits/finalizes with all-ones. **Do not use 0xdebb20e3 as the table polynomial.** APPNOTE calls it a "magic number" and never says it is the polynomial.
  - Appendix E.6.2: AE-2 entries store 0 in the CRC field.

### Z-18 — File name encoding without bit 11; Info-ZIP Unicode Path extra field (0x7075)
- **Fact (encoding):** ZIP historically supported only the original IBM PC character set, "commonly referred to as IBM Code Page 437" (D.1). If bit 11 is unset, the file name and comment SHOULD conform to the original ZIP encoding. If bit 11 is set, they MUST be UTF-8 (Unicode 4.1.0 or later), and UTF-8 data is expected not to include a BOM (D.2). Extra field 0x0008 is reserved for future language-encoding data, and its layout is undefined (D.3–D.5). Readers SHOULD support both the bit-11 method and the Info-ZIP extra fields (D.8).
- **Fact (0x7075):** Layout: tag 0x7075 ("up", 2 bytes), TSize (2), Version (1, currently 1), NameCRC32 (4), UnicodeName (variable, UTF-8, no BOM). The length of UnicodeName is TSize minus the preceding fields. NameCRC32 is the standard zip CRC32 of the header's File Name field. If that CRC check fails, the extra field SHOULD be ignored and the header File Name used. The field SHOULD NOT be used if the Version is not recognized. The comment equivalent is 0x6375 ("uc"), with ComCRC32 (§4.6.8).
- **Source:** APPNOTE Appendix D (L3706–3766), §4.6.9 (L1778–1817), §4.6.8 (L1736–1775).
- **Anchor:** "IBM Code Page 437" · "(UPath) 0x7075"
- **Notes:** APPNOTE gives no CP437 mapping table (NOT FOUND), so an implementation needs one from elsewhere. Security: 0x7075 lets the "effective" name differ from the raw header name. Path validation must run on whichever name the reader finally uses, and that choice must be the same for the duplicate-name check (*recommendation*).

### Z-19 — Other structural rules useful for a strict reader
- **Fact:**
  - Files MAY be stored in any order (§4.3.3).
  - Compression MUST NOT be applied to a local file header, an encryption header or the EOCD (§4.3.4).
  - The overall layout is: LFH/encryption header/data/descriptor repeated, then archive decryption header, archive extra data record, CD headers, zip64 EOCD record, zip64 locator, EOCD (§4.3.6).
  - The archive extra data record has signature 0x08064b50 and is used only with CD encryption (§4.3.11).
  - Date and time use MS-DOS format: years relative to 1980, 2-second precision (§4.4.6).
- **Source:** APPNOTE §4.3.3–4.3.6 (L348–384), §4.3.11 (L489–509), §4.4.6 (L888–898).
- **Anchor:** "Overall .ZIP file format:" · "(0x08064b50)"
- **Notes:** APPNOTE gives no bit-level layout of MS-DOS date and time (NOT FOUND; it only says "standard MS-DOS format").

---

## XML 1.0 (Fifth Edition, 26 Nov 2008)

### X-01 — `Char` production
- **Fact:** `[2] Char ::= #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]`, i.e. any Unicode character except the surrogate blocks, FFFE and FFFF. Processors MUST accept any character in the range of Char. The note discourages, but does **not** forbid, `[#x7F-#x84]`, `[#x86-#x9F]`, `[#xFDD0-#xFDEF]` and the `#xnFFFE–#xnFFFF` code points of every plane.
- **Source:** XML §2.2 "Characters", `#NT-Char`, `#charsets`.
- **Anchor:** "[2] Char ::= #x9 | #xA | #xD | [#x20-#xD7FF]"
- **Notes:** Every C0 control except TAB, LF and CR is illegal, including NUL. C1 controls (0x80–0x9F) are **legal** in XML 1.0. Check Char after decoding, on every character, including inside comments, PIs, CDATA and attribute values.

### X-02 — WFC: Legal Character (character references)
- **Fact:** Characters referred to using character references MUST match the production for Char. `[66] CharRef ::= '&#' [0-9]+ ';' | '&#x' [0-9a-fA-F]+ ';'`. The `&#x` form gives a hexadecimal code point and the `&#` form a decimal one.
- **Source:** XML §4.1 "Character and Entity References", `#NT-CharRef`, `#wf-Legalchar`.
- **Anchor:** "Well-formedness constraint: Legal Character"
- **Notes:** So `&#0;`, `&#xD800;`, `&#xFFFE;`, and anything above 0x10FFFF are fatal errors. `&#xD;` and `&#9;` are legal. Note the lowercase `x` only: `&#X41;` does not match the production. This is *derived* from the literal `'&#x'` plus XML §6 Notation, where a quoted 'string' "matches a literal string" exactly. Guard against overflow when parsing very long digit runs (*recommendation*).

### X-03 — `S` (white space)
- **Fact:** `[3] S ::= (#x20 | #x9 | #xD | #xA)+`. The note says #xD stays in S only for backward compatibility. Because of §2.11, a literal #xD never reaches the parser, and the only way to get #xD to match S is a character reference in an entity value literal.
- **Source:** XML §2.3 "Common Syntactic Constructs", `#NT-S`, `#sec-common-syn`.
- **Anchor:** "[3] S ::= (#x20 | #x9 | #xD | #xA)+"

### X-04 — §2.11 End-of-Line Handling (exact rule)
- **Fact:** The processor MUST behave as if it normalized all line breaks in external parsed entities (including the document entity) on input, before parsing. Both the two-character sequence `#xD #xA` and any `#xD` not followed by `#xA` become a single `#xA`.
- **Source:** XML §2.11, `#sec-line-ends`.
- **Anchor:** "translating both the two-character sequence #xD #xA"
- **Notes:** This happens **before** parsing, so it also applies inside attribute values, CDATA, comments and PIs. It does not touch character references: `&#xD;` produces a real CR (see X-05). XML 1.0 normalizes only CR and CRLF. #x85 and #x2028 are **not** line breaks in 1.0 (*derived*: the rule names only #xD and #xA).

### X-05 — §3.3.3 Attribute-Value Normalization
- **Fact:** All line breaks are first normalized to #xA (§2.11). Then, starting from an empty string, for each item of the raw value:
  - a **character reference** appends the referenced character;
  - an **entity reference** applies the algorithm recursively to the replacement text;
  - a **white-space character (#x20, #xD, #xA, #x9)** appends a space (#x20);
  - any other character is appended unchanged.

  If the attribute type is not CDATA, the processor also strips leading and trailing spaces and collapses runs of spaces. **Attributes with no declaration read SHOULD be treated as CDATA by a non-validating processor.** It is an error if an attribute value references an entity with no declaration read.
- **Source:** XML §3.3.3, `#AVNormalize`.
- **Anchor:** "append a space character (#x20) to the normalized value"
- **Notes:** The spec gives the consequence explicitly: a character reference to white space other than #x20 keeps the referenced character (#xD, #xA or #x9), while a literal white-space character becomes #x20. The worked example shows `"&#xd;&#xd;A&#xa;&#xa;B&#xd;&#xa;"` normalizing to `#xD #xD A #xA #xA B #xD #xA` for CDATA. So `&#10;` in an attribute stays LF, and a literal newline or tab becomes a space. Without a DTD, every attribute is CDATA: never trim or collapse.

### X-06 — §4.1 references and WFC: Entity Declared (exact condition)
- **Fact:** `[67] Reference ::= EntityRef | CharRef`. `[68] EntityRef ::= '&' Name ';'`, carrying [WFC: Entity Declared], [VC: Entity Declared], [WFC: Parsed Entity] and [WFC: No Recursion]. `[69] PEReference ::= '%' Name ';'`.
  - **WFC: Entity Declared applies to:** a document **without any DTD**, a document with **only an internal DTD subset that contains no parameter entity references**, or a document with **standalone='yes'**.
  - **What it requires:** for an entity reference that does not occur in the external subset or a parameter entity, the Name MUST match an entity declaration that also does not occur in the external subset or a parameter entity. Well-formed documents need not declare `amp`, `lt`, `gt`, `apos` or `quot`.
  - A general entity's declaration MUST precede any reference to it in an attribute-list default value.
  - Non-validating processors are not obliged to read declarations in parameter entities or the external subset. For such documents, "must be declared" is a WFC only if standalone='yes'.
- **Source:** XML §4.1, `#sec-references`, `#NT-EntityRef`, `#wf-entdeclared`.
- **Anchor:** "In a document without any DTD, a document with only an internal DTD subset"
- **Notes:** For a DTD-less parser, the first branch applies. Any `&name;` other than the five predefined names is a **well-formedness (fatal) error**, not something to pass through.

### X-07 — §4.6 Predefined entities
- **Fact:** The predefined general entities are `amp`, `lt`, `gt`, `apos` and `quot`. All XML processors MUST recognize them whether declared or not. If `lt` or `amp` are declared, it must be as internal entities whose replacement text is a character reference (double escaping). The example declarations give the code points: lt → `&#60;` ('<'), gt → `&#62;` ('>'), amp → `&#38;` ('&'), apos → `&#39;` ('''), quot → `&#34;` ('"').
- **Source:** XML §4.6 "Predefined Entities", `#sec-predefined-ent`. §2.4 explains when '&', '<' and '>' must be escaped.
- **Anchor:** "All XML processors MUST recognize these entities whether they are declared or not."
- **Notes:** Character references are expanded immediately and MUST be treated as character data (§4.6). So `&#60;` in content is text and never starts a tag.

### X-08 — WFC: Unique Att Spec
- **Fact:** An attribute name MUST NOT appear more than once in the same start-tag or empty-element tag. This constraint is attached to `[40] STag` and `[44] EmptyElemTag`. The order of attribute specifications is not significant.
- **Source:** XML §3.1 "Start-Tags, End-Tags, and Empty-Element Tags", `#sec-starttags`, `#uniqattspec`.
- **Anchor:** "Well-formedness constraint: Unique Att Spec"
- **Notes:** This compares raw (lexical) names. Namespaces adds a second, stricter check on expanded names (N-07). Related WFCs on `[41] Attribute`: "No External Entity References" and "No < in Attribute Values" (the replacement text of any entity referenced in an attribute value MUST NOT contain '<'). Also `[10] AttValue ::= '"' ([^<&"] | Reference)* '"' | …`, so a literal '<' in an attribute value is itself a syntax error.

### X-09 — `Name`, `NameStartChar`, `NameChar`
- **Fact:**
  - `[4] NameStartChar ::= ":" | [A-Z] | "_" | [a-z] | [#xC0-#xD6] | [#xD8-#xF6] | [#xF8-#x2FF] | [#x370-#x37D] | [#x37F-#x1FFF] | [#x200C-#x200D] | [#x2070-#x218F] | [#x2C00-#x2FEF] | [#x3001-#xD7FF] | [#xF900-#xFDCF] | [#xFDF0-#xFFFD] | [#x10000-#xEFFFF]`
  - `[4a] NameChar ::= NameStartChar | "-" | "." | [0-9] | #xB7 | [#x0300-#x036F] | [#x203F-#x2040]`
  - `[5] Name ::= NameStartChar (NameChar)*`

  Names beginning with "xml" in any case are reserved for standardization. Processors must accept the colon as a name character, although Namespaces gives it meaning.
- **Source:** XML §2.3, `#NT-NameStartChar`, `#NT-NameChar`, `#NT-Name`.
- **Anchor:** "[5] Name ::= NameStartChar (NameChar)*"
- **Notes:** #x037E (GREEK QUESTION MARK) is deliberately excluded (§2.3 explains why). The ranges are code points, so match on decoded code points, not UTF-16 units.

### X-10 — Comments
- **Fact:** `[15] Comment ::= '<!--' ((Char - '-') | ('-' (Char - '-')))* '-->'`. The string "--" MUST NOT occur within comments. A comment ending in `--->` is not well-formed (the spec gives `<!-- B+, B, or B--->` as an example). Parameter-entity references are not recognized in comments. Comments are not part of character data, and processors MAY expose them.
- **Source:** XML §2.5 "Comments", `#sec-comments`, `#NT-Comment`.
- **Anchor:** "the string \"--\" (double-hyphen) MUST NOT occur within comments"

### X-11 — Processing instructions and the "xml" target restriction
- **Fact:** `[16] PI ::= '<?' PITarget (S (Char* - (Char* '?>' Char*)))? '?>'`. `[17] PITarget ::= Name - (('X' | 'x') ('M' | 'm') ('L' | 'l'))`. Target names "XML", "xml" and so on are reserved. PIs MUST be passed to the application. Parameter-entity references are not recognized in PIs.
- **Source:** XML §2.6 "Processing Instructions", `#sec-pi`, `#NT-PI`, `#NT-PITarget`.
- **Anchor:** "[17] PITarget ::= Name - (('X' | 'x') ('M' | 'm') ('L' | 'l'))"
- **Notes:** The grammar subtracts exactly the three-letter name `xml` in any case combination. Targets that merely *start* with "xml" (e.g. `xml-stylesheet`) still match PITarget and are only "reserved" (*derived* from EBNF set difference plus §2.3). So `<?xml …?>` anywhere other than the very start of the document is a fatal error. It matches neither XMLDecl (wrong position) nor PI (excluded target) (*derived*).

### X-12 — CDATA sections
- **Fact:** `[18] CDSect ::= CDStart CData CDEnd`, `[19] CDStart ::= '<![CDATA['`, `[20] CData ::= (Char* - (Char* ']]>' Char*))`, `[21] CDEnd ::= ']]>'`. Inside a CDATA section only CDEnd is recognized as markup, so '<' and '&' appear literally and cannot be escaped there. CDATA sections cannot nest. CDATA sections may occur anywhere character data may occur.
- **Source:** XML §2.7 "CDATA Sections", `#sec-cdata-sect`, `#NT-CDSect`.
- **Anchor:** "CDATA sections cannot nest."
- **Notes:** §2.4 `[14] CharData ::= [^<&]* - ([^<&]* ']]>' [^<&]*)`: the literal string `]]>` in normal content (outside CDATA) is not allowed ('>' MUST be escaped there for compatibility). Characters inside CDATA must still match Char (X-01).

### X-13 — XML declaration, prolog and document productions
- **Fact:**
  - `[1] document ::= prolog element Misc*`
  - `[22] prolog ::= XMLDecl? Misc* (doctypedecl Misc*)?`
  - `[23] XMLDecl ::= '<?xml' VersionInfo EncodingDecl? SDDecl? S? '?>'`
  - `[24] VersionInfo ::= S 'version' Eq ("'" VersionNum "'" | '"' VersionNum '"')`
  - `[25] Eq ::= S? '=' S?`
  - `[26] VersionNum ::= '1.' [0-9]+`
  - `[27] Misc ::= Comment | PI | S`
  - `[32] SDDecl ::= S 'standalone' Eq (("'" ('yes' | 'no') "'") | ('"' ('yes' | 'no') '"'))`
  - `[80] EncodingDecl ::= S 'encoding' Eq ('"' EncName '"' | "'" EncName "'" )`
  - `[81] EncName ::= [A-Za-z] ([A-Za-z0-9._] | '-')*`

  XML documents SHOULD begin with an XML declaration. Documents SHOULD NOT specify a version other than '1.0'. A 1.0 processor processes a '1.x' document as 1.0.
- **Source:** XML §2.1 (`#NT-document`), §2.8 "Prolog and Document Type Declaration" (`#sec-prolog-dtd`, `#NT-prolog`, `#NT-XMLDecl`), §2.9 (`#NT-SDDecl`), §4.3.3 (`#NT-EncodingDecl`).
- **Anchor:** "[23] XMLDecl ::= '<?xml' VersionInfo EncodingDecl? SDDecl? S? '?>'"
- **Notes:** The pseudo-attributes have a fixed order (version, then encoding, then standalone), `version` is mandatory, and each value is case-sensitive (`'yes' | 'no'`). Exactly one root element, nested properly (§2.1). Only Comment, PI and S may follow the root (`Misc*`), so trailing text after the root element is a fatal error (*derived*).

### X-14 — `doctypedecl` production, and what "no DTD support" means for conformance
- **Fact:** `[28] doctypedecl ::= '<!DOCTYPE' S Name (S ExternalID)? S? ('[' intSubset ']' S?)? '>'`, with `[28b] intSubset ::= (markupdecl | DeclSep)*` and `[29] markupdecl ::= elementdecl | AttlistDecl | EntityDecl | NotationDecl | PI | Comment`. The doctype MUST appear before the first element.
- **Conformance:** §5.1 says non-validating processors are **REQUIRED** to check the document entity, **including the entire internal DTD subset**, for well-formedness. They are also REQUIRED to process the declarations they read in the internal subset (normalizing attributes, including internal-entity replacement text, supplying defaults).
- **Source:** XML §2.8 `#NT-doctypedecl`. §5.1 "Validating and Non-Validating Processors".
- **Anchor:** "Non-validating processors are REQUIRED to check only the document entity"
- **Notes:** **Surprising / design-relevant:** a parser that rejects any `<!DOCTYPE` is, strictly, **not** a conforming XML 1.0 processor. This is a deliberate, defensible deviation for security (it blocks XXE and entity expansion). It should be documented as such, and the error should name it ("DTD not supported"). It should not be described as "malformed XML".

### X-15 — §4.3.3 Character encoding and BOM
- **Fact:**
  - All processors MUST be able to read UTF-8 and UTF-16.
  - Entities encoded in UTF-16 MUST, and entities in UTF-8 MAY, begin with the Byte Order Mark (#xFEFF). The BOM is an encoding signature, not part of the markup or character data, and processors MUST be able to use it to tell UTF-8 from UTF-16.
  - Without external encoding information, it is a **fatal error** for an entity with an encoding declaration to be in a different encoding than the one named, or for an entity with neither a BOM nor an encoding declaration to be in anything other than UTF-8.
  - It is a fatal error if a UTF-8 entity contains ill-formed code unit sequences (Unicode §3.9).
  - Processors SHOULD match encoding names case-insensitively.
  - The terms "UTF-8" and "UTF-16" do not cover UTF-16BE, UTF-16LE or CESU-8.
  - In the document entity, the encoding declaration is part of the XML declaration.
- **Source:** XML §4.3.3 "Character Encoding in Entities", `#charencoding`. Appendix F.1 "Detection Without External Encoding Information" (`#sec-guessing`, **non-normative**) lists the BOMs: `EF BB BF` (UTF-8), `FE FF ## ##` (UTF-16 BE), `FF FE ## ##` (UTF-16 LE). It also lists the no-BOM UTF-8/ASCII signature `3C 3F 78 6D`.
- **Anchor:** "Entities encoded in UTF-16 MUST and entities encoded in UTF-8 MAY begin with the Byte Order Mark"
- **Notes:** A UTF-8-only parser should therefore: strip a leading `EF BB BF`; reject any `encoding=` other than UTF-8 (case-insensitive); reject invalid UTF-8 (overlongs, surrogates, > U+10FFFF) as fatal. It may reject UTF-16 or report it as unsupported. "Unable to process" an encoding is also a fatal error (§4.3.3). A TextDecl occurring other than at the beginning of an external entity is a fatal error (only relevant to external entities).

### X-16 — Definition of "fatal error"
- **Fact:** A fatal error MUST be detected and reported. After one, the processor MAY keep scanning for more errors, but MUST NOT continue normal processing, i.e. it MUST NOT keep passing character data and structure to the application in the normal way.
- **Source:** XML §1.2 Terminology, `#dt-fatal`.
- **Anchor:** "MUST NOT continue normal processing"
- **Notes:** This supports a fail-closed design: no partial results after a WF error.

---

## Namespaces in XML 1.0 (Third Edition, 8 Dec 2009)

### N-01 — Declaring namespaces: `xmlns` and `xmlns:prefix`
- **Fact:** A namespace binding is declared with a reserved attribute whose name is `xmlns` or begins with `xmlns:`. Productions:
  - `[1] NSAttName ::= PrefixedAttName | DefaultAttName`
  - `[2] PrefixedAttName ::= 'xmlns:' NCName` [NSC: Reserved Prefixes and Namespace Names]
  - `[3] DefaultAttName ::= 'xmlns'`
  - `[4] NCName ::= Name - (Char* ':' Char*)`

  The attribute's **normalized value** MUST be a URI reference (the namespace name) or an empty string. With PrefixedAttName the NCName is the prefix. With DefaultAttName the value is the default namespace, in the scope of the element carrying the declaration. The empty string, though a legal URI reference, cannot be a namespace name (§2.2). Relative URI references are deprecated.
- **Source:** NS §3 "Declaring Namespaces" (`#ns-decl`, `#NT-PrefixedAttName`, `#NT-DefaultAttName`, `#NT-NCName`). §2.2 "Use of URIs as Namespace Names".
- **Anchor:** "[2] PrefixedAttName ::= 'xmlns:' NCName"
- **Notes:** Declarations "may be provided directly or by default". With no DTD, only direct ones exist.

### N-02 — Reserved `xml` and `xmlns` prefixes
- **Fact (NSC: Reserved Prefixes and Namespace Names):**
  - The prefix **xml** is by definition bound to **http://www.w3.org/XML/1998/namespace**. It MAY but need not be declared, and MUST NOT be bound to any other namespace name. Other prefixes MUST NOT be bound to that namespace name, and it MUST NOT be declared as the default namespace.
  - The prefix **xmlns** is by definition bound to **http://www.w3.org/2000/xmlns/**. It MUST NOT be declared. Other prefixes MUST NOT be bound to that name, and it MUST NOT be the default namespace.
  - Element names MUST NOT have the prefix xmlns.
  - All other prefixes beginning with x-m-l in any case are reserved. Users SHOULD NOT use them, and processors **MUST NOT treat them as fatal errors**.
- **Source:** NS §3, `#xmlReserved`.
- **Anchor:** "The prefix xml is by definition bound to the namespace name"
- **Notes:** So `xmlns:xml="http://www.w3.org/XML/1998/namespace"` is allowed. `xmlns:xml="…anything else…"`, `xmlns:xmlns="…"`, `xmlns:foo="http://www.w3.org/XML/1998/namespace"`, `xmlns="http://www.w3.org/2000/xmlns/"` and `<xmlns:a/>` are all errors. Prefixes like `xmlfoo:` must be accepted.

### N-03 — NSC: Prefix Declared
- **Fact:** A namespace prefix other than `xml` or `xmlns` MUST have been declared in a namespace declaration attribute, either in the start-tag of the element where it is used or in an ancestor element. The constraint is attached to element QNames in `[12] STag`, `[13] ETag` and `[14] EmptyElemTag`, and to attribute QNames in `[15] Attribute`.
- **Source:** NS §5 "Using Qualified Names", `#nsc-NSDeclared`. QName productions are in §4: `[7] QName ::= PrefixedName | UnprefixedName`, `[8] PrefixedName ::= Prefix ':' LocalPart`, `[10] Prefix ::= NCName`, `[11] LocalPart ::= NCName`.
- **Anchor:** "Namespace constraint: Prefix Declared"
- **Notes:** Because the scope begins at the start of the start-tag (N-05), a declaration on the same tag counts for that element's own name and attributes. Collect all `xmlns*` attributes on a tag before resolving any QName on it (*derived*).

### N-04 — NSC: No Prefix Undeclaring (1.0)
- **Fact:** In a namespace declaration for a prefix (NSAttName is a PrefixedAttName), the attribute value MUST NOT be empty. By contrast, the value of a **default** namespace declaration MAY be empty, which has the same effect within its scope as having no default namespace (§6.2).
- **Source:** NS §5, `#nsc-NoPrefixUndecl`. NS §6.2 "Namespace Defaulting", `#defaulting`.
- **Anchor:** "the attribute value MUST NOT be empty"
- **Notes:** `xmlns:p=""` is an error in Namespaces 1.0. `xmlns=""` is legal and resets the default namespace (§6.2 gives the `<brandName xmlns="">` example).

### N-05 — Scoping
- **Fact:** The scope of a prefix declaration runs from the beginning of the start-tag where it appears to the end of the matching end-tag, excluding the scope of inner declarations with the same NSAttName. For an empty tag, the scope is the tag itself. A prefixed name's expanded name is (URI bound to the prefix, local part).
- **Source:** NS §6.1 "Namespace Scoping", `#scoping`.
- **Anchor:** "In the case of an empty tag, the scope is the tag itself."

### N-06 — The default namespace does not apply to attributes
- **Fact:** A default namespace declaration applies to all unprefixed **element** names in its scope. Default declarations do not apply directly to attribute names. **The namespace name for an unprefixed attribute name always has no value.** With no default declaration in scope, an unprefixed element's namespace name has no value.
- **Source:** NS §6.2 "Namespace Defaulting", `#defaulting`.
- **Anchor:** "The namespace name for an unprefixed attribute name always has no value."
- **Notes:** In XLSX, `r:id` lives in the relationships namespace, while unprefixed `ref`, `t`, `s` and so on are in no namespace, even though the element is in the SpreadsheetML default namespace (*application of the rule*).

### N-07 — NSC: Attributes Unique (expanded names)
- **Fact:** No tag may contain two attributes that (a) have identical names, or (b) have qualified names with the same local part and prefixes bound to identical namespace names. This is equivalent to requiring that no element have two attributes with the same **expanded name**. The spec's example: with `xmlns:n1` and `xmlns:n2` both bound to `http://www.w3.org`, `<bad n1:a="1" n2:a="2"/>` is illegal. With n1 bound to the same URI as the default namespace, `<good a="1" n1:a="2"/>` is legal, because the default namespace does not apply to attributes.
- **Source:** NS §6.3 "Uniqueness of Attributes", `#uniqAttrs`, `#nsc-AttrsUnique`.
- **Anchor:** "no element have two attributes with the same expanded name"
- **Notes:** This must be checked **after** prefix resolution, on top of XML's lexical Unique Att Spec (X-08).

### N-08 — Comparing namespace names
- **Fact:** Namespace URIs are compared as strings: identical only if they are the same sequence of characters. The comparison is case-sensitive, with no %-escaping done or undone. In a declaration, the URI is the attribute's **normalized value**, so character and entity references are replaced before comparison.
- **Source:** NS §2.3 "Comparing URI References".
- **Anchor:** "The comparison is case-sensitive, and no %-escaping is done or undone."
- **Notes:** So `xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` written with `&#x2F;` instead of `/` is the **same** namespace (after normalization). One differing only in case is a **different** namespace (*derived*).

### N-09 — Namespace well-formedness and processor conformance
- **Fact:** In a namespace-well-formed document, every element and attribute name matches QName, so it has zero or one colon. All other tokens that XML 1.0 requires to match Name must match NCName, so entity names, PI targets and notation names contain no colons (§7). A conforming processor MUST report namespace well-formedness violations, but is not required to check that namespace names are valid URI references (§8).
- **Source:** NS §7 "Conformance of Documents" (`#Conformance`), §8 "Conformance of Processors" (`#ProcessorConformance`).
- **Anchor:** "No entity names, processing instruction targets, or notation names contain any colons."
- **Notes:** A PI target with a colon is legal XML 1.0 but **not** namespace-well-formed. Also reject `a:b:c`, `:a` and `a:` names.

---

## Web sources

### W-01 — [MS-CFB] Compound File header signature
- **Fact:** The Compound File Header MUST be at the beginning of the file (offset 0). Its first field, **Header Signature (8 bytes)**, MUST be `0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1`. Other header fields: Header CLSID (16 bytes, all zeros); Minor Version (SHOULD be 0x003E); Major Version (MUST be 0x0003 or 0x0004); Byte Order (MUST be 0xFFFE, meaning little-endian); Sector Shift (MUST be 0x0009 for v3, giving 512-byte sectors, or 0x000C for v4).
- **Source:** [MS-CFB] §2.2 "Compound File Header" (Last updated 2021-06-24), https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/05060311-bfce-4b12-874d-71fd4ce63aea
- **Anchor:** "0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1"
- **Notes:** This is the magic for detecting encrypted OOXML (W-02) and legacy .xls. A ZIP-based .xlsx instead starts with `50 4B 03 04` (Z-01, Z-07).

### W-02 — [MS-OFFCRYPTO] How an encrypted OOXML document is stored
- **Fact:** Encrypted ECMA-376 documents use the data spaces functionality to hold the **entire document as a single stream in an OLE compound file** (§1.3.3.4). When an ECMA-376 document is encrypted (per ECMA-376 Part 2 Annex C Table C-5 BIT 0), a structured storage using the data spaces construct (§2.1) MUST be used. The streams and storages in `\0x06DataSpaces` MUST be present as in §2.1.1, unless noted otherwise (§2.3.4). Paths beginning with a backslash refer to the **root storage** of the compound file (§1.3.8), so `\EncryptedPackage` and `\EncryptionInfo` are streams in the root storage.
- **Source:** [MS-OFFCRYPTO] §1.3.3.4 "ECMA-376 Document Encryption" (Last updated 2020-11-17), https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/2995dc93-0564-468c-891a-d950464479fb. §2.3.4 "ECMA-376 Document Encryption" (Last updated 2026-02-17), …/4f2e28a1-624a-4bed-b3ec-816f0df484ff. §1.3.8 "OLE Compound File Path Encoding" (2020-11-17), …/20461932-810a-468d-827f-08c68c98b76c.
- **Anchor:** "contain the entire document as a single stream (1) in an OLE compound file"
- **Notes:** A password-protected .xlsx is therefore **not a ZIP**. A reader should check for the CFB signature first and give a specific "encrypted workbook" error, not "corrupt ZIP".

### W-03 — [MS-OFFCRYPTO] `\EncryptedPackage` and `\EncryptionInfo` streams
- **Fact:**
  - `\EncryptedPackage` is an encrypted stream of bytes holding the **entire ECMA-376 source file in compressed form** (the ZIP package). Its layout is `StreamSize` (8 bytes, unsigned: the number of bytes of encrypted data, excluding the StreamSize field) followed by `EncryptedData` (variable, encrypted with the algorithm named in `\EncryptionInfo`). The actual stream can be larger than StreamSize because of the cipher's block size (§2.3.4.4).
  - `\EncryptionInfo` holds the information used to initialize the cryptography for `\EncryptedPackage`.
  - **Standard encryption** (§2.3.4.5): `EncryptionVersionInfo.vMajor` MUST be 0x0002, 0x0003 or 0x0004, and vMinor MUST be 0x0002.
  - **Agile encryption** (§2.3.4.10): vMajor MUST be 0x0004 and vMinor MUST be 0x0004, followed by Reserved = 0x00000040.
  - **Extensible encryption** is §2.3.4.6.
- **Source:** [MS-OFFCRYPTO] §2.3.4.4 "\EncryptedPackage Stream" (2020-11-17), https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/b60c8b35-2db2-4409-8710-59d88a793f83. §2.3.4.5 (2023-05-16), …/2895eba1-acb1-4624-9bde-2cdad3fea015. §2.3.4.10 (2023-05-16), …/87020a34-e73f-4139-99bc-bbdf6cf6fa55.
- **Anchor:** "StreamSize (8 bytes)" · "Version.vMajor MUST be 0x0004 and Version.vMinor MUST be 0x0004"
- **Notes:** A reader that only detects (does not decrypt) can open the CFB, find the root entries `EncryptionInfo` and `EncryptedPackage`, read the first 4 bytes of EncryptionInfo, and report "standard" or "agile" encryption.

### W-04 — Excel specifications and limits
- **Fact:** Total rows and columns on a worksheet: **1,048,576 rows by 16,384 columns**. Total number of characters a cell can contain: **32,767 characters**. The same table gives: maximum line feeds per cell 253; column width 255 characters; row height 409 points; characters in a header or footer 255; sheets in a workbook "Limited by available memory".
- **Source:** Microsoft Support, "Excel specifications and limits", table "Worksheet and workbook specifications and limits" (Applies To: Excel for Microsoft 365, 2024, 2021, 2019, 2016; ms.date 04/13/2026), https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
- **Anchor:** "1,048,576 rows by 16,384 columns" · "32,767 characters"
- **Notes:** 16,384 columns is A…XFD (*derived*: XFD = 24·26² + 6·26 + 4 = 16,384). These are application limits, not file-format limits, but they make sensible hard caps for a defensive reader: reject cell refs beyond row 1,048,576 or column 16,384, and decide what to do with strings longer than 32,767.

### W-05 — Node.js `require()` of ES modules
- **Fact:** History table for "Loading ECMAScript modules using require()":
  - Added in **v22.0.0, v20.17.0**.
  - **v23.0.0, v22.12.0, v20.19.0:** "This feature is no longer behind the `--experimental-require-module` CLI flag" (i.e. **enabled by default from 22.12.0 on the 22.x line and from 20.19.0 on the 20.x line**).
  - **v23.5.0, v22.13.0, v20.19.0:** no experimental warning by default (`--trace-require-module` can still emit it).
  - **v23.0.0, v22.12.0:** `'module.exports'` interop export supported.
  - **v25.4.0:** no longer experimental.

  Requirements: the module is fully synchronous (no top-level await), **and** either it is `.mjs`, or it is `.js` with the nearest package.json containing `"type": "module"`, or it is `.js` whose nearest package.json does not contain `"type": "commonjs"` and which contains ES module syntax. **If the require()'d module, or the module graph it imports, contains top-level await, `ERR_REQUIRE_ASYNC_MODULE` is thrown.** `--no-require-module` disables the feature. `process.features.require_module` detects it. `--experimental-print-required-tla` prints the TLA locations.
- **Source:** Node.js docs (v26.13.0), "Modules: CommonJS modules" → "Loading ECMAScript modules using require()", https://nodejs.org/api/modules.html#loading-ecmascript-modules-using-require
- **Anchor:** "ERR_REQUIRE_ASYNC_MODULE will be thrown" · "v23.0.0, v22.12.0, v20.19.0"
- **Notes:** The result is the module namespace object, so the default export is at `.default`. `__esModule: true` is added when there is a default export. An ESM-only library can export `{ X as 'module.exports' }` to give CJS callers `require()` returning X directly, but named exports are then lost unless attached to X. To support `require()` the library graph must contain **no top-level await anywhere**.

### W-06 — Fifield: the overlapping-files technique
- **Fact:**
  - The bomb is **non-recursive**: it expands fully after one round of decompression, by **overlapping files inside the zip container**. Many central directory headers point at a single file whose data is a highly compressed "kernel".
  - The full-overlap version (all CDHs pointing to one LFH) is incompatible with some parsers because the filename cannot match for every file. UnZip warns about a "mismatching 'local' filename", and Python `zipfile` raises `BadZipFile`.
  - The improved version **quotes** each following local file header inside a DEFLATE **non-compressed block** (5-byte header) of the previous file's stream. Each LFH is then read both as structure and as file data, all streams end in the shared kernel, and output grows **quadratically** with input. The article's figures: zbsm.zip 42 kB → 5.5 GB, zblg.zip 10 MB → 281 TB, a ratio of over 28 million.
  - An **extra-field quoting** variant grows a local header's extra-field length so that it swallows the following LFHs. It works independently of the compression algorithm (e.g. with bzip2). Zip64 allows even greater expansion.
- **Source:** David Fifield, "A better zip bomb", sections "Summary", "The first insight: overlapping files", "The second insight: quoting local file headers", "Extension: extra-field quoting", https://www.bamsoftware.com/hacks/zipbomb/ (2019-07-02; updates to 2023-05-18; version at USENIX WOOT 2019).
- **Anchor:** "overlapping files inside the zip container" · "The first insight: overlapping files"
- **Notes:** A file-name mismatch between LFH and CDH is a cheap red flag. It is not sufficient, because quoted overlap keeps names consistent.

### W-07 — Fifield: recommended mitigations
- **Fact:**
  - Detecting this class of bomb is easy: **look for overlapping files**. Mark Adler wrote such a patch for Info-ZIP UnZip, which went into Debian (bug #931433) and caused some breakage with Java JARs and Firefox `omni.ja`. CVE-2019-13232 was assigned to UnZip.
  - Rejecting overlapping files does **not** by itself make untrusted zips safe: other bombs and other malicious zips exist.
  - Detection logic **must be inside the parser itself, not a separate prefilter**, because there is no single well-defined algorithm for locating the central directory, so two parsers may disagree on the contents.
  - **Summing the declared sizes does not work** to predict total output, because metadata sizes may not match actual uncompressed sizes.
  - The author's general advice: do not try to enumerate and block every known attack. **Impose external limits on time and other resources** and isolate operations on untrusted data.
  - He cites `ronomon/zip`, a restrictive JavaScript parser that **prohibits overlapping files and unused space between files**.
- **Source:** Same article, the discussion after the constructions ("Detecting the specific class of zip bomb…") and the "UnZip 6.0" and "ronomon/zip" notes in the compatibility/updates section, https://www.bamsoftware.com/hacks/zipbomb/
- **Anchor:** "look for overlapping files" · "not as a separate prefilter"
- **Notes:** Concrete reader rules that follow (*derived*):
  1. For each entry, compute the byte span `[LFH offset, LFH offset + 30 + nameLen + extraLen + compressedSize (+ data descriptor))`. Sort the spans and reject any overlap, including overlap with the central directory and EOCD.
  2. Require the LFH name to equal the CDH name.
  3. Enforce limits on actual inflated bytes while streaming, per entry and in total, using counted output rather than declared sizes.
  4. Also cap entry count and compression ratio.

### W-08 — [MS-OFFMACRO2] VBA project part: relationship type and content type
- **Fact:** VBA Project part: **Content Type `application/vnd.ms-office.vbaProject`**, Root Namespace not applicable, **Source Relationship `http://schemas.microsoft.com/office/2006/relationships/vbaProject`**. The part is specified in [MS-OVBA]. A package MUST contain at most one VBA Project part, and it MUST be the target of an **implicit** relationship from the Workbook part. A VBA Project part MUST NOT have implicit or explicit relationships to any other part in that document or in ISO/IEC 29500-1:2016.
- **Source:** [MS-OFFMACRO2] §2.2.1.3 "VBA Project" (Macro-enabled SpreadsheetML; Last updated 2019-02-14), https://learn.microsoft.com/en-us/openspecs/office_standards/ms-offmacro2/1527aa6b-c091-435f-86b3-1413abe8f8f4
- **Anchor:** "Source Relationship: http://schemas.microsoft.com/office/2006/relationships/vbaProject"
- **Notes:** To detect macros, check both the rel type in `xl/_rels/workbook.xml.rels` and the content type in `[Content_Types].xml`. Treat either one as "has VBA" (*recommendation*).
**Checked 2026-10-30 (D-1104):** [MS-OFFMACRO2] rev. 8.0 (2024-08-20) §2.2.1.3, p.15–16 — confirmed from the PDF.

### W-09 — [MS-OFFMACRO2] Macro-enabled workbook content type
- **Fact:** Workbook part (macro-enabled): Content Types **`application/vnd.ms-excel.sheet.macroEnabled.main+xml`** and **`application/vnd.ms-excel.template.macroEnabled.main+xml`**. Root Namespace `http://schemas.openxmlformats.org/spreadsheetml/2006/main`. Source Relationship `http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument`. The part is modified from ISO/IEC 29500-1:2016 Part 1 §12.3.23. It may have an **implicit** relationship to the VBA Project, Macro Sheet and International Macro Sheet parts.
- **Source:** [MS-OFFMACRO2] §2.2.1.4 "Workbook" (Last updated 2021-08-17), https://learn.microsoft.com/en-us/openspecs/office_standards/ms-offmacro2/fa1f6007-088f-4f54-ba65-83b5a9d635a6
- **Anchor:** "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
- **Notes:** The root namespace is the same as a normal workbook, so only the content type (and the presence of VBA or macrosheet parts) tells .xlsm from .xlsx.
**Checked 2026-10-30 (D-1104):** [MS-OFFMACRO2] rev. 8.0 §2.2.1.4, p.16 — confirmed from the PDF.

### W-10 — [MS-OFFMACRO2] Excel 4.0 (XLM) macro sheet parts
- **Fact:**
  - Macro Sheet: Content Type `application/vnd.ms-excel.macrosheet+xml`, Root Namespace `http://schemas.microsoft.com/office/excel/2006/main`, Source Relationship `http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet`. It MUST be the target of an **explicit** relationship from the Workbook part, and its root element MUST be `macrosheet` (§2.2.1.2).
  - International Macro Sheet: Content Type `application/vnd.ms-excel.intlmacrosheet+xml`, Source Relationship `http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet`, same root namespace (§2.2.1.1).
- **Source:** [MS-OFFMACRO2] §2.2.1.2 "Macro Sheet", https://learn.microsoft.com/en-us/openspecs/office_standards/ms-offmacro2/1de8055f-535b-4ebc-ba18-8d25e2986c89. §2.2.1.1 "International Macro Sheet" (Last updated 2019-02-14), …/5623d7b1-eb7c-4347-9bb8-cdbfa1b79e19.
- **Anchor:** "application/vnd.ms-excel.macrosheet+xml"
- **Notes:** **Inconsistency in the source:** §2.2.1.4 (Workbook) says the Workbook may have an *implicit* relationship to Macro Sheet parts, while §2.2.1.2 (Macro Sheet) says the macro sheet MUST be the target of an *explicit* relationship from the Workbook. XLM macros carry no VBA part, so a "macro present" check that looks only for vbaProject misses them.
**Checked 2026-10-30 (D-1104):** [MS-OFFMACRO2] rev. 8.0 §2.2.1.1–§2.2.1.2, p.14–15 — confirmed from the PDF (the explicit/implicit relationship inconsistency is in the PDF too).

### W-11 — RFC 1952 CRC-32 (supplementary, to fill the APPNOTE gap)
- **Fact:** The gzip CRC32 field holds a CRC of the **uncompressed data**, computed with the CRC-32 algorithm of ISO 3309 and ITU-T V.42 §8.1.1.6.2. The §8 sample code builds its table with `c = 0xedb88320L ^ (c >> 1)` and applies pre- and post-conditioning with all ones.
- **Source:** RFC 1952 §2.3.1 (member header field "CRC32") and §8 "Appendix: Sample CRC Code", https://www.rfc-editor.org/rfc/rfc1952.txt
- **Anchor:** "uncompressed data computed according to CRC-32 algorithm"
- **Notes:** This is cited only because APPNOTE §4.4.7 does not name the polynomial or the covered bytes (Z-17). That ZIP's CRC is the same algorithm, over uncompressed data, is an inference supported by APPNOTE's pre/post-conditioning description, not a sentence in APPNOTE.

---

## NOT FOUND (searched, not present in the named source)
1. **APPNOTE: CRC-32 generator polynomial, and which bytes the CRC covers.** §4.4.7 gives only the "magic number" 0xdebb20e3 and the pre/post-conditioning; I checked all "CRC" mentions. Filled from RFC 1952 (W-11).
2. **APPNOTE: algorithm for locating the EOCD** (backward scan, how to handle a comment containing the signature). Not specified. FIFIELD confirms there is no single well-defined algorithm.
3. **APPNOTE: disk numbering base (0 or 1).** Not stated.
4. **APPNOTE: path traversal (`..`) rules.** Not mentioned. Only "no drive letter, no leading slash, forward slashes" (§4.4.17.1).
5. **APPNOTE: CP437 mapping table and MS-DOS date/time bit layout.** Not given (only "IBM Code Page 437" and "standard MS-DOS format").
6. **APPNOTE: explicit statement that the Zip64 locator immediately precedes the EOCD with no gap.** Implied only by the §4.3.6 ordering.
