# Facts about other spreadsheet formats and the programs that write them

Facts behind the format plan in docs/PLAN.md (phase N) and docs/decisions.md D-N01 and D-N02. Each fact names its source and the date it was checked. All web sources were read through a web search on 2026-11-05 from the build environment, which cannot open most of these sites directly; facts marked PROVISIONAL rest on secondary sources (encyclopedias, file-format sites, vendor blogs) and are checked against a primary source at the start of the phase that uses them.

## Programs: paid or free

### FMT-01 Apple Numbers is free
The Mac App Store listing of Numbers (id409203825) gives the price as Free; Macworld reported that Apple made Pages, Numbers, Keynote, GarageBand and iMovie free for Mac and iOS (the Mac version of Numbers had cost USD 20). Sources: https://apps.apple.com/ca/app/id409203825?l=en-ca, https://www.macworld.com/article/229911/garageband-imovie-keynote-pages-and-numbers-are-now-free-for-mac-and-ios.html.

### FMT-02 Gnumeric is free software
Gnumeric is licensed under the GPL, version 2 or 3 (man page). Source: https://www.mankier.com/1/gnumeric.

### FMT-03 Quattro Pro is sold as part of WordPerfect Office
Corel WordPerfect Office 2021 (Standard, Professional, Home & Student) is sold by retailers as a paid product and includes the Quattro Pro spreadsheet. PROVISIONAL: retailer listings only, no corel.com page was reached. Sources: https://www.staples.com/corel-wordperfect-office-standard-2021-for-1-user-windows-download-esdwp2021stdef/product_24577266, https://au.finance.yahoo.com/news/wordperfect-office-2021-boosts-productivity-130100070.html.

### FMT-04 Lotus 1-2-3 is no longer sold
IBM announced the withdrawal of Lotus 1-2-3, Organizer and SmartSuite in May 2013, ended marketing in June 2013 and support in 2014 (September 30 or early October, sources differ). It was a paid product. PROVISIONAL: news reports, IBM's own notice was not reached. Sources: https://www.theregister.com/2014/10/02/so_long_lotus_123_ibm_ceases_support_after_over_30_years_of_code/, https://tech.slashdot.org/story/13/05/21/1240247/goodbye-lotus-1-2-3.

### FMT-05 Microsoft Works is no longer sold
Microsoft Works was sold from 1987 to 2009 (some sources say 2007; version 9.0 is dated 2007-09-28). Its old spreadsheet format is `.wks` (PRONOM fmt/166 lists Works Spreadsheet versions 1-5); later Works spreadsheets use the Excel format under the `.xlr` extension. PROVISIONAL. Sources: https://en.wikipedia.org/wiki/Microsoft_Works, https://www.nationalarchives.gov.uk/PRONOM/fmt/166, https://docs.fileformat.com/spreadsheet/xlr.

### FMT-06 SoftMaker PlanMaker has a paid and a free edition
PlanMaker is part of SoftMaker Office (paid) and of FreeOffice (free, reduced). Its own format is `.pmdx`; the specification is not public. No source says whether the free edition saves `.pmdx`. PROVISIONAL. Sources: https://docs.fileformat.com/spreadsheet/pmdx, https://help.softmaker.com/planmaker2024/en/working_with_other_file_formats.html.

### FMT-07 WPS Spreadsheets has a free and a premium edition
The free edition opens and edits `.et` (the native format) and converts it; Premium removes ads and unlocks extra features. PROVISIONAL: vendor pages and reviews. Sources: https://www.wps.com/blog/xls-vs-xlsx-vs-et/, https://www.wps.com/feature/et-file/.

### FMT-08 Microsoft Excel is a paid program
Excel is sold as part of Microsoft 365 and Office. Not looked up again for this file: xlvane's own scope (D-N01) treats Excel as the paid program whose OOXML formats the free core already reads and writes.

## Formats

### FMT-10 OpenDocument spreadsheets
OpenDocument is an OASIS standard and ISO/IEC 26300; ODF 1.3 was approved in 2021. `.ods` is the packaged (ZIP) form, `.fods` the single flat XML file; the standard itself names no extension for the flat form. PROVISIONAL until read in the OASIS ODF 1.3 specification. Sources: https://en.wikipedia.org/wiki/OpenDocument, https://en.wikipedia.org/wiki/OpenDocument_technical_specification.

### FMT-11 Gnumeric files
The native format is XML compressed with gzip; extensions `.gnumeric` and `.gnm`; an XSD schema exists. Source: https://help.gnome.org/users/gnumeric/stable/sect-file-formats.html.en (GNOME help), https://en.wikipedia.org/wiki/Gnumeric.

### FMT-12 Apple Numbers files
A `.numbers` file is a ZIP archive of `.iwa` entries (entry point `Index/Document.iwa`); each IWA block holds Snappy-compressed protobuf messages (Apple's Snappy framing without the CRC32 checksums). Apple publishes no specification; the schemas are taken from the iWork applications by reverse engineering, so an update of Numbers can change them. PROVISIONAL: community sources only. Sources: https://pypi.org/project/numbers-parser/3.1.0, https://andrews.substack.com/p/reverse-engineering-iwork, https://docs.rs/crate/kreuzberg/4.9.7/source/src/extractors/iwork/mod.rs.

### FMT-13 Excel binary formats
[MS-XLS] (`.xls`, BIFF8) and [MS-XLSB] (`.xlsb`) are published by Microsoft under the Open Specifications Promise, a covenant not to sue. PROVISIONAL: listed by secondary sources (SheetJS references); Microsoft's own list was not reached. Source: https://docs.sheetjs.com/docs/miscellany/references. xlvane already recognises the Compound File Binary container of `.xls` and of encrypted OOXML files ([MS-CFB], [MS-OFFCRYPTO], docs/facts/zip-xml-web.md) and rejects it with `XLV_E_UNSUPPORTED_CONTAINER`.

### FMT-14 Lotus worksheet format documentation
The Lotus Worksheet File Format documentation was dedicated to the public domain by Lotus in 1984. PROVISIONAL. Source: https://docs.sheetjs.com/docs/miscellany/references.

### FMT-15 LibreOffice import test files
The license of the files in LibreOffice's `sc/qa/unit/data` could not be established: LibreOffice build files carry MPL-2.0 headers, and the packaged data is under several licenses. Not used as a corpus until the license of each file is known (tests/corpus/LICENSES.md). Sources: https://cgit.freedesktop.org/libreoffice/core/tree/sc/CppunitTest_sc_vba_macro_test.mk, https://packages.fedoraproject.org/pkgs/libreoffice/libreoffice-data/fedora-38.html.

## tabularjs 2.2.0

### TABJS-01 What the package is
Read from the npm tarball `tabularjs@2.2.0` (`npm pack`, 2026-11-05): `README.md`, `CHANGELOG.md`, `package.json`, `src/index.d.ts`, `src/parser/`, `src/utils/zip-utils.js`. MIT license; author "Jspreadsheet Team"; runtime dependencies `jszip` ^3.10.1 and `@lemonadejs/html-to-json` ^1.0.0; first version 1.0.0 on 2025-11-28 (`npm view tabularjs time`). One default export, `tabularjs(file, options)`, that resolves to `{ worksheets[], style?, definedNames? }` with `data: any[][]` per worksheet ("Jspreadsheet Pro format" in the TSDoc). It only reads. Formats in the README: xls, xlsx, ods, Lotus (wks, wk1, wk3, wk4, 123), csv, tsv, txt, XML Spreadsheet 2003, DIF, SYLK, dBase (dbf), HTML tables. Each ZIP entry is limited to 256 MiB uncompressed (`DEFAULT_MAX_ENTRY_BYTES`).

### TABJS-02 Features named in its CHANGELOG (2.1.0 to 2.2.0)
- A file whose content belongs to another format than its extension says (an `.xlsx` renamed `.xls`, CSV or HTML saved as `.xls`) is read as what it really is.
- A password-protected `.xlsx` is reported as "Workbook is encrypted with a password" instead of a ZIP error.
- `.xls` files protected with RC4 and Excel's default password ("write-protected" or "read-only recommended" files) are decrypted; other encryption raises "Workbook is encrypted".
- CSV `options.encoding` accepts any TextDecoder label; UTF-16 files with a BOM are decoded.
- Tested against LibreOffice's import test files and openpyxl's formula translation.

Things it does that xlvane already does: shared formulas expanded to every cell (`Formula.text`, `XLV_W_SHARED_FORMULA_NOT_EXPANDED`), inline strings, phonetic runs left out of the text (D-SST-2, packages/xlvane/src/sml/shared-strings.ts), ZIP size limits.
