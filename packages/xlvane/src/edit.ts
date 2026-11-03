// Surgical editing of existing XLSX files (docs/decisions.md D-309 … D-312).
//
// Only what an edit needs is rewritten: the `sheetData` of edited sheets (untouched rows and
// cells are copied as source text, so unknown attributes and extensions stay), records appended
// to the shared string table and the styles part, and – when formulas change – the calculation
// chain and the workbook's `calcPr`. Every other part keeps its exact decompressed content.
// Offsets come from the XML parser's tag spans and refer to the text after end-of-line
// normalisation (XmlSpan).
import { XlvError, type XlvWarning } from './errors';
import { resolveLimits, type ReadLimits } from './limits';
import type { Cell, RichTextRun, Sheet, Workbook } from './model';
import { openPackage, relationshipsPartName, type OpcPackage } from './opc/package';
import { partKey, zipNameToPartName } from './opc/part-name';
import { readWorkbook, styledCellCount, toBytes } from './read';
import { formatCellRef, formatRange, MAX_COLUMNS, MAX_ROWS, parseCellRef, parseRange, type CellRange } from './sml/cell-ref';
import { classifyRelationship, relationshipTypes, SML_NS } from './sml/namespaces';
import { isDateFormat } from './sml/numfmt';
import { parseSharedStrings } from './sml/shared-strings';
import { EMPTY_STYLES, parseStyles, type Styles } from './sml/styles';
import { attr, borderXml, DEFAULT_FONT, fillXml, fontXml, formatNumber, reusableBuiltinNumFmt, stringItemXml, textElement } from './sml/write-parts';
import { checkStyleCount, mergeStyles, normalizeStyle, READING_ORDERS, stableKey, type CellStyle, type NormalizedStyle, type StyleCountLimit } from './style';
import { pickSheet, readTable } from './table';
import { readOptions } from './util/options';
import { RangeIndex } from './util/range-index';
import { suggest } from './util/suggest';
import { WarningCollector } from './util/warnings';
import { decodeXml, normalizeLineEnds, parseXml, type XmlSpan } from './xml/parser';
import { childNamed, childrenNamed, parseXmlTree, type XmlElement } from './xml/tree';
import { encodeXstring, escapeAttr, escapeText } from './xml/writer';
import { createWorkbook, type CellInput, type StoredValue } from './write';
import { openZip, type ZipEntry, type ZipReader } from './zip/reader';
import { writeZip, type ZipWriteEntry } from './zip/writer';

/**
 * One change for {@link editCells}.
 *
 * @example
 * ```ts
 * import type { CellEdit } from 'xlvane';
 *
 * const edit: CellEdit = { sheet: 'Sales', cell: 'B7', value: 1250.5, style: { fill: { color: 'FFFF00' } } };
 * ```
 */
export interface CellEdit {
  /** Sheet name (exact). */
  readonly sheet: string;
  /** A1 reference. */
  readonly cell: string;
  /** New value, as for `NewSheet.set`; `null` removes the value. */
  readonly value: CellInput;
  /** Style changes merged into the cell's current style. Default: the style stays. */
  readonly style?: CellStyle;
}

/**
 * Options for {@link editCells} and {@link appendRows}.
 *
 * @example
 * ```ts
 * import type { EditOptions } from 'xlvane';
 *
 * const options: EditOptions = { limits: { maxCells: 10_000_000 } };
 * ```
 */
export interface EditOptions {
  /**
   * Read limits for the input file; see `ReadLimits`. For `editCells` and `appendRows`,
   * `maxCells` also bounds the edit: the cells of the edited sheets plus the new cells must not
   * exceed it, and one call writes at most 1 000 000 cells unless `maxCells` is given, in which
   * case it writes at most `maxCells` (`XLV_E_LIMIT_EXCEEDED`, `details.limit` `"maxCells"`).
   */
  readonly limits?: Partial<ReadLimits>;
}

/**
 * Options for {@link appendRows}.
 *
 * @example
 * ```ts
 * import type { AppendRowsOptions } from 'xlvane';
 *
 * const options: AppendRowsOptions = { header: 1 };
 * ```
 */
export interface AppendRowsOptions extends EditOptions {
  /** Header row used to place object rows: `"auto"` (default) or a 1-based row number. */
  readonly header?: 'auto' | number;
}

/**
 * The result of an edit: the new file and warnings about it.
 *
 * @example
 * ```ts
 * import { editCells, type EditResult } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const result: EditResult = editCells(bytes, [{ sheet: 'Data', cell: 'A1', value: 'Total' }]);
 * console.log(result.bytes.length, result.warnings.map((w) => w.code));
 * ```
 */
export interface EditResult {
  /** The edited XLSX file. */
  readonly bytes: Uint8Array;
  /** For example `XLV_W_FORMULAS_NOT_RECALCULATED` when the workbook contains formulas. */
  readonly warnings: XlvWarning[];
}

/**
 * How many cells one call of {@link editCells} or {@link appendRows} writes at most when the
 * caller gives no `limits.maxCells` (A2). While the file is rebuilt each cell written costs about
 * 0.5 KB of memory, so the 5 000 000 of `DEFAULT_READ_LIMITS.maxCells` would need about 2.5 GB,
 * more than the default heap of Node.js (about 2 GB on a 64-bit machine with 16 GB of memory).
 */
export const EDIT_MAX_CELLS = 1_000_000;

/** The read limits of an edit, and how many cells it may write. */
function editLimits(input: unknown): { limits: ReadLimits; writable: number } {
  const limits = resolveLimits(input);
  const own = typeof input === 'object' && input !== null && Object.hasOwn(input, 'maxCells');
  return { limits, writable: own ? limits.maxCells : Math.min(limits.maxCells, EDIT_MAX_CELLS) };
}

/** Refuses an edit that would write more cells than one call may (A2). */
function checkWrites(count: number, writable: number, what: string, sheet?: string): void {
  if (count <= writable) return;
  throw new XlvError('XLV_E_LIMIT_EXCEEDED', `${what} would write ${String(count)} cells, more than the limit of ${String(writable)} cells per call.`, {
    location: sheet === undefined ? {} : { sheet },
    hint: `Each cell written takes about 0.5 KB of memory while the file is rebuilt. Split the cells into several calls of at most ${String(writable)}, or pass { limits: { maxCells: ${String(count)} } } if there is enough memory.`,
    details: { limit: 'maxCells', value: writable },
  });
}

/**
 * Refuses an edit after which the workbook would hold more cells (with a value or only a style)
 * than `maxCells` (A2), the limit `readWorkbook` applies to the whole workbook, so the result can be
 * read again with the same limits (D-K04). Checked before anything per cell is built.
 */
function checkCellCount(workbook: Workbook, added: ReadonlyMap<Sheet, number>, maxCells: number): void {
  let total = 0;
  for (const sheet of workbook.sheets) total += sheet.cellCount + styledCellCount(sheet) + (added.get(sheet) ?? 0);
  if (total <= maxCells) return;
  const edited = [...added.keys()];
  throw new XlvError('XLV_E_LIMIT_EXCEEDED', `After the edit, the workbook would hold ${String(total)} cells, more than the limit of ${String(maxCells)}.`, {
    location: edited.length === 1 ? { sheet: edited[0]!.name } : {},
    hint: `Cells of every sheet count, also cells that hold only a style. Write fewer new cells, or, if you trust this file and there is enough memory, raise the limit, for example { limits: { maxCells: ${String(Math.max(total, maxCells * 2))} } }.`,
    details: { limit: 'maxCells', value: maxCells },
  });
}

/** Data properties of an edit, read without running getters (planEdits validates edits). */
function dataProperty(o: object, key: string): unknown {
  return Object.getOwnPropertyDescriptor(o, key)?.value;
}

/** Number of cells per sheet that edits write where the sheet has no cell with a value. */
function newCells(workbook: Workbook, edits: readonly unknown[]): Map<Sheet, number> {
  const byName = new Map(workbook.sheets.map((sheet) => [sheet.name, sheet]));
  const seen = new Map<Sheet, Set<number>>();
  for (const raw of edits) {
    if (typeof raw !== 'object' || raw === null) continue;
    const name = dataProperty(raw, 'sheet');
    const ref = dataProperty(raw, 'cell');
    const sheet = typeof name === 'string' ? byName.get(name) : undefined;
    const address = typeof ref === 'string' ? parseCellRef(ref) : null;
    if (sheet === undefined || address === null) continue;
    let cells = seen.get(sheet);
    if (cells === undefined) {
      cells = new Set();
      seen.set(sheet, cells);
    }
    if (sheet.cell(formatCellRef(address.row, address.col)) === undefined) cells.add(address.row * (MAX_COLUMNS + 1) + address.col);
  }
  return new Map([...seen].map(([sheet, cells]) => [sheet, cells.size]));
}

// ---------------------------------------------------------------------------------------------
// Small text helpers

/** `"x:"` for a qualified name `x:local`, `""` without a prefix. */
export function prefixOf(qname: string): string {
  const colon = qname.indexOf(':');
  return colon === -1 ? '' : qname.slice(0, colon + 1);
}

/** Adds a namespace prefix to every element of a snippet generated by xlvane (never to source text). */
export function withPrefix(xml: string, prefix: string): string {
  return prefix === '' ? xml : xml.replace(/<(\/?)([A-Za-z])/g, `<$1${prefix}$2`);
}

/** Sets (or adds) an attribute in a start tag given as source text. */
export function setAttr(tag: string, name: string, value: string): string {
  const pattern = new RegExp(`(\\s${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*=\\s*)("[^"]*"|'[^']*')`);
  // A function, so that "$1" or "$&" in the value (such as "$C$14") is not a replacement pattern.
  if (pattern.test(tag)) return tag.replace(pattern, (_m, head: string) => `${head}"${escapeAttr(value)}"`);
  const close = tag.endsWith('/>') ? tag.length - 2 : tag.length - 1;
  return `${tag.slice(0, close)} ${name}="${escapeAttr(value)}"${tag.slice(close)}`;
}

/** Turns an empty-element start tag `<x .../>` into an open tag `<x ...>`. */
export function openTag(tag: string): string {
  return tag.endsWith('/>') ? `${tag.slice(0, -2).trimEnd()}>` : tag;
}

/**
 * Applies replacements given as [start, end, text] on `source`; ranges must not overlap.
 * Several insertions at one position end up in reverse order of `edits`, before a replacement
 * that starts there. The text is built in one pass, so the cost is linear in its size (D-R03).
 */
export function splice(source: string, edits: readonly [number, number, string][]): string {
  if (edits.length === 0) return source;
  // Ascending start; at one start, ascending end (insertions before the replacement), and for the
  // same span the later edit first, which is the order the earlier end-to-start splicing gave.
  const order = edits.map((edit, i) => ({ edit, i })).sort((a, b) => a.edit[0] - b.edit[0] || a.edit[1] - b.edit[1] || b.i - a.i);
  const out: string[] = [];
  let pos = 0;
  for (const { edit } of order) {
    const [start, end, text] = edit;
    if (start < pos) return spliceSlow(source, edits);
    out.push(source.slice(pos, start), text);
    pos = end;
  }
  out.push(source.slice(pos));
  return out.join('');
}

/** The earlier end-to-start splicing, kept for overlapping edits (which callers do not produce). */
function spliceSlow(source: string, edits: readonly [number, number, string][]): string {
  const sorted = [...edits].sort((a, b) => b[0] - a[0] || b[1] - a[1]);
  let out = source;
  for (const [start, end, text] of sorted) out = out.slice(0, start) + text + out.slice(end);
  return out;
}

/** An XML part as edited text, with a UTF-16 declaration corrected to UTF-8 (D-310). */
function asUtf8(text: string): string {
  return text.replace(/^(<\?xml[^>]*?encoding\s*=\s*)(["'])utf-16(?:le|be)?\2/i, '$1"UTF-8"');
}

export interface Element {
  readonly local: string;
  /** Namespace URI ('' for none). */
  readonly ns: string;
  readonly depth: number;
  readonly attrs: ReadonlyMap<string, string>;
  readonly start: XmlSpan;
  end: XmlSpan;
}

/** Scans a part and records the elements `keep` selects, with their spans. */
export function scan(text: string, part: string, limits: ReadLimits, keep: (local: string, depth: number, parent: string, ns: string) => boolean): { root: Element; elements: Element[] } {
  const elements: Element[] = [];
  const stack: (Element | null)[] = [];
  const names: string[] = [];
  let root: Element | null = null;
  parseXml(
    text,
    {
      start(ns, local, attrs, span) {
        const depth = names.length + 1;
        const parent = names[names.length - 1] ?? '';
        names.push(local);
        let element: Element | null = null;
        if (depth === 1 || keep(local, depth, parent, ns)) {
          element = { local, ns, depth, attrs, start: span, end: span };
          if (depth === 1) root = element;
          else elements.push(element);
        }
        stack.push(element);
      },
      end(_ns, _local, span) {
        names.pop();
        const element = stack.pop();
        if (element) element.end = span;
      },
      text() {},
    },
    { part, limits },
  );
  if (root === null) throw new XlvError('XLV_E_PART_INVALID', 'The part has no root element.', { location: { part }, hint: 'Re-save the file in a spreadsheet application.' });
  return { root, elements };
}

// ---------------------------------------------------------------------------------------------
// Package access

/**
 * The parts of a package being edited. Entries are inflated only when an edit reads them; parts
 * that no edit touches are copied to the new archive in their compressed form, so their bytes do
 * not change and unreferenced entries cost neither time nor memory (D-R03).
 */
export class PackageEditor {
  private readonly zip: ZipReader;
  private readonly entries: readonly ZipEntry[];
  /** Entry names by part key (and `[content_types].xml`), so lookups do not scan all entries. */
  private readonly byKey = new Map<string, ZipEntry>();
  private readonly data = new Map<string, Uint8Array>();
  private readonly changed = new Map<string, Uint8Array>();
  private readonly removed = new Set<string>();
  private readonly added: { name: string; data: Uint8Array }[] = [];
  private readonly texts = new Map<string, string>();
  private readonly encoder = new TextEncoder();

  constructor(
    bytes: Uint8Array,
    readonly limits: ReadLimits,
    readonly pkg: OpcPackage,
  ) {
    this.zip = openZip(bytes, limits);
    this.entries = this.zip.entries;
    for (const e of this.entries) {
      if (e.name === '' || e.name.endsWith('/')) continue;
      const key = e.name.toLowerCase() === '[content_types].xml' ? CONTENT_TYPES_KEY : partKey(zipNameToPartName(e.name));
      if (!this.byKey.has(key)) this.byKey.set(key, e);
    }
  }

  private entryOf(partName: string): ZipEntry | undefined {
    return this.byKey.get(partName === '/[Content_Types].xml' ? CONTENT_TYPES_KEY : partKey(partName));
  }

  private bytesOf(entry: ZipEntry): Uint8Array {
    let bytes = this.data.get(entry.name);
    if (bytes === undefined) {
      bytes = this.zip.read(entry);
      this.data.set(entry.name, bytes);
    }
    return bytes;
  }

  /** The text of a part (or of `[Content_Types].xml`), normalised and as last written. */
  text(partName: string): string | undefined {
    const cached = this.texts.get(partKey(partName));
    if (cached !== undefined) return cached;
    const entry = this.entryOf(partName);
    if (entry === undefined) return undefined;
    return normalizeLineEnds(decodeXml(this.changed.get(entry.name) ?? this.bytesOf(entry), partName));
  }

  /**
   * Writes a part. Text given in pieces is encoded piece by piece and not kept as text, so a sheet
   * of millions of cells is never one more string of its whole size (A2); reading the part again
   * decodes the written bytes.
   */
  write(partName: string, text: string | readonly string[]): void {
    let data: Uint8Array;
    if (typeof text === 'string') {
      this.texts.set(partKey(partName), text);
      data = this.encoder.encode(asUtf8(text));
    } else {
      this.texts.delete(partKey(partName));
      data = this.encodePieces(text);
    }
    const entry = this.entryOf(partName);
    if (entry) {
      this.changed.set(entry.name, data);
      return;
    }
    const name = partName.slice(1);
    const existing = this.added.find((a) => a.name === name);
    if (existing) existing.data = data;
    else this.added.push({ name, data });
  }

  private encodePieces(pieces: readonly string[]): Uint8Array {
    let size = 0;
    for (const piece of pieces) size += piece.length;
    // At most 3 bytes per UTF-16 code unit; usually far less, so start smaller and grow.
    let out = new Uint8Array(Math.max(1024, size + (size >> 2)));
    let used = 0;
    pieces.forEach((piece, i) => {
      const text = i === 0 ? asUtf8(piece) : piece;
      for (let from = 0; from < text.length; ) {
        if (out.length - used < 3 * Math.min(text.length - from, 65_536)) {
          const bigger = new Uint8Array(Math.max(out.length * 2, used + 3 * 65_536));
          bigger.set(out.subarray(0, used));
          out = bigger;
        }
        // Whole pieces of at most 64 Ki code units, cut before a low surrogate so a pair stays together.
        let to = Math.min(text.length, from + 65_536);
        if (to < text.length && (text.charCodeAt(to) & 0xfc00) === 0xdc00) to--;
        const { written } = this.encoder.encodeInto(text.slice(from, to), out.subarray(used));
        used += written;
        from = to;
      }
    });
    return out.slice(0, used);
  }

  remove(partName: string): void {
    const entry = this.entryOf(partName);
    if (entry) this.removed.add(entry.name);
  }

  hasPart(partName: string): boolean {
    return this.entryOf(partName) !== undefined || this.added.some((a) => partKey(`/${a.name}`) === partKey(partName));
  }

  build(): Uint8Array {
    const out: ZipWriteEntry[] = [];
    for (const e of this.entries) {
      if (this.removed.has(e.name)) continue;
      const changed = this.changed.get(e.name);
      if (changed !== undefined) {
        out.push({ name: e.name, data: changed });
        continue;
      }
      // Unchanged entries keep their compressed bytes; an entry already inflated is written from
      // those bytes (it was checked then), as is one whose form cannot be copied.
      const inflated = this.data.get(e.name);
      const raw = inflated === undefined ? (this.zip.raw?.(e) ?? null) : null;
      out.push(raw !== null ? { name: e.name, data: EMPTY_BYTES, raw } : { name: e.name, data: inflated ?? this.bytesOf(e) });
    }
    out.push(...this.added);
    return writeZip(out);
  }
}

const CONTENT_TYPES_KEY = '\u0000[content_types]';
const EMPTY_BYTES = new Uint8Array(0);

/** Adds an Override to `[Content_Types].xml`. */
function addContentType(edit: PackageEditor, partName: string, contentType: string): void {
  const text = edit.text('/[Content_Types].xml');
  if (text === undefined) return;
  const { root } = scan(text, '/[Content_Types].xml', edit.limits, () => false);
  const prefix = prefixOf(root.start.qname);
  const override = withPrefix(`<Override${attr('PartName', partName)}${attr('ContentType', contentType)}/>`, prefix);
  edit.write('/[Content_Types].xml', root.start.empty ? `${text.slice(0, root.start.start)}${openTag(text.slice(root.start.start, root.start.end))}${override}</${root.start.qname}>${text.slice(root.start.end)}` : splice(text, [[root.end.start, root.end.start, override]]));
}

/** Removes the Override of `partName` from `[Content_Types].xml`. */
export function removeContentType(edit: PackageEditor, partName: string): void {
  const text = edit.text('/[Content_Types].xml');
  if (text === undefined) return;
  const { elements } = scan(text, '/[Content_Types].xml', edit.limits, (local, depth) => depth === 2 && local === 'Override');
  const key = partKey(partName);
  const cut = elements.filter((e) => partKey(e.attrs.get('PartName') ?? '') === key).map((e): [number, number, string] => [e.start.start, e.end.end, '']);
  if (cut.length > 0) edit.write('/[Content_Types].xml', splice(text, cut));
}

/** Adds a relationship from `source` and returns its Id. */
function addRelationship(edit: PackageEditor, source: string, type: string, target: string): string {
  const relsPart = relationshipsPartName(source);
  const text = edit.text(relsPart) ?? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const { root, elements } = scan(text, relsPart, edit.limits, (local, depth) => depth === 2 && local === 'Relationship');
  const ids = new Set(elements.map((e) => e.attrs.get('Id') ?? ''));
  let n = elements.length + 1;
  while (ids.has(`rId${String(n)}`)) n++;
  const id = `rId${String(n)}`;
  const rel = withPrefix(`<Relationship${attr('Id', id)}${attr('Type', type)}${attr('Target', target)}/>`, prefixOf(root.start.qname));
  const updated = root.start.empty
    ? `${text.slice(0, root.start.start)}${openTag(text.slice(root.start.start, root.start.end))}${rel}</${root.start.qname}>${text.slice(root.start.end)}`
    : splice(text, [[root.end.start, root.end.start, rel]]);
  edit.write(relsPart, updated);
  return id;
}

/** Removes the relationships of `source` whose Id is in `ids`. */
export function removeRelationships(edit: PackageEditor, source: string, ids: ReadonlySet<string>): void {
  const relsPart = relationshipsPartName(source);
  const text = edit.text(relsPart);
  if (text === undefined) return;
  const { elements } = scan(text, relsPart, edit.limits, (local, depth) => depth === 2 && local === 'Relationship');
  const cut = elements.filter((e) => ids.has(e.attrs.get('Id') ?? '')).map((e): [number, number, string] => [e.start.start, e.end.end, '']);
  if (cut.length > 0) edit.write(relsPart, splice(text, cut));
}

// ---------------------------------------------------------------------------------------------
// Shared strings

class SharedStringsEditor {
  private readonly index = new Map<string, number>();
  private readonly added: string[] = [];
  private references = 0;
  private readonly existing: number;

  constructor(
    private readonly edit: PackageEditor,
    private part: string | null,
    private readonly workbookPart: string,
    private readonly conformance: 'strict' | 'transitional',
    warnings: WarningCollector,
  ) {
    const text = part === null ? undefined : edit.text(part);
    const rich = new Map<number, readonly RichTextRun[]>();
    const strings = text === undefined || part === null ? [] : parseSharedStrings(text, { part, limits: edit.limits, warnings }, rich);
    // Items are keyed by the content xlvane would write for them (stringItemXml). Rich text items
    // of the file are not indexed: their markup may hold properties xlvane does not model.
    strings.forEach((s, i) => {
      const key = textElement(s);
      if (!rich.has(i) && !this.index.has(key)) this.index.set(key, i);
    });
    this.existing = strings.length;
  }

  /** Adds a string item given as its content (stringItemXml) and returns its index. */
  add(item: string): number {
    this.references++;
    let i = this.index.get(item);
    if (i === undefined) {
      i = this.existing + this.added.length;
      this.added.push(item);
      this.index.set(item, i);
    }
    return i;
  }

  finish(): void {
    if (this.references === 0) return;
    if (this.part === null) {
      let name = '/xl/sharedStrings.xml';
      for (let n = 2; this.edit.hasPart(name); n++) name = `/xl/sharedStrings${String(n)}.xml`;
      this.part = name;
      const base = this.workbookPart.slice(0, this.workbookPart.lastIndexOf('/') + 1);
      const target = name.startsWith(base) ? name.slice(base.length) : name;
      addRelationship(this.edit, this.workbookPart, relationshipTypes('sharedStrings')[this.conformance === 'strict' ? 0 : 1], target);
      addContentType(this.edit, name, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml');
      this.edit.write(name, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${SML_NS[this.conformance]}" count="0" uniqueCount="0"></sst>`);
    }
    const part = this.part;
    const text = this.edit.text(part) ?? '';
    const { root } = scan(text, part, this.edit.limits, () => false);
    const prefix = prefixOf(root.start.qname);
    const items = this.added.map((item) => withPrefix(`<si>${item}</si>`, prefix)).join('');
    const count = Number(root.attrs.get('count') ?? this.existing) + this.references;
    let tag = setAttr(text.slice(root.start.start, root.start.end), 'count', String(Number.isFinite(count) ? count : this.existing + this.references));
    tag = setAttr(tag, 'uniqueCount', String(this.existing + this.added.length));
    const updated = root.start.empty
      ? `${text.slice(0, root.start.start)}${openTag(tag)}${items}</${root.start.qname}>${text.slice(root.start.end)}`
      : splice(text, [
          [root.start.start, root.start.end, tag],
          [root.end.start, root.end.start, items],
        ]);
    this.edit.write(part, updated);
  }
}

// ---------------------------------------------------------------------------------------------
// Styles

/** Containers of the styles part in schema order (W2-10). */
const STYLE_CONTAINERS = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];
/** Records prepended when a container has to be created, so that index 0 keeps its meaning (D-311). */
const CONTAINER_DEFAULTS: Readonly<Record<string, readonly string[]>> = {
  fonts: [fontXml({})],
  fills: ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'],
  borders: ['<border><left/><right/><top/><bottom/><diagonal/></border>'],
  cellXfs: ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'],
};
const COUNT_LIMITS: Readonly<Record<string, StyleCountLimit>> = { numFmts: 'customNumFmts', fonts: 'fonts', fills: 'fills', borders: 'borders', cellXfs: 'cellXfs' };
const ITEM_NAMES: Readonly<Record<string, string>> = { numFmts: 'numFmt', fonts: 'font', fills: 'fill', borders: 'border', cellXfs: 'xf' };

class StylesEditor {
  private readonly styles: Styles;
  private readonly root: XmlElement | null;
  private readonly ns: string;
  private readonly existingCount = new Map<string, number>();
  private readonly additions = new Map<string, string[]>();
  private readonly addedIds = new Map<string, Map<string, number>>();
  private readonly numFmtIds = new Map<string, number>();
  private nextNumFmtId = 164;
  private readonly cache = new Map<string, number>();

  constructor(
    private readonly edit: PackageEditor,
    private part: string | null,
    private readonly workbookPart: string,
    private readonly conformance: 'strict' | 'transitional',
    warnings: WarningCollector,
  ) {
    const text = part === null ? undefined : edit.text(part);
    if (text === undefined || part === null) {
      this.styles = EMPTY_STYLES;
      this.root = null;
      this.ns = SML_NS[conformance];
      return;
    }
    this.styles = parseStyles(text, { part, limits: edit.limits, warnings });
    this.root = parseXmlTree(text, { part, limits: edit.limits });
    this.ns = this.root.ns;
    for (const container of ['numFmts', 'fonts', 'fills', 'borders', 'cellXfs']) {
      const el = childNamed(this.root, this.ns, container);
      this.existingCount.set(container, el === undefined ? -1 : childrenNamed(el, this.ns, ITEM_NAMES[container] ?? '').length);
    }
    const numFmts = childNamed(this.root, this.ns, 'numFmts');
    for (const nf of numFmts === undefined ? [] : childrenNamed(numFmts, this.ns, 'numFmt')) {
      const id = Number(nf.attrs.get('numFmtId'));
      const code = nf.attrs.get('formatCode');
      if (!Number.isInteger(id) || code === undefined) continue;
      if (!this.numFmtIds.has(code)) this.numFmtIds.set(code, id);
      this.nextNumFmtId = Math.max(this.nextNumFmtId, id + 1);
    }
  }

  /** `true` if the styles part has a cell format at this index. */
  hasCellFormat(index: number): boolean {
    return index < (this.existingCount.get('cellXfs') ?? -1);
  }

  /** The cell style of a cell format index, as the reader resolves it. */
  styleOf(index: number): CellStyle {
    return this.styles.cellStyle(index, null);
  }

  private add(container: string, xml: string): number {
    let ids = this.addedIds.get(container);
    if (!ids) {
      ids = new Map();
      this.addedIds.set(container, ids);
    }
    const known = ids.get(xml);
    if (known !== undefined) return known;
    let list = this.additions.get(container);
    if (!list) {
      list = [];
      this.additions.set(container, list);
    }
    const existing = this.existingCount.get(container) ?? -1;
    const base = existing >= 0 ? existing : (CONTAINER_DEFAULTS[container]?.length ?? 0);
    const id = base + list.length;
    // Only records added by this edit are checked, so files that already have more stay editable.
    const limit = COUNT_LIMITS[container];
    if (limit !== undefined) checkStyleCount(limit, id + 1);
    list.push(xml);
    ids.set(xml, id);
    return id;
  }

  private numFmtId(code: string | number): number {
    // A built-in id refers to the format the file gives that id, implied or declared (E-41).
    if (typeof code === 'number') return code;
    const builtin = reusableBuiltinNumFmt(code);
    if (builtin !== undefined) return builtin;
    let id = this.numFmtIds.get(code);
    if (id === undefined) {
      id = this.nextNumFmtId++;
      this.numFmtIds.set(code, id);
      this.add('numFmts', `<numFmt${attr('numFmtId', id)}${attr('formatCode', encodeXstring(code))}/>`);
    }
    return id;
  }

  /** A cell format: the format at `base` with the parts of `patch` replaced or merged (D-311). */
  derive(base: number, patch: NormalizedStyle): number {
    const key = `${String(base)}|${stableKey(patch)}`;
    const cached = this.cache.get(key);
    if (cached !== undefined) return cached;
    const cellXfs = this.root === null ? undefined : childNamed(this.root, this.ns, 'cellXfs');
    const baseXf = cellXfs === undefined ? undefined : childrenNamed(cellXfs, this.ns, 'xf')[base];
    const attrs = new Map<string, string>();
    for (const [k, v] of baseXf?.attrs ?? []) if (!k.startsWith('{')) attrs.set(k, v);
    for (const [k, v] of [['numFmtId', '0'], ['fontId', '0'], ['fillId', '0'], ['borderId', '0']] as const) if (!attrs.has(k)) attrs.set(k, v);
    // The style already in the file is kept even where it is outside the limits for new styles.
    const current = normalizeStyle(this.styleOf(base), 'existing');
    const merged = mergeStyles(current, patch);

    if (patch.font) {
      const defaults = this.styles.defaultFont();
      const fontDefaults = { name: defaults?.name ?? DEFAULT_FONT.name, size: defaults?.size ?? DEFAULT_FONT.size };
      attrs.set('fontId', String(this.add('fonts', fontXml(merged.font ?? {}, fontDefaults))));
      attrs.set('applyFont', '1');
    }
    if (patch.fill && merged.fill) {
      attrs.set('fillId', String(this.add('fills', fillXml(merged.fill))));
      attrs.set('applyFill', '1');
    }
    if (patch.border && merged.border) {
      attrs.set('borderId', String(this.add('borders', borderXml(merged.border))));
      attrs.set('applyBorder', '1');
    }
    if (patch.numFmt !== undefined) {
      attrs.set('numFmtId', String(this.numFmtId(patch.numFmt)));
      attrs.set('applyNumberFormat', '1');
    }
    let children = '';
    const alignmentAttrs = new Map<string, string>();
    const baseAlignment = baseXf === undefined ? undefined : childNamed(baseXf, this.ns, 'alignment');
    for (const [k, v] of baseAlignment?.attrs ?? []) if (!k.startsWith('{')) alignmentAttrs.set(k, v);
    if (patch.alignment) {
      const a = patch.alignment;
      if (a.horizontal !== undefined) alignmentAttrs.set('horizontal', a.horizontal);
      if (a.vertical !== undefined) alignmentAttrs.set('vertical', a.vertical);
      if (a.wrapText) alignmentAttrs.set('wrapText', '1');
      if (a.indent !== undefined) alignmentAttrs.set('indent', String(a.indent));
      if (a.textRotation !== undefined) alignmentAttrs.set('textRotation', String(a.textRotation));
      if (a.shrinkToFit) alignmentAttrs.set('shrinkToFit', '1');
      if (a.justifyLastLine) alignmentAttrs.set('justifyLastLine', '1');
      if (a.readingOrder !== undefined) alignmentAttrs.set('readingOrder', String(READING_ORDERS.indexOf(a.readingOrder)));
      attrs.set('applyAlignment', '1');
    }
    if (alignmentAttrs.size > 0) children += `<alignment${[...alignmentAttrs].map(([k, v]) => attr(k, v)).join('')}/>`;
    const protection = baseXf === undefined ? undefined : childNamed(baseXf, this.ns, 'protection');
    const protectionAttrs = new Map<string, string>();
    for (const [k, v] of protection?.attrs ?? []) if (!k.startsWith('{')) protectionAttrs.set(k, v);
    if (patch.protection) {
      // Defaults (locked, not hidden; [MS-OI29500] §2.1.714) are written explicitly when the
      // base format has the other value, so the patch always wins (W2-63).
      for (const key of ['locked', 'hidden'] as const) {
        const v = patch.protection[key];
        if (v !== undefined) protectionAttrs.set(key, v ? '1' : '0');
      }
      attrs.set('applyProtection', '1');
    }
    if (protectionAttrs.size > 0) children += `<protection${[...protectionAttrs].map(([k, v]) => attr(k, v)).join('')}/>`;
    const xf = `<xf${[...attrs].map(([k, v]) => attr(k, v)).join('')}${children === '' ? '/>' : `>${children}</xf>`}`;
    const id = this.add('cellXfs', xf);
    this.cache.set(key, id);
    return id;
  }

  finish(): void {
    if (this.additions.size === 0) return;
    if (this.part === null) {
      let name = '/xl/styles.xml';
      for (let n = 2; this.edit.hasPart(name); n++) name = `/xl/styles${String(n)}.xml`;
      this.part = name;
      const base = this.workbookPart.slice(0, this.workbookPart.lastIndexOf('/') + 1);
      addRelationship(this.edit, this.workbookPart, relationshipTypes('styles')[this.conformance === 'strict' ? 0 : 1], name.startsWith(base) ? name.slice(base.length) : name);
      addContentType(this.edit, name, 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml');
      this.edit.write(name, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="${this.ns}"></styleSheet>`);
    }
    const part = this.part;
    const text = this.edit.text(part) ?? '';
    const { root, elements } = scan(text, part, this.edit.limits, (_local, depth) => depth === 2);
    const prefix = prefixOf(root.start.qname);
    const changes: [number, number, string][] = [];
    let rootTag = text.slice(root.start.start, root.start.end);
    let appendAtEnd = '';
    for (const [container, items] of this.additions) {
      const element = elements.find((e) => e.local === container);
      if (element) {
        const existing = this.existingCount.get(container) ?? 0;
        const tag = setAttr(text.slice(element.start.start, element.start.end), 'count', String(existing + items.length));
        const body = withPrefix(items.join(''), prefix);
        if (element.start.empty) changes.push([element.start.start, element.start.end, `${openTag(tag)}${body}</${element.start.qname}>`]);
        else changes.push([element.start.start, element.start.end, tag], [element.end.start, element.end.start, body]);
        continue;
      }
      const all = [...(CONTAINER_DEFAULTS[container] ?? []), ...items];
      const created = withPrefix(`<${container}${attr('count', all.length)}>${all.join('')}</${container}>`, prefix);
      const order = STYLE_CONTAINERS.indexOf(container);
      const next = elements.find((e) => STYLE_CONTAINERS.indexOf(e.local) > order);
      if (next) changes.push([next.start.start, next.start.start, created]);
      else if (root.start.empty) appendAtEnd += created;
      else changes.push([root.end.start, root.end.start, created]);
    }
    if (root.start.empty) {
      rootTag = openTag(rootTag);
      changes.push([root.start.start, root.start.end, `${rootTag}${appendAtEnd}</${root.start.qname}>`]);
    }
    this.edit.write(part, splice(text, changes));
  }
}

// ---------------------------------------------------------------------------------------------
// Worksheets

interface PlannedCell {
  readonly row: number;
  readonly col: number;
  readonly ref: string;
  /** New cell XML (unprefixed), or `null` to remove the cell. */
  readonly xml: string | null;
}

function valueParts(v: Exclude<StoredValue, { kind: 'formula' }>, sst: SharedStringsEditor, inFormula: boolean): { t: string; v: string } {
  switch (v.kind) {
    case 'number':
      return { t: '', v: formatNumber(v.n) };
    case 'string':
      return inFormula ? { t: 'str', v: escapeText(encodeXstring(v.s)) } : { t: 's', v: String(sst.add(stringItemXml(v))) };
    case 'boolean':
      return { t: 'b', v: v.b ? '1' : '0' };
    case 'error':
      return { t: 'e', v: escapeText(v.e) };
  }
}

function newCellXml(ref: string, value: StoredValue | null, xf: number, sst: SharedStringsEditor): string {
  const s = xf === 0 ? '' : attr('s', xf);
  if (value === null) return `<c r="${ref}"${s}/>`;
  if (value.kind === 'formula') {
    const f = `<f>${escapeText(encodeXstring(value.text))}</f>`;
    if (value.result === null) return `<c r="${ref}"${s}>${f}</c>`;
    const r = valueParts(value.result, sst, true);
    return `<c r="${ref}"${s}${r.t === '' ? '' : attr('t', r.t)}>${f}<v>${r.v}</v></c>`;
  }
  const r = valueParts(value, sst, false);
  return `<c r="${ref}"${s}${r.t === '' ? '' : attr('t', r.t)}><v>${r.v}</v></c>`;
}

/** A growable list of 32-bit integers, so that a sheet's cells cost a few numbers each (A2, D-K04). */
export class IntList {
  private data = new Int32Array(1024);
  length = 0;

  push(value: number): void {
    if (this.length === this.data.length) {
      const bigger = new Int32Array(this.data.length * 2);
      bigger.set(this.data);
      this.data = bigger;
    }
    this.data[this.length++] = value;
  }

  get(index: number): number {
    return this.data[index]!;
  }

  set(index: number, value: number): void {
    this.data[index] = value;
  }
}

/** Integers stored per row of {@link SheetScan.rows}. */
const ROW_FIELDS = 8;
/** Integers stored per cell of {@link SheetScan.cells}. */
const CELL_FIELDS = 5;

/**
 * The structure of a worksheet part that editing needs, kept as numbers: per row (ROW_FIELDS
 * integers) its number, the offsets of its start tag, of its end tag, its flags (bit 0: has `r`,
 * bit 1: empty-element tag, from bit 2: length of its qualified name), its first cell and its
 * number of cells; per cell (CELL_FIELDS integers) its column, the offsets of its start tag and
 * of its end, and its flags (bit 0: has `r`, from bit 1: length of its qualified name). Objects
 * per cell made editing a sheet of millions of cells run out of memory (A2).
 */
interface SheetScan {
  readonly rootQname: string;
  readonly sheetData: { start: XmlSpan; end: XmlSpan } | null;
  readonly dimension: { start: XmlSpan; ref: string | undefined } | null;
  readonly rows: IntList;
  readonly rowCount: number;
  readonly cells: IntList;
  /** `true` when a row has no `r` attribute: then every row is written with references. */
  readonly missingRowRefs: boolean;
  /** Cell format of new cells by column (index = column), 0 for none (W2-03). */
  readonly columnStyles: Int32Array;
  /** Rows with their own format (`row@s` with `customFormat`), whose new cells take format 0 (D-A04). */
  readonly rowsWithFormat: ReadonlySet<number>;
}

function parseIndex(value: string | undefined, digits: number): number {
  const v = value?.trim() ?? '';
  return v.length > 0 && v.length <= digits && /^[0-9]+$/.test(v) ? Number(v) : Number.NaN;
}

/**
 * Scans a worksheet part once for {@link editSheetXml} and {@link newCellStyle}. Rows and cells
 * are numbered as the reader numbers them (D-105); the reader skips rows and cells with invalid
 * references (D-903), and editing around them could put new cells in the wrong place, so a sheet
 * that has any is not edited.
 */
function scanSheet(text: string, part: string, sheet: string, limits: ReadLimits, styles: StylesEditor): SheetScan {
  const invalid = (ref: string): never => {
    throw new XlvError('XLV_E_INVALID_CELL_REF', `Sheet "${sheet}" contains the invalid reference "${ref.slice(0, 20)}", so it cannot be edited safely.`, {
      location: { part, sheet },
      hint: 'Open and re-save the file in a spreadsheet application, which drops the damaged cells, then edit the re-saved file.',
      details: { ref: ref.slice(0, 40) },
    });
  };
  const names: string[] = [];
  /** Whether each open element is a recorded row (depth 3) or cell (depth 4). */
  const recorded: boolean[] = [];
  let rootQname: string | null = null;
  let sheetDataStart: XmlSpan | null = null;
  let sheetData: { start: XmlSpan; end: XmlSpan } | null = null;
  let dimension: { start: XmlSpan; ref: string | undefined } | null = null;
  const rows = new IntList();
  const cells = new IntList();
  let rowCount = 0;
  let currentRow = 0;
  let lastCol = 0;
  let missingRowRefs = false;
  const cols: { min: number; max: number; style: number }[] = [];
  const rowsWithFormat = new Set<number>();
  parseXml(
    text,
    {
      start(_ns, local, attrs, span) {
        const depth = names.length + 1;
        const parent = names[names.length - 1] ?? '';
        names.push(local);
        let keep = false;
        if (depth === 1) rootQname = span.qname;
        else if (depth === 2 && local === 'sheetData' && sheetDataStart === null) sheetDataStart = span;
        else if (depth === 2 && local === 'dimension' && dimension === null) dimension = { start: span, ref: attrs.get('ref') };
        else if (depth === 3 && local === 'col' && parent === 'cols') {
          const [min, max, style] = [parseIndex(attrs.get('min'), 10), parseIndex(attrs.get('max'), 10), parseIndex(attrs.get('style'), 10)];
          if (Number.isInteger(min) && Number.isInteger(max) && Number.isInteger(style) && style !== 0 && styles.hasCellFormat(style)) cols.push({ min, max, style });
        } else if (depth === 3 && local === 'row' && parent === 'sheetData') {
          const r = attrs.get('r');
          const n = r === undefined ? currentRow + 1 : /^[0-9]{1,7}$/.test(r.trim()) ? Number(r.trim()) : Number.NaN;
          if (!Number.isInteger(n) || n < 1 || n > MAX_ROWS) invalid(r ?? String(n));
          currentRow = n;
          lastCol = 0;
          if (r === undefined) missingRowRefs = true;
          const custom = attrs.get('customFormat')?.trim();
          if (attrs.has('s') && (custom === '1' || custom === 'true')) rowsWithFormat.add(n);
          rows.push(n);
          rows.push(span.start);
          rows.push(span.end);
          rows.push(span.end);
          rows.push(span.end);
          rows.push((r === undefined ? 0 : 1) | (span.empty ? 2 : 0) | (span.qname.length << 2));
          rows.push(cells.length / CELL_FIELDS);
          rows.push(0);
          rowCount++;
          keep = true;
        } else if (depth === 4 && local === 'c' && parent === 'row' && recorded[depth - 2] === true) {
          const r = attrs.get('r');
          const address = r === undefined || r === '' ? null : parseCellRef(r);
          if (r !== undefined && r !== '' && address === null) invalid(r);
          const col = address?.col ?? lastCol + 1;
          lastCol = col;
          cells.push(col);
          cells.push(span.start);
          cells.push(span.end);
          cells.push(span.end);
          cells.push((r === undefined ? 0 : 1) | (span.qname.length << 1));
          rows.set((rowCount - 1) * ROW_FIELDS + 7, rows.get((rowCount - 1) * ROW_FIELDS + 7) + 1);
          keep = true;
        }
        recorded.push(keep);
      },
      end(_ns, local, span) {
        const depth = names.length;
        names.pop();
        const keep = recorded.pop() === true;
        if (depth === 2 && local === 'sheetData' && sheetData === null && sheetDataStart !== null) sheetData = { start: sheetDataStart, end: span };
        else if (keep && depth === 3) {
          rows.set((rowCount - 1) * ROW_FIELDS + 3, span.start);
          rows.set((rowCount - 1) * ROW_FIELDS + 4, span.end);
        } else if (keep && depth === 4) cells.set(cells.length - 2, span.end);
      },
      text() {},
    },
    { part, limits },
  );
  if (rootQname === null) throw new XlvError('XLV_E_PART_INVALID', 'The part has no root element.', { location: { part }, hint: 'Re-save the file in a spreadsheet application.' });
  // The first col element that covers a column gives its style. Each column is assigned once, so
  // overlapping col elements cost no more than the columns they cover (A3).
  const columnStyles = new Int32Array(MAX_COLUMNS + 2);
  const next = new Int32Array(MAX_COLUMNS + 2);
  for (let i = 0; i < next.length; i++) next[i] = i;
  const unset = (c: number): number => {
    let root = c;
    while (next[root] !== root) root = next[root]!;
    while (next[c] !== root) {
      const up = next[c]!;
      next[c] = root;
      c = up;
    }
    return root;
  };
  for (const c of cols) {
    const min = Math.max(c.min, 1);
    const max = Math.min(c.max, MAX_COLUMNS);
    if (min > max) continue;
    for (let col = unset(min); col <= max; col = unset(col + 1)) {
      columnStyles[col] = c.style;
      next[col] = col + 1;
    }
  }
  return { rootQname, sheetData, dimension, rows, rowCount, cells, missingRowRefs, columnStyles, rowsWithFormat };
}

/**
 * The cell format a new cell of a sheet starts from: the `style` of the `col` element that covers
 * its column, which ECMA-376 Part 1 §18.3.1.13 says affects cells not yet allocated (W2-03). A row
 * with its own format (`row@s` with `customFormat`) keeps the earlier behaviour, format 0, since
 * which of the two wins is not stated in the sources read (D-A04). Indexes beyond the cell
 * formats of the file are ignored.
 */
function newCellStyle(scan: SheetScan, row: number, col: number): number {
  return scan.rowsWithFormat.has(row) ? 0 : (scan.columnStyles[col] ?? 0);
}

/** Joins many small strings without keeping one string per piece until the end. */
export class TextBuilder {
  private readonly pieces: string[] = [];
  private readonly joined: string[] = [];

  add(text: string): void {
    this.pieces.push(text);
    if (this.pieces.length >= 4096) {
      this.joined.push(this.pieces.join(''));
      this.pieces.length = 0;
    }
  }

  /** The text as a list of strings of a few thousand pieces each. */
  finish(): string[] {
    this.joined.push(this.pieces.join(''));
    this.pieces.length = 0;
    return this.joined;
  }
}

/** Inserts `r="…"` after the qualified name of a start tag given as source text. */
function insertRef(tag: string, qnameLength: number, ref: string): string {
  const at = 1 + qnameLength;
  return `${tag.slice(0, at)} r="${ref}"${tag.slice(at)}`;
}

/** The edited text of a worksheet part, as pieces for {@link PackageEditor.write}. */
function editSheetXml(text: string, part: string, sheet: string, scan: SheetScan, planned: readonly PlannedCell[]): string[] {
  const prefix = prefixOf(scan.rootQname);
  const sheetData = scan.sheetData;
  if (!sheetData) {
    throw new XlvError('XLV_E_PART_INVALID', `The worksheet part of sheet "${sheet}" has no sheetData element.`, {
      location: { part, sheet },
      hint: 'The sheet part is damaged. Re-save the file in a spreadsheet application.',
    });
  }
  const { rows, cells, rowCount } = scan;

  const byRow = new Map<number, PlannedCell[]>();
  for (const p of planned) {
    const list = byRow.get(p.row) ?? [];
    byRow.set(p.row, list);
    list.push(p);
  }
  for (const list of byRow.values()) list.sort((a, b) => a.col - b.col);
  const emitted = new Set<number>();
  // A set of the existing row numbers and an index into the sorted new ones: a search of every
  // row per edited row, and shift(), made many edits on a long sheet quadratic (D-H01).
  const existingRows = new Set<number>();
  for (let i = 0; i < rowCount; i++) existingRows.add(rows.get(i * ROW_FIELDS));
  const pending = [...byRow.keys()].filter((n) => !existingRows.has(n)).sort((a, b) => a - b);
  let next = 0;
  const p = (xml: string): string => withPrefix(xml, prefix);
  const out = new TextBuilder();
  const newRow = (n: number): void => {
    out.add(p(`<row r="${String(n)}">`));
    for (const c of byRow.get(n) ?? []) if (c.xml !== null) out.add(p(c.xml));
    out.add(`</${prefix}row>`);
  };

  for (let i = 0; i < rowCount; i++) {
    const at = i * ROW_FIELDS;
    const number = rows.get(at);
    const tagStart = rows.get(at + 1);
    const tagEnd = rows.get(at + 2);
    const endStart = rows.get(at + 3);
    const endEnd = rows.get(at + 4);
    const flags = rows.get(at + 5);
    const firstCell = rows.get(at + 6);
    const cellCount = rows.get(at + 7);
    const hasR = (flags & 1) === 1;
    while (next < pending.length && (pending[next] ?? 0) < number) newRow(pending[next++] ?? 0);
    const edits = byRow.get(number);
    let needsRefs = scan.missingRowRefs;
    for (let k = 0; !needsRefs && k < cellCount; k++) needsRefs = (cells.get((firstCell + k) * CELL_FIELDS + 4) & 1) === 0;
    if (!edits && hasR && !needsRefs) {
      out.add(text.slice(tagStart, endEnd));
      continue;
    }
    let tag = text.slice(tagStart, tagEnd);
    if (!hasR) tag = insertRef(tag, flags >>> 2, String(number));
    const empty = (flags & 2) === 2;
    out.add(empty ? openTag(tag) : tag);
    const edited = new Set((edits ?? []).map((c) => c.col));
    const toInsert = emitted.has(number) ? [] : (edits ?? []).filter((c) => c.xml !== null);
    emitted.add(number);
    let k = 0;
    for (let j = 0; j < cellCount; j++) {
      const c = (firstCell + j) * CELL_FIELDS;
      const col = cells.get(c);
      while (k < toInsert.length && (toInsert[k]?.col ?? 0) < col) out.add(p(toInsert[k++]?.xml ?? ''));
      if (edited.has(col)) continue;
      const cellFlags = cells.get(c + 4);
      if ((cellFlags & 1) === 1) out.add(text.slice(cells.get(c + 1), cells.get(c + 3)));
      else {
        // Only the start tag gets the reference; the rest of the cell is copied as it is.
        out.add(insertRef(text.slice(cells.get(c + 1), cells.get(c + 2)), cellFlags >>> 1, formatCellRef(number, col)));
        out.add(text.slice(cells.get(c + 2), cells.get(c + 3)));
      }
    }
    while (k < toInsert.length) out.add(p(toInsert[k++]?.xml ?? ''));
    out.add(empty ? `</${text.slice(tagStart + 1, tagStart + 1 + (flags >>> 2))}>` : text.slice(endStart, endEnd));
  }
  while (next < pending.length) newRow(pending[next++] ?? 0);
  const inner = out.finish();

  const changes: [number, number, readonly string[]][] = [];
  if (sheetData.start.empty) changes.push([sheetData.start.start, sheetData.start.end, [openTag(text.slice(sheetData.start.start, sheetData.start.end)), ...inner, `</${sheetData.start.qname}>`]]);
  else changes.push([sheetData.start.end, sheetData.end.start, inner]);

  const dimension = scan.dimension;
  if (dimension) {
    const old = parseRange(dimension.ref ?? '');
    const values = planned.filter((c) => c.xml !== null);
    if (values.length > 0) {
      let minRow = old?.start.row ?? Infinity;
      let minCol = old?.start.col ?? Infinity;
      let maxRow = old?.end.row ?? 0;
      let maxCol = old?.end.col ?? 0;
      for (const c of values) {
        minRow = Math.min(minRow, c.row);
        minCol = Math.min(minCol, c.col);
        maxRow = Math.max(maxRow, c.row);
        maxCol = Math.max(maxCol, c.col);
      }
      const ref = minRow === maxRow && minCol === maxCol ? formatCellRef(minRow, minCol) : `${formatCellRef(minRow, minCol)}:${formatCellRef(maxRow, maxCol)}`;
      changes.push([dimension.start.start, dimension.start.end, [setAttr(text.slice(dimension.start.start, dimension.start.end), 'ref', ref)]]);
    }
  }
  // The two changes (sheetData and dimension) never overlap.
  changes.sort((a, b) => a[0] - b[0]);
  const pieces: string[] = [];
  let pos = 0;
  for (const [start, end, replacement] of changes) {
    pieces.push(text.slice(pos, start), ...replacement);
    pos = end;
  }
  pieces.push(text.slice(pos));
  return pieces;
}

// ---------------------------------------------------------------------------------------------
// Workbook part

const CALC_PR_LATER = ['oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];

/** Sets `calcPr@fullCalcOnLoad="1"`, adding `calcPr` in schema order if needed (W2-22, PROVISIONAL). */
export function requestFullCalculation(edit: PackageEditor, workbookPart: string): void {
  const text = edit.text(workbookPart);
  if (text === undefined) return;
  const { root, elements } = scan(text, workbookPart, edit.limits, (_local, depth) => depth === 2);
  const calcPr = elements.find((e) => e.local === 'calcPr');
  if (calcPr) {
    if (calcPr.attrs.get('fullCalcOnLoad') === '1') return;
    edit.write(workbookPart, splice(text, [[calcPr.start.start, calcPr.start.end, setAttr(text.slice(calcPr.start.start, calcPr.start.end), 'fullCalcOnLoad', '1')]]));
    return;
  }
  const created = withPrefix('<calcPr fullCalcOnLoad="1"/>', prefixOf(root.start.qname));
  const next = elements.find((e) => CALC_PR_LATER.includes(e.local));
  edit.write(workbookPart, splice(text, [[next ? next.start.start : root.end.start, next ? next.start.start : root.end.start, created]]));
}

/** Ranges with their text as written, found by a cell they contain (A3). */
class RangeLookup {
  private readonly index = new RangeIndex();
  private readonly refs = new Map<CellRange, string>();

  add(range: CellRange, ref: string): void {
    this.index.add(range);
    this.refs.set(range, ref);
  }

  /** A range that contains the cell and that `accept` (when given) accepts, with its text. */
  at(row: number, col: number, accept?: (range: CellRange) => boolean): { ref: string; range: CellRange } | undefined {
    const cell = { row, col };
    const range = this.index.find({ start: cell, end: cell }, accept);
    return range === undefined ? undefined : { ref: this.refs.get(range) ?? '', range };
  }
}

/**
 * editCells writes plain formulas only, and does not change cells whose formula covers other
 * cells: an array formula's range (P1 §18.17.6.3) or a shared formula master with dependents
 * (E-28). Replacing a cell that shares a master's formula is allowed: its own value overrides
 * the group (E-28).
 */
function checkFormulaGroups(sheet: Sheet, ref: string, at: { row: number; col: number }, value: StoredValue | null, cache: Map<string, RangeLookup>): void {
  const refuse = (message: string, hint: string) => new XlvError('XLV_E_NOT_SUPPORTED', message, { location: { part: sheet.part ?? '', sheet: sheet.name, cell: ref }, hint, details: { cell: ref } });
  if (value !== null && value.kind === 'formula' && (value.group !== undefined || value.master !== undefined)) {
    throw refuse(`editCells writes plain formulas only; the value for ${ref} is an array or shared formula.`, 'Write { formula: "…" } for each cell, or build the sheet with createWorkbook to write array and shared formulas.');
  }
  let arrays = cache.get(sheet.name);
  if (arrays === undefined) {
    // An index of the array formulas, so that each edit costs about the same however many the
    // sheet has (A3).
    arrays = new RangeLookup();
    for (const c of sheet.cells()) {
      const range = c.formula?.kind === 'array' && c.formula.ref !== null ? parseRange(c.formula.ref) : null;
      if (range !== null) arrays.add(range, c.formula!.ref!);
    }
    cache.set(sheet.name, arrays);
  }
  const array = arrays.at(at.row, at.col);
  if (array !== undefined) {
    throw refuse(`${ref} belongs to the array formula ${array.ref}, which editCells cannot change cell by cell.`, `Change the array formula in a spreadsheet application, or rebuild the sheet with createWorkbook; cells outside ${array.ref} can be edited.`);
  }
  const f = sheet.cell(ref)?.formula;
  const range = f?.kind === 'shared' && f.ref !== null ? parseRange(f.ref) : null;
  if (range !== null && (range.start.row !== range.end.row || range.start.col !== range.end.col)) {
    throw refuse(`${ref} is the master of a shared formula used by cells in ${f!.ref!}; replacing it would leave them without a formula.`, `Edit the other cells of ${f!.ref!} first, or rebuild the sheet with createWorkbook.`);
  }
}

// ---------------------------------------------------------------------------------------------
// Tables of edited sheets (D-R03)

/** Rows that appendRows adds below the last used row of a sheet. */
interface AppendedRows {
  readonly sheet: string;
  /** The last used row before the append (0 for an empty sheet). */
  readonly lastRow: number;
  readonly count: number;
}

/** Column names of a table: 1 to 255 characters on one line, unique ignoring case (TD-04). */
const MAX_COLUMN_NAME = 255;

/**
 * Keeps the Table Definition parts of an edited sheet in step with its cells:
 * - a header cell of a table holds the column name (OI §2.1.681 i, TD-05), so editing it renames
 *   `tableColumn@name`; a value that is not text, an empty or multi-line name and a name another
 *   column of the table already has are refused (`XLV_E_INVALID_VALUE`). Structured references
 *   to the old name in formulas are not rewritten;
 * - rows appended directly below a table's last row (appendRows) join the table: `ref` and an
 *   `autoFilter` that covers the header and data rows grow with them, as when rows are inserted
 *   inside a table (TD-11). A table whose totals row is the last used row is refused instead of
 *   moving its totals row, with a hint to insert the rows above the totals row with insertRows.
 */
function syncTables(edit: PackageEditor, plan: EditPlan, appended: AppendedRows | null): void {
  const sheet = plan.sheet;
  const tables = sheet.tables;
  for (const table of tables) {
    const range = parseRange(table.ref.trim().toUpperCase());
    if (range === null) continue;
    const header = table.headerRowCount > 0 ? range.start.row : null;
    const totals = table.totalsRowCount > 0 ? 1 : 0;
    const where = { part: table.part, sheet: sheet.name };

    const renames = new Map<number, string>();
    if (header !== null) {
      for (const [ref, c] of plan.cells) {
        if (c.row !== header || c.col < range.start.col || c.col > range.end.col) continue;
        const index = c.col - range.start.col;
        const value = c.value;
        const name = value !== null && value.kind === 'string' ? value.s : null;
        const location = { part: sheet.part ?? '', sheet: sheet.name, cell: ref };
        if (name === null || name.length === 0 || name.length > MAX_COLUMN_NAME || /[\r\n]/.test(name)) {
          throw new XlvError('XLV_E_INVALID_VALUE', `${ref} is a header cell of table "${table.displayName}", so it must hold the column name as text of 1 to ${String(MAX_COLUMN_NAME)} characters on one line.`, {
            location,
            hint: 'Write the new column name as a string, for example "Revenue". To change the data, edit the rows below the header.',
            details: { table: table.displayName, cell: ref },
          });
        }
        renames.set(index, name);
      }
      if (renames.size > 0) {
        const names = table.columns.map((column, i) => renames.get(i) ?? column.name);
        const seen = new Map<string, number>();
        names.forEach((name, i) => {
          const key = name.toLowerCase();
          const other = seen.get(key);
          if (other !== undefined && (renames.has(i) || renames.has(other))) {
            const cell = formatCellRef(header, range.start.col + (renames.has(i) ? i : other));
            throw new XlvError('XLV_E_INVALID_VALUE', `Table "${table.displayName}" would have two columns named "${name.slice(0, 40)}".`, {
              location: { part: sheet.part ?? '', sheet: sheet.name, cell },
              hint: 'Column names must be unique in a table (case is ignored). Choose another name, for example with a suffix such as "_2".',
              details: { table: table.displayName, cell },
            });
          }
          seen.set(key, i);
        });
      }
    }

    let newEnd: number | null = null;
    if (appended !== null && appended.count > 0 && range.end.row >= appended.lastRow && range.start.row <= appended.lastRow) {
      if (totals === 1) {
        throw new XlvError('XLV_E_NOT_SUPPORTED', `Table "${table.displayName}" of sheet "${sheet.name}" ends with a totals row on row ${String(range.end.row)}, so rows appended below it would fall outside the table.`, {
          location: { ...where, cell: formatCellRef(range.end.row, range.start.col) },
          hint: `Insert the rows above the totals row with insertRows(bytes, { sheet: "${sheet.name}", at: ${String(range.end.row)}, count: ${String(appended.count)} }), which grows the table, then fill them with editCells.`,
          details: { table: table.displayName },
        });
      }
      const end = appended.lastRow + appended.count;
      if (end > range.end.row) {
        newEnd = end;
        const grown: CellRange = { start: range.start, end: { row: end, col: range.end.col } };
        for (const other of tables) {
          if (other === table) continue;
          const r = parseRange(other.ref.trim().toUpperCase());
          if (r !== null && r.start.row <= grown.end.row && r.end.row >= grown.start.row && r.start.col <= grown.end.col && r.end.col >= grown.start.col) {
            throw new XlvError('XLV_E_NOT_SUPPORTED', `Growing table "${table.displayName}" to ${formatRange(grown)} for the appended rows would overlap table "${other.displayName}" (${other.ref}).`, {
              location: where,
              hint: 'Tables must not overlap. Append the rows to another sheet, or move one of the tables first.',
              details: { table: table.displayName },
            });
          }
        }
      }
    }
    if (renames.size === 0 && newEnd === null) continue;

    const text = edit.text(table.part);
    if (text === undefined) continue;
    const { root, elements } = scan(text, table.part, edit.limits, (local, depth, parent) => (depth === 2 && local === 'autoFilter') || (depth === 3 && local === 'tableColumn' && parent === 'tableColumns'));
    const edits: [number, number, string][] = [];
    if (newEnd !== null) {
      const end = newEnd;
      const grow = (r: CellRange): string => formatRange({ start: r.start, end: { row: end, col: r.end.col } });
      edits.push([root.start.start, root.start.end, setAttr(text.slice(root.start.start, root.start.end), 'ref', grow(range))]);
      const filter = elements.find((e) => e.local === 'autoFilter');
      const filterRef = filter?.attrs.get('ref');
      const filterRange = filterRef === undefined ? null : parseRange(filterRef.trim().toUpperCase());
      // A filter that ends on the table's last row grows with it; any other filter is left alone.
      if (filter !== undefined && filterRange !== null && filterRange.end.row === range.end.row) {
        edits.push([filter.start.start, filter.start.end, setAttr(text.slice(filter.start.start, filter.start.end), 'ref', grow(filterRange))]);
      }
    }
    elements
      .filter((e) => e.local === 'tableColumn')
      .forEach((e, i) => {
        const name = renames.get(i);
        if (name !== undefined) edits.push([e.start.start, e.start.end, setAttr(text.slice(e.start.start, e.start.end), 'name', name)]);
      });
    edit.write(table.part, splice(text, edits));
  }
}

// ---------------------------------------------------------------------------------------------
// Entry points

interface EditPlan {
  readonly sheet: Sheet;
  readonly cells: Map<string, { row: number; col: number; value: StoredValue | null; style: NormalizedStyle | null; date: string | number | null }>;
}

export function inputBytes(input: unknown): Uint8Array {
  if (input instanceof Uint8Array || input instanceof ArrayBuffer || ArrayBuffer.isView(input)) return toBytes(input);
  throw new XlvError('XLV_E_INVALID_INPUT', 'Editing needs the bytes of the original file.', {
    hint: 'Pass the file contents as a Uint8Array or ArrayBuffer (not a workbook from readWorkbook): unchanged parts are copied from the original bytes.',
  });
}

function isFormula(value: StoredValue | null): boolean {
  return value !== null && value.kind === 'formula';
}

/** Validates and converts edits with the writer's rules, collecting them per sheet. */
function planEdits(workbook: Workbook, edits: readonly unknown[]): Map<string, EditPlan> {
  const scratch = createWorkbook({ dateSystem: workbook.dateSystem });
  const scratchSheet = scratch.addSheet('Edit');
  const plans = new Map<string, EditPlan>();
  const arrayRanges = new Map<string, RangeLookup>();
  const byName = new Map(workbook.sheets.map((sheet) => [sheet.name, sheet]));
  edits.forEach((raw, i) => {
    const o = readOptions(raw, ['sheet', 'cell', 'value', 'style'], `edits[${String(i)}]`);
    const sheetName = o.get('sheet');
    if (typeof sheetName !== 'string') {
      throw new XlvError('XLV_E_INVALID_OPTION', `edits[${String(i)}] has no sheet name.`, { hint: 'Give each edit a sheet, for example { sheet: "Data", cell: "B2", value: 5 }.', details: { option: 'sheet' } });
    }
    const sheet = byName.get(sheetName) ?? workbook.getSheet(sheetName);
    if (sheet.kind !== 'worksheet' || sheet.part === null) {
      throw new XlvError('XLV_E_INVALID_INPUT', `Sheet "${sheet.name}" is a ${sheet.kind} without editable cells.`, { location: { sheet: sheet.name }, hint: 'Edit cells of worksheets only.' });
    }
    const cellRef = o.get('cell');
    const address = typeof cellRef === 'string' ? parseCellRef(cellRef) : null;
    if (address === null) {
      throw new XlvError('XLV_E_INVALID_CELL_REF', `"${String(cellRef).slice(0, 40)}" is not a valid cell reference.`, {
        location: { sheet: sheet.name },
        hint: 'Pass an A1-style reference such as "B7" (column letters A to XFD, rows 1 to 1048576).',
        details: { ref: String(cellRef).slice(0, 40) },
      });
    }
    const ref = formatCellRef(address.row, address.col);
    const value = o.get('value') ?? null;
    let style: NormalizedStyle | null = null;
    try {
      // Style errors carry the location of the edit, as value errors do (D-R03).
      style = o.has('style') ? normalizeStyle(o.get('style')) : null;
      scratchSheet.set(ref, value as CellInput, undefined);
    } catch (error) {
      if (error instanceof XlvError) {
        throw new XlvError(error.code, error.message, { location: { part: sheet.part, sheet: sheet.name, cell: ref }, hint: error.hint, details: error.details, cause: error });
      }
      throw error;
    }
    const stored = scratchSheet.cellMap.get(address.row)?.get(address.col);
    checkFormulaGroups(sheet, ref, address, stored?.value ?? null, arrayRanges);
    const scratchStyle = scratch.styleOf(stored?.style ?? 0);
    scratchSheet.set(ref, null, {});
    const dateValue = value instanceof Date ? value : typeof value === 'object' && value !== null && 'result' in value && (value as { result?: unknown }).result instanceof Date ? ((value as { result: Date }).result) : null;
    if (dateValue !== null && workbook.conformance === 'strict' && dateValue.getTime() < Date.UTC(1900, 2, 1)) {
      // Strict and Transitional serials differ before 1900-03-01 (D-106); the writer uses Transitional.
      throw new XlvError('XLV_E_INVALID_VALUE', `The date ${dateValue.toISOString()} for ${ref} cannot be written to a Strict workbook.`, {
        location: { part: sheet.part, sheet: sheet.name, cell: ref },
        hint: 'Dates before 1900-03-01 are not supported when editing Strict workbooks. Write the date as text.',
      });
    }
    let plan = plans.get(sheet.name);
    if (!plan) {
      plan = { sheet, cells: new Map() };
      plans.set(sheet.name, plan);
    }
    plan.cells.set(ref, { row: address.row, col: address.col, value: stored?.value ?? null, style, date: scratchStyle.numFmt ?? null });
  });
  return plans;
}

function applyEdits(bytes: Uint8Array, workbook: Workbook, plans: Map<string, EditPlan>, limits: ReadLimits, appended: AppendedRows | null = null): EditResult {
  const pkg = openPackage(openZip(bytes, limits), limits);
  const edit = new PackageEditor(bytes, limits, pkg);
  const warnings = new WarningCollector();
  const officeDocument = pkg.relationships(null).find((r) => classifyRelationship(r.type)?.name === 'officeDocument');
  const workbookPart = officeDocument?.target ?? '/xl/workbook.xml';
  const rels = pkg.relationships(workbookPart);
  const relOf = (name: string) => rels.find((r) => !r.external && classifyRelationship(r.type)?.name === name);
  const sstRel = relOf('sharedStrings');
  const stylesRel = relOf('styles');
  const sst = new SharedStringsEditor(edit, sstRel !== undefined && pkg.hasPart(sstRel.target) ? sstRel.target : null, workbookPart, workbook.conformance, warnings);
  const styles = new StylesEditor(edit, stylesRel !== undefined && pkg.hasPart(stylesRel.target) ? stylesRel.target : null, workbookPart, workbook.conformance, warnings);

  let formulasChanged = false;
  for (const plan of plans.values()) {
    const part = plan.sheet.part ?? '';
    const planned: PlannedCell[] = [];
    const text = edit.text(part);
    const sheetScan = text === undefined ? null : scanSheet(text, part, plan.sheet.name, limits, styles);
    // A cell that holds only a style keeps that style when it gets a value (D-E05). Only the cells
    // of this edit are kept, not every styled cell of the sheet.
    let styled: Map<string, Cell> | null = null;
    const styledCell = (ref: string): Cell | undefined => {
      if (styled === null) {
        const wanted = new Set<string>();
        for (const r of plan.cells.keys()) if (plan.sheet.cell(r) === undefined) wanted.add(r);
        styled = new Map();
        if (wanted.size > 0) for (const cell of plan.sheet.styledCells()) if (wanted.has(cell.ref)) styled.set(cell.ref, cell);
      }
      return styled.get(ref);
    };
    // Merged ranges in an index, so that each cell costs about the same however many the sheet has (A3).
    let merges: RangeLookup | null = null;
    for (const [ref, c] of plan.cells) {
      if (c.value !== null && plan.sheet.merges.length > 0) {
        if (merges === null) {
          merges = new RangeLookup();
          for (const m of plan.sheet.merges) {
            const range = parseRange(m.trim().toUpperCase());
            if (range !== null) merges.add(range, m);
          }
        }
        // Content of a merged range is stored in its top-left cell (E-35); a value elsewhere is not shown.
        const merge = merges.at(c.row, c.col, (r) => c.row !== r.start.row || c.col !== r.start.col);
        if (merge !== undefined) {
          const topLeft = formatCellRef(merge.range.start.row, merge.range.start.col);
          // Warnings of one sheet are grouped; `cells` lists the cells and the message names the first.
          warnings.add('XLV_W_MERGED_CELL_WRITE', `A value was written into a merged range outside its top-left cell (first: ${ref} in ${merge.ref}). Spreadsheet applications show only the top-left cell (${topLeft}), so the value will not be visible. Write it to the top-left cell, or unmerge the range.`, { part, sheet: plan.sheet.name, cell: ref });
        }
      }
      const existing: Cell | undefined = plan.sheet.cell(ref) ?? styledCell(ref);
      if (isFormula(c.value) || existing?.formula) formulasChanged = true;
      if (c.value === null && c.style === null) {
        planned.push({ row: c.row, col: c.col, ref, xml: null });
        continue;
      }
      // A new cell starts from its column's default style (col@style, W2-03).
      const base = existing?.styleIndex ?? (sheetScan === null ? 0 : newCellStyle(sheetScan, c.row, c.col));
      let patch: NormalizedStyle = c.style ?? {};
      const existingIsDate = existing?.numFmt != null && isDateFormat(existing.numFmt);
      if (c.date !== null && patch.numFmt === undefined && !existingIsDate) patch = { ...patch, numFmt: c.date };
      const xf = Object.keys(patch).length === 0 ? base : styles.derive(base, patch);
      planned.push({ row: c.row, col: c.col, ref, xml: newCellXml(ref, c.value, xf, sst) });
    }
    if (text === undefined || sheetScan === null) continue;
    syncTables(edit, plan, appended !== null && appended.sheet === plan.sheet.name ? appended : null);
    edit.write(part, editSheetXml(text, part, plan.sheet.name, sheetScan, planned));
  }
  sst.finish();
  styles.finish();

  let hasFormulas = formulasChanged;
  for (const sheet of workbook.sheets) {
    if (hasFormulas) break;
    for (const cell of sheet.cells()) {
      if (cell.formula !== null) {
        hasFormulas = true;
        break;
      }
    }
  }
  if (hasFormulas) {
    requestFullCalculation(edit, workbookPart);
    warnings.add('XLV_W_FORMULAS_NOT_RECALCULATED', 'The workbook contains formulas. xlvane does not recalculate them: cached results may be out of date until the file is opened in a spreadsheet application, which is asked to recalculate on load.', { part: workbookPart });
  }
  if (formulasChanged) {
    const chains = rels.filter((r) => !r.external && classifyRelationship(r.type)?.name === 'calcChain');
    for (const chain of chains) {
      edit.remove(chain.target);
      removeContentType(edit, chain.target);
    }
    if (chains.length > 0) removeRelationships(edit, workbookPart, new Set(chains.map((c) => c.id)));
  }
  return { bytes: edit.build(), warnings: warnings.list() as XlvWarning[] };
}

/**
 * Changes cells of an existing XLSX file and returns the new file. Only the edited sheets'
 * cell data changes, plus new entries in the shared strings and styles: charts, comments,
 * conditional formatting, data validation, defined names and every other part stay exactly as
 * they were. Each edit sets a value (as `NewSheet.set` does, with the same validation) and can
 * merge a style into the cell's current style. When the workbook contains formulas, the file
 * asks spreadsheet applications to recalculate on load (`XLV_W_FORMULAS_NOT_RECALCULATED`).
 * Editing a table's header cell renames the table column, so the value must be a unique,
 * non-empty name (`XLV_E_INVALID_VALUE` otherwise). A value written into a merged range outside
 * its top-left cell gives `XLV_W_MERGED_CELL_WRITE`.
 *
 * @example
 * ```ts
 * import { editCells } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const { bytes: updated, warnings } = editCells(bytes, [
 *   { sheet: 'Sales', cell: 'B7', value: 1250.5 },
 *   { sheet: 'Sales', cell: 'C7', value: 'checked', style: { font: { color: '008000' } } },
 * ]);
 * console.log(updated.length, warnings.length);
 * ```
 */
export function editCells(input: Uint8Array | ArrayBuffer, edits: readonly CellEdit[], options?: EditOptions): EditResult {
  const bytes = inputBytes(input);
  if (!Array.isArray(edits)) {
    throw new XlvError('XLV_E_INVALID_INPUT', 'editCells expects an array of edits.', { hint: 'Pass an array such as [{ sheet: "Data", cell: "A1", value: 1 }].' });
  }
  const o = readOptions(options, ['limits'], 'editCells options');
  const { limits, writable } = editLimits(o.get('limits'));
  // The size of the edit is checked before anything per cell is built (A2).
  checkWrites(edits.length, writable, `${String(edits.length)} edits`);
  // Each edit's own properties are read once (getters run here), so the cells counted are the
  // cells planned (D-K04).
  const copies = edits.map((e: unknown) => (typeof e === 'object' && e !== null && !Array.isArray(e) ? { ...e } : e)) as unknown as readonly CellEdit[];
  const workbook = readWorkbook(bytes, { limits });
  checkCellCount(workbook, newCells(workbook, copies), limits.maxCells);
  return applyEdits(bytes, workbook, planEdits(workbook, copies), limits);
}

/**
 * Appends rows below the last used row of a sheet (the rows of merged ranges count as used) in
 * an existing XLSX file, editing the file surgically like {@link editCells}. Object rows are
 * placed under the matching header columns (unknown keys throw `XLV_E_INVALID_ROWS` with
 * suggestions). Array rows start at the first used column, or, when the new rows extend a table
 * (one that reaches the last used row), at that table's first column; when several such tables
 * start in different columns the call is refused with `XLV_E_INVALID_ARGUMENT` (pass object rows
 * instead). A table whose last row is the sheet's last used row grows to include the new rows
 * (its range and filter); if that table has a totals row the call is refused with
 * `XLV_E_NOT_SUPPORTED` and a hint to insert the rows above the totals row with `insertRows`.
 * Calculated columns are not filled in the new rows. One call writes at most 1 000 000 cells
 * unless `limits.maxCells` is given (see `EditOptions`).
 *
 * @example
 * ```ts
 * import { appendRows } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const { bytes: updated } = appendRows(bytes, 'Orders', [
 *   { Customer: 'Alice', Total: 120.5 },
 *   { Customer: 'Bob', Total: 80 },
 * ]);
 * // Array rows that extend a table start at the table's first column.
 * const { bytes: more } = appendRows(updated, 'Orders', [['Carol', 64]]);
 * console.log(more.length);
 * ```
 */
export function appendRows(
  input: Uint8Array | ArrayBuffer,
  sheet: string,
  rows: readonly (readonly CellInput[])[] | readonly Readonly<Record<string, CellInput>>[],
  options?: AppendRowsOptions,
): EditResult {
  const bytes = inputBytes(input);
  const o = readOptions(options, ['limits', 'header'], 'appendRows options');
  const { limits, writable } = editLimits(o.get('limits'));
  const workbook = readWorkbook(bytes, { limits });
  const target = pickSheet(workbook, sheet);
  if (!Array.isArray(rows)) {
    throw new XlvError('XLV_E_INVALID_ROWS', 'appendRows expects an array of rows.', { hint: 'Pass an array of objects or arrays.' });
  }
  // The cells to write are counted before any is built (A2).
  let values = 0;
  for (const row of rows as readonly unknown[]) {
    if (Array.isArray(row)) {
      for (const value of row as readonly unknown[]) if (value !== null && value !== undefined) values++;
    } else if (typeof row === 'object' && row !== null) {
      for (const key of Object.keys(row)) if ((row as Record<string, unknown>)[key] !== null && (row as Record<string, unknown>)[key] !== undefined) values++;
    }
  }
  checkWrites(values, writable, `Appending ${String(rows.length)} rows`, target.name);
  checkCellCount(workbook, new Map([[target, values]]), limits.maxCells);
  let lastRow = 0;
  let firstCol = Infinity;
  for (const cell of target.cells()) {
    lastRow = Math.max(lastRow, cell.row);
    firstCol = Math.min(firstCol, cell.col);
  }
  if (firstCol === Infinity) firstCol = 1;
  // A merged range shows as one cell, so the rows it covers are used: a value appended inside it
  // would be hidden (E-35). The new rows start below every merge, also one without a value.
  for (const ref of target.merges) {
    const merge = parseRange(ref);
    if (merge !== null && merge.end.row > lastRow) lastRow = merge.end.row;
  }
  if (lastRow + rows.length > MAX_ROWS) {
    throw new XlvError('XLV_E_LIMIT_EXCEEDED', `Appending ${String(rows.length)} rows after row ${String(lastRow)} would pass the last row of the sheet.`, {
      location: { sheet: target.name },
      hint: `A sheet has ${String(MAX_ROWS)} rows. Append fewer rows or use another sheet.`,
      details: { limit: 'rows', value: MAX_ROWS },
    });
  }
  // Rows appended below a table that reaches the last used row join it (syncTables), so array rows
  // start at its first column rather than at the sheet's first used column (B8).
  let arrayCol = firstCol;
  if ((rows as readonly unknown[]).some((row) => Array.isArray(row))) {
    const joined = target.tables
      .map((table) => ({ table, range: parseRange(table.ref.trim().toUpperCase()) }))
      .filter((t): t is { table: (typeof target.tables)[number]; range: CellRange } => t.range !== null && lastRow > 0 && t.range.end.row >= lastRow && t.range.start.row <= lastRow);
    const starts = new Set(joined.map((t) => t.range.start.col));
    if (starts.size > 1) {
      const names = joined.map((t) => `"${t.table.displayName}" (${t.table.ref})`).join(', ');
      throw new XlvError('XLV_E_INVALID_ARGUMENT', `Tables ${names} of sheet "${target.name}" reach the last used row ${String(lastRow)} and start in different columns, so it is not clear in which column array rows start.`, {
        location: { sheet: target.name },
        hint: 'Pass the rows as objects keyed by header text, for example [{ Customer: "Dan", Total: 42 }], so that each value goes under its column (use { header: n } to name the header row).',
        details: { tables: joined.map((t) => t.table.displayName), row: lastRow },
      });
    }
    const [start] = starts;
    if (start !== undefined) arrayCol = start;
  }
  let columns: Map<string, number> | null = null;
  const edits: CellEdit[] = [];
  rows.forEach((row: unknown, i) => {
    const r = lastRow + 1 + i;
    if (Array.isArray(row)) {
      row.forEach((value: unknown, k) => {
        if (value !== null && value !== undefined) edits.push({ sheet: target.name, cell: formatCellRef(r, arrayCol + k), value: value as CellInput });
      });
      return;
    }
    if (typeof row !== 'object' || row === null) {
      throw new XlvError('XLV_E_INVALID_ROWS', `Row ${String(i + 1)} is not an object or an array.`, { location: { sheet: target.name }, hint: 'Pass objects keyed by header text, or arrays of values.', details: { row: i + 1 } });
    }
    if (columns === null) {
      const header = o.get('header') ?? 'auto';
      const table = readTable(workbook, { sheet: target.name, header: header as 'auto' | number });
      if (table.headerRow === null) {
        throw new XlvError('XLV_E_HEADER_NOT_FOUND', `Sheet "${target.name}" has no header row to place object rows under.`, {
          location: { sheet: target.name },
          hint: 'Pass the header row with { header: 1 }, or pass rows as arrays of values.',
          details: { row: 0 },
        });
      }
      columns = new Map(table.columns.map((c) => [c.key, c.col]));
    }
    const known = columns;
    for (const key of Object.keys(row)) {
      const col = known.get(key);
      if (col === undefined) {
        const close = suggest(key, [...known.keys()]);
        throw new XlvError('XLV_E_INVALID_ROWS', `Row ${String(i + 1)} has the key "${key.slice(0, 40)}", which is not a column of sheet "${target.name}".`, {
          location: { sheet: target.name },
          hint: `${close.length > 0 ? `Did you mean "${close[0] ?? ''}"? ` : ''}Columns: ${[...known.keys()].join(', ')}.`,
          details: { row: i + 1, key: key.slice(0, 40) },
        });
      }
      const value = (row as Record<string, unknown>)[key];
      if (value !== null && value !== undefined) edits.push({ sheet: target.name, cell: formatCellRef(r, col), value: value as CellInput });
    }
  });
  if (edits.length === 0) return { bytes: bytes.slice(), warnings: [] };
  return applyEdits(bytes, workbook, planEdits(workbook, edits), limits, { sheet: target.name, lastRow, count: rows.length });
}

