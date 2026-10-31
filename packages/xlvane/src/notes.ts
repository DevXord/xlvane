// Cell notes, the "comments" of ECMA-376 (Part 1 §18.7, §12.3.3; Part 4 §8.2, §15.4.1.1,
// §19.4; [MS-OI29500] §2.1.689 … §2.1.692; docs/facts/sml-write.md W2-45 … W2-48;
// docs/decisions.md D-A12). The text and author of a note are stored in the sheet's Comments
// part; the box that shows it is a VML shape in a VML Drawing part, which the reader ignores.
import { XlvError } from './errors';
import type { RichTextRun } from './model';
import { parseCellRef, formatCellRef } from './sml/cell-ref';
import { SML_NS } from './sml/namespaces';
import type { PartContext } from './sml/shared-strings';
import { parseRunFont } from './sml/styles';
import { decodeXstring } from './sml/xstring';
import { childNamed, childrenNamed, parseXmlTree, type XmlElement } from './xml/tree';

/**
 * A note (comment) of a cell, read from the sheet's Comments part. `text` is the plain text;
 * `runs` holds the formatted parts when the text is rich text. Excel usually starts the text
 * with the author's name in bold, as part of the text itself.
 *
 * @example
 * ```ts
 * import { readWorkbook, type Note } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const notes: readonly Note[] = readWorkbook(bytes).getSheet('Sales').notes;
 * for (const note of notes) console.log(note.ref, note.author, note.text);
 * ```
 */
export interface Note {
  /** The cell the note belongs to, for example `"B4"`. */
  readonly ref: string;
  /** The text of the note, without formatting. */
  readonly text: string;
  /** The author as stored in the file (may be `""`), or `null` if the file names none. */
  readonly author: string | null;
  /** The formatted parts of the text, when it is rich text. */
  readonly runs?: readonly RichTextRun[];
  /** The settings of the box that shows the note, read from the VML part, when the file has one for this cell. */
  readonly box?: NoteBox;
}

/**
 * The box that shows a note (a VML text box shape, Part 4 §19.1.2.22, §19.4; docs/facts/sml-write.md
 * W2-69). `margins` are the inner margins in millimetres, in the order of Part 4 §19.1.2.22
 * (left, top, right, bottom); with `autoMargins` the application calculates them instead
 * (ST_InsetMode `auto`, §19.2.3.17). `locked` and `lockText` apply only on a protected sheet
 * (§19.4.2.37, §19.4.2.38; omitted means locked). `anchor` tells how the box follows the cells:
 * `"twoCell"` moves and sizes with them, `"oneCell"` moves but keeps its size, `"absolute"` does
 * neither (x:MoveWithCells / x:SizeWithCells, whose meaning Excel reverses, [MS-OI29500]
 * §2.1.1881, §2.1.1886).
 *
 * @example
 * ```ts
 * import { readWorkbook, type NoteBox } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const box: NoteBox | undefined = readWorkbook(bytes).getSheet('Sales').notes[0]?.box;
 * console.log(box?.margins, box?.locked);
 * ```
 */
export interface NoteBox {
  readonly autoMargins: boolean;
  readonly margins: NoteMargins | null;
  readonly locked: boolean;
  readonly lockText: boolean;
  readonly anchor: 'twoCell' | 'oneCell' | 'absolute';
}

/**
 * Inner margins of a note box in millimetres (left, top, right, bottom; Part 4 §19.1.2.22).
 *
 * @example
 * ```ts
 * import type { NoteMargins } from 'xlvane';
 *
 * const margins: NoteMargins = { left: 1.3, top: 1.3, right: 2.5, bottom: 2.5 };
 * console.log(margins.left);
 * ```
 */
export interface NoteMargins {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

/**
 * Settings for the box of a new note; every one is optional. Defaults: margins calculated by
 * the application (`autoMargins`, unless `margins` is given), box and text locked, `anchor`
 * `"absolute"`.
 *
 * @example
 * ```ts
 * import { createWorkbook, type NoteBoxInput } from 'xlvane';
 *
 * const box: NoteBoxInput = { margins: { left: 2, top: 1, right: 2, bottom: 1 }, lockText: false, anchor: 'twoCell' };
 * createWorkbook().addSheet('Data').note('A1', { text: 'Editable on a protected sheet', box });
 * ```
 */
export interface NoteBoxInput {
  readonly autoMargins?: boolean;
  readonly margins?: NoteMargins;
  readonly locked?: boolean;
  readonly lockText?: boolean;
  readonly anchor?: 'twoCell' | 'oneCell' | 'absolute';
}

/**
 * A note for the writers: plain text, or an object with `text` or rich text `runs` (one of the
 * two) and an optional `author`. Limits follow [MS-OI29500]: at most 32 767 characters of text
 * (§2.1.678 b) in at most 32 767 runs (§2.1.630 a), and an author of at most 65 535 characters
 * (§2.1.689 a). Without an author, the note is stored with the author `""`. `box` sets the
 * box that shows the note (see {@link NoteBoxInput}).
 *
 * @example
 * ```ts
 * import { createWorkbook, type NoteInput } from 'xlvane';
 *
 * const plain: NoteInput = 'Checked by finance';
 * const rich: NoteInput = { runs: [{ text: 'Ann: ', font: { bold: true } }, { text: 'please confirm' }], author: 'Ann' };
 * createWorkbook().addSheet('Data').set('A1', 1250).note('A1', plain).note('B1', rich);
 * ```
 */
export type NoteInput = string | { readonly text?: string; readonly runs?: readonly RichTextRun[]; readonly author?: string; readonly box?: NoteBoxInput };

/** Comments part content type (Part 1 §12.3.3, W2-46). */
export const COMMENTS_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
/** VML Drawing part content type (Part 4 §8.2, W2-47). */
export const VML_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.vmlDrawing';
/** Relationship type of a threaded comments part ([MS-XLSX] §2.1.17). */
export const THREADED_COMMENTS_REL = 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment';
/** Most characters of an author ([MS-OI29500] §2.1.689 a). */
export const MAX_AUTHOR_LENGTH = 65_535;

/**
 * Reads a Comments part (CT_Comments, W2-45). A comment whose `ref` is not a single cell, or
 * whose `authorId` names no author, is read with a warning (skipped, or with author `null`).
 * Malformed XML or another root element gives a warning and no notes: notes are annotations, so
 * the sheet stays readable without them. Limits and DTD rules of the XML parser still apply.
 */
export function parseComments(xml: string, ctx: PartContext & { readonly sheet: string }): Note[] {
  const warn = (message: string, cell?: string): void => ctx.warnings.add('XLV_W_INVALID_VALUE', message, { part: ctx.part, sheet: ctx.sheet, ...(cell === undefined ? {} : { cell }) });
  let root: XmlElement;
  try {
    root = parseXmlTree(xml, ctx);
  } catch (error) {
    if (error instanceof XlvError && error.code === 'XLV_E_XML_MALFORMED') {
      warn('The comments part of the sheet is not well-formed XML; its notes are ignored.');
      return [];
    }
    throw error;
  }
  const ns = root.ns;
  if (root.local !== 'comments' || (ns !== SML_NS.strict && ns !== SML_NS.transitional)) {
    warn(`The comments part of the sheet has the root element <${root.local}>; its notes are ignored.`);
    return [];
  }
  const authors = childrenNamed(childNamed(root, ns, 'authors') ?? root, ns, 'author').map((a) => decodeXstring(a.text));
  const list = childNamed(root, ns, 'commentList');
  const notes: Note[] = [];
  if (list === undefined) return notes;
  for (const c of childrenNamed(list, ns, 'comment')) {
    const rawRef = c.attrs.get('ref') ?? '';
    const address = parseCellRef(rawRef.trim().toUpperCase());
    if (address === null) {
      warn(`A note refers to "${rawRef.slice(0, 40)}", which is not a single cell; it is skipped.`);
      continue;
    }
    const ref = formatCellRef(address.row, address.col);
    const idText = c.attrs.get('authorId');
    const id = idText === undefined || !/^\d+$/.test(idText.trim()) ? -1 : Number(idText.trim());
    const author = authors[id];
    if (author === undefined) warn(`The note of ${ref} names author ${String(idText ?? '(none)').slice(0, 20)}, which is not in the list of authors; its author is unknown.`, ref);
    const { text, runs } = readRst(childNamed(c, ns, 'text'), ns, (m) => warn(m, ref));
    notes.push(Object.freeze(runs === undefined ? { ref, text, author: author ?? null } : { ref, text, author: author ?? null, runs }));
  }
  return notes;
}

/**
 * Reads CT_Rst (`t?`, `r*`, `rPh*`): the text is the own `t` and the `t` of every run;
 * phonetic runs are not part of it (as for shared strings, D-SST-2).
 */
function readRst(el: XmlElement | undefined, ns: string, warn: (message: string) => void): { text: string; runs?: readonly RichTextRun[] } {
  if (el === undefined) return { text: '' };
  let text = '';
  const runs: RichTextRun[] = [];
  let rich = false;
  for (const child of el.children) {
    if (child.ns !== ns) continue;
    if (child.local === 't') {
      const t = decodeXstring(child.text);
      text += t;
      runs.push({ text: t });
    } else if (child.local === 'r') {
      rich = true;
      const t = decodeXstring(childNamed(child, ns, 't')?.text ?? '');
      text += t;
      const rPr = childNamed(child, ns, 'rPr');
      const font = rPr === undefined ? {} : parseRunFont(rPr, ns, warn);
      runs.push(Object.keys(font).length > 0 ? Object.freeze({ text: t, font }) : Object.freeze({ text: t }));
    }
  }
  return rich ? { text, runs: Object.freeze(runs) } : { text };
}

const VML_NS = 'urn:schemas-microsoft-com:vml';
const OFFICE_NS = 'urn:schemas-microsoft-com:office:office';
const EXCEL_NS = 'urn:schemas-microsoft-com:office:excel';

/** Millimetres per unit of the VML length units (Part 4 §19.1; px at 96 per inch, a convention, W2-69). */
const MM_PER_UNIT: Readonly<Record<string, number>> = { mm: 1, cm: 10, in: 25.4, pt: 25.4 / 72, pc: 25.4 / 6, px: 25.4 / 96 };
/** Part 4 §19.1.2.22: the default inset, "0.1in, 0.05in, 0.1in, 0.05in", in millimetres. */
const DEFAULT_INSET_MM = [2.54, 1.27, 2.54, 1.27] as const;

/** ST_TrueFalseBlank (Part 4 §20.1.2.6): blank and t/true mean true. */
function trueFalseBlank(text: string): boolean | null {
  const v = text.trim().toLowerCase();
  if (v === '' || v === 't' || v === 'true') return true;
  if (v === 'f' || v === 'false') return false;
  return null;
}

/** Parses the `inset` of a text box into millimetres; `null` when a value has no or an unknown unit. */
function parseInset(inset: string): NoteMargins | null {
  const parts = inset.split(/\s*,\s*|\s+/);
  const out: number[] = [];
  for (let i = 0; i < 4; i++) {
    const p = (parts[i] ?? '').trim();
    if (p === '') {
      out.push(DEFAULT_INSET_MM[i]!);
      continue;
    }
    const m = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+))(mm|cm|in|pt|pc|px)$/.exec(p);
    if (m === null) return null;
    out.push(Math.round(Number(m[1]) * MM_PER_UNIT[m[2]!]! * 10000) / 10000);
  }
  return { left: out[0]!, top: out[1]!, right: out[2]!, bottom: out[3]! };
}

/**
 * Reads the note boxes of a VML Drawing part (Part 4 §19.4; W2-69): for every shape with an
 * `x:ClientData` of type `Note`, its cell (`x:Row`, `x:Column`, 0-based) and its settings.
 * Malformed VML gives a warning and no boxes; the notes stay readable without them.
 */
export function parseNoteBoxes(xml: string, ctx: PartContext & { readonly sheet: string }): Map<string, NoteBox> {
  const boxes = new Map<string, NoteBox>();
  const warn = (message: string): void => ctx.warnings.add('XLV_W_INVALID_VALUE', message, { part: ctx.part, sheet: ctx.sheet });
  let root: XmlElement;
  try {
    root = parseXmlTree(xml, ctx);
  } catch (error) {
    if (error instanceof XlvError && error.code === 'XLV_E_XML_MALFORMED') {
      warn('The VML drawing of the sheet is not well-formed XML; the settings of its note boxes are not read.');
      return boxes;
    }
    throw error;
  }
  const visit = (el: XmlElement): void => {
    if (el.ns === VML_NS && el.local === 'shape') {
      const data = el.children.find((c) => c.ns === EXCEL_NS && c.local === 'ClientData');
      if (data !== undefined && (data.attrs.get('ObjectType') ?? '').trim() === 'Note') readShape(el, data);
      return;
    }
    for (const c of el.children) visit(c);
  };
  const readShape = (shape: XmlElement, data: XmlElement): void => {
    const child = (name: string): XmlElement | undefined => data.children.find((c) => c.ns === EXCEL_NS && c.local === name);
    const row = Number((child('Row')?.text ?? '').trim());
    const col = Number((child('Column')?.text ?? '').trim());
    if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0 || row >= 1_048_576 || col >= 16_384 || child('Row') === undefined || child('Column') === undefined) return;
    const flag = (name: string): boolean => {
      const el = child(name);
      if (el === undefined) return true;
      const v = trueFalseBlank(el.text);
      if (v === null) warn(`The note box of row ${String(row + 1)}, column ${String(col + 1)} has x:${name} "${el.text.trim().slice(0, 20)}"; it is read as true.`);
      return v ?? true;
    };
    const move = child('MoveWithCells') !== undefined;
    const size = child('SizeWithCells') !== undefined;
    const textbox = shape.children.find((c) => c.ns === VML_NS && c.local === 'textbox');
    const insetText = textbox?.attrs.get('inset');
    let margins: NoteMargins | null = null;
    if (insetText !== undefined) {
      margins = parseInset(insetText);
      if (margins === null) warn(`The note box of row ${String(row + 1)}, column ${String(col + 1)} has margins "${insetText.slice(0, 40)}" without a known unit; they are not read.`);
    }
    // The shape's o:insetmode decides (Excel reads it from the parent, [MS-OI29500] §2.1.1811 f/i);
    // its default is custom (Part 4 §19.1.2.22, §19.2.3.17).
    const mode = (shape.attrs.get(`{${OFFICE_NS}}insetmode`) ?? textbox?.attrs.get(`{${OFFICE_NS}}insetmode`) ?? 'custom').trim();
    boxes.set(formatCellRef(row + 1, col + 1), Object.freeze({
      autoMargins: mode === 'auto',
      margins,
      locked: flag('Locked'),
      lockText: flag('LockText'),
      anchor: move && size ? 'absolute' : move || size ? 'oneCell' : 'twoCell',
    }));
  };
  visit(root);
  return boxes;
}
