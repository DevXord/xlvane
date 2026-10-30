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
}

/**
 * A note for the writers: plain text, or an object with `text` or rich text `runs` (one of the
 * two) and an optional `author`. Limits follow [MS-OI29500]: at most 32 767 characters of text
 * (§2.1.678 b) in at most 32 767 runs (§2.1.630 a), and an author of at most 65 535 characters
 * (§2.1.689 a). Without an author, the note is stored with the author `""`.
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
export type NoteInput = string | { readonly text?: string; readonly runs?: readonly RichTextRun[]; readonly author?: string };

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
