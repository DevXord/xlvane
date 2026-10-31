// Conversion between the exceljs model and xlvane (docs/decisions.md D-503, D-504).
import {
  createWorkbook,
  readWorkbook,
  type CellInput,
  type CellStyle,
  type ConditionalFormatRuleInput,
  type DataValidationRule,
  type HeaderFooter as XlvHeaderFooter,
  type PageSetupOptions,
  type PrintTitles,
  type SheetProtectionOptions,
  type Color as XlvColor,
  type ColumnOptions,
  type FontStyle,
  type Formula as XlvFormula,
  type FrozenPanes,
  type HyperlinkTarget,
  type NewWorkbook,
  type Note as XlvNote,
  type NoteBox,
  type NoteBoxInput,
  type NoteInput,
  type RichTextRun,
  type SheetPropertiesOptions,
  type SheetView,
  type SheetViewOptions,
  type WorkbookViewOptions,
  type Workbook as XlvWorkbook,
  type WorkbookProperties,
  XlvError,
} from 'xlvane';
import { imageAnchorIn, imageBytes, tableInput, type ImageData } from './media';
import { columnLetter, decodeAddress, isJsonValue, notSupported, ValueType, type Cell, type Column, type Worksheet } from './model';
import { conditionalFormatsIn, dataValidationsIn, headerFooterIn, pageSetupIn, protectionIn, rowBreaksIn } from './sheet-rules';
import type { Alignment, Borders, CellValue, Color, Comment, Fill, Font, Style, WorksheetView } from './types';

/**
 * Format codes exceljs writes as built-in ids (lib/xlsx/defaultnumformats.js, used by
 * NumFmtXform.getDefaultFmtId, lib/xlsx/xform/style/numfmt-xform.js L55; X4-01). Every id here is
 * one xlvane accepts as a built-in id (D-A05). `General` (0) means no format.
 */
const BUILTIN_CODES: ReadonlyMap<string, number> = new Map([
  ['0', 1],
  ['0.00', 2],
  ['#,##0', 3],
  ['#,##0.00', 4],
  ['0%', 9],
  ['0.00%', 10],
  ['0.00E+00', 11],
  ['# ?/?', 12],
  ['# ??/??', 13],
  ['mm-dd-yy', 14],
  ['d-mmm-yy', 15],
  ['d-mmm', 16],
  ['mmm-yy', 17],
  ['h:mm AM/PM', 18],
  ['h:mm:ss AM/PM', 19],
  ['h:mm', 20],
  ['h:mm:ss', 21],
  ['m/d/yy "h":mm', 22],
  ['#,##0 ;(#,##0)', 37],
  ['#,##0 ;[Red](#,##0)', 38],
  ['#,##0.00 ;(#,##0.00)', 39],
  ['#,##0.00 ;[Red](#,##0.00)', 40],
  ['mm:ss', 45],
  ['[h]:mm:ss', 46],
  ['mmss.0', 47],
  ['##0.0E+0', 48],
  ['@', 49],
]);

/** exceljs writes dates without a number format with built-in format 14 (styles-xform.js L260; X4-02). */
export const DATE_NUMFMT = 14;

// ---------------------------------------------------------------------------------------------
// exceljs → xlvane

function colorIn(c: Partial<Color> | undefined, where: string): XlvColor | undefined {
  if (c === undefined || c === null) return undefined;
  if (typeof c.argb === 'string') return c.argb;
  if (typeof c.theme === 'number') return c.tint ? { theme: c.theme, tint: c.tint } : { theme: c.theme };
  if (typeof c.indexed === 'number') return { indexed: c.indexed };
  if (Object.keys(c).length === 0) return undefined;
  return notSupported(`The colour ${JSON.stringify(c).slice(0, 60)} in ${where}`);
}

/** A colour of a conditional format or another setting given as an exceljs colour object. */
export function colorInFor(c: unknown, where: string): XlvColor {
  const out = colorIn(c as Partial<Color> | undefined, where);
  return out === undefined ? notSupported(`A missing or empty colour in ${where}`) : out;
}

function fontIn(f: Partial<Font> | undefined): CellStyle['font'] {
  if (!f) return undefined;
  const out: Record<string, unknown> = {};
  if (f.name !== undefined) out['name'] = f.name;
  if (f.size !== undefined) out['size'] = f.size;
  if (f.bold) out['bold'] = true;
  if (f.italic) out['italic'] = true;
  if (f.strike) out['strike'] = true;
  if (f.underline && f.underline !== 'none') out['underline'] = f.underline === true ? 'single' : f.underline;
  const color = colorIn(f.color, 'font.color');
  if (color !== undefined) out['color'] = color;
  // The remaining font properties of exceljs (font-xform.js L20–37; D-E04), also in rich-text runs.
  if (f.vertAlign) out['vertAlign'] = f.vertAlign;
  if (typeof f.family === 'number') out['family'] = f.family;
  if (f.scheme) out['scheme'] = f.scheme;
  if (typeof f.charset === 'number') out['charset'] = f.charset;
  for (const key of ['outline', 'shadow', 'condense', 'extend'] as const) if ((f as Record<string, unknown>)[key]) out[key] = true;
  return out as CellStyle['font'];
}

function fillIn(f: Fill | undefined): CellStyle['fill'] {
  if (!f) return undefined;
  if (f.type === 'gradient') return gradientIn(f as unknown as Record<string, unknown>);
  if (f.type !== 'pattern') return notSupported(`A fill of type "${String(f.type)}"`);
  if (!f.pattern || f.pattern === 'none') return undefined;
  const out: Record<string, unknown> = { pattern: f.pattern };
  const fg = colorIn(f.fgColor, 'fill.fgColor');
  if (fg !== undefined) out['fgColor'] = fg;
  const bg = colorIn(f.bgColor, 'fill.bgColor');
  if (bg !== undefined) out['bgColor'] = bg;
  return out as CellStyle['fill'];
}

/**
 * An exceljs gradient fill (fill-xform.js L132–258; X4-27): `angle` becomes a linear gradient
 * with its `degree`; `path` takes `center.left`/`top`, and `right`/`bottom` default to them.
 */
function gradientIn(f: Record<string, unknown>): CellStyle['fill'] {
  const stops = Array.isArray(f['stops']) ? (f['stops'] as { position?: unknown; color?: unknown }[]) : [];
  const out: Record<string, unknown> = {
    type: 'gradient',
    stops: stops.map((st, i) => ({ position: st.position, color: colorInFor(st.color, `fill.stops[${String(i)}].color`) })),
  };
  if (f['gradient'] === 'path') {
    out['gradient'] = 'path';
    const c = (f['center'] ?? {}) as Record<string, unknown>;
    const left = typeof c['left'] === 'number' ? c['left'] : 0;
    const top = typeof c['top'] === 'number' ? c['top'] : 0;
    out['left'] = left;
    out['right'] = typeof c['right'] === 'number' ? c['right'] : left;
    out['top'] = top;
    out['bottom'] = typeof c['bottom'] === 'number' ? c['bottom'] : top;
  } else {
    out['gradient'] = 'linear';
    if (typeof f['degree'] === 'number') out['degree'] = f['degree'];
  }
  return out as CellStyle['fill'];
}

function borderIn(b: Partial<Borders> | undefined): CellStyle['border'] {
  if (!b) return undefined;
  const out: Record<string, unknown> = {};
  for (const edge of ['top', 'right', 'bottom', 'left', 'diagonal'] as const) {
    const e = b[edge];
    if (!e || !e.style) continue;
    const color = colorIn(e.color, `border.${edge}.color`);
    out[edge] = color === undefined ? { style: e.style } : { style: e.style, color };
  }
  // exceljs writes diagonalUp/diagonalDown only with a diagonal style (border-xform.js L122–149).
  if (b.diagonal?.style) {
    if (b.diagonal.up) out['diagonalUp'] = true;
    if (b.diagonal.down) out['diagonalDown'] = true;
  }
  return Object.keys(out).length > 0 ? (out as CellStyle['border']) : undefined;
}

/** exceljs `utils.validInt`: `parseInt(value, 10)`, or 0 when that is NaN. */
function validInt(value: unknown): number {
  const i = Number.parseInt(String(value), 10);
  return Number.isNaN(i) ? 0 : i;
}

function alignmentIn(a: Partial<Alignment> | undefined): CellStyle['alignment'] {
  if (!a) return undefined;
  const out: Record<string, unknown> = {};
  if (a.horizontal) out['horizontal'] = a.horizontal;
  // exceljs calls the centre "middle" (index.d.ts Alignment L282); SpreadsheetML calls it "center".
  if (a.vertical) out['vertical'] = a.vertical === 'middle' ? 'center' : a.vertical;
  if (a.wrapText) out['wrapText'] = true;
  if (a.shrinkToFit) out['shrinkToFit'] = true;
  // exceljs validInt (utils.js L135): parseInt, NaN → 0; the indent is at least 0 (alignment-xform.js; X4-27).
  const indent = Math.max(0, validInt(a.indent));
  if (indent) out['indent'] = indent;
  // exceljs keeps "vertical" (→ 255) or a whole number from −90 to 90 (−90…−1 → 90 − tr) and
  // drops other values (alignment-xform.js validation.textRotation, textRotationXform; X4-27).
  if (a.textRotation === 'vertical') out['textRotation'] = 255;
  else if (a.textRotation !== undefined && a.textRotation !== null) {
    const tr = validInt(a.textRotation);
    if (tr > 0 && tr <= 90) out['textRotation'] = tr;
    else if (tr < 0 && tr >= -90) out['textRotation'] = 90 - tr;
  }
  if (a.readingOrder === 'ltr') out['readingOrder'] = 'leftToRight';
  else if (a.readingOrder === 'rtl') out['readingOrder'] = 'rightToLeft';
  return Object.keys(out).length > 0 ? (out as CellStyle['alignment']) : undefined;
}

/** exceljs `{ locked, hidden }` (protection-xform.js; X4-27); only values that differ from the defaults matter. */
function cellProtectionIn(p: unknown): CellStyle['protection'] {
  if (p === undefined || p === null || typeof p !== 'object') return undefined;
  const r = p as Record<string, unknown>;
  const out: Record<string, boolean> = {};
  if (r['locked'] === false) out['locked'] = false;
  if (r['hidden'] === true) out['hidden'] = true;
  return Object.keys(out).length > 0 ? out : undefined;
}

/** Converts an exceljs style to an xlvane CellStyle (unsupported parts throw XLV_E_NOT_SUPPORTED). */
export function styleIn(s: Style): CellStyle {
  const out: Record<string, unknown> = {};
  const protection = cellProtectionIn(s.protection);
  if (protection) out['protection'] = protection;
  const font = fontIn(s.font);
  if (font && Object.keys(font).length > 0) out['font'] = font;
  const fill = fillIn(s.fill);
  if (fill) out['fill'] = fill;
  const border = borderIn(s.border);
  if (border) out['border'] = border;
  const alignment = alignmentIn(s.alignment);
  if (alignment) out['alignment'] = alignment;
  if (s.numFmt && s.numFmt !== 'General') out['numFmt'] = BUILTIN_CODES.get(s.numFmt) ?? s.numFmt;
  return out as CellStyle;
}

/** True when exceljs gives the cell the date format: a Date value or a formula with a Date result (cell-xform.js L36). */
export function isDateCell(cell: Cell): boolean {
  const v = cell.value;
  if (v instanceof Date) return true;
  return isRecord(v) && ('formula' in v || 'sharedFormula' in v) && v['result'] instanceof Date;
}

/** The style written for a cell: its style, with format 14 for a date without a format (X4-02). */
export function cellStyleIn(cell: Cell, style: CellStyle): CellStyle {
  return style.numFmt === undefined && isDateCell(cell) ? { ...style, numFmt: DATE_NUMFMT } : style;
}

const DEFAULT_COLUMN_WIDTH = 9;

/**
 * The `col` options for a column, as exceljs `Column.toModel` (lib/doc/column.js L265; X4-04):
 * a column is written unless it is a default one (width unset or 9, visible, no style), and a
 * written column without a width gets width 9. `withStyle` false drops the style, as the
 * streaming writer does without `useStyles`.
 */
export function columnIn(column: Column, withStyle: boolean): ColumnOptions | undefined {
  const s = column.style;
  const styled = !!(s.font || s.numFmt || s.alignment || s.border || s.fill || s.protection);
  const customWidth = column.width !== undefined && column.width !== DEFAULT_COLUMN_WIDTH;
  if (!customWidth && !column.hidden && !styled && !column.outlineLevel) return undefined;
  const style = withStyle && styled ? styleIn(s) : undefined;
  return {
    width: column.width ?? DEFAULT_COLUMN_WIDTH,
    ...(column.hidden ? { hidden: true } : {}),
    ...(style && Object.keys(style).length > 0 ? { style } : {}),
    // col-xform.js L32–36: outlineLevel and the computed collapsed flag (X4-28).
    ...(column.outlineLevel ? { outlineLevel: column.outlineLevel } : {}),
    ...(column.collapsed ? { collapsed: true } : {}),
  };
}

function runsIn(runs: unknown[], address: string): RichTextRun[] {
  if (runs.length === 0) return notSupported(`An empty { richText: [] } value in cell ${address}`);
  return runs.map((r) => {
    if (!isRecord(r)) return notSupported(`A rich text run that is not an object in cell ${address}`);
    const text = r['text'] === undefined || r['text'] === null ? '' : String(r['text']);
    const font = fontIn(r['font'] as Partial<Font> | undefined);
    return font && Object.keys(font).length > 0 ? { text, font: font as FontStyle } : { text };
  });
}

function isRecord(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !(v instanceof Date) && !Array.isArray(v);
}

/** The `result` part of a formula value for the writer. */
export function resultIn(r: unknown): { result?: CellInput } {
  if (r === undefined || r === null) return {};
  if (isRecord(r) && typeof r['error'] === 'string') return { result: { error: r['error'] as '#N/A' } } as { result: CellInput };
  return { result: r as number };
}

/** The cached result of a formula as a plain value (`null` without one). */
function resultValue(r: unknown): CellInput {
  return resultIn(r).result ?? null;
}

/**
 * The range of each shared formula of a worksheet, by master address, as exceljs computes it on
 * write (cell-xform.js L94–160; X4-29): from the master to the last of its dependents, or the
 * `ref` of a `{ formula, ref }` master that is not an array formula. Only groups the core can write
 * as shared formulas are returned.
 */
export function sharedFormulaRanges(ws: Worksheet): Map<string, string> {
  type Box = { top: number; left: number; bottom: number; right: number };
  const bounds = new Map<string, Box>();
  const arrays: Box[] = [];
  const grow = (master: string, row: number, col: number): void => {
    const b = bounds.get(master);
    if (!b) {
      const m = decodeAddress(master);
      bounds.set(master, { top: Math.min(m.row, row), left: Math.min(m.col, col), bottom: Math.max(m.row, row), right: Math.max(m.col, col) });
      return;
    }
    b.top = Math.min(b.top, row);
    b.left = Math.min(b.left, col);
    b.bottom = Math.max(b.bottom, row);
    b.right = Math.max(b.right, col);
  };
  const box = (ref: string, fallback: string): Box => {
    const [a, b = a] = ref.split(':');
    const s = decodeAddress(a ?? fallback);
    const e = decodeAddress(b ?? fallback);
    return { top: Math.min(s.row, e.row), left: Math.min(s.col, e.col), bottom: Math.max(s.row, e.row), right: Math.max(s.col, e.col) };
  };
  const formulaMasters = new Set<string>();
  for (const row of ws._rows) {
    for (const cell of row?._cells ?? []) {
      const v = cell?.value;
      if (!cell || !isRecord(v)) continue;
      if (typeof v['sharedFormula'] === 'string') grow(v['sharedFormula'].replace(/\$/g, '').toUpperCase(), cell.row, cell.col);
      else if (typeof v['formula'] === 'string') {
        formulaMasters.add(cell.address);
        if (typeof v['ref'] !== 'string') continue;
        const r = box(v['ref'], cell.address);
        if (v['shareType'] === 'array') arrays.push(r);
        else {
          grow(cell.address, r.top, r.left);
          grow(cell.address, r.bottom, r.right);
        }
      }
    }
  }
  // A group is written as a shared formula only when the core can write it: its master holds a
  // formula, is the top-left cell of the range, and the range overlaps no other group or array
  // formula (F2-15). Other groups, which files can contain, are written as ordinary formulas by
  // valueIn (docs/decisions.md D-R01).
  const groups = [...bounds].filter(([master, b]) => {
    if (!formulaMasters.has(master)) return false;
    const m = decodeAddress(master);
    return m.row === b.top && m.col === b.left;
  });
  const overlaps = (a: Box, b: Box): boolean => a.top <= b.bottom && b.top <= a.bottom && a.left <= b.right && b.left <= a.right;
  const bad = new Set<string>();
  const sorted = [...groups].sort((x, y) => x[1].top - y[1].top);
  for (let i = 0; i < sorted.length; i++) {
    const [mi, bi] = sorted[i]!;
    for (let j = i + 1; j < sorted.length && sorted[j]![1].top <= bi.bottom; j++) {
      const [mj, bj] = sorted[j]!;
      if (overlaps(bi, bj)) {
        bad.add(mi);
        bad.add(mj);
      }
    }
    if (arrays.some((a) => overlaps(a, bi))) bad.add(mi);
  }
  const out = new Map<string, string>();
  for (const [master, b] of groups) {
    if (!bad.has(master)) out.set(master, `${columnLetter(b.left)}${String(b.top)}:${columnLetter(b.right)}${String(b.bottom)}`);
  }
  return out;
}

/** Converts a cell value for the writer; a `{ text, hyperlink }` value gives its text (the link is {@link linkIn}). */
export function valueIn(cell: Cell, sharedRanges?: ReadonlyMap<string, string>): CellInput {
  const v: CellValue = cell.value;
  if (v === null || v === undefined) return null;
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean' || v instanceof Date) return v;
  if (isRecord(v)) {
    if (cell.type === ValueType.Hyperlink) {
      // exceljs stores the link text in the shared string table (cell-xform.js L80–83; X4-10).
      const t = v['text'];
      if (isRecord(t) && Array.isArray(t['richText'])) return { richText: runsIn(t['richText'] as unknown[], cell.address) };
      return String(t);
    }
    // Shared and array formulas (D-E02; X4-29): a dependent names its master, a master gets the
    // range of its dependents, an array formula its `ref`.
    if (typeof v['sharedFormula'] === 'string') {
      const master = v['sharedFormula'].replace(/\$/g, '').toUpperCase();
      if (sharedRanges === undefined || sharedRanges.has(master)) return { sharedFormula: master, ...resultIn(v['result']) } as CellInput;
      // A group the core cannot write as shared (see sharedFormulaRanges): the formula of this
      // cell, translated from its master as exceljs's `cell.formula` does, or else its result.
      const text = cell.formula;
      if (text !== undefined) return { formula: text, ...resultIn(v['result']) } as CellInput;
      return resultValue(v['result']);
    }
    if (typeof v['formula'] === 'string') {
      const base = { formula: v['formula'], ...resultIn(v['result']) };
      if (v['shareType'] === 'array' && typeof v['ref'] === 'string') return { ...base, array: v['ref'] } as CellInput;
      const shared = sharedRanges?.get(cell.address);
      return (shared === undefined ? base : { ...base, shared }) as CellInput;
    }
    if (typeof v['error'] === 'string') return { error: v['error'] as '#N/A' };
    if (Array.isArray(v['richText'])) return { richText: runsIn(v['richText'] as unknown[], cell.address) };
  }
  // exceljs JSON values are written as their JSON text (lib/doc/cell.js L1014; X4-26).
  if (isJsonValue(v)) return JSON.stringify(v);
  return notSupported(`The value of cell ${cell.address}`);
}

/**
 * The link of a `{ text, hyperlink, tooltip }` value for `NewSheet.link`, or `null`. A target that
 * starts with `#` is a place in the workbook and is written as `location` without a relationship;
 * any other target is an external `url`. exceljs writes every target as an external relationship
 * (worksheet-xform.js L158–167); xlvane follows the specification instead (D-A08, X4-10).
 */
export function linkIn(cell: Cell): HyperlinkTarget | null {
  if (cell.type !== ValueType.Hyperlink) return null;
  const v = cell.value as Record<string, unknown>;
  const target = String(v['hyperlink']);
  const tooltip = v['tooltip'] ? { tooltip: String(v['tooltip']) } : {};
  return target.startsWith('#') ? { location: target.slice(1), ...tooltip } : { url: target, ...tooltip };
}

/**
 * The note of a cell for `NewSheet.note`, or `null`. exceljs writes every note with the author
 * "Author" (comments-xform.js L29–31; X4-14), so xlvane does too, and with the box settings of
 * `Note.DEFAULT_CONFIGS` merged with the note's `margins`, `protection` and `editAs`
 * (lib/doc/note.js L51–64; X4-31; D-E07).
 */
export function noteIn(cell: Cell): NoteInput | null {
  const note = cell.note;
  if (note === undefined) return null;
  if (typeof note === 'string') return { text: note, author: NOTE_AUTHOR, box: noteBoxIn(undefined) };
  const texts = note.texts ?? [];
  const box = noteBoxIn(note);
  return texts.length === 0 ? { text: '', author: NOTE_AUTHOR, box } : { runs: runsIn(texts, cell.address), author: NOTE_AUTHOR, box };
}

/** The author exceljs writes for every note. */
export const NOTE_AUTHOR = 'Author';

/** exceljs's default note box (lib/doc/note.js L51–64): automatic margins, inset 0.13/0.13/0.25/0.25 cm. */
const DEFAULT_INSET_CM = [0.13, 0.13, 0.25, 0.25] as const;
const EDIT_AS_IN = { absolute: 'absolute', oneCells: 'oneCell', twoCells: 'twoCell' } as const;
const EDIT_AS_OUT = { absolute: 'absolute', oneCell: 'oneCells', twoCell: 'twoCells' } as const;

/**
 * The box settings of an exceljs note. exceljs gives `inset` in centimetres and writes them ×10 as
 * millimetres (vml-textbox-xform.js L8–35); `editAs` 'absolute' writes MoveWithCells and
 * SizeWithCells, 'oneCells' only SizeWithCells, 'twoCells' neither (vml-client-data-xform.js,
 * vml-position-xform.js). Locked and LockText are locked unless given as "False".
 */
function noteBoxIn(note: Comment | undefined): NoteBoxInput {
  const inset = note?.margins?.inset ?? DEFAULT_INSET_CM;
  const mm = (i: number): number => Math.round(Number(inset[i] ?? DEFAULT_INSET_CM[i]) * 10 * 1e6) / 1e6;
  const unlocked = (v: unknown): boolean => v === 'False' || v === false || v === 'false';
  return {
    autoMargins: (note?.margins?.insetmode ?? 'auto') === 'auto',
    margins: { left: mm(0), top: mm(1), right: mm(2), bottom: mm(3) },
    locked: !unlocked(note?.protection?.locked),
    lockText: !unlocked(note?.protection?.lockText),
    anchor: EDIT_AS_IN[note?.editAs ?? 'absolute'] ?? 'absolute',
  };
}

/**
 * A note read by xlvane as exceljs gives it back: a string when it is one run without a font,
 * otherwise `{ texts }` with the box settings from the VML part (`margins` in centimetres,
 * `protection` as "True"/"False", `editAs`), as lib/doc/note.js `set model` and the VML xforms
 * read them (X4-14, X4-31). exceljs also adds the raw `x:Anchor` text as `anchor`; xlvane does not.
 */
export function noteOut(note: XlvNote): string | Comment {
  const runs = note.runs;
  if (!runs || runs.length === 0 || (runs.length === 1 && !runs[0]!.font)) return runs?.length === 1 ? runs[0]!.text : note.text;
  const out: Comment = { texts: richTextOut(runs).richText };
  const box = note.box;
  if (box !== undefined) Object.assign(out, noteBoxOut(box));
  return out;
}

function noteBoxOut(box: NoteBox): Pick<Comment, 'margins' | 'editAs' | 'protection'> {
  const cm = (v: number): number => Number((v / 10).toFixed(2));
  const flag = (v: boolean): 'True' | 'False' => (v ? 'True' : 'False');
  return {
    margins: { insetmode: box.autoMargins ? 'auto' : 'custom', inset: box.margins === null ? [...DEFAULT_INSET_CM] : [cm(box.margins.left), cm(box.margins.top), cm(box.margins.right), cm(box.margins.bottom)] },
    editAs: EDIT_AS_OUT[box.anchor],
    protection: { locked: flag(box.locked), lockText: flag(box.lockText) },
  };
}

/** Sheet layout from an exceljs view and `properties.tabColor`, applied to a new xlvane sheet. */
export interface SheetLayoutIn {
  readonly freeze: FrozenPanes | null;
  readonly view: SheetViewOptions;
  readonly tabColor: XlvColor | undefined;
}

/**
 * The first exceljs view and the tab colour (X4-11, X4-12). xlvane writes one view: frozen panes
 * (`xSplit`/`ySplit` in columns and rows) or a split window (`xSplit`/`ySplit` in 1/20 point, as
 * exceljs), plus `zoomScale`, `showGridLines`, `showRowColHeaders`, `rightToLeft` and `style`.
 * `topLeftCell`, `activeCell` and `activePane` are not written; xlvane writes the default
 * top-left cell and active pane.
 */
export function layoutIn(ws: Worksheet): SheetLayoutIn {
  const v = ws.views[0];
  const view: { -readonly [K in keyof SheetViewOptions]: SheetViewOptions[K] } = {};
  let freeze: FrozenPanes | null = null;
  if (v) {
    if (typeof v.zoomScale === 'number' && v.zoomScale !== 100) view.zoom = v.zoomScale;
    if (v.showGridLines === false) view.showGridLines = false;
    if (v['showRowColHeaders'] === false) view.showRowColHeaders = false;
    if (v['rightToLeft'] === true) view.rightToLeft = true;
    const style = v['style'];
    if (style === 'pageBreakPreview' || style === 'pageLayout') view.layout = style;
    // sheet-view-xform.js L26–99 (X4-11): further attributes written when they differ from the defaults.
    if (typeof v['zoomScaleNormal'] === 'number' && v['zoomScaleNormal'] !== 100 && v['zoomScaleNormal'] !== 0) view.zoomScaleNormal = v['zoomScaleNormal'];
    if (v['showRuler'] === false) view.showRuler = false;
    if (typeof v['activeCell'] === 'string' && v['activeCell'] !== '') view.activeCell = v['activeCell'];
    if (v.state !== 'frozen' && v.state !== 'split' && typeof v['topLeftCell'] === 'string' && v['topLeftCell'] !== '') view.topLeftCell = v['topLeftCell'];
    if (v.state === 'frozen') freeze = { rows: v.ySplit ?? 0, cols: v.xSplit ?? 0 };
    else if (v.state === 'split' && ((v.xSplit ?? 0) > 0 || (v.ySplit ?? 0) > 0)) view.split = { x: (v.xSplit ?? 0) / 20, y: (v.ySplit ?? 0) / 20 };
  }
  let tab = ws.properties['tabColor'] as Partial<Color> | undefined;
  // The exceljs README sets a red tab with the 7-digit argb "FFC0000", which exceljs writes as
  // given (sheet-properties-xform.js L27, color-xform.js L19–20) although an ARGB value has 8 hex digits (ST_UnsignedIntHex).
  // xlvane completes such a value with a final 0 ("FFC00000"), keeping the alpha FF (D-R01).
  if (tab && typeof tab.argb === 'string' && /^[0-9A-Fa-f]{7}$/.test(tab.argb)) tab = { ...tab, argb: `${tab.argb}0` };
  return { freeze, view, tabColor: colorIn(tab, `the tab colour of worksheet "${ws.name}"`) };
}

/** Workbook-level data written with the sheets: document properties and defined names. */
export interface WorkbookMeta {
  readonly properties?: WorkbookProperties;
  readonly definedNames?: readonly { name: string; ranges: string[] }[];
  /** The workbook's images, by exceljs image id (`workbook.addImage`). */
  readonly media?: readonly ImageData[];
  /** `false` writes no cell or column styles, as exceljs `xlsx.write(…, { useStyles: false })` (dates keep format 14). */
  readonly useStyles?: boolean;
  /** `workbook.views` (workbook windows). */
  readonly views?: readonly Record<string, unknown>[];
  /** `workbook.calcProperties`. */
  readonly calcProperties?: Record<string, unknown>;
}

/**
 * exceljs workbook views as core windows: exceljs writes `x`, `y`, `width`, `height` with the
 * defaults 0, 0, 12000, 24000 (workbook-view-xform.js L4–16; X4-28).
 */
export function bookViewsIn(views: readonly Record<string, unknown>[]): WorkbookViewOptions[] {
  return views.map((v) => {
    const n = (key: string, dflt?: number): number | undefined => (typeof v[key] === 'number' && Number.isFinite(v[key]) ? Math.trunc(v[key]) : dflt);
    const out: Record<string, unknown> = { x: n('x', 0) || 0, y: n('y', 0) || 0, width: n('width', 12000) || 12000, height: n('height', 24000) || 24000 };
    const firstSheet = n('firstSheet');
    if (firstSheet !== undefined) out['firstSheet'] = firstSheet;
    const activeTab = n('activeTab');
    if (activeTab !== undefined) out['activeTab'] = activeTab;
    if (v['visibility'] === 'hidden' || v['visibility'] === 'veryHidden') out['visibility'] = v['visibility'];
    return out as WorkbookViewOptions;
  });
}

/** exceljs `worksheet.properties` as core sheet properties (sheet-format-properties-xform.js; X4-28). */
export function sheetPropertiesIn(props: Record<string, unknown>): SheetPropertiesOptions {
  const out: Record<string, unknown> = {};
  if (typeof props['defaultRowHeight'] === 'number' && props['defaultRowHeight'] !== 15) out['defaultRowHeight'] = props['defaultRowHeight'];
  for (const key of ['outlineLevelRow', 'outlineLevelCol'] as const) if (typeof props[key] === 'number' && props[key] > 0) out[key] = props[key];
  const outline = props['outlineProperties'];
  if (outline && typeof outline === 'object') {
    const o = outline as Record<string, unknown>;
    if (o['summaryBelow'] === false) out['summaryBelow'] = false;
    if (o['summaryRight'] === false) out['summaryRight'] = false;
  }
  return out as SheetPropertiesOptions;
}

/**
 * Runs `run`; an XlvError without a sheet in its location gets the worksheet (and the cell), so
 * an error found while writing says where the value came from (D-R01).
 */
function at<T>(sheet: string, cell: string | undefined, run: () => T): T {
  try {
    return run();
  } catch (error) {
    if (!(error instanceof XlvError) || error.location.sheet !== undefined) throw error;
    const where = cell === undefined ? `Worksheet "${sheet}"` : `Worksheet "${sheet}", cell ${cell}`;
    throw new XlvError(error.code, `${where}: ${error.message}`, {
      location: { ...error.location, sheet, ...(cell === undefined ? {} : { cell }) },
      hint: error.hint,
      details: error.details,
      cause: error,
    });
  }
}

/** Builds an xlvane workbook from the exceljs model (D-503). */
export function toNewWorkbook(sheets: readonly Worksheet[], date1904: boolean, meta: WorkbookMeta = {}): NewWorkbook {
  const book = createWorkbook({ ...(date1904 ? { dateSystem: '1904' as const } : {}), ...(meta.properties ? { properties: meta.properties } : {}) });
  if (sheets.length === 0) {
    // exceljs writes a workbook without sheets; xlvane requires one (D-505).
    notSupported('Writing a workbook without worksheets');
  }
  const coreImages = new Map<number, number>();
  const coreBytes: Uint8Array[] = [];
  const imageId = (id: number): number => {
    let core = coreImages.get(id);
    if (core === undefined) {
      const image = meta.media?.[id];
      if (!image) {
        throw new XlvError('XLV_E_INVALID_OPTION', `Image ${String(id)} was not added to this workbook.`, { hint: 'Use the id that workbook.addImage() returned.', details: { option: 'imageId', value: id } });
      }
      // Identical images under different ids become one core image, so a drawing has one
      // relationship per image file (exceljs writes one per use, X4-22).
      const bytes = imageBytes(image);
      const same = coreBytes.findIndex((b) => b.length === bytes.length && b.every((x, i) => x === bytes[i]));
      if (same === -1) {
        core = book.addImage(bytes);
        coreBytes[core] = bytes;
      } else core = same;
      coreImages.set(id, core);
    }
    return core;
  };
  if (meta.views && meta.views.length > 0) book.views(bookViewsIn(meta.views));
  // exceljs writes calcPr always, with calcId 171027 (workbook-calc-properties-xform.js); xlvane
  // writes it only for fullCalcOnLoad, without an application's calcId (X4-28).
  if (meta.calcProperties?.['fullCalcOnLoad']) book.calculation({ fullCalcOnLoad: true });
  for (const ws of sheets) at(ws.name, undefined, () => {
    const sheet = book.addSheet(ws.name, { state: ws.state });
    const useStyles = meta.useStyles !== false;
    const sharedRanges = sharedFormulaRanges(ws);
    const props = sheetPropertiesIn(ws.properties);
    if (Object.keys(props).length > 0) sheet.properties(props);
    ws.columns.forEach((column) => {
      const options = columnIn(column, useStyles);
      if (options) sheet.column(column.number, options);
    });
    const styleCache = new Map<Style, CellStyle>();
    const styleOf = (s: Style): CellStyle => {
      let c = styleCache.get(s);
      if (!c) {
        c = styleIn(s);
        styleCache.set(s, c);
      }
      return c;
    };
    ws._rows.forEach((row) => {
      if (!row) return;
      if (row.height !== undefined || row.hidden || row.outlineLevel) {
        sheet.row(row.number, {
          ...(row.height === undefined ? {} : { height: row.height }),
          ...(row.hidden ? { hidden: true } : {}),
          // row-xform.js L49–54: outlineLevel and the computed collapsed flag (X4-28).
          ...(row.outlineLevel ? { outlineLevel: row.outlineLevel } : {}),
          ...(row.collapsed ? { collapsed: true } : {}),
        });
      }
      row._cells.forEach((cell) => {
        if (cell) at(ws.name, cell.address, () => writeCell(cell));
      });
    });
    function writeCell(cell: Cell): void {
      const note = noteIn(cell);
      if (note) sheet.note(cell.address, note);
      const style = useStyles ? cellStyleIn(cell, styleOf(cell.style)) : isDateCell(cell) ? { numFmt: DATE_NUMFMT } : {};
      const value = cell.type === ValueType.Merge ? null : valueIn(cell, sharedRanges);
      if (value === null && Object.keys(style).length === 0) return;
      sheet.set(cell.address, value, style);
      const link = cell.type === ValueType.Merge ? null : linkIn(cell);
      if (link) sheet.link(cell.address, link);
    }
    for (const m of ws._merges.values()) sheet.merge(`${columnLetter(m.left)}${String(m.top)}:${columnLetter(m.right)}${String(m.bottom)}`);
    const layout = layoutIn(ws);
    if (layout.freeze) sheet.freeze(layout.freeze);
    if (Object.keys(layout.view).length > 0) sheet.view(layout.view);
    if (layout.tabColor !== undefined) sheet.tabColor(layout.tabColor);
    const filter = autoFilterRange(ws.autoFilter);
    if (filter) sheet.autoFilter(filter);
    const rules = sheetRulesIn(ws);
    for (const v of rules.dataValidations) sheet.validate(v.sqref, v.rule);
    for (const cf of rules.conditionalFormats) sheet.conditionalFormat(cf.sqref, cf.rules);
    if (rules.protect) sheet.protect(rules.protect);
    sheet.pageSetup(rules.pageSetup);
    if (rules.headerFooter) sheet.headerFooter(rules.headerFooter);
    if (rules.pageBreaks.length > 0) sheet.pageBreaks({ rows: rules.pageBreaks });
    if (rules.printArea) sheet.printArea(rules.printArea);
    if (rules.printTitles) sheet.printTitles(rules.printTitles);
    // Tables (D-D04): the cells are written above; the core adds the header names, the totals
    // formulas and the table part. exceljs writes the ref computed at the last validate/commit.
    for (const table of ws.getTables()) sheet.addTable(tableInput(table));
    // Pictures and the background: each exceljs image becomes one core image, stored once (D-D03).
    for (const m of ws._media) {
      const id = imageId(m.imageId);
      if (m.type === 'background') sheet.background(id);
      else if (m.range) {
        const { anchor, hyperlink } = imageAnchorIn(ws, m.range);
        sheet.addImage(id, anchor, hyperlink === undefined ? {} : { hyperlink });
      }
    }
  });
  // exceljs writes each name's ranges joined by commas, without a sheet scope (defined-name-xform.js L5).
  for (const d of meta.definedNames ?? []) book.defineName(d.name, d.ranges.join(','));
  return book;
}

/**
 * The validations, conditional formats, protection and printing settings of a worksheet in the
 * form of the core API (D-C06). exceljs always writes page margins and page setup
 * (worksheet-xform.js L337–339), so `pageSetup` is always given.
 */
export function sheetRulesIn(ws: Worksheet): {
  dataValidations: { sqref: string; rule: DataValidationRule }[];
  conditionalFormats: { sqref: string; rules: ConditionalFormatRuleInput[] }[];
  protect: SheetProtectionOptions | null;
  pageSetup: PageSetupOptions;
  headerFooter: XlvHeaderFooter | null;
  pageBreaks: number[];
  printArea: string[] | null;
  printTitles: PrintTitles | null;
} {
  const page = pageSetupIn(ws.pageSetup);
  return {
    dataValidations: dataValidationsIn(ws.dataValidations),
    conditionalFormats: conditionalFormatsIn(ws.conditionalFormattings),
    protect: ws.sheetProtection ? protectionIn(ws.sheetProtection) : null,
    pageSetup: page.options,
    headerFooter: headerFooterIn(ws.headerFooter),
    pageBreaks: rowBreaksIn(ws.rowBreaks),
    printArea: page.printArea,
    printTitles: page.printTitles,
  };
}

function refOf(x: string | { row: number; column: number }): string {
  return typeof x === 'string' ? x : `${columnLetter(x.column)}${String(x.row)}`;
}

/** The autofilter range in A1 form (exceljs accepts a string or { from, to }). */
export function autoFilterRange(f: Worksheet['autoFilter']): string | null {
  if (!f) return null;
  if (typeof f === 'string') return f;
  return `${refOf(f.from)}:${refOf(f.to)}`;
}

// ---------------------------------------------------------------------------------------------
// xlvane → exceljs

/** Converts a colour read by xlvane to the exceljs format. */
export function colorOut(c: XlvColor | undefined): Partial<Color> | undefined {
  if (c === undefined) return undefined;
  if (typeof c === 'string') return { argb: c };
  if ('theme' in c) return c.tint === undefined ? { theme: c.theme } : { theme: c.theme, tint: c.tint };
  if ('indexed' in c) return { indexed: c.indexed };
  return undefined;
}

function fontOut(f: FontStyle): Partial<Font> {
  const font: Partial<Font> = {};
  if (f.name !== undefined) font.name = f.name;
  if (f.size !== undefined) font.size = f.size;
  if (f.bold) font.bold = true;
  if (f.italic) font.italic = true;
  if (f.strike) font.strike = true;
  if (f.underline) font.underline = f.underline === true ? 'single' : f.underline;
  const color = colorOut(f.color);
  if (color) font.color = color;
  if (f.vertAlign && f.vertAlign !== 'baseline') font.vertAlign = f.vertAlign;
  if (f.family !== undefined) font.family = f.family;
  if (f.scheme !== undefined) font.scheme = f.scheme;
  if (f.charset !== undefined) font.charset = f.charset;
  if (f.outline) font.outline = true;
  for (const key of ['shadow', 'condense', 'extend'] as const) if (f[key]) (font as Record<string, unknown>)[key] = true;
  return font;
}

/**
 * An exceljs view from the view read by xlvane, shaped as exceljs reads one (sheet-view-xform.js
 * L145–192; X4-11): a missing `zoomScaleNormal` reads as 100; a frozen view gets the default
 * `topLeftCell`; a split view gets `xSplit`/`ySplit` in 1/20 point and the active pane xlvane
 * writes; a normal view has no `topLeftCell`, as in exceljs.
 */
export function viewOut(view: SheetView, freeze: FrozenPanes | null): WorksheetView {
  const common: WorksheetView = {
    workbookViewId: 0,
    rightToLeft: view.rightToLeft,
    showRuler: view.showRuler,
    showRowColHeaders: view.showRowColHeaders,
    showGridLines: view.showGridLines,
    zoomScale: view.zoom,
    // exceljs reads a missing zoomScaleNormal as 100.
    zoomScaleNormal: view.zoomScaleNormal || 100,
  };
  const style = view.layout === 'normal' ? {} : { style: view.layout };
  if (freeze) {
    const topLeftCell = `${columnLetter(freeze.cols + 1)}${String(freeze.rows + 1)}`;
    // xlvane writes the top-left cell as the active cell of frozen panes when none is given.
    const active = view.activeCell !== null && view.activeCell !== topLeftCell ? { activeCell: view.activeCell } : {};
    return { ...common, state: 'frozen', xSplit: freeze.cols, ySplit: freeze.rows, topLeftCell, ...active, ...style };
  }
  const active = view.activeCell !== null ? { activeCell: view.activeCell } : {};
  if (view.split) {
    const { x, y } = view.split;
    const activePane = x > 0 && y > 0 ? 'bottomRight' : y > 0 ? 'bottomLeft' : 'topRight';
    return { ...common, state: 'split', xSplit: Math.round(x * 20), ySplit: Math.round(y * 20), topLeftCell: undefined, activePane, ...active, ...style };
  }
  return { ...common, state: 'normal', ...active, ...style };
}

/**
 * A formula read by xlvane in exceljs's shape (cell-xform.js L460–470; X4-29): a shared-formula
 * master `{ formula, result, ref, shareType: 'shared' }`, a dependent `{ sharedFormula: master,
 * result }` (its `cell.formula` is translated), an array formula `{ formula, result, ref,
 * shareType: 'array' }`, otherwise `{ formula, result }`.
 */
export function formulaOut(f: XlvFormula, plain: unknown, ref: string, masters: ReadonlyMap<number, string>): CellValue {
  const result = plain === null || plain === undefined ? {} : { result: plain };
  if (f.kind === 'shared' && f.sharedIndex !== null) {
    const master = masters.get(f.sharedIndex);
    if (f.ref !== null) return { formula: f.text ?? '', ...result, ref: f.ref, shareType: 'shared' } as CellValue;
    if (master !== undefined && master !== ref) return { sharedFormula: master, ...result } as CellValue;
  }
  if (f.kind === 'array' && f.ref !== null) return { formula: f.text ?? '', ...result, ref: f.ref, shareType: 'array' } as CellValue;
  return { formula: f.text ?? '', ...result } as CellValue;
}

/** Converts rich text runs read by xlvane to the exceljs `{ richText }` value. */
export function richTextOut(runs: readonly RichTextRun[]): { richText: { text: string; font?: Partial<Font> }[] } {
  return { richText: runs.map((r) => (r.font ? { font: fontOut(r.font), text: r.text } : { text: r.text })) };
}

/** Converts a style read by xlvane to the exceljs format. */
export function styleOut(s: CellStyle, numFmt: string | null): Style {
  const out: Style = {};
  if (numFmt) out.numFmt = numFmt;
  if (s.font) out.font = fontOut(s.font);
  if (s.fill && 'pattern' in s.fill) {
    const fill: Fill = { type: 'pattern', pattern: s.fill.pattern };
    const fg = colorOut(s.fill.fgColor);
    if (fg) fill.fgColor = fg;
    const bg = colorOut(s.fill.bgColor);
    if (bg) fill.bgColor = bg;
    out.fill = fill;
  } else if (s.fill && 'type' in s.fill && s.fill.type === 'gradient') {
    // As exceljs reads a gradient (fill-xform.js L201–229): `angle` with `degree`, or `path` with
    // `center` (right/bottom only when they differ from left/top).
    const g = s.fill;
    const fill: Record<string, unknown> = { type: 'gradient' };
    if (g.gradient === 'path') {
      const center: Record<string, number> = { left: g.left ?? 0, top: g.top ?? 0 };
      if ((g.right ?? 0) !== (g.left ?? 0)) center['right'] = g.right ?? 0;
      if ((g.bottom ?? 0) !== (g.top ?? 0)) center['bottom'] = g.bottom ?? 0;
      fill['gradient'] = 'path';
      fill['center'] = center;
    } else {
      fill['gradient'] = 'angle';
      fill['degree'] = g.degree ?? 0;
    }
    fill['stops'] = g.stops.map((st) => ({ position: st.position, color: colorOut(st.color) ?? {} }));
    out.fill = fill as unknown as Fill;
  }
  if (s.border) {
    const border: Partial<Borders> = {};
    for (const edge of ['top', 'right', 'bottom', 'left'] as const) {
      const e = s.border[edge];
      if (!e) continue;
      const color = colorOut(e.color);
      border[edge] = color ? { style: e.style, color } : { style: e.style };
    }
    const d = s.border.diagonal;
    if (d) {
      const color = colorOut(d.color);
      border.diagonal = { style: d.style, ...(color ? { color } : {}), up: s.border.diagonalUp === true, down: s.border.diagonalDown === true };
    }
    out.border = border;
  }
  if (s.alignment) {
    const a = s.alignment;
    const alignment: Partial<Alignment> = {};
    if (a.horizontal && a.horizontal !== 'general') alignment.horizontal = a.horizontal;
    if (a.vertical) alignment.vertical = a.vertical === 'center' ? 'middle' : a.vertical;
    if (a.wrapText) alignment.wrapText = true;
    // Office accepts indents up to 255 ([MS-OI29500] §2.1.693 b), the limit the writer enforces, so
    // a larger indent from a file is left out and the loaded workbook can be written again (D-E06).
    if (a.indent && a.indent <= 255) alignment.indent = a.indent;
    if (a.textRotation !== undefined) {
      const tr = a.textRotation;
      alignment.textRotation = tr === 255 ? 'vertical' : tr > 90 && tr <= 180 ? 90 - tr : tr;
    }
    if (a.shrinkToFit) alignment.shrinkToFit = true;
    // exceljs reads every readingOrder other than 2 as "ltr" (alignment-xform.js parseOpen).
    if (a.readingOrder !== undefined) alignment.readingOrder = a.readingOrder === 'rightToLeft' ? 'rtl' : 'ltr';
    out.alignment = alignment;
  }
  if (s.protection && (s.protection.locked === false || s.protection.hidden === true)) {
    out.protection = { locked: s.protection.locked !== false, hidden: s.protection.hidden === true };
  }
  return out;
}

/** Reads a file with xlvane; returns the read model for `fillFromRead`. */
export function readBytes(bytes: Uint8Array): XlvWorkbook {
  return readWorkbook(bytes);
}
