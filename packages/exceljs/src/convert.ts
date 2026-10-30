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
  type SheetProtectionFlags,
  type Color as XlvColor,
  type ColumnOptions,
  type FontStyle,
  type FrozenPanes,
  type HyperlinkTarget,
  type NewWorkbook,
  type Note as XlvNote,
  type NoteInput,
  type RichTextRun,
  type SheetView,
  type SheetViewOptions,
  type Workbook as XlvWorkbook,
  type WorkbookProperties,
} from 'xlvane';
import { columnLetter, notSupported, ValueType, type Cell, type Column, type Worksheet } from './model';
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
  if (f.vertAlign) notSupported('Superscript and subscript fonts (font.vertAlign)');
  if (f.outline) notSupported('Outline fonts (font.outline)');
  const out: Record<string, unknown> = {};
  if (f.name !== undefined) out['name'] = f.name;
  if (f.size !== undefined) out['size'] = f.size;
  if (f.bold) out['bold'] = true;
  if (f.italic) out['italic'] = true;
  if (f.strike) out['strike'] = true;
  if (f.underline && f.underline !== 'none') out['underline'] = f.underline === true ? 'single' : f.underline;
  const color = colorIn(f.color, 'font.color');
  if (color !== undefined) out['color'] = color;
  // family, scheme and charset describe the font for substitution; they are not written (D-504).
  return out as CellStyle['font'];
}

function fillIn(f: Fill | undefined): CellStyle['fill'] {
  if (!f) return undefined;
  if (f.type === 'gradient') return notSupported('Gradient fills');
  if (f.type !== 'pattern') return notSupported(`A fill of type "${String(f.type)}"`);
  if (!f.pattern || f.pattern === 'none') return undefined;
  const out: Record<string, unknown> = { pattern: f.pattern };
  const fg = colorIn(f.fgColor, 'fill.fgColor');
  if (fg !== undefined) out['fgColor'] = fg;
  const bg = colorIn(f.bgColor, 'fill.bgColor');
  if (bg !== undefined) out['bgColor'] = bg;
  return out as CellStyle['fill'];
}

function borderIn(b: Partial<Borders> | undefined): CellStyle['border'] {
  if (!b) return undefined;
  if (b.diagonal && (b.diagonal.style || b.diagonal.up || b.diagonal.down)) notSupported('Diagonal borders');
  const out: Record<string, unknown> = {};
  for (const edge of ['top', 'right', 'bottom', 'left'] as const) {
    const e = b[edge];
    if (!e || !e.style) continue;
    const color = colorIn(e.color, `border.${edge}.color`);
    out[edge] = color === undefined ? { style: e.style } : { style: e.style, color };
  }
  return Object.keys(out).length > 0 ? (out as CellStyle['border']) : undefined;
}

function alignmentIn(a: Partial<Alignment> | undefined): CellStyle['alignment'] {
  if (!a) return undefined;
  if (a.shrinkToFit) notSupported('alignment.shrinkToFit');
  if (a.readingOrder) notSupported('alignment.readingOrder');
  const out: Record<string, unknown> = {};
  if (a.horizontal) out['horizontal'] = a.horizontal;
  // exceljs calls the centre "middle" (index.d.ts Alignment L282); SpreadsheetML calls it "center".
  if (a.vertical) out['vertical'] = a.vertical === 'middle' ? 'center' : a.vertical;
  if (a.wrapText) out['wrapText'] = true;
  if (a.indent) out['indent'] = a.indent;
  if (a.textRotation !== undefined) out['textRotation'] = a.textRotation === 'vertical' ? 255 : a.textRotation;
  return Object.keys(out).length > 0 ? (out as CellStyle['alignment']) : undefined;
}

/** Converts an exceljs style to an xlvane CellStyle (unsupported parts throw XLV_E_NOT_SUPPORTED). */
export function styleIn(s: Style): CellStyle {
  if (s.protection !== undefined && s.protection !== null) notSupported('Cell protection');
  const out: Record<string, unknown> = {};
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
  if (!customWidth && !column.hidden && !styled) return undefined;
  const style = withStyle && styled ? styleIn(s) : undefined;
  return {
    width: column.width ?? DEFAULT_COLUMN_WIDTH,
    ...(column.hidden ? { hidden: true } : {}),
    ...(style && Object.keys(style).length > 0 ? { style } : {}),
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

/** Converts a cell value for the writer; a `{ text, hyperlink }` value gives its text (the link is {@link linkIn}). */
export function valueIn(cell: Cell): CellInput {
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
    if ('sharedFormula' in v) return notSupported('Writing { sharedFormula } values (write the formula text with { formula })');
    if (typeof v['formula'] === 'string') {
      const r = v['result'];
      if (r === undefined || r === null) return { formula: v['formula'] };
      if (isRecord(r) && typeof r['error'] === 'string') return { formula: v['formula'], result: { error: r['error'] as '#N/A' } };
      return { formula: v['formula'], result: r as number };
    }
    if (typeof v['error'] === 'string') return { error: v['error'] as '#N/A' };
    if (Array.isArray(v['richText'])) return { richText: runsIn(v['richText'] as unknown[], cell.address) };
  }
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
 * "Author" (comments-xform.js L29–31; X4-14), so xlvane does too. `margins`, `protection` and
 * `editAs` are not written (D-A14).
 */
export function noteIn(cell: Cell): NoteInput | null {
  const note = cell.note;
  if (note === undefined) return null;
  if (typeof note === 'string') return { text: note, author: NOTE_AUTHOR };
  const texts = note.texts ?? [];
  return texts.length === 0 ? { text: '', author: NOTE_AUTHOR } : { runs: runsIn(texts, cell.address), author: NOTE_AUTHOR };
}

/** The author exceljs writes for every note. */
export const NOTE_AUTHOR = 'Author';

/**
 * A note read by xlvane as exceljs gives it back: a string when it is one run without a font,
 * otherwise `{ texts }` (lib/doc/note.js `set model`; X4-14). exceljs also merges the VML box
 * settings into the object; xlvane does not read the VML (D-A12), so they are absent.
 */
export function noteOut(note: XlvNote): string | Comment {
  const runs = note.runs;
  if (!runs || runs.length === 0) return note.text;
  if (runs.length === 1 && !runs[0]!.font) return runs[0]!.text;
  return { texts: richTextOut(runs).richText };
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
    if (v.state === 'frozen') freeze = { rows: v.ySplit ?? 0, cols: v.xSplit ?? 0 };
    else if (v.state === 'split' && ((v.xSplit ?? 0) > 0 || (v.ySplit ?? 0) > 0)) view.split = { x: (v.xSplit ?? 0) / 20, y: (v.ySplit ?? 0) / 20 };
  }
  const tab = ws.properties['tabColor'] as Partial<Color> | undefined;
  return { freeze, view, tabColor: colorIn(tab, `the tab colour of worksheet "${ws.name}"`) };
}

/** Workbook-level data written with the sheets: document properties and defined names. */
export interface WorkbookMeta {
  readonly properties?: WorkbookProperties;
  readonly definedNames?: readonly { name: string; ranges: string[] }[];
}

/** Builds an xlvane workbook from the exceljs model (D-503). */
export function toNewWorkbook(sheets: readonly Worksheet[], date1904: boolean, meta: WorkbookMeta = {}): NewWorkbook {
  const book = createWorkbook({ ...(date1904 ? { dateSystem: '1904' as const } : {}), ...(meta.properties ? { properties: meta.properties } : {}) });
  if (sheets.length === 0) {
    // exceljs writes a workbook without sheets; xlvane requires one (D-505).
    notSupported('Writing a workbook without worksheets');
  }
  for (const ws of sheets) {
    const sheet = book.addSheet(ws.name, { state: ws.state });
    ws.columns.forEach((column) => {
      const options = columnIn(column, true);
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
      if (row.height !== undefined || row.hidden) {
        sheet.row(row.number, { ...(row.height === undefined ? {} : { height: row.height }), ...(row.hidden ? { hidden: true } : {}) });
      }
      row._cells.forEach((cell) => {
        if (!cell) return;
        const note = noteIn(cell);
        if (note) sheet.note(cell.address, note);
        const style = cellStyleIn(cell, styleOf(cell.style));
        const value = cell.type === ValueType.Merge ? null : valueIn(cell);
        if (value === null && Object.keys(style).length === 0) return;
        sheet.set(cell.address, value, style);
        const link = cell.type === ValueType.Merge ? null : linkIn(cell);
        if (link) sheet.link(cell.address, link);
      });
    });
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
  }
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
  protect: SheetProtectionFlags | null;
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
  return font;
}

/**
 * An exceljs view from the view read by xlvane, shaped as exceljs reads one (sheet-view-xform.js
 * L145–192; X4-11): `showRuler` and `zoomScaleNormal` are not read by xlvane and get their
 * defaults; a frozen view gets the default `topLeftCell`; a split view gets `xSplit`/`ySplit` in
 * 1/20 point and the active pane xlvane writes.
 */
export function viewOut(view: SheetView, freeze: FrozenPanes | null): WorksheetView {
  const common: WorksheetView = {
    workbookViewId: 0,
    rightToLeft: view.rightToLeft,
    showRuler: true,
    showRowColHeaders: view.showRowColHeaders,
    showGridLines: view.showGridLines,
    zoomScale: view.zoom,
    zoomScaleNormal: 100,
  };
  const style = view.layout === 'normal' ? {} : { style: view.layout };
  if (freeze) {
    return { ...common, state: 'frozen', xSplit: freeze.cols, ySplit: freeze.rows, topLeftCell: `${columnLetter(freeze.cols + 1)}${String(freeze.rows + 1)}`, ...style };
  }
  if (view.split) {
    const { x, y } = view.split;
    const activePane = x > 0 && y > 0 ? 'bottomRight' : y > 0 ? 'bottomLeft' : 'topRight';
    return { ...common, state: 'split', xSplit: Math.round(x * 20), ySplit: Math.round(y * 20), topLeftCell: undefined, activePane, ...style };
  }
  return { ...common, state: 'normal', ...style };
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
  }
  if (s.border) {
    const border: Partial<Borders> = {};
    for (const edge of ['top', 'right', 'bottom', 'left'] as const) {
      const e = s.border[edge];
      if (!e) continue;
      const color = colorOut(e.color);
      border[edge] = color ? { style: e.style, color } : { style: e.style };
    }
    out.border = border;
  }
  if (s.alignment) {
    const a = s.alignment;
    const alignment: Partial<Alignment> = {};
    if (a.horizontal && a.horizontal !== 'general') alignment.horizontal = a.horizontal;
    if (a.vertical) alignment.vertical = a.vertical === 'center' ? 'middle' : a.vertical;
    if (a.wrapText) alignment.wrapText = true;
    if (a.indent) alignment.indent = a.indent;
    if (a.textRotation !== undefined) alignment.textRotation = a.textRotation === 255 ? 'vertical' : a.textRotation;
    out.alignment = alignment;
  }
  return out;
}

/** Reads a file with xlvane; returns the read model for `fillFromRead`. */
export function readBytes(bytes: Uint8Array): XlvWorkbook {
  return readWorkbook(bytes);
}
