// Streaming writer: rows are pulled from (async) iterables while the output stream is read, so
// memory stays constant however many rows are written (docs/decisions.md D-403, D-404). Values
// and styles are validated with the rules of the batch writer (write.ts).
import { checkDefinedName, checkUniqueName, type DefinedNameOptions, type NormalizedDefinedName } from './defined-names';
import { APP_PROPERTIES_PART, appPropertiesXml, CORE_PROPERTIES_PART, corePropertiesXml, type WorkbookProperties } from './doc-properties';
import { XlvError } from './errors';
import type { HyperlinkTarget } from './hyperlinks';
import { printNames, printXml, sheetProtectionXml, type HeaderFooter, type PageBreaks, type PageSetupOptions, type PrintTitles, type SheetProtectionOptions } from './printing';
import type { DateSystem, SheetState } from './model';
import type { NoteInput } from './notes';
import type { ConditionalFormatRuleInput, DataValidationRule } from './sheet-rules';
import { formatCellRef, MAX_COLUMNS, MAX_ROWS, type CellRange } from './sml/cell-ref';
import {
  buildStyles,
  cellXml,
  contentTypesXml,
  nameTitles,
  createSharedStrings,
  rootRelsXml,
  rowStart,
  sheetHead,
  sheetTail,
  hyperlinksXml,
  sheetRulesXml,
  noteParts,
  type NoteParts,
  SheetRels,
  sheetRelsPartName,
  workbookRelsXml,
  workbookXml,
} from './sml/write-parts';
import { normalizeStyle, type CellStyle, type Color, type NormalizedStyle } from './style';
import type { TableColumnSpec } from './table';
import { readOptions } from './util/options';
import {
  convertCellInput,
  createWorkbook,
  DATE_FORMAT,
  DATE_TIME_FORMAT,
  sortedNotes,
  type CellInput,
  type ColumnOptions,
  type FreezeOptions,
  type SheetViewOptions,
  type NewSheet,
} from './write';
import { createZipStream, type ZipStreamEntry } from './zip/stream';

/**
 * A row for {@link writeWorkbookStream}: an array of values (column A first), or an object with
 * the values and optional formatting.
 *
 * @example
 * ```ts
 * import type { StreamRowInput } from 'xlvane';
 *
 * const plain: StreamRowInput = ['Alice', 42, new Date(Date.UTC(2024, 0, 31))];
 * const header: StreamRowInput = { values: ['Name', 'Total', 'Date'], style: { font: { bold: true } }, height: 20 };
 * const linked: StreamRowInput = { values: ['Docs', 7], hyperlinks: [{ url: 'https://example.com' }], notes: [undefined, 'Estimate'] };
 * ```
 */
export type StreamRowInput =
  | readonly CellInput[]
  | {
      readonly values: readonly CellInput[];
      /** Style for every cell of the row (used where `styles` has no entry). */
      readonly style?: CellStyle;
      /** Style per cell, by position. */
      readonly styles?: readonly (CellStyle | undefined)[];
      /** Row height in points. */
      readonly height?: number;
      readonly hidden?: boolean;
      /** Hyperlink per cell, by position, as for `NewSheet.link`. */
      readonly hyperlinks?: readonly (HyperlinkTarget | null | undefined)[];
      /**
       * Note per cell, by position, as for `NewSheet.note`. Notes are kept in memory until the
       * sheet ends, because they are written to their own parts after the rows.
       */
      readonly notes?: readonly (NoteInput | null | undefined)[];
    };

/**
 * A sheet of {@link writeWorkbookStream}. Layout is given up front; rows are pulled lazily.
 *
 * @example
 * ```ts
 * import type { StreamSheetSpec } from 'xlvane';
 *
 * const sheet: StreamSheetSpec = { name: 'Data', freeze: { rows: 1 }, autoFilter: true, rows: [['id'], [1], [2]] };
 * ```
 */
export interface StreamSheetSpec {
  readonly name: string;
  readonly state?: SheetState;
  /** Column options by position (column A first); `undefined` entries keep the default. */
  readonly columns?: readonly (ColumnOptions | undefined)[];
  readonly freeze?: FreezeOptions;
  /** Colour of the sheet tab, as for `NewSheet.tabColor`. */
  readonly tabColor?: Color;
  /** Display settings, as for `NewSheet.view`. */
  readonly view?: SheetViewOptions;
  /** `true`: a filter over all written rows and columns; or a range such as `"A1:D100"`. */
  readonly autoFilter?: boolean | string;
  /** Ranges to merge, for example `["A1:C1"]`. */
  readonly merges?: readonly string[];
  /**
   * Hyperlinks of the sheet, each a cell or range (`ref`) with a `HyperlinkTarget`; checked like
   * `NewSheet.link`. They are given up front because they are written after the rows.
   */
  readonly hyperlinks?: readonly (HyperlinkTarget & { readonly ref: string })[];
  /** Sheet protection, as for `NewSheet.protect`. */
  readonly protect?: SheetProtectionOptions;
  /** Page setup, as for `NewSheet.pageSetup`. */
  readonly pageSetup?: PageSetupOptions;
  /** Headers and footers, as for `NewSheet.headerFooter`. */
  readonly headerFooter?: HeaderFooter;
  /** Manual page breaks, as for `NewSheet.pageBreaks`. */
  readonly pageBreaks?: PageBreaks;
  /** Print area, as for `NewSheet.printArea`. */
  readonly printArea?: string | readonly string[];
  /** Print titles, as for `NewSheet.printTitles`. */
  readonly printTitles?: PrintTitles;
  /** Data validations, each a range (`sqref`) with a rule, as for `NewSheet.validate`. */
  readonly dataValidations?: readonly (DataValidationRule & { readonly sqref: string })[];
  /** Conditional formats, each a range with its rules, as for `NewSheet.conditionalFormat`. */
  readonly conditionalFormats?: readonly { readonly sqref: string; readonly rules: readonly ConditionalFormatRuleInput[] }[];
  /** The rows, from row 1. Any iterable or async iterable (arrays, generators, streams of rows). */
  readonly rows: Iterable<StreamRowInput> | AsyncIterable<StreamRowInput>;
}

/**
 * Options for {@link writeWorkbookStream}.
 *
 * @example
 * ```ts
 * import type { WriteWorkbookStreamOptions } from 'xlvane';
 *
 * const options: WriteWorkbookStreamOptions = { sheets: [{ name: 'Data', rows: [[1, 2, 3]] }] };
 * ```
 */
export interface WriteWorkbookStreamOptions {
  /**
   * The sheets. An array is checked before the stream starts; an (async) iterable is read lazily,
   * so sheets can be added while earlier ones are being written.
   */
  readonly sheets: readonly StreamSheetSpec[] | Iterable<StreamSheetSpec> | AsyncIterable<StreamSheetSpec>;
  /** Default `"1900"`. */
  readonly dateSystem?: DateSystem;
  /**
   * `false` (default): strings are stored in their cells (inline strings), so memory stays
   * constant. `true`: strings go to a shared string table, which is kept in memory until the end
   * (smaller files when values repeat).
   */
  readonly sharedStrings?: boolean;
  /** Document properties, written to `docProps/core.xml` (see `WorkbookProperties`). */
  readonly properties?: WorkbookProperties;
  /**
   * Defined names, as for `NewWorkbook.defineName`. With sheets given as an iterable, a name's
   * sheet is checked when all sheets are known, at the end of the stream.
   */
  readonly definedNames?: readonly StreamDefinedName[];
}

/**
 * A defined name for {@link WriteWorkbookStreamOptions.definedNames}.
 *
 * @example
 * ```ts
 * import type { StreamDefinedName } from 'xlvane';
 *
 * const name: StreamDefinedName = { name: 'Items', refersTo: 'Data!$A$1:$A$1000' };
 * ```
 */
export interface StreamDefinedName extends DefinedNameOptions {
  readonly name: string;
  /** Formula text with or without `=`. */
  readonly refersTo: string;
}

const CHUNK_CHARS = 64 * 1024;
const encoder = new TextEncoder();

interface PreparedSheet {
  readonly spec: StreamSheetSpec;
  readonly sheet: NewSheet;
  readonly autoFilter: boolean | CellRange;
  readonly rows: Iterable<StreamRowInput> | AsyncIterable<StreamRowInput>;
  /** Filled in while the sheet is written. */
  filterRange: CellRange | null;
  /** Relationships of the sheet part (hyperlinks, notes), written after the sheet. */
  readonly rels: SheetRels;
  /** The note parts, known when the sheet has been written. */
  notes: NoteParts | null;
}

function isIterable(value: unknown): value is Iterable<unknown> | AsyncIterable<unknown> {
  return (
    typeof value === 'object' &&
    value !== null &&
    (typeof (value as Iterable<unknown>)[Symbol.iterator] === 'function' || typeof (value as AsyncIterable<unknown>)[Symbol.asyncIterator] === 'function')
  );
}

async function* iterate<T>(source: Iterable<T> | AsyncIterable<T>): AsyncGenerator<T> {
  if (typeof (source as AsyncIterable<T>)[Symbol.asyncIterator] === 'function') yield* source as AsyncIterable<T>;
  else yield* source as Iterable<T>;
}

/**
 * Writes a workbook as a stream of XLSX bytes. Rows are pulled from each sheet's iterable only
 * as fast as the stream is read, so files with millions of rows need constant memory (with the
 * default inline strings). Values and styles follow the rules of `NewSheet.set`; an invalid
 * value errors the stream with an `XlvError` naming the sheet and cell. The layout (columns,
 * frozen panes, merges, filter) is given up front.
 *
 * @example
 * ```ts
 * import { writeWorkbookStream } from 'xlvane';
 *
 * function* rows() {
 *   yield { values: ['id', 'value'], style: { font: { bold: true } } };
 *   for (let i = 1; i <= 1_000_000; i++) yield [i, Math.sqrt(i)];
 * }
 * const stream = writeWorkbookStream({ sheets: [{ name: 'Data', freeze: { rows: 1 }, rows: rows() }] });
 * const bytes = new Uint8Array(await new Response(stream).arrayBuffer());
 * ```
 */
export function writeWorkbookStream(options: WriteWorkbookStreamOptions): ReadableStream<Uint8Array> {
  const o = readOptions(options, ['sheets', 'dateSystem', 'sharedStrings', 'properties', 'definedNames'], 'writeWorkbookStream options');
  const dateSystem = o.get('dateSystem');
  const book = createWorkbook({
    ...(dateSystem === undefined ? {} : { dateSystem: dateSystem as DateSystem }),
    ...(o.has('properties') ? { properties: o.get('properties') as WorkbookProperties } : {}),
  });
  const core = corePropertiesXml(book.properties);
  const rawNames = o.get('definedNames') ?? [];
  if (!Array.isArray(rawNames)) {
    throw new XlvError('XLV_E_INVALID_OPTION', 'The option "definedNames" must be an array.', { hint: 'For example definedNames: [{ name: "Items", refersTo: "Data!$A$1:$A$10" }].', details: { option: 'definedNames' } });
  }
  // Names are checked now; the sheet of a local name once the sheets are known.
  const names: NormalizedDefinedName[] = [];
  rawNames.forEach((raw: unknown, i) => {
    const r = readOptions(raw, ['name', 'refersTo', 'sheet', 'hidden', 'comment'], `definedNames[${String(i)}]`);
    const { name, refersTo, ...rest } = Object.fromEntries(r) as Record<string, unknown>;
    const n = checkDefinedName(name, refersTo, rest, null);
    checkUniqueName(n, names);
    names.push(n);
  });
  const addNames = (): void => {
    for (const n of names) book.defineName(n.name, n.refersTo, { ...(n.sheet === null ? {} : { sheet: n.sheet }), hidden: n.hidden, ...(n.comment === null ? {} : { comment: n.comment }) });
  };
  const sharedOption = o.get('sharedStrings') ?? false;
  if (typeof sharedOption !== 'boolean') {
    throw new XlvError('XLV_E_INVALID_OPTION', 'The option "sharedStrings" must be true or false.', { hint: 'Use sharedStrings: true or leave it out.', details: { option: 'sharedStrings' } });
  }
  const shared: boolean = sharedOption;
  const specs = o.get('sheets');
  const dynamic = !Array.isArray(specs);
  if (dynamic ? !isIterable(specs) : (specs as unknown[]).length === 0) {
    throw new XlvError('XLV_E_INVALID_INPUT', 'The workbook has no sheets.', { hint: 'Pass sheets: [{ name: "Sheet1", rows: [...] }].' });
  }
  const prepare = (raw: unknown, i: number): PreparedSheet => {
    const s = readOptions(raw, ['name', 'state', 'columns', 'freeze', 'tabColor', 'view', 'autoFilter', 'merges', 'hyperlinks', 'protect', 'pageSetup', 'headerFooter', 'pageBreaks', 'printArea', 'printTitles', 'dataValidations', 'conditionalFormats', 'rows'], `sheets[${String(i)}]`);
    const state = s.get('state');
    const sheet = book.addSheet(s.get('name') as string, state === undefined ? undefined : { state: state as SheetState });
    const columns = s.get('columns');
    if (columns !== undefined) {
      if (!Array.isArray(columns)) throw new XlvError('XLV_E_INVALID_OPTION', `sheets[${String(i)}].columns must be an array.`, { hint: 'Pass column options by position, e.g. [{ width: 20 }, undefined, { hidden: true }].', details: { option: 'columns' } });
      columns.forEach((c: unknown, k) => {
        if (c !== undefined && c !== null) sheet.column(k + 1, c as ColumnOptions);
      });
    }
    const freeze = s.get('freeze');
    if (freeze !== undefined) sheet.freeze(freeze as FreezeOptions);
    const tabColor = s.get('tabColor');
    if (tabColor !== undefined) sheet.tabColor(tabColor as Color);
    const view = s.get('view');
    if (view !== undefined) sheet.view(view as SheetViewOptions);
    if (s.has('protect')) sheet.protect(s.get('protect') as SheetProtectionOptions);
    if (s.has('pageSetup')) sheet.pageSetup(s.get('pageSetup') as PageSetupOptions);
    if (s.has('headerFooter')) sheet.headerFooter(s.get('headerFooter') as HeaderFooter);
    if (s.has('pageBreaks')) sheet.pageBreaks(s.get('pageBreaks') as PageBreaks);
    if (s.has('printArea')) sheet.printArea(s.get('printArea') as string | readonly string[]);
    if (s.has('printTitles')) sheet.printTitles(s.get('printTitles') as PrintTitles);
    const listOf = (key: string, example: string): unknown[] => {
      const v = s.get(key) ?? [];
      if (!Array.isArray(v)) throw new XlvError('XLV_E_INVALID_OPTION', `sheets[${String(i)}].${key} must be an array.`, { hint: `For example ${key}: [${example}].`, details: { option: key } });
      return v.map((item: unknown) => {
        if (typeof item !== 'object' || item === null) throw new XlvError('XLV_E_INVALID_OPTION', `sheets[${String(i)}].${key} must hold objects.`, { hint: `For example ${example}.`, details: { option: key } });
        return item;
      });
    };
    for (const v of listOf('dataValidations', '{ sqref: "A2:A9", type: "list", formula1: "\\"a,b\\"" }')) {
      const { sqref, ...rule } = v as { sqref?: unknown };
      sheet.validate(sqref as string, rule as DataValidationRule);
    }
    for (const v of listOf('conditionalFormats', '{ sqref: "A2:A9", rules: [{ type: "cellIs", operator: "greaterThan", formulas: [0], style: { font: { bold: true } } }] }')) {
      const c = readOptions(v, ['sqref', 'rules'], `sheets[${String(i)}].conditionalFormats`);
      sheet.conditionalFormat(c.get('sqref') as string, c.get('rules') as ConditionalFormatRuleInput[]);
    }
    const merges = s.get('merges');
    if (merges !== undefined) {
      if (!Array.isArray(merges)) throw new XlvError('XLV_E_INVALID_OPTION', `sheets[${String(i)}].merges must be an array of ranges.`, { hint: 'For example merges: ["A1:C1"].', details: { option: 'merges' } });
      for (const m of merges) sheet.merge(m as string);
    }
    const filter = s.get('autoFilter') ?? false;
    let autoFilter: boolean | CellRange = false;
    if (typeof filter === 'string') {
      sheet.autoFilter(filter);
      autoFilter = sheet.filter ?? false;
    } else if (typeof filter === 'boolean') autoFilter = filter;
    else throw new XlvError('XLV_E_INVALID_OPTION', `sheets[${String(i)}].autoFilter must be true, false or a range.`, { hint: 'Use autoFilter: true for a filter over all rows.', details: { option: 'autoFilter' } });
    const rows = s.get('rows');
    if (!isIterable(rows)) {
      throw new XlvError('XLV_E_TABLE_INVALID', `sheets[${String(i)}].rows is not an iterable.`, { hint: 'Pass an array, a generator or an async iterable of rows.', details: { row: 0 } });
    }
    const links = s.get('hyperlinks');
    if (links !== undefined) {
      if (!Array.isArray(links)) throw new XlvError('XLV_E_INVALID_OPTION', `sheets[${String(i)}].hyperlinks must be an array.`, { hint: 'For example hyperlinks: [{ ref: "A1", url: "https://example.com" }].', details: { option: 'hyperlinks' } });
      for (const l of links) {
        if (typeof l !== 'object' || l === null) throw new XlvError('XLV_E_INVALID_OPTION', `sheets[${String(i)}].hyperlinks must hold objects.`, { hint: 'For example { ref: "A1", url: "https://example.com" }.', details: { option: 'hyperlinks' } });
        const { ref, ...target } = l as { ref?: unknown };
        sheet.link(ref as string, target);
      }
    }
    return { spec: raw as StreamSheetSpec, sheet, autoFilter, rows: rows as Iterable<StreamRowInput>, filterRange: null, rels: new SheetRels(), notes: null };
  };
  const noVisibleSheet = (): XlvError =>
    new XlvError('XLV_E_INVALID_INPUT', 'All sheets of the workbook are hidden.', { hint: 'A workbook needs at least one visible sheet.' });
  const prepared: PreparedSheet[] = dynamic ? [] : (specs as unknown[]).map(prepare);
  if (!dynamic && !prepared.some((p) => p.sheet.state === 'visible')) throw noVisibleSheet();
  if (!dynamic) addNames();

  const strings = shared ? createSharedStrings() : null;
  const styleCache = new WeakMap<object, NormalizedStyle>();
  const dateStyle = new Map<string, number>();
  const normalized = (style: CellStyle): NormalizedStyle => {
    let n = styleCache.get(style);
    if (n === undefined) {
      n = normalizeStyle(style);
      styleCache.set(style, n);
    }
    return n;
  };

  /** First 1024-id block of VML shape ids for the next sheet with notes. */
  let shapeBlock = 1;

  async function* sheetData(p: PreparedSheet, selected: boolean, n: number): AsyncGenerator<Uint8Array> {
    const name = p.sheet.name;
    let buffer = sheetHead({ freeze: p.sheet.frozen, columns: p.sheet.columnMap, tabColor: p.sheet.tab, view: p.sheet.viewSettings, print: p.sheet.print }, selected, null);
    let r = 0;
    let maxCol = 0;
    for await (const raw of iterate(p.rows)) {
      r++;
      if (r > MAX_ROWS) {
        throw new XlvError('XLV_E_LIMIT_EXCEEDED', `Sheet "${name}" has more than ${String(MAX_ROWS)} rows.`, {
          location: { sheet: name },
          hint: 'Split the rows over several sheets.',
          details: { limit: 'rows', value: MAX_ROWS },
        });
      }
      let values: readonly unknown[];
      let rowStyle: CellStyle | undefined;
      let styles: readonly (CellStyle | undefined)[] | undefined;
      let info: { height: number | null; hidden: boolean } | undefined;
      if (Array.isArray(raw)) values = raw;
      else {
        const ro = readOptions(raw, ['values', 'style', 'styles', 'height', 'hidden', 'hyperlinks', 'notes'], `row ${String(r)} of sheet "${name}"`);
        const v = ro.get('values');
        if (!Array.isArray(v)) {
          throw new XlvError('XLV_E_TABLE_INVALID', `Row ${String(r)} of sheet "${name}" is not an array of values.`, { location: { sheet: name }, hint: 'Yield arrays of values, or { values: [...] } objects.', details: { row: r } });
        }
        values = v;
        rowStyle = ro.get('style') as CellStyle | undefined;
        styles = ro.get('styles') as (CellStyle | undefined)[] | undefined;
        const height = ro.get('height');
        const hidden = ro.get('hidden');
        if (height !== undefined || hidden !== undefined) {
          // Validated with the batch writer's rules, then dropped from the scratch sheet.
          p.sheet.row(r, { ...(height === undefined ? {} : { height: height as number }), ...(hidden === undefined ? {} : { hidden: hidden as boolean }) });
          info = p.sheet.rowMap.get(r);
          p.sheet.rowMap.delete(r);
        }
        for (const key of ['hyperlinks', 'notes'] as const) {
          const list = ro.get(key);
          if (list === undefined) continue;
          if (!Array.isArray(list)) {
            throw new XlvError('XLV_E_INVALID_OPTION', `${key} of row ${String(r)} of sheet "${name}" must be an array, by column position.`, {
              location: { sheet: name },
              hint: key === 'notes' ? 'For example notes: [undefined, "Checked"] for a note on column B.' : 'For example hyperlinks: [{ url: "https://example.com" }] for a link on column A.',
              details: { option: key },
            });
          }
          list.forEach((item: unknown, c) => {
            if (item === undefined || item === null) return;
            const ref = formatCellRef(r, c + 1);
            if (key === 'notes') p.sheet.note(ref, item as NoteInput);
            else p.sheet.link(ref, item as HyperlinkTarget);
          });
        }
      }
      if (values.length > MAX_COLUMNS) {
        throw new XlvError('XLV_E_LIMIT_EXCEEDED', `Row ${String(r)} of sheet "${name}" has more than ${String(MAX_COLUMNS)} values.`, {
          location: { sheet: name },
          hint: 'A sheet has 16 384 columns.',
          details: { limit: 'columns', value: MAX_COLUMNS },
        });
      }
      let cells = '';
      for (let c = 0; c < values.length; c++) {
        const ref = formatCellRef(r, c + 1);
        const converted = convertCellInput(values[c], { sheet: name, dateSystem: book.dateSystem }, ref);
        const style = styles?.[c] ?? rowStyle;
        let base = style === undefined ? {} : normalized(style);
        let id = style === undefined ? 0 : book.intern(base);
        if (converted.isDate && base.numFmt === undefined) {
          const key = `${String(id)}|${converted.hasTime ? 't' : 'd'}`;
          let dated = dateStyle.get(key);
          if (dated === undefined) {
            base = { ...base, numFmt: converted.hasTime ? DATE_TIME_FORMAT : DATE_FORMAT };
            dated = book.intern(base);
            dateStyle.set(key, dated);
          }
          id = dated;
        }
        if (converted.stored === null && id === 0) continue;
        cells += cellXml(ref, { value: converted.stored, style: id }, null, strings);
        if (c + 1 > maxCol) maxCol = c + 1;
      }
      const start = rowStart(r, info);
      if (cells !== '') buffer += `${start}>${cells}</row>`;
      else if (info !== undefined) buffer += `${start}/>`;
      if (buffer.length >= CHUNK_CHARS) {
        yield encoder.encode(buffer);
        buffer = '';
      }
    }
    if (p.autoFilter === true) p.filterRange = r > 0 && maxCol > 0 ? { start: { row: 1, col: 1 }, end: { row: r, col: maxCol } } : null;
    else if (p.autoFilter !== false) p.filterRange = p.autoFilter;
    p.notes = noteParts(sortedNotes(p.sheet.noteMap), n, p.rels, shapeBlock);
    p.sheet.noteMap.clear();
    if (p.notes !== null) shapeBlock += p.notes.blocks;
    const rules = sheetRulesXml({ conditionalFormats: p.sheet.conditionalList, dataValidations: p.sheet.validationList }, n);
    buffer += sheetTail(
      p.filterRange,
      p.sheet.mergeList,
      rules.main + hyperlinksXml([...p.sheet.linkMap.values()], p.rels) + printXml(p.sheet.print) + (p.notes?.legacyDrawing ?? '') + rules.ext,
      sheetProtectionXml(p.sheet.print.protection),
    );
    yield encoder.encode(buffer);
  }

  /** The Comments and VML parts of a sheet with notes, after the sheet has been written. */
  function* noteEntries(p: PreparedSheet): Generator<ZipStreamEntry> {
    const parts = p.notes;
    if (parts === null) return;
    yield { name: parts.comments.name, data: encoder.encode(parts.comments.xml) };
    yield { name: parts.vml.name, data: encoder.encode(parts.vml.xml) };
  }
  const noteSheets = (): number[] => prepared.flatMap((p, i) => (p.notes === null ? [] : [i + 1]));

  /**
   * The sheet's relationships part, if it has relationships. The ZIP stream reads an entry's data
   * completely before it asks for the next entry, so the sheet has been written by now.
   */
  function* sheetRelsEntry(p: PreparedSheet, n: number): Generator<ZipStreamEntry> {
    const xml = p.rels.xml();
    if (xml !== null) yield { name: sheetRelsPartName(n), data: encoder.encode(xml) };
  }

  /** The parts written after the sheets: filter ranges and styles are known only then. */
  function* closingParts(): Generator<ZipStreamEntry> {
    yield {
      name: 'xl/workbook.xml',
      data: encoder.encode(workbookXml(prepared.map((p) => ({ name: p.sheet.name, state: p.sheet.state, autoFilter: p.filterRange })), book.dateSystem, allNames())),
    };
    yield { name: 'xl/_rels/workbook.xml.rels', data: encoder.encode(workbookRelsXml(prepared.length, shared)) };
    yield { name: 'xl/styles.xml', data: encoder.encode(buildStyles(book.styles(), true, book.dxfStyles()).xml) };
    if (core !== null) yield { name: CORE_PROPERTIES_PART, data: encoder.encode(core) };
    const app = appXml();
    if (app !== null) yield { name: APP_PROPERTIES_PART, data: encoder.encode(app) };
    if (strings) yield { name: 'xl/sharedStrings.xml', data: encoder.encode(strings.xml()) };
  }
  /** `docProps/app.xml` (W2-52); known once the sheets and names are. */
  const appXml = (): string | null => appPropertiesXml(book.properties, prepared.map((p) => p.sheet.name), nameTitles(allNames()));
  /** The user's defined names and the print area and titles of every sheet (SR-30). */
  const allNames = (): NormalizedDefinedName[] => [...book.definedNames, ...prepared.flatMap((p) => printNames(p.sheet.name, p.sheet.print))];
  async function* entries(): AsyncGenerator<ZipStreamEntry> {
    // The package-level parts follow the sheets: which sheets have notes, and so which parts the
    // content types must name, is known only after the rows (D-A12); sheets given as an iterable
    // can be added while earlier ones are written (D-409). ZIP entries may be stored in any order
    // (APPNOTE §4.3.3, Z-19).
    let selected = false;
    const sheets: AsyncIterable<unknown> | Iterable<unknown> = dynamic ? iterate(specs as Iterable<unknown>) : prepared;
    let n = 0;
    for await (const raw of sheets) {
      let p: PreparedSheet;
      if (dynamic) {
        p = prepare(raw, prepared.length);
        prepared.push(p);
      } else p = raw as PreparedSheet;
      n++;
      const select = !selected && p.sheet.state === 'visible';
      if (select) selected = true;
      yield { name: `xl/worksheets/sheet${String(n)}.xml`, data: sheetData(p, select, n) };
      // The ZIP stream reads an entry's data completely before it asks for the next entry, so the
      // sheet has been written (and its notes and relationships are known) by now.
      yield* noteEntries(p);
      yield* sheetRelsEntry(p, n);
    }
    if (prepared.length === 0) throw new XlvError('XLV_E_INVALID_INPUT', 'The workbook has no sheets.', { hint: 'Add at least one sheet before closing the stream.' });
    if (!selected) throw noVisibleSheet();
    if (dynamic) addNames();
    const app = appXml() !== null;
    yield { name: '[Content_Types].xml', data: encoder.encode(contentTypesXml(prepared.length, shared, core !== null, app, noteSheets())) };
    yield { name: '_rels/.rels', data: encoder.encode(rootRelsXml(core !== null, app)) };
    yield* closingParts();
  }
  return createZipStream(entries());
}

/**
 * Options for {@link writeTableStream}.
 *
 * @example
 * ```ts
 * import type { WriteTableStreamOptions } from 'xlvane';
 *
 * const options: WriteTableStreamOptions = { columns: ['id', { key: 'total', numFmt: '#,##0.00', width: 12 }], sheet: 'Orders' };
 * ```
 */
export interface WriteTableStreamOptions {
  /** Columns in order (required: rows are not looked at in advance). */
  readonly columns: readonly TableColumnSpec[];
  /** Default `"Sheet1"`. */
  readonly sheet?: string;
  /** `"report"` (default): bold header with fill and border, frozen header row, filter. `"plain"`: values only. */
  readonly preset?: 'report' | 'plain';
  readonly dateSystem?: DateSystem;
  readonly sharedStrings?: boolean;
}

const REPORT_HEADER: CellStyle = { font: { bold: true }, fill: { color: 'D9E1F2' }, border: { bottom: { style: 'thin' } } };

/**
 * Streams row objects into a formatted XLSX table, like `writeTable` but with constant memory.
 * Because rows are not read in advance, the columns must be given, and widths are not fitted to
 * the content (set `width` per column).
 *
 * @example
 * ```ts
 * import { writeTableStream } from 'xlvane';
 *
 * async function* orders() {
 *   for (let id = 1; id <= 500_000; id++) yield { id, customer: `C${id % 100}`, total: id * 1.5 };
 * }
 * const stream = writeTableStream(orders(), { columns: ['id', 'customer', { key: 'total', numFmt: '#,##0.00' }] });
 * ```
 */
export function writeTableStream(rows: Iterable<object> | AsyncIterable<object>, options: WriteTableStreamOptions): ReadableStream<Uint8Array> {
  const o = readOptions(options, ['columns', 'sheet', 'preset', 'dateSystem', 'sharedStrings'], 'writeTableStream options');
  const spec = o.get('columns');
  if (!Array.isArray(spec) || spec.length === 0) {
    throw new XlvError('XLV_E_INVALID_OPTION', 'writeTableStream needs the "columns" option.', {
      hint: 'Rows are not read in advance, so list the columns, for example columns: ["id", "name", { key: "total", numFmt: "#,##0.00" }].',
      details: { option: 'columns' },
    });
  }
  const columns = spec.map((entry: unknown, i) => {
    if (typeof entry === 'string') return { key: entry, header: entry, numFmt: undefined as string | undefined, width: undefined as number | undefined };
    const c = readOptions(entry, ['key', 'header', 'numFmt', 'width'], `columns[${String(i)}]`);
    const key = c.get('key');
    if (typeof key !== 'string') throw new XlvError('XLV_E_INVALID_OPTION', `columns[${String(i)}] has no "key".`, { hint: 'Give each column the property it shows, e.g. { key: "total" }.', details: { option: 'columns' } });
    return { key, header: (c.get('header') as string | undefined) ?? key, numFmt: c.get('numFmt') as string | undefined, width: c.get('width') as number | undefined };
  });
  const preset = o.get('preset') ?? 'report';
  if (preset !== 'report' && preset !== 'plain') {
    throw new XlvError('XLV_E_INVALID_OPTION', 'The option "preset" must be "report" or "plain".', { hint: 'Use preset: "report" or "plain".', details: { option: 'preset' } });
  }
  if (!isIterable(rows)) {
    throw new XlvError('XLV_E_TABLE_INVALID', 'writeTableStream expects an iterable of row objects.', { hint: 'Pass an array, a generator or an async iterable.', details: { row: 0 } });
  }
  const cellStyles = columns.map((c) => (c.numFmt === undefined ? undefined : ({ numFmt: c.numFmt } as CellStyle)));
  async function* table(): AsyncGenerator<StreamRowInput> {
    yield { values: columns.map((c) => c.header), ...(preset === 'report' ? { style: REPORT_HEADER } : {}) };
    let n = 0;
    for await (const row of iterate(rows as Iterable<object>)) {
      n++;
      if (typeof row !== 'object' || row === null || Array.isArray(row)) {
        throw new XlvError('XLV_E_TABLE_INVALID', `Row ${String(n)} is not an object.`, { hint: 'Yield objects keyed by column, such as { id: 1, name: "Alice" }.', details: { row: n + 1 } });
      }
      const values = columns.map((c): CellInput => {
        const v = Object.prototype.hasOwnProperty.call(row, c.key) ? (row as Record<string, unknown>)[c.key] : undefined;
        if (v === undefined || v === null) return null;
        if (typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean' || v instanceof Date) return v;
        throw new XlvError('XLV_E_TABLE_INVALID', `Row ${String(n)}, column "${c.key}" holds ${typeof v}, which is not a cell value.`, {
          hint: 'Cell values are numbers, strings, booleans, Dates or null.',
          details: { row: n + 1, key: c.key },
        });
      });
      yield { values, styles: cellStyles };
    }
  }
  const shared = o.get('sharedStrings');
  const dateSystem = o.get('dateSystem');
  return writeWorkbookStream({
    ...(dateSystem === undefined ? {} : { dateSystem: dateSystem as DateSystem }),
    ...(shared === undefined ? {} : { sharedStrings: shared as boolean }),
    sheets: [
      {
        name: (o.get('sheet') as string | undefined) ?? 'Sheet1',
        columns: columns.map((c) => (c.width === undefined ? undefined : { width: c.width })),
        ...(preset === 'report' ? { freeze: { rows: 1 }, autoFilter: true } : {}),
        rows: table(),
      },
    ],
  });
}
