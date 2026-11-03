// Charts on worksheets: DrawingML chart parts (ECMA-376 Part 1 §21.2) reached through the Drawings
// part of a sheet. Facts: docs/facts/charts.md CH-01 … CH-35; decisions: docs/decisions.md D-K01.
// Writing covers column, bar, line, pie, scatter and area charts in the element order of the
// schema (CH-06 … CH-27); reading reports the chart groups, series references, titles, legend and
// axes of any chart, never evaluates a reference and never reads a cache into the model.
import { XlvError } from './errors';
import type { ImageAnchor, ImageAnchorInput } from './images';
import type { ReadLimits } from './limits';
import type { OpcPackage } from './opc/package';
import { partKey } from './opc/part-name';
import type { CellRange } from './sml/cell-ref';
import { formatCellRef } from './sml/cell-ref';
import { tokenize, type CellPart, type PrefixPart } from './sml/formula-shift';
import { decodeXstring } from './sml/xstring';
import type { WarningCollector } from './util/warnings';
import { clip } from './util/clip';
import { decodeXml } from './xml/parser';
import { childNamed, childrenNamed, parseXmlTree, type XmlElement } from './xml/tree';
import { encodeXstring, escapeText, firstLoneSurrogate } from './xml/writer';

/**
 * Chart types `NewSheet.addChart` writes. `column` and `bar` are both a `barChart` (vertical and
 * horizontal bars), the others are `lineChart`, `pieChart`, `scatterChart` and `areaChart`
 * (docs/facts/charts.md CH-10).
 *
 * @example
 * ```ts
 * import type { ChartType } from 'xlvane';
 *
 * const type: ChartType = 'column';
 * ```
 */
export type ChartType = 'column' | 'bar' | 'line' | 'pie' | 'scatter' | 'area';

/**
 * Where a chart's legend is shown (ST_LegendPos `r`, `b`, `t`, `l`; CH-27).
 *
 * @example
 * ```ts
 * import type { ChartLegendPosition } from 'xlvane';
 *
 * const legend: ChartLegendPosition = 'bottom';
 * ```
 */
export type ChartLegendPosition = 'right' | 'bottom' | 'top' | 'left';

/**
 * One series of a chart for `NewSheet.addChart`:
 * - `values`: a sheet-qualified range of one row or one column, such as `"'Sales'!B2:B13"` (for
 *   scatter charts the y values);
 * - `categories`: the same for the category labels (for scatter charts the x values);
 * - `name`: the series name, either literal text or a reference to one cell such as `"Sales!B1"`;
 *   text that reads as a sheet-qualified cell reference is always taken as a reference;
 * - `color`: the colour as six hex digits `RRGGBB` (a leading `#` is accepted). Without it the
 *   series takes the next colour of the default palette (CH-35). Pie charts colour each slice
 *   from the palette and do not take `color`.
 *
 * @example
 * ```ts
 * import type { ChartSeriesInput } from 'xlvane';
 *
 * const series: ChartSeriesInput = { name: 'Sales!B1', categories: 'Sales!A2:A13', values: 'Sales!B2:B13', color: '2E75B6' };
 * ```
 */
export interface ChartSeriesInput {
  readonly values: string;
  readonly categories?: string;
  readonly name?: string;
  readonly color?: string;
}

/**
 * Options of one chart axis: its title, and for the value axis (`y`) the fixed minimum and
 * maximum of its scale (CH-26).
 *
 * @example
 * ```ts
 * import type { ChartAxisInput } from 'xlvane';
 *
 * const y: ChartAxisInput = { title: 'Revenue', min: 0, max: 100 };
 * ```
 */
export interface ChartAxisInput {
  readonly title?: string;
  readonly min?: number;
  readonly max?: number;
}

/**
 * A chart for `NewSheet.addChart`. `anchor` places it like a picture (see `ImageAnchorInput`).
 * `stacked` (column, bar, line and area only) stacks the series, or with `'percent'` stacks them
 * to 100 %. `legend` defaults to `'right'`; `false` leaves the legend out. `axes.x` is the
 * category axis (the x values of a scatter chart; drawn vertically for `bar`), `axes.y` the
 * value axis; pie charts have no axes. Options left out (`undefined`) take their defaults;
 * `null` is refused. A chart has 1 to 255 series (a limit of xlvane), each an object (no holes
 * in the array), a pie chart exactly one with at most 1 000 values; titles and literal series
 * names are 1 to 255 characters without control characters. `writeWorkbook` writes the cached
 * values of at most 4 000 000 cells for all charts of a workbook together (a limit of xlvane).
 * The anchor is copied when the chart is added. A sheet name that reads as a cell reference is
 * given in quotes (`'S1'!B2:B5`); without quotes it must be exactly the name of a sheet.
 *
 * @example
 * ```ts
 * import type { ChartInput } from 'xlvane';
 *
 * const chart: ChartInput = {
 *   type: 'line',
 *   anchor: 'E2:L18',
 *   title: 'Monthly sales',
 *   series: [{ name: 'Sales!B1', categories: 'Sales!A2:A13', values: 'Sales!B2:B13' }],
 *   axes: { y: { title: 'EUR', min: 0 } },
 * };
 * ```
 */
export interface ChartInput {
  readonly type: ChartType;
  readonly anchor: ImageAnchorInput;
  readonly series: readonly ChartSeriesInput[];
  readonly title?: string;
  readonly legend?: ChartLegendPosition | false;
  readonly stacked?: boolean | 'percent';
  readonly axes?: { readonly x?: { readonly title?: string }; readonly y?: ChartAxisInput };
  /** Name of the chart object, by default "Chart N". */
  readonly name?: string;
  /** Alternative text of the chart object. */
  readonly description?: string;
}

/**
 * The type of a chart group as read: a {@link ChartType} (`column` or `bar` from `barDir`), or
 * `other:` and the element name for chart types xlvane does not model, such as
 * `'other:radarChart'` or `'other:bar3DChart'`.
 *
 * @example
 * ```ts
 * import type { ChartKind } from 'xlvane';
 *
 * const kinds: ChartKind[] = ['column', 'line', 'other:doughnutChart'];
 * ```
 */
export type ChartKind = ChartType | `other:${string}`;

/**
 * A title or series name as read: literal `text`, or a `ref`erence to a cell (the formula text as
 * stored, for example `Sheet1!$B$1`; never evaluated).
 *
 * @example
 * ```ts
 * import type { ChartText } from 'xlvane';
 *
 * const show = (t: ChartText | null): string => (t === null ? '(none)' : 'text' in t ? t.text : `=${t.ref}`);
 * ```
 */
export type ChartText = { readonly text: string } | { readonly ref: string };

/**
 * An axis of a chart as read: its title and the fixed `min` and `max` of its scale, or `null`
 * where the file sets none.
 *
 * @example
 * ```ts
 * import type { ChartAxis } from 'xlvane';
 *
 * const axis: ChartAxis = { title: { text: 'EUR' }, min: 0, max: null };
 * ```
 */
export interface ChartAxis {
  readonly title: ChartText | null;
  readonly min: number | null;
  readonly max: number | null;
}

/**
 * A series of a chart as read. `categories` and `values` are the reference texts as stored (the
 * x and y values of a scatter series), or `null` when the series has none or holds literal data.
 *
 * @example
 * ```ts
 * import { readWorkbook, type ChartSeries } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const series: readonly ChartSeries[] = readWorkbook(bytes).getSheet('Report').charts[0]?.series ?? [];
 * for (const s of series) console.log(s.type, s.values);
 * ```
 */
export interface ChartSeries {
  readonly type: ChartKind;
  readonly name: ChartText | null;
  readonly categories: string | null;
  readonly values: string | null;
}

/**
 * A chart on a worksheet (a `graphicFrame` of the sheet's drawing and its chart part, CH-02,
 * CH-04). `types` lists the chart groups of the plot area in order; a combination chart has
 * several, for example `['column', 'line']`. `stacked` comes from the first group that has a
 * grouping (`true` for `stacked`, `'percent'` for `percentStacked`, `false` otherwise; `null`
 * for pie, scatter and unmodelled types). `title` is `null` without a title and for an automatic
 * title whose text the file does not hold. `legend` is `null` without a legend. `axes.x` is the
 * category axis (the x axis of a scatter chart) and `axes.y` the value axis of the first group
 * with axes. Charts on chart sheets are not read.
 *
 * @example
 * ```ts
 * import { readWorkbook, type SheetChart } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const charts: readonly SheetChart[] = readWorkbook(bytes).getSheet('Report').charts;
 * for (const c of charts) console.log(c.types.join('+'), c.series.map((s) => s.values));
 * ```
 */
export interface SheetChart {
  readonly name: string;
  readonly description: string | null;
  readonly anchor: ImageAnchor;
  readonly types: readonly ChartKind[];
  readonly stacked: boolean | 'percent' | null;
  readonly title: ChartText | null;
  readonly legend: ChartLegendPosition | 'topRight' | null;
  readonly axes: { readonly x: ChartAxis | null; readonly y: ChartAxis | null };
  readonly series: readonly ChartSeries[];
}

/** Content type of a chart part (CH-01). */
export const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
/** Chart namespace (Transitional) and the `graphicData` uri of a chart frame (CH-01, CH-04). */
export const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Most series of a chart and characters of a title or literal series name (limits of xlvane). */
export const MAX_CHART_SERIES = 255;
const MAX_CHART_TEXT = 255;
/** Most values of a pie series: each slice gets a `dPt` with its colour (CH-21). */
const MAX_PIE_POINTS = 1000;
/**
 * Most cells the caches of all charts of one workbook written by `writeWorkbook` cover (a limit
 * of xlvane). A cached point takes about 42 bytes of XML (a 200 000-point scatter chart is a
 * 16.8 MB part), so the chart parts stay below about 170 MB, within the reader's default
 * `maxEntryBytes`. The streaming writer writes no caches (CH-20) and has no such limit.
 */
export const MAX_CHART_CACHE_CELLS = 4_000_000;
/** Longest reference text accepted, far above any sheet name (31) plus two cells. */
const MAX_REF_LENGTH = 512;

/**
 * Default series colours: accent1 … accent6 of the theme part of `chartTitle_withTitle.xlsx`
 * (Microsoft Excel 15.0300) in the POI corpus (CH-35). Series and pie slices take them in order
 * and start again after the sixth.
 */
export const CHART_PALETTE: readonly string[] = ['5B9BD5', 'ED7D31', 'A5A5A5', 'FFC000', '4472C4', '70AD47'];

// ---------------------------------------------------------------------------------------------
// Validation

/** A validated reference: the sheet as written by the caller (resolved at write time) and the cells. */
export interface ChartRef {
  readonly sheet: string;
  readonly range: CellRange;
  /**
   * `true` when the sheet name was given without quotes although it reads as a cell or R1C1
   * reference (such as `S1!$A$1`); it must then be exactly the name of a sheet of the workbook.
   */
  readonly unquoted?: boolean;
}

/** A validated series. */
export interface NormalizedSeries {
  readonly name: { readonly text: string } | { readonly ref: ChartRef } | null;
  readonly values: ChartRef;
  readonly categories: ChartRef | null;
  readonly color: string | null;
}

/** A validated chart, without its anchor. */
export interface NormalizedChart {
  readonly type: ChartType;
  readonly title: string | null;
  readonly legend: ChartLegendPosition | null;
  readonly stacked: boolean | 'percent';
  readonly series: readonly NormalizedSeries[];
  readonly xTitle: string | null;
  readonly yTitle: string | null;
  readonly min: number | null;
  readonly max: number | null;
  readonly name: string | null;
  readonly description: string | null;
}

const TYPES: ReadonlySet<string> = new Set(['column', 'bar', 'line', 'pie', 'scatter', 'area']);
const LEGENDS: ReadonlySet<string> = new Set(['right', 'bottom', 'top', 'left']);
const CONTROL = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/;
/** Titles and literal series names are one line: every C0 control is refused (paragraphs are R6-11). */
const LINE_CONTROL = /[\u0000-\u001f\ufffe\uffff]/;

function chartError(message: string, hint: string, sheet: string, option: string): XlvError {
  return new XlvError('XLV_E_INVALID_CHART', message, { location: { sheet }, hint, details: { option } });
}

function keysOf(value: unknown, allowed: readonly string[], what: string, example: string, sheet: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) throw chartError(`The ${what} must be an object.`, `For example ${example}.`, sheet, what);
  for (const key of Object.keys(value)) {
    if (!allowed.includes(key)) throw chartError(`Unknown ${what} option "${key.slice(0, 40)}".`, `Use ${allowed.join(', ')}.`, sheet, key.slice(0, 40));
  }
  return value as Record<string, unknown>;
}

/**
 * Parses a sheet-qualified reference such as `'Sales'!B2:B13` or `Sales!$B$1` (an optional
 * leading `=` is ignored) with the formula tokenizer (F2-11); `null` if the text is not one. An
 * unquoted sheet name of ASCII letters, digits and `_` that the tokenizer reads as a cell or R1C1
 * reference, such as `S1!$A$1` as LibreOffice writes it back, gives a reference marked
 * `unquoted`, which `resolveChartSheets` accepts only for a sheet of exactly that name.
 */
export function parseChartRef(text: string): ChartRef | null {
  if (text.length > MAX_REF_LENGTH) return null;
  const body = text.startsWith('=') ? text.slice(1) : text;
  const ref = refOfTokens(tokenize(body));
  if (ref !== null) return ref;
  const bare = /^([A-Za-z_][A-Za-z0-9_]*)!/.exec(body);
  if (bare === null) return null;
  const quoted = refOfTokens(tokenize(`'${bare[1]!}'${body.slice(bare[1]!.length)}`));
  return quoted === null ? null : { ...quoted, unquoted: true };
}

function refOfTokens(parts: ReturnType<typeof tokenize>): ChartRef | null {
  if (parts === null || (parts.length !== 2 && parts.length !== 4)) return null;
  const [prefix, a, colon, b] = parts as [PrefixPart, CellPart, string | undefined, CellPart | undefined];
  if (typeof prefix !== 'object' || prefix.kind !== 'prefix' || prefix.external || prefix.sheets.length !== 1) return null;
  if (typeof a !== 'object' || a.kind !== 'cell') return null;
  if (parts.length === 4 && (colon !== ':' || typeof b !== 'object' || b.kind !== 'cell')) return null;
  const end = b ?? a;
  return {
    sheet: prefix.sheets[0]!,
    range: { start: { row: Math.min(a.row, end.row), col: Math.min(a.col, end.col) }, end: { row: Math.max(a.row, end.row), col: Math.max(a.col, end.col) } },
  };
}

function checkText(value: unknown, what: string, max: number, sheet: string): string {
  if (typeof value !== 'string' || value === '' || value.length > max || LINE_CONTROL.test(value) || firstLoneSurrogate(value) !== -1) {
    throw chartError(`The chart ${what} must be a text of 1 to ${String(max)} characters on one line without control characters.`, `For example ${what}: "Sales".`, sheet, what);
  }
  return value;
}

function checkRange(value: unknown, what: string, sheet: string): ChartRef {
  const ref = typeof value === 'string' ? parseChartRef(value.trim()) : null;
  if (ref === null) {
    throw chartError(`The chart series ${what} ${typeof value === 'string' ? `"${value.slice(0, 60)}"` : ''} is not a sheet-qualified range.`, `Give the sheet and the cells, for example ${what}: "'Sales'!B2:B13".`, sheet, what);
  }
  if (ref.range.start.row !== ref.range.end.row && ref.range.start.col !== ref.range.end.col) {
    throw chartError(`The chart series ${what} must be one row or one column, not a block of cells.`, `For example "'Sales'!B2:B13" (a column) or "'Sales'!B2:M2" (a row).`, sheet, what);
  }
  return ref;
}

function cellCount(r: ChartRef): number {
  return (r.range.end.row - r.range.start.row + 1) * (r.range.end.col - r.range.start.col + 1);
}

function normalizeSeries(raw: unknown, t: ChartType, sheet: string): NormalizedSeries {
  const s = keysOf(raw, ['values', 'categories', 'name', 'color'], 'series', '{ values: "Sales!B2:B13" }', sheet);
  const values = checkRange(s['values'], 'values', sheet);
  const categories = s['categories'] === undefined ? null : checkRange(s['categories'], 'categories', sheet);
  let name: NormalizedSeries['name'] = null;
  if (s['name'] !== undefined) {
    const ref = typeof s['name'] === 'string' ? parseChartRef(s['name'].trim()) : null;
    if (ref !== null && cellCount(ref) !== 1) throw chartError('A series name reference must be one cell.', 'For example name: "Sales!B1", or a literal name such as name: "Sales".', sheet, 'name');
    name = ref === null ? { text: checkText(s['name'], 'name', MAX_CHART_TEXT, sheet) } : { ref };
  }
  let color: string | null = null;
  if (s['color'] !== undefined) {
    if (t === 'pie') throw chartError('The slices of a pie chart take their colours from the palette; a pie series has no color.', 'Leave color out of the pie series.', sheet, 'color');
    const c = s['color'];
    if (typeof c !== 'string' || !/^#?[0-9A-Fa-f]{6}$/.test(c)) throw chartError('A series color must be six hex digits RRGGBB.', 'For example color: "2E75B6".', sheet, 'color');
    color = c.replace('#', '').toUpperCase();
  }
  if (t === 'pie' && cellCount(values) > MAX_PIE_POINTS) throw chartError(`A pie chart has at most ${String(MAX_PIE_POINTS)} values.`, 'Use fewer cells, or a column chart for many values.', sheet, 'values');
  return { name, values, categories, color };
}

/**
 * Validates the options of `NewSheet.addChart` except the anchor (checked by the caller) and
 * the sheets of the references (checked when the workbook is written). Throws
 * `XLV_E_INVALID_CHART`.
 */
export function normalizeChart(input: unknown, sheet: string): NormalizedChart {
  const o = keysOf(input, ['type', 'anchor', 'series', 'title', 'legend', 'stacked', 'axes', 'name', 'description'], 'chart', '{ type: "column", anchor: "E2:L18", series: [{ values: "Sales!B2:B13" }] }', sheet);
  const type = o['type'];
  if (typeof type !== 'string' || !TYPES.has(type)) throw chartError(`The chart type ${JSON.stringify(String(type).slice(0, 20))} is not supported.`, 'Use "column", "bar", "line", "pie", "scatter" or "area".', sheet, 'type');
  const t = type as ChartType;
  const list = o['series'];
  if (!Array.isArray(list) || list.length === 0 || list.length > MAX_CHART_SERIES) {
    throw chartError(`A chart needs 1 to ${String(MAX_CHART_SERIES)} series.`, 'For example series: [{ name: "Sales!B1", categories: "Sales!A2:A13", values: "Sales!B2:B13" }].', sheet, 'series');
  }
  if (t === 'pie' && list.length !== 1) throw chartError('A pie chart has exactly one series.', 'Give one series, or use a column or bar chart to compare several.', sheet, 'series');
  const series: NormalizedSeries[] = [];
  for (let i = 0; i < list.length; i++) {
    // A hole would be skipped by map() and leave a chart with fewer series than given.
    if (!Object.hasOwn(list, i)) throw chartError(`Series ${String(i + 1)} of the chart is missing (the series array has a hole).`, 'Give an object such as { values: "Sales!B2:B13" } for every series.', sheet, 'series');
    series.push(normalizeSeries(list[i], t, sheet));
  }
  const title = o['title'] === undefined ? null : checkText(o['title'], 'title', MAX_CHART_TEXT, sheet);
  // Only the documented values or undefined: null is not "use the default".
  const legend = o['legend'] === undefined ? 'right' : o['legend'];
  if (legend !== false && (typeof legend !== 'string' || !LEGENDS.has(legend))) throw chartError('The chart legend must be "right", "bottom", "top", "left" or false.', 'For example legend: "bottom", or legend: false for no legend.', sheet, 'legend');
  const stacked = o['stacked'] === undefined ? false : o['stacked'];
  if (stacked !== false && stacked !== true && stacked !== 'percent') throw chartError('stacked must be true, false or "percent".', 'For example stacked: true.', sheet, 'stacked');
  if (stacked !== false && (t === 'pie' || t === 'scatter')) throw chartError(`A ${t} chart cannot be stacked.`, 'Use stacked with column, bar, line or area charts.', sheet, 'stacked');
  let xTitle: string | null = null;
  let yTitle: string | null = null;
  let min: number | null = null;
  let max: number | null = null;
  if (o['axes'] !== undefined) {
    if (t === 'pie') throw chartError('A pie chart has no axes.', 'Leave axes out of a pie chart.', sheet, 'axes');
    const axes = keysOf(o['axes'], ['x', 'y'], 'axes', '{ y: { title: "EUR", min: 0 } }', sheet);
    if (axes['x'] !== undefined) {
      const x = keysOf(axes['x'], ['title'], 'x axis', '{ title: "Month" }', sheet);
      if (x['title'] !== undefined) xTitle = checkText(x['title'], 'axis title', MAX_CHART_TEXT, sheet);
    }
    if (axes['y'] !== undefined) {
      const y = keysOf(axes['y'], ['title', 'min', 'max'], 'y axis', '{ title: "EUR", min: 0, max: 100 }', sheet);
      if (y['title'] !== undefined) yTitle = checkText(y['title'], 'axis title', MAX_CHART_TEXT, sheet);
      for (const k of ['min', 'max'] as const) {
        const v = y[k];
        if (v === undefined) continue;
        if (typeof v !== 'number' || !Number.isFinite(v)) throw chartError(`The y axis ${k} must be a finite number.`, `For example ${k}: 0.`, sheet, k);
        if (k === 'min') min = v;
        else max = v;
      }
      if (min !== null && max !== null && min >= max) throw chartError('The y axis min must be less than its max.', 'For example { min: 0, max: 100 }.', sheet, 'min');
    }
  }
  const objectText = (key: 'name' | 'description'): string | null => {
    const v = o[key];
    if (v === undefined) return null;
    // xsd:string attributes of cNvPr, as for pictures (TD-35, D-D05).
    if (typeof v !== 'string' || v.length > 2084 || CONTROL.test(v) || firstLoneSurrogate(v) !== -1 || (key === 'name' && v === '')) {
      throw chartError(`The chart ${key} must be a text of ${key === 'name' ? '1' : '0'} to 2084 characters without control characters.`, `For example ${key}: "Sales chart".`, sheet, key);
    }
    return v;
  };
  return { type: t, title, legend: legend === false ? null : (legend as ChartLegendPosition), stacked: stacked as boolean | 'percent', series, xTitle, yTitle, min, max, name: objectText('name'), description: objectText('description') };
}

/**
 * Replaces the sheet of every reference by the workbook's sheet of that name (names compare
 * without regard to case; an `unquoted` reference must name a sheet exactly). Throws
 * `XLV_E_INVALID_CHART` for a sheet the workbook does not have.
 */
export function resolveChartSheets(chart: NormalizedChart, sheets: readonly string[], sheet: string): NormalizedChart {
  const byKey = new Map(sheets.map((s) => [s.toLowerCase(), s] as const));
  const fix = (r: ChartRef): ChartRef => {
    if (r.unquoted === true) {
      if (sheets.includes(r.sheet)) return { sheet: r.sheet, range: r.range };
      const quoted = chartRefText({ sheet: r.sheet, range: r.range });
      throw new XlvError(
        'XLV_E_INVALID_CHART',
        `A chart on sheet "${sheet}" refers to sheet "${r.sheet.slice(0, 100)}" without quotes. Such a name reads as a cell or R1C1 reference, so it is accepted only as exactly the name of a sheet of the workbook, and the workbook has no sheet "${r.sheet.slice(0, 100)}".`,
        { location: { sheet }, hint: `Quote the sheet name, for example "${quoted}", or use the exact name of one of the sheets: ${sheets.slice(0, 10).map((s) => `"${s}"`).join(', ')}${sheets.length > 10 ? ', …' : ''}.`, details: { option: 'series', sheet: r.sheet.slice(0, 100) } },
      );
    }
    const name = byKey.get(r.sheet.toLowerCase());
    if (name === undefined) {
      throw new XlvError('XLV_E_INVALID_CHART', `A chart on sheet "${sheet}" refers to sheet "${r.sheet.slice(0, 100)}", which the workbook does not have.`, {
        location: { sheet },
        hint: `Use one of the sheets: ${sheets.slice(0, 10).map((s) => `"${s}"`).join(', ')}${sheets.length > 10 ? ', …' : ''}.`,
        details: { option: 'series', sheet: r.sheet.slice(0, 100) },
      });
    }
    return { sheet: name, range: r.range };
  };
  return {
    ...chart,
    series: chart.series.map((s) => ({ ...s, values: fix(s.values), categories: s.categories === null ? null : fix(s.categories), name: s.name !== null && 'ref' in s.name ? { ref: fix(s.name.ref) } : s.name })),
  };
}

// ---------------------------------------------------------------------------------------------
// Writing

/**
 * Checks the cells the caches of `charts` (all charts of a workbook, in order) would cover against
 * {@link MAX_CHART_CACHE_CELLS} before any chart part is built. Throws `XLV_E_INVALID_CHART`.
 */
export function checkChartCacheCells(charts: readonly { readonly chart: NormalizedChart; readonly sheet: string }[]): void {
  let total = 0;
  for (const { chart, sheet } of charts) {
    for (const s of chart.series) total += cellCount(s.values) + (s.categories === null ? 0 : cellCount(s.categories)) + (s.name !== null && 'ref' in s.name ? 1 : 0);
    if (total > MAX_CHART_CACHE_CELLS) {
      throw new XlvError('XLV_E_INVALID_CHART', `The charts of the workbook would hold cached values of ${String(total)} cells (counted up to a chart on sheet "${sheet}"), more than the limit of ${String(MAX_CHART_CACHE_CELLS)} cells.`, {
        location: { sheet },
        hint: 'Refer charts to fewer cells, or write the workbook with writeWorkbookStream, which writes charts without cached values.',
        details: { option: 'series', limit: MAX_CHART_CACHE_CELLS, value: total },
      });
    }
  }
}

/** A cell value for the caches (CH-19): a number with its number format, or text. */
export type ChartCacheValue = { readonly n: number; readonly format: string } | { readonly s: string };

/** Looks up the value of a cell of the workbook being written; `null` for an empty cell. */
export type ChartCells = (sheet: string, row: number, col: number) => ChartCacheValue | null;

/**
 * Quotes a sheet name for a reference (F2-11): unquoted only for ASCII letters, digits and `_`
 * starting with a letter or `_`, not `TRUE` or `FALSE`, not letters followed by digits (such as
 * `XFE1` or `Sheet1`, which can read as a cell even beyond the grid) and not read as a cell or
 * R1C1 reference by the formula tokenizer, the form Excel writes for names such as `Data`
 * (CH-18); quoted otherwise, which is valid for every name (F2-11).
 */
export function chartSheetPrefix(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name) && !/^(true|false)$/i.test(name) && !/^[A-Za-z]+[0-9]+$/.test(name)) {
    const parts = tokenize(`${name}!A1`);
    const prefix = parts?.[0];
    if (parts?.length === 2 && typeof prefix === 'object' && prefix.kind === 'prefix' && prefix.sheets[0] === name) return `${name}!`;
  }
  return `'${name.replace(/'/g, "''")}'!`;
}

/** The `f` text of a reference: sheet prefix and absolute cells, e.g. `Sales!$B$2:$B$13` (CH-18). */
export function chartRefText(r: ChartRef): string {
  const abs = (row: number, col: number): string => formatCellRef(row, col).replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');
  const a = abs(r.range.start.row, r.range.start.col);
  const single = r.range.start.row === r.range.end.row && r.range.start.col === r.range.end.col;
  return `${chartSheetPrefix(r.sheet)}${single ? a : `${a}:${abs(r.range.end.row, r.range.end.col)}`}`;
}

const val = (local: string, v: string | number): string => `<c:${local} val="${String(v)}"/>`;
const solid = (color: string): string => `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
const numText = (n: number): string => String(n).replace('e+', 'E').replace('e-', 'E-');

/** Literal rich text (CT_Tx `rich`, CH-22): one paragraph with one run. */
function richXml(text: string, rotate: boolean): string {
  return `<c:tx><c:rich><a:bodyPr${rotate ? ' rot="-5400000" vert="horz"' : ''}/><a:lstStyle/><a:p><a:r><a:t>${escapeText(text)}</a:t></a:r></a:p></c:rich></c:tx>`;
}

/** CT_Title: `tx`, `overlay` (CH-22, CH-25); the vertical axis title is rotated as excelize and LibreOffice write it. */
function titleXml(text: string | null, rotate = false): string {
  return text === null ? '' : `<c:title>${richXml(text, rotate)}${val('overlay', 0)}</c:title>`;
}

/** The cells of a reference in order (one row or one column). */
function cellsOf(r: ChartRef, cells: ChartCells): (ChartCacheValue | null)[] {
  const out: (ChartCacheValue | null)[] = [];
  for (let row = r.range.start.row; row <= r.range.end.row; row++) {
    for (let col = r.range.start.col; col <= r.range.end.col; col++) out.push(cells(r.sheet, row, col));
  }
  return out;
}

/** CT_StrData: `ptCount`, then a `pt` for every cell with a value (CH-19). */
function strCache(values: readonly (ChartCacheValue | null)[]): string {
  let xml = `<c:strCache>${val('ptCount', values.length)}`;
  values.forEach((v, i) => {
    if (v !== null) xml += `<c:pt idx="${String(i)}"><c:v>${escapeText(encodeXstring('s' in v ? v.s : numText(v.n)))}</c:v></c:pt>`;
  });
  return `${xml}</c:strCache>`;
}

/** CT_NumData: `formatCode` (of the first number), `ptCount`, then a `pt` for every number (CH-19). */
function numCache(values: readonly (ChartCacheValue | null)[]): string {
  const first = values.find((v): v is { n: number; format: string } => v !== null && 'n' in v);
  let xml = `<c:numCache><c:formatCode>${escapeText(encodeXstring(first?.format ?? 'General'))}</c:formatCode>${val('ptCount', values.length)}`;
  values.forEach((v, i) => {
    if (v !== null && 'n' in v) xml += `<c:pt idx="${String(i)}"><c:v>${numText(v.n)}</c:v></c:pt>`;
  });
  return `${xml}</c:numCache>`;
}

/**
 * A data source: `numRef` or `strRef` with the reference and, when `cells` is given, its cache.
 * `kind` `'auto'` (categories) is a `numRef` when every non-empty cell holds a number and a
 * `strRef` otherwise; without cells (the streaming writer) it is a `strRef`.
 */
function dataXml(local: string, r: ChartRef, kind: 'num' | 'str' | 'auto', cells: ChartCells | null): string {
  const values = cells === null ? null : cellsOf(r, cells);
  const k = kind !== 'auto' ? kind : values !== null && values.some((v) => v !== null) && values.every((v) => v === null || 'n' in v) ? 'num' : 'str';
  const cache = values === null ? '' : k === 'num' ? numCache(values) : strCache(values);
  return `<c:${local}><c:${k}Ref><c:f>${escapeText(chartRefText(r))}</c:f>${cache}</c:${k}Ref></c:${local}>`;
}

/** CT_SerTx: literal `v` or a `strRef` to one cell (CH-17). */
function seriesNameXml(name: NormalizedSeries['name'], cells: ChartCells | null): string {
  if (name === null) return '';
  if ('text' in name) return `<c:tx><c:v>${escapeText(encodeXstring(name.text))}</c:v></c:tx>`;
  const cache = cells === null ? '' : strCache(cellsOf(name.ref, cells));
  return `<c:tx><c:strRef><c:f>${escapeText(chartRefText(name.ref))}</c:f>${cache}</c:strRef></c:tx>`;
}

/** One series in the order of its type's sequence (CH-16), with an explicit colour (CH-21). */
function seriesXml(type: ChartType, s: NormalizedSeries, i: number, cells: ChartCells | null): string {
  const color = s.color ?? CHART_PALETTE[i % CHART_PALETTE.length]!;
  let xml = `<c:ser>${val('idx', i)}${val('order', i)}${seriesNameXml(s.name, cells)}`;
  const cat = s.categories === null ? '' : dataXml(type === 'scatter' ? 'xVal' : 'cat', s.categories, type === 'scatter' ? 'num' : 'auto', cells);
  const values = dataXml(type === 'scatter' ? 'yVal' : 'val', s.values, 'num', cells);
  if (type === 'line') {
    // CT_LineSer: spPr (the line colour on a:ln), marker (none), cat, val, smooth.
    xml += `<c:spPr><a:ln>${solid(color)}</a:ln></c:spPr><c:marker>${val('symbol', 'none')}</c:marker>${cat}${values}${val('smooth', 0)}`;
  } else if (type === 'scatter') {
    // CT_ScatterSer: markers only (no line), filled with the series colour (CH-14, CH-21).
    xml += `<c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:marker>${val('symbol', 'circle')}${val('size', 5)}<c:spPr>${solid(color)}<a:ln>${solid(color)}</a:ln></c:spPr></c:marker>${cat}${values}${val('smooth', 0)}`;
  } else if (type === 'pie') {
    // CT_PieSer: one dPt per value with its colour, since varyColors colours each slice (CH-13, CH-21).
    let points = '';
    for (let k = 0; k < cellCount(s.values); k++) points += `<c:dPt>${val('idx', k)}<c:spPr>${solid(CHART_PALETTE[k % CHART_PALETTE.length]!)}</c:spPr></c:dPt>`;
    xml += `${points}${cat}${values}`;
  } else {
    // CT_BarSer and CT_AreaSer: spPr with the fill, cat, val.
    xml += `<c:spPr>${solid(color)}</c:spPr>${cat}${values}`;
  }
  return `${xml}</c:ser>`;
}

/** Axis ids of the one chart group of a part: fixed, unique within the part (CH-24, PROVISIONAL). */
const AXIS_IDS = [500000001, 500000002] as const;

function axisXml(local: 'catAx' | 'valAx', id: number, cross: number, pos: 'b' | 'l', title: string | null, rotate: boolean, opts: { gridlines?: boolean; crossBetween?: 'between' | 'midCat'; min?: number | null; max?: number | null }): string {
  // EG_AxShared order (CH-23); CT_Scaling has max before min (CH-26).
  const scaling = `<c:scaling>${val('orientation', 'minMax')}${opts.max == null ? '' : val('max', numText(opts.max))}${opts.min == null ? '' : val('min', numText(opts.min))}</c:scaling>`;
  let xml = `<c:${local}>${val('axId', id)}${scaling}${val('delete', 0)}${val('axPos', pos)}${opts.gridlines ? '<c:majorGridlines/>' : ''}${titleXml(title, rotate)}`;
  xml += `<c:numFmt formatCode="General" sourceLinked="1"/>${val('tickLblPos', 'nextTo')}${val('crossAx', cross)}${val('crosses', 'autoZero')}`;
  xml += local === 'catAx' ? `${val('auto', 1)}${val('lblAlgn', 'ctr')}${val('lblOffset', 100)}` : val('crossBetween', opts.crossBetween ?? 'between');
  return `${xml}</c:${local}>`;
}

/**
 * The chart part `xl/charts/chartN.xml` (CH-06 … CH-28): `roundedCorners`, then `chart` with the
 * title, `autoTitleDeleted`, the plot area (one chart group and its axes), the legend,
 * `plotVisOnly` and `dispBlanksAs`. `cells` fills the caches (CH-19); the streaming writer passes
 * `null` and writes references without caches (CH-20). No Office extensions (CH-32).
 */
export function chartPartXml(chart: NormalizedChart, cells: ChartCells | null): string {
  const t = chart.type;
  const series = chart.series.map((s, i) => seriesXml(t, s, i, cells)).join('');
  const grouping = chart.stacked === 'percent' ? 'percentStacked' : chart.stacked ? 'stacked' : t === 'column' || t === 'bar' ? 'clustered' : 'standard';
  const [x, y] = AXIS_IDS;
  const axIds = val('axId', x) + val('axId', y);
  // CH-10 … CH-15, CH-28: varyColors written explicitly, 0 except for pie.
  let group: string;
  if (t === 'column' || t === 'bar') group = `<c:barChart>${val('barDir', t === 'bar' ? 'bar' : 'col')}${val('grouping', grouping)}${val('varyColors', 0)}${series}${val('gapWidth', 150)}${chart.stacked ? val('overlap', 100) : ''}${axIds}</c:barChart>`;
  else if (t === 'line') group = `<c:lineChart>${val('grouping', grouping)}${val('varyColors', 0)}${series}${val('marker', 1)}${axIds}</c:lineChart>`;
  else if (t === 'area') group = `<c:areaChart>${val('grouping', grouping)}${val('varyColors', 0)}${series}${axIds}</c:areaChart>`;
  else if (t === 'scatter') group = `<c:scatterChart>${val('scatterStyle', 'lineMarker')}${val('varyColors', 0)}${series}${axIds}</c:scatterChart>`;
  else group = `<c:pieChart>${val('varyColors', 1)}${series}${val('firstSliceAng', 0)}</c:pieChart>`;
  let axes = '';
  if (t !== 'pie') {
    // The category axis at the bottom and the value axis at the left, also for horizontal bars,
    // as excelize writes and LibreOffice writes back (CH-23, PROVISIONAL R6-14).
    const between = t === 'area' || t === 'scatter' ? 'midCat' : 'between';
    axes =
      t === 'scatter'
        ? axisXml('valAx', x, y, 'b', chart.xTitle, false, { crossBetween: between }) + axisXml('valAx', y, x, 'l', chart.yTitle, true, { gridlines: true, crossBetween: between, min: chart.min, max: chart.max })
        : axisXml('catAx', x, y, 'b', chart.xTitle, t === 'bar', {}) + axisXml('valAx', y, x, 'l', chart.yTitle, t !== 'bar', { gridlines: true, crossBetween: between, min: chart.min, max: chart.max });
  }
  const legend = chart.legend === null ? '' : `<c:legend>${val('legendPos', { right: 'r', bottom: 'b', top: 't', left: 'l' }[chart.legend])}${val('overlay', 0)}</c:legend>`;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">${val('roundedCorners', 0)}<c:chart>` +
    `${titleXml(chart.title)}${val('autoTitleDeleted', chart.title === null ? 1 : 0)}<c:plotArea><c:layout/>${group}${axes}</c:plotArea>${legend}${val('plotVisOnly', 1)}${val('dispBlanksAs', 'gap')}</c:chart></c:chartSpace>`
  );
}

/** Part name of chart `n` (1-based, numbered across the workbook; CH-01). */
export function chartPartName(n: number): string {
  return `xl/charts/chart${String(n)}.xml`;
}

// ---------------------------------------------------------------------------------------------
// Reading

/** What a chart part holds, shared by every frame that shows it. */
export type ParsedChart = Omit<SheetChart, 'name' | 'description' | 'anchor'>;

/** Chart parts already read, by part key, so that a part used by several frames is parsed once. */
export type ChartCache = Map<string, ParsedChart | null>;

/** Chart group elements of CT_PlotArea (CH-09). */
const GROUPS: ReadonlySet<string> = new Set(['areaChart', 'area3DChart', 'lineChart', 'line3DChart', 'stockChart', 'radarChart', 'scatterChart', 'pieChart', 'pie3DChart', 'doughnutChart', 'barChart', 'bar3DChart', 'ofPieChart', 'surfaceChart', 'surface3DChart', 'bubbleChart']);
const KINDS: ReadonlyMap<string, ChartType> = new Map([
  ['lineChart', 'line'],
  ['pieChart', 'pie'],
  ['scatterChart', 'scatter'],
  ['areaChart', 'area'],
]);
const LEGEND_POSITIONS: ReadonlyMap<string, ChartLegendPosition | 'topRight'> = new Map([
  ['r', 'right'],
  ['b', 'bottom'],
  ['t', 'top'],
  ['l', 'left'],
  ['tr', 'topRight'],
]);
const AXES: ReadonlySet<string> = new Set(['catAx', 'valAx', 'dateAx', 'serAx']);

const valOf = (el: XmlElement | undefined): string | undefined => el?.attrs.get('val');

/** The `f` of a `numRef`, `strRef` or `multiLvlStrRef` child of a data source (CH-18), or `null`. */
function refOf(source: XmlElement | undefined): string | null {
  if (source === undefined) return null;
  for (const local of ['numRef', 'strRef', 'multiLvlStrRef']) {
    const ref = childNamed(source, CHART_NS, local);
    const f = ref === undefined ? undefined : childNamed(ref, CHART_NS, 'f');
    if (f !== undefined) return f.text;
  }
  return null;
}

/**
 * The text of a CT_Tx or CT_SerTx (CH-17, CH-22): `rich` text (the `a:t` of every run and field,
 * paragraphs joined by a line feed, PROVISIONAL R6-11), a literal `v`, or the `f` of a `strRef`.
 */
function textOf(tx: XmlElement | undefined): ChartText | null {
  if (tx === undefined) return null;
  const v = childNamed(tx, CHART_NS, 'v');
  if (v !== undefined) return { text: decodeXstring(v.text) };
  const ref = childNamed(tx, CHART_NS, 'strRef');
  const f = ref === undefined ? undefined : childNamed(ref, CHART_NS, 'f');
  if (f !== undefined) return { ref: f.text };
  const rich = childNamed(tx, CHART_NS, 'rich');
  if (rich === undefined) return null;
  const paragraphs = childrenNamed(rich, A_NS, 'p').map((p) =>
    p.children
      .map((run) => (run.ns !== A_NS ? '' : run.local === 'br' ? '\n' : run.local === 'r' || run.local === 'fld' ? (childNamed(run, A_NS, 't')?.text ?? '') : ''))
      .join(''),
  );
  return { text: paragraphs.join('\n') };
}

function titleOf(el: XmlElement | undefined): ChartText | null {
  return el === undefined ? null : textOf(childNamed(el, CHART_NS, 'tx'));
}

function numberOf(el: XmlElement | undefined): number | null {
  const t = (valOf(el) ?? '').trim();
  return /^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(t) && Number.isFinite(Number(t)) ? Number(t) : null;
}

function axisOf(el: XmlElement | undefined): ChartAxis | null {
  if (el === undefined) return null;
  const scaling = childNamed(el, CHART_NS, 'scaling');
  return Object.freeze({
    title: titleOf(childNamed(el, CHART_NS, 'title')),
    min: scaling === undefined ? null : numberOf(childNamed(scaling, CHART_NS, 'min')),
    max: scaling === undefined ? null : numberOf(childNamed(scaling, CHART_NS, 'max')),
  });
}

/** The model of a `chartSpace` element; children are found by name, so order does not matter (CH-29). */
function chartOf(root: XmlElement): ParsedChart | null {
  const chart = childNamed(root, CHART_NS, 'chart');
  const plotArea = chart === undefined ? undefined : childNamed(chart, CHART_NS, 'plotArea');
  if (chart === undefined || plotArea === undefined) return null;
  const axes = new Map<string, XmlElement>();
  for (const el of plotArea.children) {
    if (el.ns !== CHART_NS || !AXES.has(el.local)) continue;
    const id = valOf(childNamed(el, CHART_NS, 'axId'));
    if (id !== undefined && !axes.has(id)) axes.set(id, el);
  }
  const types: ChartKind[] = [];
  const series: ChartSeries[] = [];
  let stacked: boolean | 'percent' | null = null;
  let axisPair: { x: ChartAxis | null; y: ChartAxis | null } | null = null;
  for (const group of plotArea.children) {
    if (group.ns !== CHART_NS || !GROUPS.has(group.local)) continue;
    // barDir is required; its schema default is col (CH-10).
    const kind: ChartKind = group.local === 'barChart' ? (valOf(childNamed(group, CHART_NS, 'barDir')) === 'bar' ? 'bar' : 'column') : (KINDS.get(group.local) ?? `other:${group.local}`);
    types.push(kind);
    if (stacked === null && (group.local === 'barChart' || group.local === 'lineChart' || group.local === 'areaChart')) {
      const g = valOf(childNamed(group, CHART_NS, 'grouping'));
      stacked = g === 'stacked' ? true : g === 'percentStacked' ? 'percent' : false;
    }
    const scatter = group.local === 'scatterChart' || group.local === 'bubbleChart';
    for (const ser of childrenNamed(group, CHART_NS, 'ser')) {
      series.push(
        Object.freeze({
          type: kind,
          name: textOf(childNamed(ser, CHART_NS, 'tx')),
          categories: refOf(childNamed(ser, CHART_NS, scatter ? 'xVal' : 'cat')),
          values: refOf(childNamed(ser, CHART_NS, scatter ? 'yVal' : 'val')),
        }),
      );
    }
    const ids = childrenNamed(group, CHART_NS, 'axId').map((a) => valOf(a));
    if (axisPair === null && ids.length >= 2) {
      const pair = ids.slice(0, 2).map((id) => (id === undefined ? undefined : axes.get(id)));
      // The category axis is x; with two value axes (scatter) the first one is (CH-14, CH-23).
      const xi = pair.findIndex((a) => a !== undefined && a.local !== 'valAx');
      const xIndex = xi === -1 ? 0 : xi;
      axisPair = { x: axisOf(pair[xIndex]), y: axisOf(pair[1 - xIndex]) };
    }
  }
  const legend = childNamed(chart, CHART_NS, 'legend');
  return Object.freeze({
    types: Object.freeze(types),
    stacked,
    title: titleOf(childNamed(chart, CHART_NS, 'title')),
    // An element without legendPos takes the schema default r (CH-27).
    legend: legend === undefined ? null : (LEGEND_POSITIONS.get(valOf(childNamed(legend, CHART_NS, 'legendPos')) ?? 'r') ?? 'right'),
    axes: Object.freeze(axisPair ?? { x: null, y: null }),
    series: Object.freeze(series),
  });
}

/**
 * Elements of a chart part that are never read: the cached and literal values of data sources
 * (CH-18, CH-19) and extensions. They are parsed under the XML limits but not built, so a chart
 * over many cells reads in time and memory linear in the part and within `maxXmlTreeElements`.
 */
const NOT_READ: ReadonlySet<string> = new Set(['numCache', 'strCache', 'multiLvlStrCache', 'numLit', 'strLit', 'extLst']);
const skipNotRead = (ns: string, local: string): boolean => ns === CHART_NS && NOT_READ.has(local);

/**
 * Reads the chart part `target` once per workbook (`cache`). A chart is optional content, so a
 * part that cannot be read is skipped with a warning and `null`, and the workbook is read:
 * - `XLV_W_PART_MISSING`: the part does not exist;
 * - `XLV_W_INVALID_VALUE`: the part's content type is not the chart content type (CH-01), its
 *   bytes are not valid UTF-8 or UTF-16 or declare another encoding, it is not well-formed XML,
 *   or it has no `chartSpace` with a chart and a plot area;
 * - `XLV_W_LIMIT_REACHED`: the part's XML is over a limit of `ctx.limits` that applies to one
 *   part (`maxXmlTreeElements`, `maxXmlDepth`, `maxXmlAttributes`, `maxXmlNameLength`,
 *   `maxXmlTextLength`). Cached and literal values are parsed under these limits but do not count
 *   against `maxXmlTreeElements`.
 *
 * Errors of the package still throw as for every part: ZIP damage (`XLV_E_ZIP_CORRUPT`), the
 * budgets of the whole file (`maxEntryBytes`, `maxTotalBytes`, `maxCompressionRatio`, as
 * `XLV_E_LIMIT_EXCEEDED`), and a DTD or entity declaration (`XLV_E_XML_DTD_FORBIDDEN`,
 * `XLV_E_XML_ENTITY_FORBIDDEN`).
 */
export function readChartPart(pkg: OpcPackage, target: string, sheet: string, ctx: { readonly limits: ReadLimits; readonly warnings: WarningCollector }, cache: ChartCache): ParsedChart | null {
  const key = partKey(target);
  const cached = cache.get(key);
  if (cached !== undefined || cache.has(key)) return cached ?? null;
  cache.set(key, null);
  const where = { part: target, sheet };
  if (!pkg.hasPart(target)) {
    ctx.warnings.add('XLV_W_PART_MISSING', 'A chart of the sheet refers to a chart part that does not exist; the chart is skipped.', where);
    return null;
  }
  const type = pkg.contentType(target);
  if (type !== CHART_CONTENT_TYPE) {
    ctx.warnings.add('XLV_W_INVALID_VALUE', `A chart of the sheet refers to a part of content type "${clip(type ?? '(none)')}", not a chart part; the chart is skipped.`, where);
    return null;
  }
  // ZIP damage and the budgets of the whole file throw here, as for every part.
  const bytes = pkg.readBytes(target)!;
  let root: XmlElement;
  try {
    root = parseXmlTree(decodeXml(bytes, target), { part: target, limits: ctx.limits, skip: skipNotRead });
  } catch (error) {
    if (error instanceof XlvError && (error.code === 'XLV_E_XML_MALFORMED' || error.code === 'XLV_E_XML_ENCODING')) {
      ctx.warnings.add('XLV_W_INVALID_VALUE', `A chart part is not ${error.code === 'XLV_E_XML_ENCODING' ? 'encoded in UTF-8 or UTF-16' : 'well-formed XML'}; the chart is skipped.`, where);
      return null;
    }
    if (error instanceof XlvError && error.code === 'XLV_E_LIMIT_EXCEEDED') {
      const limit = typeof error.details['limit'] === 'string' ? error.details['limit'] : 'a read limit';
      ctx.warnings.add('XLV_W_LIMIT_REACHED', `A chart part is over the limit ${limit}; the chart is skipped. If you trust this file, raise limits.${limit} to read it.`, where);
      return null;
    }
    throw error;
  }
  const parsed = root.ns === CHART_NS && root.local === 'chartSpace' ? chartOf(root) : null;
  if (parsed === null) ctx.warnings.add('XLV_W_INVALID_VALUE', `A chart part has no chart (root <${root.local.slice(0, 40)}>, or no chart with a plot area); the chart is skipped.`, where);
  cache.set(key, parsed);
  return parsed;
}
