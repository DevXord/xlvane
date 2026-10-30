// Data validations, conditional formatting, sheet protection and printing for @xlvane/exceljs:
// the exceljs 4.4.0 model (docs/facts/exceljs.md X4-16 … X4-19) mapped to the xlvane core
// (docs/decisions.md D-C01 … D-C04) and back. Where exceljs writes data the specification or
// Excel rejects, xlvane writes valid data instead (D-C06, docs/exceljs-compat.md).
import type {
  CfValueInput,
  ConditionalFormat as XlvConditionalFormat,
  ConditionalFormatRule as XlvConditionalFormatRule,
  ConditionalFormatRuleInput,
  DataValidation as XlvDataValidation,
  DataValidationRule,
  HeaderFooter as XlvHeaderFooter,
  PageSetup as XlvPageSetup,
  PageSetupOptions,
  PrintTitles,
  SheetProtection as XlvSheetProtection,
  SheetProtectionFlags,
} from 'xlvane';
import { colorInFor, colorOut, styleIn, styleOut } from './convert';
import { columnLetter, notSupported } from './model';
import type { ConditionalFormattingRule, ConditionalFormattingOptions, DataValidation, HeaderFooter, PageSetup, RowBreak, Style } from './types';

// ---------------------------------------------------------------------------------------------
// Addresses

const MAX_ROW = 1_048_576;
const MAX_COL = 16_384;

function colNumber(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

/** A range given as `A1`, `A1:B2`, `A:C` or `1:3` (with or without `$`), or `null`. */
function parseRange(ref: string): { top: number; left: number; bottom: number; right: number } | null {
  const parts = ref.replace(/\$/g, '').toUpperCase().split(':');
  if (parts.length < 1 || parts.length > 2) return null;
  const ends = parts.map((p) => /^([A-Z]{0,3})(\d{0,7})$/.exec(p));
  if (ends.some((m) => m === null || (m[1] === '' && m[2] === ''))) return null;
  const [a, b] = [ends[0]!, ends[ends.length - 1]!];
  const col = (m: RegExpExecArray, dflt: number): number => (m[1] === '' ? dflt : colNumber(m[1]!));
  const row = (m: RegExpExecArray, dflt: number): number => (m[2] === '' ? dflt : Number(m[2]));
  const top = row(a, 1);
  const left = col(a, 1);
  const bottom = row(b, parts.length === 1 ? top : MAX_ROW);
  const right = col(b, parts.length === 1 ? left : MAX_COL);
  if (top < 1 || left < 1 || bottom > MAX_ROW || right > MAX_COL) return null;
  return { top: Math.min(top, bottom), left: Math.min(left, right), bottom: Math.max(top, bottom), right: Math.max(left, right) };
}

function address(row: number, col: number): string {
  return `${columnLetter(col)}${String(row)}`;
}

/**
 * A range list in the A1 form a worksheet stores (`A:A` becomes `A1:A1048576`, `$` removed);
 * parts that are not ranges are passed on for the core to reject with its own error.
 */
function sqrefIn(ref: string): string {
  return ref
    .split(/\s+/)
    .filter((p) => p !== '')
    .map((p) => {
      const r = parseRange(p);
      return r ? rangeText(r) : p;
    })
    .join(' ');
}

function rangeText(r: { top: number; left: number; bottom: number; right: number }): string {
  const a = address(r.top, r.left);
  return r.top === r.bottom && r.left === r.right ? a : `${a}:${address(r.bottom, r.right)}`;
}

// ---------------------------------------------------------------------------------------------
// Data validations (X4-16; D-C01)

/**
 * Cells per sheet up to which validations read from a file are given to every cell they cover,
 * as exceljs does (data-validations-xform.js L229–239). Larger ranges, such as whole columns,
 * keep their range as the key; `find()` and `cell.dataValidation` still answer for every cell.
 */
const MAX_EXPANDED_CELLS = 100_000;

/**
 * The data validations of a worksheet (`worksheet.dataValidations`), keyed by cell address or
 * range as in exceljs (lib/doc/data-validations.js). `remove()` deletes the entry; exceljs
 * leaves an empty entry behind, which it writes as a validation without a type.
 *
 * @example
 * ```ts
 * import ExcelJS from '@xlvane/exceljs';
 *
 * const sheet = new ExcelJS.Workbook().addWorksheet('Form');
 * sheet.dataValidations.add('B2:B20', { type: 'list', allowBlank: true, formulae: ['"Yes,No"'] });
 * console.log(sheet.getCell('B5').dataValidation?.type); // list
 * ```
 */
export class DataValidations {
  model: Record<string, DataValidation | undefined>;

  constructor(model?: Record<string, DataValidation | undefined>) {
    this.model = model ?? {};
  }

  add(address: string, validation: DataValidation): DataValidation {
    this.model[address] = validation;
    return validation;
  }

  /** The validation of a cell: its own entry, or the entry of a range that contains it. */
  find(address: string): DataValidation | undefined {
    const own = this.model[address];
    if (own !== undefined) return own;
    const cell = parseRange(address);
    if (!cell) return undefined;
    for (const [key, v] of Object.entries(this.model)) {
      if (v === undefined || !key.includes(':')) continue;
      for (const part of key.split(/\s+/)) {
        const r = parseRange(part);
        if (r && cell.top >= r.top && cell.top <= r.bottom && cell.left >= r.left && cell.left <= r.right) return v;
      }
    }
    return undefined;
  }

  remove(address: string): void {
    delete this.model[address];
  }
}

function stableKey(v: unknown): string {
  if (v instanceof Date) return `D${String(v.getTime())}`;
  if (Array.isArray(v)) return `[${v.map(stableKey).join(',')}]`;
  if (v !== null && typeof v === 'object') {
    return `{${Object.keys(v)
      .sort()
      .filter((k) => (v as Record<string, unknown>)[k] !== undefined)
      .map((k) => `${k}:${stableKey((v as Record<string, unknown>)[k])}`)
      .join(',')}}`;
  }
  return JSON.stringify(v) ?? 'u';
}

/**
 * Groups equal validations of single cells into rectangles that never overlap (exceljs's
 * optimisation, data-validations-xform.js L24–105, can produce overlapping ranges; X4-16).
 */
function rectangles(cells: readonly { row: number; col: number }[]): string[] {
  const left = new Set(cells.map((c) => `${String(c.row)},${String(c.col)}`));
  const has = (r: number, c: number): boolean => left.has(`${String(r)},${String(c)}`);
  const sorted = [...cells].sort((a, b) => a.row - b.row || a.col - b.col);
  const out: string[] = [];
  for (const { row, col } of sorted) {
    if (!has(row, col)) continue;
    let right = col;
    while (has(row, right + 1)) right++;
    let bottom = row;
    for (;;) {
      let full = true;
      for (let c = col; c <= right; c++) if (!has(bottom + 1, c)) full = false;
      if (!full) break;
      bottom++;
    }
    for (let r = row; r <= bottom; r++) for (let c = col; c <= right; c++) left.delete(`${String(r)},${String(c)}`);
    out.push(rangeText({ top: row, left: col, bottom, right }));
  }
  return out;
}

function formulaIn(f: unknown): string | number | Date {
  if (typeof f === 'number' || typeof f === 'string' || f instanceof Date) return f;
  return String(f);
}

/** The core rule of an exceljs data validation (X4-16). */
function validationIn(v: DataValidation): DataValidationRule {
  const type = v.type === 'any' ? 'none' : v.type;
  const rule: Record<string, unknown> = { type };
  // exceljs writes no operator for list and custom (data-validations-xform.js L125), and the
  // core refuses one there ([MS-XLSX] §2.6.5); the default operator is between.
  const takesOperator = type !== 'none' && type !== 'list' && type !== 'custom';
  const operator = takesOperator ? (v.operator ?? 'between') : undefined;
  if (takesOperator && v.operator !== undefined) rule['operator'] = v.operator;
  const formulae = Array.isArray(v.formulae) ? v.formulae : [];
  if (type !== 'none' && formulae[0] !== undefined) rule['formula1'] = formulaIn(formulae[0]);
  if ((operator === 'between' || operator === 'notBetween') && formulae[1] !== undefined) rule['formula2'] = formulaIn(formulae[1]);
  for (const k of ['allowBlank', 'showInputMessage', 'showErrorMessage'] as const) if (v[k]) rule[k] = true;
  for (const k of ['errorStyle', 'errorTitle', 'error', 'promptTitle', 'prompt'] as const) if (v[k]) rule[k] = v[k];
  return rule as unknown as DataValidationRule;
}

/** The validations of a worksheet as core `validate(sqref, rule)` calls. */
export function dataValidationsIn(dv: DataValidations): { sqref: string; rule: DataValidationRule }[] {
  const groups = new Map<string, { rule: DataValidation; cells: { row: number; col: number }[]; ranges: string[] }>();
  for (const [key, v] of Object.entries(dv.model)) {
    if (v === undefined) continue;
    const id = stableKey(v);
    let g = groups.get(id);
    if (!g) {
      g = { rule: v, cells: [], ranges: [] };
      groups.set(id, g);
    }
    const single = !key.includes(':') && !/\s/.test(key) ? parseRange(key) : null;
    if (single) g.cells.push({ row: single.top, col: single.left });
    else g.ranges.push(...key.split(/\s+/).filter((s) => s !== ''));
  }
  return [...groups.values()].map((g) => ({ sqref: sqrefIn([...rectangles(g.cells), ...g.ranges].join(' ')), rule: validationIn(g.rule) }));
}

/** exceljs utils.excelToDate (lib/utils/utils.js L58–61). */
function excelToDate(v: number, date1904: boolean): Date {
  return new Date(Math.round((v - 25569 + (date1904 ? 1462 : 0)) * 24 * 3600 * 1000));
}

/** The exceljs model of a validation read by xlvane, typed as exceljs reads it (data-validations-xform.js L165–264). */
function validationOut(d: XlvDataValidation, date1904: boolean): DataValidation {
  const type = d.type === 'none' ? 'any' : d.type;
  const out: Record<string, unknown> = { type };
  if (type !== 'any' && d.allowBlank) out['allowBlank'] = true;
  if (d.showInputMessage) out['showInputMessage'] = true;
  if (d.showErrorMessage) out['showErrorMessage'] = true;
  const formulae: unknown[] = [];
  for (const f of [d.formula1, d.formula2]) {
    if (f === null) continue;
    if (type === 'whole' || type === 'textLength') formulae.push(Number.parseInt(f, 10));
    else if (type === 'decimal') formulae.push(Number.parseFloat(f));
    else if (type === 'date') formulae.push(excelToDate(Number.parseFloat(f), date1904));
    else formulae.push(f);
  }
  if (formulae.length > 0) {
    if (type !== 'any' && type !== 'list' && type !== 'custom') out['operator'] = d.operator ?? 'between';
    out['formulae'] = formulae;
  }
  if (d.errorStyle !== 'stop') out['errorStyle'] = d.errorStyle;
  for (const k of ['promptTitle', 'prompt', 'errorTitle', 'error'] as const) if (d[k] !== null) out[k] = d[k];
  return out as unknown as DataValidation;
}

/** The `worksheet.dataValidations` model of validations read by xlvane. */
export function dataValidationsOut(list: readonly XlvDataValidation[], date1904: boolean): Record<string, DataValidation> {
  const model: Record<string, DataValidation> = {};
  let budget = MAX_EXPANDED_CELLS;
  for (const d of list) {
    const v = validationOut(d, date1904);
    for (const part of d.sqref.split(/\s+/)) {
      const r = parseRange(part);
      if (!r) continue;
      const size = (r.bottom - r.top + 1) * (r.right - r.left + 1);
      if (size > budget) {
        model[part] = v;
        continue;
      }
      budget -= size;
      for (let row = r.top; row <= r.bottom; row++) for (let col = r.left; col <= r.right; col++) model[address(row, col)] = v;
    }
  }
  return model;
}

// ---------------------------------------------------------------------------------------------
// Conditional formatting (X4-17; D-C02)

const TEXT_OPERATORS = new Set(['containsText', 'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors']);
const EXT_ICON_SETS = new Set(['3Triangles', '3Stars', '5Boxes']);

/** The top-left cell of the first range of `ref` (exceljs Range.tl, used in cf-rule-xform.js L23, L46). */
function topLeft(ref: string): string {
  const r = parseRange(ref.split(/\s+/)[0] ?? '');
  return r ? address(r.top, r.left) : 'A1';
}

/** A text for a formula string literal: quotes doubled. exceljs inserts the text as it is (X4-17). */
function quoted(text: string): string {
  return `"${text.replace(/"/g, '""')}"`;
}

/** The formula exceljs generates for a text rule (cf-rule-xform.js L18–39). */
function textFormula(operator: string, text: string, tl: string): string | undefined {
  switch (operator) {
    case 'containsText':
      return `NOT(ISERROR(SEARCH(${quoted(text)},${tl})))`;
    case 'containsBlanks':
      return `LEN(TRIM(${tl}))=0`;
    case 'notContainsBlanks':
      return `LEN(TRIM(${tl}))>0`;
    case 'containsErrors':
      return `ISERROR(${tl})`;
    case 'notContainsErrors':
      return `NOT(ISERROR(${tl}))`;
    default:
      return undefined;
  }
}

/** The formula exceljs generates for a timePeriod rule (cf-rule-xform.js L41–72). */
function timePeriodFormula(period: string, tl: string): string | undefined {
  switch (period) {
    case 'thisWeek':
      return `AND(TODAY()-ROUNDDOWN(${tl},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${tl},0)-TODAY()<=7-WEEKDAY(TODAY()))`;
    case 'lastWeek':
      return `AND(TODAY()-ROUNDDOWN(${tl},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${tl},0)<(WEEKDAY(TODAY())+7))`;
    case 'nextWeek':
      return `AND(ROUNDDOWN(${tl},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${tl},0)-TODAY()<(15-WEEKDAY(TODAY())))`;
    case 'yesterday':
      return `FLOOR(${tl},1)=TODAY()-1`;
    case 'today':
      return `FLOOR(${tl},1)=TODAY()`;
    case 'tomorrow':
      return `FLOOR(${tl},1)=TODAY()+1`;
    case 'last7Days':
      return `AND(TODAY()-FLOOR(${tl},1)<=6,FLOOR(${tl},1)<=TODAY())`;
    case 'lastMonth':
      return `AND(MONTH(${tl})=MONTH(EDATE(TODAY(),0-1)),YEAR(${tl})=YEAR(EDATE(TODAY(),0-1)))`;
    case 'thisMonth':
      return `AND(MONTH(${tl})=MONTH(TODAY()),YEAR(${tl})=YEAR(TODAY()))`;
    case 'nextMonth':
      return `AND(MONTH(${tl})=MONTH(EDATE(TODAY(),0+1)),YEAR(${tl})=YEAR(EDATE(TODAY(),0+1)))`;
    default:
      return undefined;
  }
}

function cfvoIn(c: { type: string; value?: unknown }): CfValueInput {
  const out: Record<string, unknown> = { type: c.type };
  if (c.value !== undefined && c.value !== null) out['value'] = typeof c.value === 'number' ? c.value : String(c.value);
  return out as unknown as CfValueInput;
}

/** A rule style: font, fill, border and number format (the parts a differential format has). */
function dxfStyleIn(s: Style | undefined): ConditionalFormatRuleInput['style'] {
  if (!s) return undefined;
  const full = styleIn({ font: s.font, fill: s.fill, border: s.border, numFmt: s.numFmt });
  return Object.keys(full).length > 0 ? full : undefined;
}

function formulasIn(f: unknown): (string | number)[] | undefined {
  if (!Array.isArray(f) || f.length === 0) return undefined;
  return f.map((x) => (typeof x === 'number' ? x : String(x)));
}

// TODO(phase C2, x14): data bars without `gradient: true`, custom icon sets and the icon sets
// 3Triangles, 3Stars and 5Boxes need the x14 extensions of [MS-XLSX]; until then they throw.

/** The core rule of an exceljs rule of the range `ref` (X4-17). */
function cfRuleIn(rule: ConditionalFormattingRule, ref: string): ConditionalFormatRuleInput {
  const r = rule as unknown as Record<string, unknown>;
  const out: Record<string, unknown> = {};
  if (typeof r['priority'] === 'number') out['priority'] = r['priority'];
  const style = dxfStyleIn(r['style'] as Style | undefined);
  const formulas = formulasIn(r['formulae']);
  const tl = topLeft(ref);
  switch (rule.type) {
    case 'expression':
      out['type'] = 'expression';
      if (formulas) out['formulas'] = formulas;
      break;
    case 'cellIs':
      out['type'] = 'cellIs';
      out['operator'] = r['operator'];
      if (formulas) out['formulas'] = formulas;
      break;
    case 'top10':
      out['type'] = 'top10';
      // cf-rule-xform.js L180: rank defaults to 10.
      out['rank'] = typeof r['rank'] === 'number' ? r['rank'] : 10;
      if (r['percent']) out['percent'] = true;
      if (r['bottom']) out['bottom'] = true;
      break;
    case 'aboveAverage':
      out['type'] = 'aboveAverage';
      if (r['aboveAverage'] === false) out['aboveAverage'] = false;
      break;
    case 'colorScale': {
      const cfvo = Array.isArray(r['cfvo']) ? (r['cfvo'] as { type: string; value?: unknown }[]) : [];
      const colors = Array.isArray(r['color']) ? (r['color'] as unknown[]) : [];
      out['type'] = 'colorScale';
      out['colorScale'] = { cfvo: cfvo.map(cfvoIn), colors: colors.map((c, i) => colorInFor(c, `colorScale color ${String(i + 1)}`)) };
      return out as unknown as ConditionalFormatRuleInput;
    }
    case 'dataBar': {
      if (r['gradient'] !== true) notSupported('Data bars without gradient (the x14 data bar of conditional formatting)');
      const cfvo = Array.isArray(r['cfvo']) ? (r['cfvo'] as { type: string; value?: unknown }[]) : [];
      out['type'] = 'dataBar';
      out['dataBar'] = { cfvo: cfvo.map(cfvoIn), color: colorInFor(r['color'], 'the data bar colour') };
      return out as unknown as ConditionalFormatRuleInput;
    }
    case 'iconSet': {
      const name = typeof r['iconSet'] === 'string' ? r['iconSet'] : '3TrafficLights';
      if (r['custom'] || EXT_ICON_SETS.has(name)) notSupported(`The icon set ${r['custom'] ? 'with custom icons' : name} (an x14 icon set of conditional formatting)`);
      const cfvo = Array.isArray(r['cfvo']) ? (r['cfvo'] as { type: string; value?: unknown }[]) : [];
      // exceljs calls the default set "3TrafficLights" (icon-set-xform.js L20); its name in
      // ST_IconSetType is 3TrafficLights1.
      const iconSet: Record<string, unknown> = { iconSet: name === '3TrafficLights' ? '3TrafficLights1' : name, cfvo: cfvo.map(cfvoIn) };
      if (r['reverse']) iconSet['reverse'] = true;
      if (r['showValue'] === false) iconSet['showValue'] = false;
      out['type'] = 'iconSet';
      out['iconSet'] = iconSet;
      return out as unknown as ConditionalFormatRuleInput;
    }
    case 'containsText': {
      // exceljs writes the operator as the rule type (cf-rule-xform.js L238); without an operator
      // it writes a rule with no type, so xlvane uses containsText.
      const operator = typeof r['operator'] === 'string' ? r['operator'] : 'containsText';
      if (!TEXT_OPERATORS.has(operator)) notSupported(`The containsText operator "${operator}"`);
      out['type'] = operator;
      const text = typeof r['text'] === 'string' ? r['text'] : '';
      if (operator === 'containsText') {
        out['operator'] = 'containsText';
        out['text'] = text;
      }
      const f = formulas ?? [textFormula(operator, text, tl)!];
      out['formulas'] = f;
      break;
    }
    case 'timePeriod': {
      out['type'] = 'timePeriod';
      out['timePeriod'] = r['timePeriod'];
      const f = formulas ?? (typeof r['timePeriod'] === 'string' ? [timePeriodFormula(r['timePeriod'], tl)].filter((x): x is string => x !== undefined) : []);
      if (f.length > 0) out['formulas'] = f;
      break;
    }
    default:
      return notSupported(`Conditional formatting rules of type "${String(rule.type)}"`);
  }
  if (style) out['style'] = style;
  return out as unknown as ConditionalFormatRuleInput;
}

/** The core `conditionalFormat(sqref, rules)` calls of a worksheet. */
export function conditionalFormatsIn(list: readonly ConditionalFormattingOptions[]): { sqref: string; rules: ConditionalFormatRuleInput[] }[] {
  return list.map((cf) => ({ sqref: sqrefIn(cf.ref), rules: cf.rules.map((rule) => cfRuleIn(rule, cf.ref)) }));
}

function cfvoOut(c: { type: string; value: string | null }): { type: string; value?: number | string } {
  if (c.value === null) return { type: c.type };
  const n = Number(c.value);
  return { type: c.type, value: c.value.trim() !== '' && Number.isFinite(n) ? n : c.value };
}

/** An exceljs rule from a rule read by xlvane, shaped as exceljs reads one (cf-rule-xform.js L269–280; X4-17). */
function cfRuleOut(r: XlvConditionalFormatRule): ConditionalFormattingRule {
  const out: Record<string, unknown> = { type: r.type, priority: r.priority };
  if (r.style) out['style'] = styleOut(r.style, typeof r.style.numFmt === 'string' ? r.style.numFmt : null);
  if (r.formulas.length > 0) out['formulae'] = [...r.formulas];
  switch (r.type) {
    case 'cellIs':
      out['operator'] = r.operator;
      break;
    case 'top10':
      out['rank'] = r.rank ?? 10;
      out['percent'] = r.percent;
      out['bottom'] = r.bottom;
      break;
    case 'aboveAverage':
      out['aboveAverage'] = r.aboveAverage;
      break;
    case 'containsText':
    case 'containsBlanks':
    case 'notContainsBlanks':
    case 'containsErrors':
    case 'notContainsErrors':
      out['type'] = 'containsText';
      out['operator'] = r.type;
      if (r.text !== null) out['text'] = r.text;
      break;
    case 'timePeriod':
      out['timePeriod'] = r.timePeriod;
      break;
    case 'colorScale':
      out['cfvo'] = r.colorScale!.cfvo.map(cfvoOut);
      out['color'] = r.colorScale!.colors.map((c) => colorOut(c));
      break;
    case 'dataBar':
      out['cfvo'] = r.dataBar!.cfvo.map(cfvoOut);
      out['color'] = colorOut(r.dataBar!.color);
      out['gradient'] = true;
      break;
    case 'iconSet':
      out['iconSet'] = r.iconSet!.iconSet === '3TrafficLights1' ? '3TrafficLights' : r.iconSet!.iconSet;
      out['reverse'] = r.iconSet!.reverse;
      out['showValue'] = r.iconSet!.showValue;
      out['cfvo'] = r.iconSet!.cfvo.map(cfvoOut);
      break;
    default:
      break;
  }
  return out as unknown as ConditionalFormattingRule;
}

/** `worksheet.conditionalFormattings` from the conditional formats read by xlvane. */
export function conditionalFormatsOut(list: readonly XlvConditionalFormat[]): ConditionalFormattingOptions[] {
  return list.map((cf) => ({ ref: cf.sqref, rules: cf.rules.map(cfRuleOut) }));
}

// ---------------------------------------------------------------------------------------------
// Sheet protection (X4-18; D-C03)

const PROTECTION_FLAGS = [
  'objects',
  'scenarios',
  'selectLockedCells',
  'selectUnlockedCells',
  'formatCells',
  'formatColumns',
  'formatRows',
  'insertColumns',
  'insertRows',
  'insertHyperlinks',
  'deleteColumns',
  'deleteRows',
  'sort',
  'autoFilter',
  'pivotTables',
] as const;

/**
 * Core protection flags from exceljs options. exceljs options say what users may still do
 * (`formatCells: true` = allowed), the file's attributes say what is locked (P1 §18.3.1.85,
 * SR-20), and exceljs writes each option as the negated attribute (sheet-protection-xform.js
 * L17–48; X4-18), so every given option becomes its negation.
 */
export function protectionIn(model: Record<string, unknown>): SheetProtectionFlags {
  const out: Record<string, boolean> = {};
  for (const k of PROTECTION_FLAGS) if (typeof model[k] === 'boolean') out[k] = !model[k];
  return out as SheetProtectionFlags;
}

const LOCKED_BY_DEFAULT = new Set(['formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows', 'insertHyperlinks', 'deleteColumns', 'deleteRows', 'sort', 'autoFilter', 'pivotTables']);

/** `worksheet.sheetProtection` from the protection read by xlvane, as exceljs reads it (sheet-protection-xform.js L50–82). */
export function protectionOut(p: XlvSheetProtection): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  if (p.sheet) out['sheet'] = true;
  for (const k of PROTECTION_FLAGS) {
    const locked = p[k];
    if (LOCKED_BY_DEFAULT.has(k)) {
      if (!locked) out[k] = true;
    } else if (locked) out[k] = false;
  }
  if (p.algorithmName !== null) {
    out['algorithmName'] = p.algorithmName;
    out['hashValue'] = p.hashValue;
    out['saltValue'] = p.saltValue;
    out['spinCount'] = p.spinCount;
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// Page setup, header and footer, page breaks (X4-19; D-C04)

/** The default `worksheet.pageSetup` of exceljs (lib/doc/worksheet.js L64–90). */
export function defaultPageSetup(options: Partial<PageSetup> | undefined): PageSetup {
  const o = options ?? {};
  return {
    margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    orientation: 'portrait',
    horizontalDpi: 4294967295,
    verticalDpi: 4294967295,
    fitToPage: !!((o.fitToWidth || o.fitToHeight) && !o.scale),
    pageOrder: 'downThenOver',
    blackAndWhite: false,
    draft: false,
    cellComments: 'None',
    errors: 'displayed',
    scale: 100,
    fitToWidth: 1,
    fitToHeight: 1,
    paperSize: undefined,
    showRowColHeaders: false,
    showGridLines: false,
    firstPageNumber: undefined,
    horizontalCentered: false,
    verticalCentered: false,
    rowBreaks: null,
    colBreaks: null,
    ...o,
  };
}

/** The default `worksheet.headerFooter` of exceljs (lib/doc/worksheet.js L94–106). */
export function defaultHeaderFooter(options: Partial<HeaderFooter> | undefined): HeaderFooter {
  return {
    differentFirst: false,
    differentOddEven: false,
    oddHeader: null,
    oddFooter: null,
    evenHeader: null,
    evenFooter: null,
    firstHeader: null,
    firstFooter: null,
    ...options,
  };
}

const CELL_COMMENTS: Record<string, PageSetupOptions['cellComments']> = {
  None: 'none',
  none: 'none',
  atEnd: 'atEnd',
  // exceljs spells the value "asDisplyed" (page-setup-xform.js L18) and drops "asDisplayed".
  asDisplyed: 'asDisplayed',
  asDisplayed: 'asDisplayed',
};

/** The core options of an exceljs `pageSetup`, with the print area and titles. */
export function pageSetupIn(ps: PageSetup): { options: PageSetupOptions; printArea: string[] | null; printTitles: PrintTitles | null } {
  const o: Record<string, unknown> = {};
  if (ps.margins) o['margins'] = ps.margins;
  if (ps.orientation === 'portrait' || ps.orientation === 'landscape') o['orientation'] = ps.orientation;
  if (typeof ps.paperSize === 'number') o['paperSize'] = ps.paperSize;
  for (const k of ['scale', 'fitToWidth', 'fitToHeight', 'copies', 'horizontalDpi', 'verticalDpi'] as const) if (typeof ps[k] === 'number') o[k] = ps[k];
  if (typeof ps.firstPageNumber === 'number') o['firstPageNumber'] = ps.firstPageNumber;
  if (ps.fitToPage) o['fitToPage'] = true;
  if (ps.pageOrder === 'overThenDown' || ps.pageOrder === 'downThenOver') o['pageOrder'] = ps.pageOrder;
  if (ps.blackAndWhite) o['blackAndWhite'] = true;
  if (ps.draft) o['draft'] = true;
  const comments = typeof ps.cellComments === 'string' ? CELL_COMMENTS[ps.cellComments] : undefined;
  if (comments) o['cellComments'] = comments;
  if (ps.errors === 'displayed' || ps.errors === 'blank' || ps.errors === 'dash' || ps.errors === 'NA') o['errors'] = ps.errors;
  const printOptions: Record<string, boolean> = {};
  if (ps.showGridLines) printOptions['gridLines'] = true;
  if (ps.showRowColHeaders) printOptions['headings'] = true;
  if (ps.horizontalCentered) printOptions['horizontalCentered'] = true;
  if (ps.verticalCentered) printOptions['verticalCentered'] = true;
  if (Object.keys(printOptions).length > 0) o['printOptions'] = printOptions;
  // Several print areas are joined with "&&" in exceljs (workbook-xform.js L45).
  const printArea = typeof ps.printArea === 'string' && ps.printArea !== '' ? ps.printArea.split('&&').filter((a) => a !== '') : null;
  // exceljs writes either setting as a range of the sheet and reads rows and columns back by their
  // form (workbook-xform.js L58–78, L211–226), so a column range given as printTitlesRow ends up
  // as column titles; xlvane sorts them by form as well.
  const titles: Record<string, string> = {};
  for (const v of [ps.printTitlesRow, ps.printTitlesColumn]) {
    if (typeof v !== 'string' || v === '') continue;
    if (/^\$?\d+(:\$?\d+)?$/.test(v)) titles['rows'] = v.replace(/\$/g, '');
    else if (/^\$?[A-Za-z]{1,3}(:\$?[A-Za-z]{1,3})?$/.test(v)) titles['columns'] = v.replace(/\$/g, '');
    else notSupported(`The print titles "${v.slice(0, 40)}" (use rows such as "1:2" or columns such as "A:B")`);
  }
  return { options: o as PageSetupOptions, printArea, printTitles: Object.keys(titles).length > 0 ? titles : null };
}

/** The core header and footer of an exceljs `headerFooter`, or `null` when nothing is set. */
export function headerFooterIn(hf: HeaderFooter | null | undefined): XlvHeaderFooter | null {
  if (!hf) return null;
  const out: Record<string, unknown> = {};
  for (const k of ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter'] as const) {
    const v = hf[k];
    if (typeof v === 'string' && v !== '') out[k] = v;
  }
  if (Object.keys(out).length === 0) return null;
  if (hf.differentFirst) out['differentFirst'] = true;
  if (hf.differentOddEven) out['differentOddEven'] = true;
  return out as XlvHeaderFooter;
}

/** `worksheet.pageSetup` from the page setup read by xlvane, as exceljs reads it (worksheet-xform.js L390–399; page-setup-xform.js L72–89). */
export function pageSetupOut(ps: XlvPageSetup, printArea: readonly string[] | null, printTitles: { rows: string | null; columns: string | null } | null): PageSetup {
  const out: PageSetup = {
    fitToPage: ps.fitToPage,
    margins: ps.margins ? { ...ps.margins } : undefined,
    paperSize: ps.paperSize === 1 ? undefined : ps.paperSize,
    orientation: ps.orientation === 'landscape' ? 'landscape' : 'portrait',
    horizontalDpi: ps.horizontalDpi ?? 4294967295,
    verticalDpi: ps.verticalDpi ?? 4294967295,
    pageOrder: ps.pageOrder,
    blackAndWhite: ps.blackAndWhite,
    draft: ps.draft,
    cellComments: ps.cellComments === 'none' ? 'None' : ps.cellComments,
    errors: ps.errors,
    scale: ps.scale,
    fitToWidth: ps.fitToWidth,
    fitToHeight: ps.fitToHeight,
    firstPageNumber: ps.firstPageNumber ?? 1,
    useFirstPageNumber: ps.firstPageNumber !== null,
    usePrinterDefaults: false,
    copies: ps.copies,
    showRowColHeaders: ps.printOptions.headings,
    showGridLines: ps.printOptions.gridLines,
    horizontalCentered: ps.printOptions.horizontalCentered,
    verticalCentered: ps.printOptions.verticalCentered,
  };
  // workbook-xform.js L187–226: print areas joined by "&&", titles without "$".
  if (printArea && printArea.length > 0) out.printArea = printArea.map((a) => a.replace(/\$/g, '')).join('&&');
  if (printTitles?.rows) out.printTitlesRow = printTitles.rows.replace(/\$/g, '');
  if (printTitles?.columns) out.printTitlesColumn = printTitles.columns.replace(/\$/g, '');
  return out;
}

/** `worksheet.headerFooter` from the header and footer read by xlvane (`null` without one, as exceljs). */
export function headerFooterOut(hf: Required<XlvHeaderFooter> | null): HeaderFooter | null {
  if (!hf) return null;
  const out: HeaderFooter = {};
  if (hf.differentFirst) out.differentFirst = true;
  if (hf.differentOddEven) out.differentOddEven = true;
  for (const k of ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter'] as const) if (hf[k] !== '') out[k] = hf[k];
  return out;
}

/**
 * Row breaks as exceljs `row.addPageBreak` stores them (lib/doc/row.js L150–162): `id` is the
 * row after which the page ends, so the core break is at `id + 1`. exceljs writes `max` 16838,
 * one more than the last column index allows; the core writes the break itself.
 */
export function rowBreaksIn(breaks: readonly RowBreak[] | undefined): number[] {
  return (breaks ?? []).map((b) => b.id + 1).filter((n) => Number.isInteger(n) && n >= 2 && n <= MAX_ROW);
}

/** `worksheet.rowBreaks` from the breaks read by xlvane; exceljs drops them on read (X4-19). */
export function rowBreaksOut(rows: readonly number[]): RowBreak[] {
  return rows.map((r) => ({ id: r - 1, max: MAX_COL - 1, man: 1 }));
}
