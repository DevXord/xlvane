// Data validations and conditional formatting of a worksheet (ECMA-376 Part 1 §18.3.1.10,
// §18.3.1.11, §18.3.1.15, §18.3.1.16, §18.3.1.18, §18.3.1.28, §18.3.1.32, §18.3.1.33,
// §18.3.1.43–§18.3.1.45, §18.3.1.49, §18.18; [MS-OI29500] §2.1.595 … §2.1.625, §2.1.1085,
// §2.1.1092; docs/facts/sheet-rules.md SR-01 … SR-16; docs/decisions.md D-C01, D-C02).
// This module validates what callers give the writer and reads the elements back; the XML is
// written by sml/write-parts.ts. Enumerations are those of the Transitional sml.xsd, which are
// identical in Strict (SR-01, SR-07). Formulas are formula text without a leading "=": xlvane
// stores them as given and never evaluates them.
import { XlvError } from './errors';
import { dateToSerial } from './sml/dates';
import { formatRange, parseRange } from './sml/cell-ref';
import { formatNumber } from './sml/write-parts';
import { decodeXstring } from './sml/xstring';
import { normalizeColor, normalizeStyle, type CellStyle, type Color, type NormalizedStyle } from './style';
import type { DateSystem } from './model';
import { suggest } from './util/suggest';
import { firstLoneSurrogate } from './xml/writer';
import { childrenNamed, type XmlElement } from './xml/tree';

// ---------------------------------------------------------------------------------------------
// Enumerations (sml.xsd, SR-01 and SR-07)

/**
 * ST_DataValidationType (§18.18.21): what a data validation checks. `none` validates nothing but can show an input message.
 *
 * @example
 * ```ts
 * import type { DataValidationType } from 'xlvane';
 *
 * const type: DataValidationType = 'list';
 * ```
 */
export type DataValidationType = 'none' | 'whole' | 'decimal' | 'list' | 'date' | 'time' | 'textLength' | 'custom';
/**
 * ST_DataValidationOperator (§18.18.20): the comparison of a `whole`, `decimal`, `date`, `time` or `textLength` validation.
 *
 * @example
 * ```ts
 * import type { DataValidationOperator } from 'xlvane';
 *
 * const op: DataValidationOperator = 'greaterThanOrEqual';
 * ```
 */
export type DataValidationOperator = 'between' | 'notBetween' | 'equal' | 'notEqual' | 'lessThan' | 'lessThanOrEqual' | 'greaterThan' | 'greaterThanOrEqual';
/**
 * ST_DataValidationErrorStyle (§18.18.18): the icon of the error alert.
 *
 * @example
 * ```ts
 * import type { DataValidationErrorStyle } from 'xlvane';
 *
 * const style: DataValidationErrorStyle = 'warning';
 * ```
 */
export type DataValidationErrorStyle = 'stop' | 'warning' | 'information';
/**
 * ST_DataValidationImeMode (§18.18.19): how a data validation sets the input method editor.
 *
 * @example
 * ```ts
 * import type { DataValidationImeMode } from 'xlvane';
 *
 * const mode: DataValidationImeMode = 'noControl';
 * ```
 */
export type DataValidationImeMode =
  | 'noControl'
  | 'off'
  | 'on'
  | 'disabled'
  | 'hiragana'
  | 'fullKatakana'
  | 'halfKatakana'
  | 'fullAlpha'
  | 'halfAlpha'
  | 'fullHangul'
  | 'halfHangul';
/**
 * ST_CfType (§18.18.12): the type of a conditional formatting rule.
 *
 * @example
 * ```ts
 * import type { ConditionalFormatType } from 'xlvane';
 *
 * const type: ConditionalFormatType = 'cellIs';
 * ```
 */
export type ConditionalFormatType =
  | 'expression'
  | 'cellIs'
  | 'colorScale'
  | 'dataBar'
  | 'iconSet'
  | 'top10'
  | 'uniqueValues'
  | 'duplicateValues'
  | 'containsText'
  | 'notContainsText'
  | 'beginsWith'
  | 'endsWith'
  | 'containsBlanks'
  | 'notContainsBlanks'
  | 'containsErrors'
  | 'notContainsErrors'
  | 'timePeriod'
  | 'aboveAverage';
/**
 * ST_ConditionalFormattingOperator (§18.18.15): the comparison of a `cellIs` rule, or the text test of the text rules.
 *
 * @example
 * ```ts
 * import type { ConditionalFormatOperator } from 'xlvane';
 *
 * const op: ConditionalFormatOperator = 'between';
 * ```
 */
export type ConditionalFormatOperator =
  | 'lessThan'
  | 'lessThanOrEqual'
  | 'equal'
  | 'notEqual'
  | 'greaterThanOrEqual'
  | 'greaterThan'
  | 'between'
  | 'notBetween'
  | 'containsText'
  | 'notContains'
  | 'beginsWith'
  | 'endsWith';
/**
 * ST_TimePeriod (§18.18.82): the period of a `timePeriod` rule.
 *
 * @example
 * ```ts
 * import type { TimePeriod } from 'xlvane';
 *
 * const period: TimePeriod = 'last7Days';
 * ```
 */
export type TimePeriod = 'today' | 'yesterday' | 'tomorrow' | 'last7Days' | 'thisMonth' | 'lastMonth' | 'nextMonth' | 'thisWeek' | 'lastWeek' | 'nextWeek';
/**
 * ST_CfvoType (§18.18.13): the kind of a threshold of a colour scale, data bar or icon set.
 *
 * @example
 * ```ts
 * import type { CfValueType } from 'xlvane';
 *
 * const type: CfValueType = 'percentile';
 * ```
 */
export type CfValueType = 'num' | 'percent' | 'max' | 'min' | 'formula' | 'percentile';
/**
 * ST_IconSetType (§18.18.42): the icons of an icon set rule.
 *
 * @example
 * ```ts
 * import type { IconSetType } from 'xlvane';
 *
 * const icons: IconSetType = '3Arrows';
 * ```
 */
export type IconSetType =
  | '3Arrows'
  | '3ArrowsGray'
  | '3Flags'
  | '3TrafficLights1'
  | '3TrafficLights2'
  | '3Signs'
  | '3Symbols'
  | '3Symbols2'
  | '4Arrows'
  | '4ArrowsGray'
  | '4RedToBlack'
  | '4Rating'
  | '4TrafficLights'
  | '5Arrows'
  | '5ArrowsGray'
  | '5Rating'
  | '5Quarters';

const DV_TYPES: readonly DataValidationType[] = ['none', 'whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom'];
const DV_OPERATORS: readonly DataValidationOperator[] = ['between', 'notBetween', 'equal', 'notEqual', 'lessThan', 'lessThanOrEqual', 'greaterThan', 'greaterThanOrEqual'];
const DV_ERROR_STYLES: readonly DataValidationErrorStyle[] = ['stop', 'warning', 'information'];
const DV_IME_MODES: readonly DataValidationImeMode[] = ['noControl', 'off', 'on', 'disabled', 'hiragana', 'fullKatakana', 'halfKatakana', 'fullAlpha', 'halfAlpha', 'fullHangul', 'halfHangul'];
const CF_TYPES: readonly ConditionalFormatType[] = [
  'expression',
  'cellIs',
  'colorScale',
  'dataBar',
  'iconSet',
  'top10',
  'uniqueValues',
  'duplicateValues',
  'containsText',
  'notContainsText',
  'beginsWith',
  'endsWith',
  'containsBlanks',
  'notContainsBlanks',
  'containsErrors',
  'notContainsErrors',
  'timePeriod',
  'aboveAverage',
];
const CF_OPERATORS: readonly ConditionalFormatOperator[] = [
  'lessThan',
  'lessThanOrEqual',
  'equal',
  'notEqual',
  'greaterThanOrEqual',
  'greaterThan',
  'between',
  'notBetween',
  'containsText',
  'notContains',
  'beginsWith',
  'endsWith',
];
const TEXT_OPERATORS: readonly ConditionalFormatOperator[] = ['containsText', 'notContains', 'beginsWith', 'endsWith'];
const TIME_PERIODS: readonly TimePeriod[] = ['today', 'yesterday', 'tomorrow', 'last7Days', 'thisMonth', 'lastMonth', 'nextMonth', 'thisWeek', 'lastWeek', 'nextWeek'];
const CFVO_TYPES: readonly CfValueType[] = ['num', 'percent', 'max', 'min', 'formula', 'percentile'];
const ICON_SETS: readonly IconSetType[] = [
  '3Arrows',
  '3ArrowsGray',
  '3Flags',
  '3TrafficLights1',
  '3TrafficLights2',
  '3Signs',
  '3Symbols',
  '3Symbols2',
  '4Arrows',
  '4ArrowsGray',
  '4RedToBlack',
  '4Rating',
  '4TrafficLights',
  '5Arrows',
  '5ArrowsGray',
  '5Rating',
  '5Quarters',
];

/** Types whose `operator` gives the comparison and whose `formula2` is the upper bound (SR-03). */
const DV_COMPARING: readonly DataValidationType[] = ['whole', 'decimal', 'date', 'time', 'textLength'];
/** Rule types whose criteria are given entirely by formulas (SR-11; generated formulas: C2). */
const CF_NEEDS_FORMULA: readonly ConditionalFormatType[] = [
  'expression',
  'cellIs',
  'containsText',
  'notContainsText',
  'beginsWith',
  'endsWith',
  'containsBlanks',
  'notContainsBlanks',
  'containsErrors',
  'notContainsErrors',
  'timePeriod',
];
const CF_TEXT_TYPES: readonly ConditionalFormatType[] = ['containsText', 'notContainsText', 'beginsWith', 'endsWith'];
const CF_GRAPHIC_TYPES: readonly ConditionalFormatType[] = ['colorScale', 'dataBar', 'iconSet'];

// ---------------------------------------------------------------------------------------------
// Limits ([MS-OI29500], SR-05, SR-06, SR-14)

/** `prompt` of a data validation: at most 255 characters ([MS-OI29500] §2.1.613 a). */
const MAX_PROMPT = 255;
/** `error`: at most 224 characters (§2.1.613 c). */
const MAX_ERROR = 224;
/** `errorTitle` and `promptTitle`: at most 32 characters (§2.1.613 d, e). */
const MAX_TITLE = 32;
/** References in `dataValidation@sqref`: 1 to 32767 (§2.1.613 b). */
const MAX_DV_REFS = 32_767;
/** References in any other ST_Sqref, such as `conditionalFormatting@sqref`: 8192 (§2.1.1092 a). */
const MAX_SQREF_REFS = 8192;
/** `dataValidation` elements of a sheet: at most 65534 (§2.1.614 g). */
export const MAX_DATA_VALIDATIONS = 65_534;
/** Length of an ST_Formula: at most 8192 characters (§2.1.1085 a). */
const MAX_FORMULA = 8192;
/** `dataBar@minLength` and `@maxLength`: at most 100 (§2.1.611). */
const MAX_BAR_LENGTH = 100;
/** cfvo and color children of a colour scale: 2 (schema minimum) to 3 (§2.1.601 a, b). */
const MAX_SCALE_POINTS = 3;
/** `cfRule@priority` is an xsd:int; [MS-XLSX] §2.6.27 requires it to be greater than 0. */
const MAX_PRIORITY = 2_147_483_647;

// ---------------------------------------------------------------------------------------------
// Public shapes

/**
 * A data validation rule for {@link NewSheet.validate}: what may be entered in a range, and the
 * messages shown. Formulas are formula text (a leading "=" is removed), numbers, or `Date`s
 * (stored as serial numbers). A list of fixed values is a formula holding one quoted string,
 * such as `'"Yes,No"'`; a list from cells is a reference such as `'Lists!$A$1:$A$5'`.
 *
 * Rules (SR-03): `formula1` is required except for `type: 'none'`; `whole`, `decimal`, `date`,
 * `time` and `textLength` compare with `operator` (default `between`), and `between` /
 * `notBetween` need `formula2` as the upper bound; `list`, `custom` and `none` take no operator.
 * The messages and the error alert are shown only when `showInputMessage` / `showErrorMessage`
 * are `true` (schema defaults `false`). Limits ([MS-OI29500] §2.1.613): `prompt` 255 characters,
 * `error` 224, `promptTitle` and `errorTitle` 32; formulas 8192.
 *
 * @example
 * ```ts
 * import { createWorkbook, type DataValidationRule } from 'xlvane';
 *
 * const yesNo: DataValidationRule = { type: 'list', formula1: '"Yes,No"', showErrorMessage: true, error: 'Choose Yes or No.' };
 * const amount: DataValidationRule = { type: 'decimal', operator: 'between', formula1: 0, formula2: 1000, showInputMessage: true, prompt: '0 to 1000' };
 * createWorkbook().addSheet('Form').validate('B2:B50', yesNo).validate('C2:C50', amount);
 * ```
 */
export interface DataValidationRule {
  readonly type: DataValidationType;
  readonly operator?: DataValidationOperator;
  readonly formula1?: string | number | Date;
  readonly formula2?: string | number | Date;
  readonly allowBlank?: boolean;
  /** Stored as the `showDropDown` attribute; the meaning of its value is not settled (SR-04). */
  readonly showDropDown?: boolean;
  readonly showInputMessage?: boolean;
  readonly showErrorMessage?: boolean;
  readonly errorStyle?: DataValidationErrorStyle;
  readonly imeMode?: DataValidationImeMode;
  readonly errorTitle?: string;
  readonly error?: string;
  readonly promptTitle?: string;
  readonly prompt?: string;
}

/**
 * A data validation of a worksheet, as read or as written: the range (`sqref`, space-separated
 * ranges) and the rule with every attribute resolved to its schema default. `operator` is `null`
 * for types that take none.
 *
 * @example
 * ```ts
 * import { readWorkbook, type DataValidation } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const rules: readonly DataValidation[] = readWorkbook(bytes).getSheet('Form').dataValidations;
 * for (const r of rules) console.log(r.sqref, r.type, r.formula1);
 * ```
 */
export interface DataValidation {
  readonly sqref: string;
  readonly type: DataValidationType;
  readonly operator: DataValidationOperator | null;
  readonly formula1: string | null;
  readonly formula2: string | null;
  readonly allowBlank: boolean;
  readonly showDropDown: boolean;
  readonly showInputMessage: boolean;
  readonly showErrorMessage: boolean;
  readonly errorStyle: DataValidationErrorStyle;
  readonly imeMode: DataValidationImeMode;
  readonly errorTitle: string | null;
  readonly error: string | null;
  readonly promptTitle: string | null;
  readonly prompt: string | null;
}

/**
 * A value object (`cfvo`, §18.3.1.11) of a colour scale, data bar or icon set: a threshold given
 * as the minimum or maximum of the range, a number, a percent, a percentile or a formula.
 *
 * @example
 * ```ts
 * import type { CfValueInput } from 'xlvane';
 *
 * const low: CfValueInput = { type: 'min' };
 * const half: CfValueInput = { type: 'percentile', value: 50 };
 * ```
 */
export interface CfValueInput {
  readonly type: CfValueType;
  readonly value?: string | number;
  /** Icon sets: `false` compares with "greater than" instead of "greater than or equal". */
  readonly gte?: boolean;
}

/**
 * A value object as read: `value` is the text of `val` (`null` when absent), `gte` defaults to
 * `true`.
 *
 * @example
 * ```ts
 * import type { CfValue } from 'xlvane';
 *
 * const v: CfValue = { type: 'percent', value: '33', gte: true };
 * console.log(v.value);
 * ```
 */
export interface CfValue {
  readonly type: CfValueType;
  readonly value: string | null;
  readonly gte: boolean;
}

/**
 * A conditional formatting rule for {@link NewSheet.conditionalFormat}. `style` (font, fill,
 * border and number format) is the format applied where the rule is true; it is stored as a
 * differential format (`dxf`). Rules of the types `expression`, `cellIs`, the text, blank, error
 * and `timePeriod` types need their criteria as `formulas` (1 to 3, each up to 8192 characters);
 * relative references in them are taken from the top-left cell of the range in spreadsheet
 * applications (SR-11, not stated by the standard). `cellIs` needs an `operator`, `timePeriod`
 * a `timePeriod`; `colorScale`, `dataBar` and `iconSet` need the object of the same name and take
 * no style. Priorities are unique on a sheet, 1 being applied first; rules without one get the
 * next free number.
 *
 * @example
 * ```ts
 * import { createWorkbook, type ConditionalFormatRuleInput } from 'xlvane';
 *
 * const high: ConditionalFormatRuleInput = { type: 'cellIs', operator: 'greaterThan', formulas: [1000], style: { fill: { pattern: 'solid', fgColor: 'FFC7CE' } } };
 * const scale: ConditionalFormatRuleInput = { type: 'colorScale', colorScale: { cfvo: [{ type: 'min' }, { type: 'max' }], colors: ['F8696B', '63BE7B'] } };
 * createWorkbook().addSheet('Sales').conditionalFormat('B2:B100', [high, scale]);
 * ```
 */
export interface ConditionalFormatRuleInput {
  readonly type: ConditionalFormatType;
  readonly priority?: number;
  readonly stopIfTrue?: boolean;
  readonly style?: CellStyle;
  readonly formulas?: readonly (string | number)[];
  readonly operator?: ConditionalFormatOperator;
  readonly text?: string;
  readonly timePeriod?: TimePeriod;
  /** top10: the n of "top n". */
  readonly rank?: number;
  readonly percent?: boolean;
  readonly bottom?: boolean;
  /** aboveAverage: `false` for "below average" (schema default `true`). */
  readonly aboveAverage?: boolean;
  readonly equalAverage?: boolean;
  readonly stdDev?: number;
  readonly colorScale?: { readonly cfvo: readonly CfValueInput[]; readonly colors: readonly Color[] };
  readonly dataBar?: { readonly cfvo: readonly [CfValueInput, CfValueInput]; readonly color: Color; readonly minLength?: number; readonly maxLength?: number; readonly showValue?: boolean };
  readonly iconSet?: { readonly iconSet?: IconSetType; readonly cfvo: readonly CfValueInput[]; readonly reverse?: boolean; readonly showValue?: boolean; readonly percent?: boolean };
}

/**
 * A conditional formatting rule as read (or as written): every attribute resolved to its schema
 * default, `style` the differential format (`null` without one), the colour scale, data bar or
 * icon set `null` for other types.
 *
 * @example
 * ```ts
 * import { readWorkbook, type ConditionalFormatRule } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const rules: ConditionalFormatRule[] = readWorkbook(bytes).getSheet('Sales').conditionalFormats.flatMap((c) => [...c.rules]);
 * for (const r of rules) console.log(r.priority, r.type, r.formulas);
 * ```
 */
export interface ConditionalFormatRule {
  readonly type: ConditionalFormatType;
  readonly priority: number;
  readonly stopIfTrue: boolean;
  readonly style: CellStyle | null;
  readonly formulas: readonly string[];
  readonly operator: ConditionalFormatOperator | null;
  readonly text: string | null;
  readonly timePeriod: TimePeriod | null;
  readonly rank: number | null;
  readonly percent: boolean;
  readonly bottom: boolean;
  readonly aboveAverage: boolean;
  readonly equalAverage: boolean;
  readonly stdDev: number | null;
  readonly colorScale: { readonly cfvo: readonly CfValue[]; readonly colors: readonly Color[] } | null;
  readonly dataBar: { readonly cfvo: readonly CfValue[]; readonly color: Color; readonly minLength: number; readonly maxLength: number; readonly showValue: boolean } | null;
  readonly iconSet: { readonly iconSet: IconSetType; readonly cfvo: readonly CfValue[]; readonly reverse: boolean; readonly showValue: boolean; readonly percent: boolean } | null;
}

/**
 * The conditional formatting of a range (`conditionalFormatting`, §18.3.1.18): its range
 * (space-separated ranges) and its rules in file order.
 *
 * @example
 * ```ts
 * import { readWorkbook, type ConditionalFormat } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const formats: readonly ConditionalFormat[] = readWorkbook(bytes).getSheet('Sales').conditionalFormats;
 * console.log(formats.map((f) => `${f.sqref}: ${f.rules.length} rules`));
 * ```
 */
export interface ConditionalFormat {
  readonly sqref: string;
  readonly rules: readonly ConditionalFormatRule[];
}

/** A rule of the writer: the normalised rule and its style as a differential format. */
export interface StoredCfRule extends Omit<ConditionalFormatRule, 'style'> {
  readonly style: NormalizedStyle | null;
}

/** A conditional format of the writer, with the dxf index of each rule (`null` without style). */
export interface StoredConditionalFormat {
  readonly sqref: string;
  readonly rules: readonly StoredCfRule[];
  readonly dxfIds: readonly (number | null)[];
}

// ---------------------------------------------------------------------------------------------
// Validation of caller input

interface Where {
  readonly sheet: string;
  readonly dateSystem: DateSystem;
}

function fail(where: Where, option: string, message: string, hint: string): XlvError {
  return new XlvError('XLV_E_INVALID_OPTION', message, { location: { sheet: where.sheet }, hint, details: { option } });
}

function readKeys(value: unknown, known: readonly string[], what: string, where: Where): Map<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value) || value instanceof Date) {
    throw fail(where, what, `${what} must be an object.`, `Pass an object with some of these properties: ${known.join(', ')}.`);
  }
  const out = new Map<string, unknown>();
  for (const key of Object.keys(value)) {
    if (!known.includes(key)) {
      const close = suggest(key, known);
      throw fail(where, `${what}.${key.slice(0, 40)}`, `Unknown property "${key.slice(0, 40)}" in ${what}.`, `${close.length > 0 ? `Did you mean "${close[0] ?? ''}"? ` : ''}Known properties: ${known.join(', ')}.`);
    }
    const v = (value as Record<string, unknown>)[key];
    if (v !== undefined) out.set(key, v);
  }
  return out;
}

function enumValue<T extends string>(value: unknown, allowed: readonly T[], option: string, where: Where): T {
  if (typeof value === 'string' && (allowed as readonly string[]).includes(value)) return value as T;
  const shown = typeof value === 'string' ? JSON.stringify(value.slice(0, 40)) : String(value);
  const close = typeof value === 'string' ? suggest(value, allowed) : [];
  throw fail(where, option, `The ${option} ${shown} is not valid.`, `${close.length > 0 ? `Did you mean "${close[0] ?? ''}"? ` : ''}Allowed values: ${allowed.join(', ')}.`);
}

function boolValue(value: unknown, option: string, where: Where): boolean {
  if (typeof value === 'boolean') return value;
  throw fail(where, option, `The ${option} must be true or false.`, `Use ${option}: true.`);
}

function intValue(value: unknown, option: string, min: number, max: number, where: Where): number {
  if (typeof value === 'number' && Number.isInteger(value) && value >= min && value <= max) return value;
  throw fail(where, option, `The ${option} ${String(value).slice(0, 40)} is not valid.`, `Use a whole number from ${String(min)} to ${String(max)}.`);
}

function textValue(value: unknown, option: string, max: number, where: Where): string {
  if (typeof value !== 'string') throw fail(where, option, `The ${option} must be a string.`, `For example ${option}: "text".`);
  if (value.length > max) throw fail(where, option, `The ${option} has ${String(value.length)} characters; at most ${String(max)} are allowed.`, `Shorten the ${option} to ${String(max)} characters or fewer.`);
  if (firstLoneSurrogate(value) !== -1) throw fail(where, option, `The ${option} contains an unpaired surrogate.`, 'Remove the broken character.');
  return value;
}

/** Formula text from a caller value: text without a leading "=", a number, or a date (serial). */
function formulaValue(value: unknown, option: string, where: Where): string {
  let text: string;
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw fail(where, option, `The ${option} must be a finite number.`, `Pass a number, or formula text such as "A1*2".`);
    text = formatNumber(value);
  } else if (value instanceof Date) {
    const time = value.getTime();
    if (Number.isNaN(time)) throw fail(where, option, `The ${option} is an invalid Date.`, 'Pass a valid Date.');
    try {
      text = formatNumber(dateToSerial(value, where.dateSystem));
    } catch {
      throw fail(where, option, `The date ${value.toISOString()} in ${option} is outside the ${where.dateSystem} date system.`, 'Use a date from 1900 (or 1904 with that date system) to 9999.');
    }
  } else if (typeof value === 'string') {
    text = value.startsWith('=') ? value.slice(1) : value;
    if (text.trim() === '') throw fail(where, option, `The ${option} is empty.`, 'Pass formula text such as "A1>0", a number, or "\\"Yes,No\\"" for a list.');
  } else {
    throw fail(where, option, `The ${option} must be formula text, a number or a Date.`, 'For example "A1>0", 10 or new Date(Date.UTC(2024, 0, 31)).');
  }
  if (text.length > MAX_FORMULA) throw fail(where, option, `The ${option} has ${String(text.length)} characters; at most ${String(MAX_FORMULA)} are allowed.`, `Shorten the formula to ${String(MAX_FORMULA)} characters or fewer.`);
  if (firstLoneSurrogate(text) !== -1) throw fail(where, option, `The ${option} contains an unpaired surrogate.`, 'Remove the broken character.');
  return text;
}

/**
 * Normalises a range list (ST_Sqref, §18.18.76: space-separated references) to ranges in
 * canonical form, at most `max` of them. Throws `XLV_E_INVALID_CELL_REF` for a bad range.
 */
export function normalizeSqref(sqref: unknown, max: number, where: Where): string {
  const parts = typeof sqref === 'string' ? sqref.trim().split(/\s+/) : [];
  if (parts.length === 0 || parts[0] === '') {
    throw new XlvError('XLV_E_INVALID_CELL_REF', `"${String(sqref).slice(0, 40)}" is not a range.`, {
      location: { sheet: where.sheet },
      hint: 'Pass a cell or range such as "B2:B100", or several separated by spaces: "B2:B10 D2:D10".',
      details: { ref: String(sqref).slice(0, 40) },
    });
  }
  if (parts.length > max) throw fail(where, 'sqref', `The range list has ${String(parts.length)} ranges; at most ${String(max)} are allowed.`, 'Combine neighbouring cells into larger ranges, such as "A1:A100".');
  return parts
    .map((p) => {
      const range = parseRange(p.toUpperCase().replace(/\$/g, ''));
      if (range === null) {
        throw new XlvError('XLV_E_INVALID_CELL_REF', `"${p.slice(0, 40)}" is not a valid range.`, {
          location: { sheet: where.sheet },
          hint: 'Use A1-style ranges such as "B2:B100" (columns A to XFD, rows 1 to 1048576), separated by spaces.',
          details: { ref: p.slice(0, 40) },
        });
      }
      return formatRange(range);
    })
    .join(' ');
}

const DV_KEYS = ['type', 'operator', 'formula1', 'formula2', 'allowBlank', 'showDropDown', 'showInputMessage', 'showErrorMessage', 'errorStyle', 'imeMode', 'errorTitle', 'error', 'promptTitle', 'prompt'];

/** Validates a data validation rule (SR-02 … SR-06); throws `XLV_E_INVALID_OPTION`. */
export function normalizeDataValidation(sqref: unknown, rule: unknown, where: Where): DataValidation {
  const range = normalizeSqref(sqref, MAX_DV_REFS, where);
  const o = readKeys(rule, DV_KEYS, 'data validation', where);
  if (!o.has('type')) throw fail(where, 'type', 'The data validation has no type.', `Pass type: one of ${DV_TYPES.join(', ')}.`);
  const type = enumValue(o.get('type'), DV_TYPES, 'type', where);
  const comparing = DV_COMPARING.includes(type);
  let operator: DataValidationOperator | null = null;
  if (o.has('operator')) {
    if (!comparing) throw fail(where, 'operator', `A data validation of type "${type}" takes no operator.`, `Remove the operator; it applies to ${DV_COMPARING.join(', ')}.`);
    operator = enumValue(o.get('operator'), DV_OPERATORS, 'operator', where);
  } else if (comparing) operator = 'between';
  const f1 = o.get('formula1');
  const f2 = o.get('formula2');
  if (type === 'none') {
    if (f1 !== undefined || f2 !== undefined) throw fail(where, f1 !== undefined ? 'formula1' : 'formula2', 'A data validation of type "none" takes no formula1 or formula2.', 'Remove the formulas, or choose a type such as "list" or "whole".');
  } else if (f1 === undefined) {
    throw fail(where, 'formula1', `A data validation of type "${type}" needs formula1.`, type === 'list' ? 'For a fixed list pass formula1: "\\"Yes,No\\"", for a list from cells formula1: "Lists!$A$1:$A$5".' : 'Pass the value or formula to compare with, for example formula1: 10.');
  }
  const ranged = operator === 'between' || operator === 'notBetween';
  if (ranged && f2 === undefined) throw fail(where, 'formula2', `The operator "${String(operator)}" needs formula2, the upper bound.`, 'Pass formula2, or choose another operator such as "greaterThan".');
  if (!ranged && f2 !== undefined) throw fail(where, 'formula2', `formula2 is used only with the operators "between" and "notBetween".`, 'Remove formula2, or use operator: "between".');
  const flag = (k: string): boolean => (o.has(k) ? boolValue(o.get(k), k, where) : false);
  const text = (k: string, max: number): string | null => (o.has(k) ? textValue(o.get(k), k, max, where) : null);
  return Object.freeze({
    sqref: range,
    type,
    operator,
    formula1: f1 === undefined ? null : formulaValue(f1, 'formula1', where),
    formula2: f2 === undefined ? null : formulaValue(f2, 'formula2', where),
    allowBlank: flag('allowBlank'),
    showDropDown: flag('showDropDown'),
    showInputMessage: flag('showInputMessage'),
    showErrorMessage: flag('showErrorMessage'),
    errorStyle: o.has('errorStyle') ? enumValue(o.get('errorStyle'), DV_ERROR_STYLES, 'errorStyle', where) : 'stop',
    imeMode: o.has('imeMode') ? enumValue(o.get('imeMode'), DV_IME_MODES, 'imeMode', where) : 'noControl',
    errorTitle: text('errorTitle', MAX_TITLE),
    error: text('error', MAX_ERROR),
    promptTitle: text('promptTitle', MAX_TITLE),
    prompt: text('prompt', MAX_PROMPT),
  });
}

const CF_KEYS = [
  'type',
  'priority',
  'stopIfTrue',
  'style',
  'formulas',
  'operator',
  'text',
  'timePeriod',
  'rank',
  'percent',
  'bottom',
  'aboveAverage',
  'equalAverage',
  'stdDev',
  'colorScale',
  'dataBar',
  'iconSet',
];

function cfValues(value: unknown, option: string, min: number, max: number, where: Where, inIconSet: boolean): CfValue[] {
  if (!Array.isArray(value) || value.length < min || value.length > max) {
    throw fail(where, option, `${option} must be an array of ${min === max ? String(min) : `${String(min)} to ${String(max)}`} value objects.`, 'For example [{ type: "min" }, { type: "max" }].');
  }
  return value.map((v: unknown, i) => {
    const o = readKeys(v, ['type', 'value', 'gte'], `${option}[${String(i)}]`, where);
    const type = enumValue(o.get('type'), CFVO_TYPES, `${option}[${String(i)}].type`, where);
    // In Office the thresholds of an icon set cannot be min or max ([MS-OI29500] §2.1.596 b).
    if (inIconSet && (type === 'min' || type === 'max')) {
      throw fail(where, `${option}[${String(i)}].type`, `An icon set threshold cannot be of type "${type}".`, 'Use { type: "percent", value: 0 } for the lowest threshold.');
    }
    const raw = o.get('value');
    if ((type === 'num' || type === 'percent' || type === 'percentile' || type === 'formula') && raw === undefined) {
      throw fail(where, `${option}[${String(i)}].value`, `A value object of type "${type}" needs a value.`, `For example { type: "${type}", value: ${type === 'formula' ? '"$A$1"' : '50'} }.`);
    }
    const out: CfValue = Object.freeze({
      type,
      value: raw === undefined ? null : formulaValue(raw, `${option}[${String(i)}].value`, where),
      gte: o.has('gte') ? boolValue(o.get('gte'), `${option}[${String(i)}].gte`, where) : true,
    });
    if (!inIconSet && !out.gte) throw fail(where, `${option}[${String(i)}].gte`, 'gte applies to icon sets only.', 'Remove gte.');
    return out;
  });
}

/** Validates one conditional formatting rule (SR-08 … SR-15); `priority` is assigned by the caller. */
export function normalizeCfRule(rule: unknown, index: number, where: Where): Omit<StoredCfRule, 'priority'> & { readonly priority: number | null } {
  const what = `rules[${String(index)}]`;
  const o = readKeys(rule, CF_KEYS, what, where);
  const type = enumValue(o.get('type'), CF_TYPES, `${what}.type`, where);
  const only = (key: string, types: readonly ConditionalFormatType[]): boolean => {
    if (!o.has(key)) return false;
    if (!types.includes(type)) throw fail(where, `${what}.${key}`, `The rule property "${key}" does not apply to rules of type "${type}".`, `Remove "${key}"; it applies to ${types.join(', ')}.`);
    return true;
  };
  const graphic = CF_GRAPHIC_TYPES.includes(type);
  let style: NormalizedStyle | null = null;
  if (o.has('style')) {
    if (graphic) throw fail(where, `${what}.style`, `A rule of type "${type}" takes no style.`, 'Remove the style; colour scales, data bars and icon sets draw their own colours.');
    const raw = o.get('style');
    if (typeof raw === 'object' && raw !== null) {
      for (const k of Object.keys(raw)) {
        if (k !== 'font' && k !== 'fill' && k !== 'border' && k !== 'numFmt') {
          throw fail(where, `${what}.style.${k.slice(0, 40)}`, `The style of a conditional format cannot set "${k.slice(0, 40)}".`, 'Use font, fill, border and numFmt.');
        }
      }
    }
    style = normalizeStyle(raw);
  }
  let formulas: string[] = [];
  if (o.has('formulas')) {
    if (graphic) throw fail(where, `${what}.formulas`, `A rule of type "${type}" takes no formulas.`, 'Give the thresholds in the cfvo value objects instead.');
    const raw = o.get('formulas');
    if (!Array.isArray(raw) || raw.length < 1 || raw.length > 3) throw fail(where, `${what}.formulas`, 'formulas must be an array of 1 to 3 formulas.', 'For example formulas: ["$B2>100"].');
    formulas = raw.map((f: unknown, i) => formulaValue(f, `${what}.formulas[${String(i)}]`, where));
  } else if (CF_NEEDS_FORMULA.includes(type)) {
    throw fail(
      where,
      `${what}.formulas`,
      `A rule of type "${type}" needs its criteria in formulas.`,
      type === 'cellIs' ? 'For example formulas: [100] with operator: "greaterThan".' : 'Pass the condition as a formula for the top-left cell of the range, for example formulas: ["ISNUMBER(SEARCH(\\"x\\",A1))"].',
    );
  }
  let operator: ConditionalFormatOperator | null = null;
  if (type === 'cellIs') {
    // Office requires operator exactly for cellIs rules ([MS-OI29500] §2.1.595 b).
    if (!o.has('operator')) throw fail(where, `${what}.operator`, 'A rule of type "cellIs" needs an operator.', `Use one of ${CF_OPERATORS.slice(0, 8).join(', ')}.`);
    operator = enumValue(o.get('operator'), CF_OPERATORS.slice(0, 8), `${what}.operator`, where);
  } else if (only('operator', CF_TEXT_TYPES)) operator = enumValue(o.get('operator'), TEXT_OPERATORS, `${what}.operator`, where);
  let timePeriod: TimePeriod | null = null;
  if (type === 'timePeriod') {
    // timePeriod is required exactly for timePeriod rules (§2.1.595 c).
    if (!o.has('timePeriod')) throw fail(where, `${what}.timePeriod`, 'A rule of type "timePeriod" needs timePeriod.', `Use one of ${TIME_PERIODS.join(', ')}.`);
    timePeriod = enumValue(o.get('timePeriod'), TIME_PERIODS, `${what}.timePeriod`, where);
  } else only('timePeriod', ['timePeriod']);
  const text = only('text', CF_TEXT_TYPES) ? textValue(o.get('text'), `${what}.text`, MAX_FORMULA, where) : null;
  const rank = only('rank', ['top10']) ? intValue(o.get('rank'), `${what}.rank`, 0, 4_294_967_295, where) : null;
  const flag = (k: string, types: readonly ConditionalFormatType[], dflt: boolean): boolean => (only(k, types) ? boolValue(o.get(k), `${what}.${k}`, where) : dflt);
  const percent = flag('percent', ['top10'], false);
  const bottom = flag('bottom', ['top10'], false);
  const aboveAverage = flag('aboveAverage', ['aboveAverage'], true);
  const equalAverage = flag('equalAverage', ['aboveAverage'], false);
  const stdDev = only('stdDev', ['aboveAverage']) ? intValue(o.get('stdDev'), `${what}.stdDev`, -2_147_483_648, 2_147_483_647, where) : null;
  for (const key of CF_GRAPHIC_TYPES) {
    if (type !== key) only(key, [key]);
    else if (!o.has(key)) throw fail(where, `${what}.${key}`, `A rule of type "${type}" needs ${key}.`, `Pass ${key}: { cfvo: [...], ... } (see ConditionalFormatRuleInput).`);
  }
  let colorScale: StoredCfRule['colorScale'] = null;
  if (type === 'colorScale') {
    const c = readKeys(o.get('colorScale'), ['cfvo', 'colors'], `${what}.colorScale`, where);
    const cfvo = cfValues(c.get('cfvo'), `${what}.colorScale.cfvo`, 2, MAX_SCALE_POINTS, where, false);
    const colors = c.get('colors');
    // Office: 2 or 3 of each, and as many colours as value objects ([MS-OI29500] §2.1.601).
    if (!Array.isArray(colors) || colors.length !== cfvo.length) {
      throw fail(where, `${what}.colorScale.colors`, `colorScale.colors must have one colour per value object (${String(cfvo.length)}).`, 'For example colors: ["F8696B", "63BE7B"] for two value objects.');
    }
    colorScale = Object.freeze({ cfvo, colors: colors.map((x: unknown, i) => normalizeColor(x, `${what}.colorScale.colors[${String(i)}]`)) });
  }
  let dataBar: StoredCfRule['dataBar'] = null;
  if (type === 'dataBar') {
    const d = readKeys(o.get('dataBar'), ['cfvo', 'color', 'minLength', 'maxLength', 'showValue'], `${what}.dataBar`, where);
    const cfvo = cfValues(d.get('cfvo'), `${what}.dataBar.cfvo`, 2, 2, where, false);
    if (!d.has('color')) throw fail(where, `${what}.dataBar.color`, 'A data bar needs a color.', 'For example color: "638EC6".');
    const len = (k: string, dflt: number): number => (d.has(k) ? intValue(d.get(k), `${what}.dataBar.${k}`, 0, MAX_BAR_LENGTH, where) : dflt);
    dataBar = Object.freeze({
      cfvo,
      color: normalizeColor(d.get('color'), `${what}.dataBar.color`),
      minLength: len('minLength', 10),
      maxLength: len('maxLength', 90),
      showValue: d.has('showValue') ? boolValue(d.get('showValue'), `${what}.dataBar.showValue`, where) : true,
    });
  }
  let iconSet: StoredCfRule['iconSet'] = null;
  if (type === 'iconSet') {
    const s = readKeys(o.get('iconSet'), ['iconSet', 'cfvo', 'reverse', 'showValue', 'percent'], `${what}.iconSet`, where);
    const b = (k: string, dflt: boolean): boolean => (s.has(k) ? boolValue(s.get(k), `${what}.iconSet.${k}`, where) : dflt);
    iconSet = Object.freeze({
      iconSet: s.has('iconSet') ? enumValue(s.get('iconSet'), ICON_SETS, `${what}.iconSet.iconSet`, where) : '3TrafficLights1',
      cfvo: cfValues(s.get('cfvo'), `${what}.iconSet.cfvo`, 2, 5, where, true),
      reverse: b('reverse', false),
      showValue: b('showValue', true),
      percent: b('percent', true),
    });
  }
  return {
    type,
    priority: o.has('priority') ? intValue(o.get('priority'), `${what}.priority`, 1, MAX_PRIORITY, where) : null,
    stopIfTrue: o.has('stopIfTrue') ? boolValue(o.get('stopIfTrue'), `${what}.stopIfTrue`, where) : false,
    style,
    formulas,
    operator,
    text,
    timePeriod,
    rank,
    percent,
    bottom,
    aboveAverage,
    equalAverage,
    stdDev,
    colorScale,
    dataBar,
    iconSet,
  };
}

/**
 * Validates the rules of one `conditionalFormat` call and gives each a priority: the given one,
 * which must not be used on the sheet yet, or the next number after the highest one used.
 */
export function normalizeConditionalFormat(sqref: unknown, rules: unknown, where: Where, used: Set<number>): { readonly sqref: string; readonly rules: StoredCfRule[] } {
  const range = normalizeSqref(sqref, MAX_SQREF_REFS, where);
  if (!Array.isArray(rules) || rules.length === 0) throw fail(where, 'rules', 'conditionalFormat needs an array with at least one rule.', 'For example [{ type: "cellIs", operator: "greaterThan", formulas: [100], style: { font: { bold: true } } }].');
  const normalized = rules.map((r: unknown, i) => normalizeCfRule(r, i, where));
  const taken = new Set(used);
  for (const [i, r] of normalized.entries()) {
    if (r.priority === null) continue;
    if (taken.has(r.priority)) throw fail(where, `rules[${String(i)}].priority`, `The priority ${String(r.priority)} is already used on the sheet.`, 'Leave priority out to get the next free one, or choose a number not used yet.');
    taken.add(r.priority);
  }
  let next = Math.max(0, ...taken) + 1;
  const out = normalized.map((r) => {
    if (r.priority !== null) return Object.freeze({ ...r, priority: r.priority });
    if (next > MAX_PRIORITY) throw fail(where, 'priority', 'No priority is left on the sheet.', 'Give the rules priorities explicitly.');
    taken.add(next);
    return Object.freeze({ ...r, priority: next++ });
  });
  for (const p of taken) used.add(p);
  return { sqref: range, rules: out };
}

// ---------------------------------------------------------------------------------------------
// Reading

type Warn = (message: string) => void;

const parseBool = (raw: string | undefined, dflt: boolean): boolean => {
  if (raw === undefined) return dflt;
  const s = raw.trim();
  if (s === 'true' || s === '1') return true;
  if (s === 'false' || s === '0') return false;
  return dflt;
};

const parseInteger = (raw: string | undefined, min: number, max: number): number | null => {
  if (raw === undefined || !/^\s*[-+]?\d{1,10}\s*$/.test(raw)) return null;
  const n = Number(raw.trim());
  return n >= min && n <= max ? n : null;
};

/** A space-separated range list as read, canonical; `null` (with a warning) when invalid or too long. */
function readSqref(raw: string | undefined, max: number, what: string, warn: Warn): string | null {
  const parts = (raw ?? '').trim().split(/\s+/).filter((p) => p !== '');
  if (parts.length === 0) {
    warn(`A ${what} has no range (sqref); it is ignored.`);
    return null;
  }
  if (parts.length > max) {
    warn(`A ${what} lists ${String(parts.length)} ranges, more than the ${String(max)} Office allows; it is ignored.`);
    return null;
  }
  const out: string[] = [];
  for (const p of parts) {
    const r = parseRange(p.replace(/\$/g, ''));
    if (r === null) {
      warn(`A ${what} has the invalid range "${p.slice(0, 40)}"; it is ignored.`);
      return null;
    }
    out.push(formatRange(r));
  }
  return out.join(' ');
}

/** Reads `dataValidations` (§18.3.1.33); invalid entries are skipped with a warning. */
export function parseDataValidations(el: XmlElement, ns: string, warn: Warn): DataValidation[] {
  const out: DataValidation[] = [];
  for (const dv of childrenNamed(el, ns, 'dataValidation')) {
    if (out.length >= MAX_DATA_VALIDATIONS) {
      warn(`The sheet has more than ${String(MAX_DATA_VALIDATIONS)} data validations; the rest are ignored.`);
      break;
    }
    const sqref = readSqref(dv.attrs.get('sqref'), MAX_DV_REFS, 'data validation', warn);
    if (sqref === null) continue;
    const rawType = dv.attrs.get('type') ?? 'none';
    const type = DV_TYPES.find((t) => t === rawType);
    if (type === undefined) {
      warn(`A data validation has the unknown type "${rawType.slice(0, 40)}"; it is ignored.`);
      continue;
    }
    const pick = <T extends string>(name: string, allowed: readonly T[], dflt: T): T => {
      const raw = dv.attrs.get(name);
      if (raw === undefined) return dflt;
      const found = allowed.find((a) => a === raw);
      if (found === undefined) warn(`A data validation has the unknown ${name} "${raw.slice(0, 40)}"; "${dflt}" is used.`);
      return found ?? dflt;
    };
    const text = (name: string): string | null => {
      const raw = dv.attrs.get(name);
      return raw === undefined ? null : decodeXstring(raw);
    };
    const formula = (name: string): string | null => {
      const f = childrenNamed(dv, ns, name)[0];
      return f === undefined ? null : decodeXstring(f.text);
    };
    // formula2 is used only by "between" and "notBetween" (§18.3.1.45, SR-03); LibreOffice writes
    // one for other rules too, which is ignored here.
    const operator = DV_COMPARING.includes(type) ? pick('operator', DV_OPERATORS, 'between') : null;
    out.push(
      Object.freeze({
        sqref,
        type,
        operator,
        formula1: type === 'none' ? null : formula('formula1'),
        formula2: operator === 'between' || operator === 'notBetween' ? formula('formula2') : null,
        allowBlank: parseBool(dv.attrs.get('allowBlank'), false),
        showDropDown: parseBool(dv.attrs.get('showDropDown'), false),
        showInputMessage: parseBool(dv.attrs.get('showInputMessage'), false),
        showErrorMessage: parseBool(dv.attrs.get('showErrorMessage'), false),
        errorStyle: pick('errorStyle', DV_ERROR_STYLES, 'stop'),
        imeMode: pick('imeMode', DV_IME_MODES, 'noControl'),
        errorTitle: text('errorTitle'),
        error: text('error'),
        promptTitle: text('promptTitle'),
        prompt: text('prompt'),
      }),
    );
  }
  return out;
}

function readCfValues(parent: XmlElement, ns: string, warn: Warn): CfValue[] | null {
  const out: CfValue[] = [];
  for (const v of childrenNamed(parent, ns, 'cfvo')) {
    const raw = v.attrs.get('type') ?? '';
    const type = CFVO_TYPES.find((t) => t === raw);
    if (type === undefined) {
      warn(`A conditional format value object has the unknown type "${raw.slice(0, 40)}"; the rule is ignored.`);
      return null;
    }
    const val = v.attrs.get('val');
    // The value of min and max is not used (SR-12); it is reported only for the other types.
    out.push(Object.freeze({ type, value: val === undefined || type === 'min' || type === 'max' ? null : decodeXstring(val), gte: parseBool(v.attrs.get('gte'), true) }));
  }
  return out;
}

function readColors(parent: XmlElement, ns: string, warn: Warn, readColor: (el: XmlElement) => Color | undefined): Color[] | null {
  const out: Color[] = [];
  for (const c of childrenNamed(parent, ns, 'color')) {
    const color = readColor(c);
    if (color === undefined) {
      warn('A conditional format colour is not valid; the rule is ignored.');
      return null;
    }
    out.push(color);
  }
  return out;
}

/**
 * Reads one `conditionalFormatting` element (§18.3.1.18) with its rules; `dxf(i)` resolves a
 * differential format. Rules with an unknown type or without a valid priority are skipped with
 * a warning; `null` when no rule is left or the range is invalid.
 */
export function parseConditionalFormatting(
  el: XmlElement,
  ns: string,
  dxf: (index: number) => CellStyle | undefined,
  readColor: (el: XmlElement) => Color | undefined,
  warn: Warn,
): ConditionalFormat | null {
  const sqref = readSqref(el.attrs.get('sqref'), MAX_SQREF_REFS, 'conditional format', warn);
  if (sqref === null) return null;
  const rules: ConditionalFormatRule[] = [];
  for (const r of childrenNamed(el, ns, 'cfRule')) {
    const rawType = r.attrs.get('type') ?? '';
    const type = CF_TYPES.find((t) => t === rawType);
    if (type === undefined) {
      warn(`A conditional formatting rule has the unknown type "${rawType.slice(0, 40)}"; it is ignored.`);
      continue;
    }
    const priority = parseInteger(r.attrs.get('priority'), -2_147_483_648, MAX_PRIORITY);
    if (priority === null) {
      warn('A conditional formatting rule has no valid priority; it is ignored.');
      continue;
    }
    let style: CellStyle | null = null;
    const rawDxf = r.attrs.get('dxfId');
    if (rawDxf !== undefined) {
      const id = parseInteger(rawDxf, 0, 4_294_967_294);
      const found = id === null ? undefined : dxf(id);
      if (found === undefined) warn(`A conditional formatting rule refers to the differential format "${rawDxf.slice(0, 20)}", which does not exist; the rule is read without a format.`);
      else style = found;
    }
    const enumAttr = <T extends string>(name: string, allowed: readonly T[]): T | null => {
      const raw = r.attrs.get(name);
      if (raw === undefined) return null;
      const found = allowed.find((a) => a === raw);
      if (found === undefined) warn(`A conditional formatting rule has the unknown ${name} "${raw.slice(0, 40)}"; it is ignored.`);
      return found ?? null;
    };
    let colorScale: ConditionalFormatRule['colorScale'] = null;
    let dataBar: ConditionalFormatRule['dataBar'] = null;
    let iconSet: ConditionalFormatRule['iconSet'] = null;
    const sub = (name: string): XmlElement | undefined => childrenNamed(r, ns, name)[0];
    const cs = type === 'colorScale' ? sub('colorScale') : undefined;
    const db = type === 'dataBar' ? sub('dataBar') : undefined;
    const is = type === 'iconSet' ? sub('iconSet') : undefined;
    if (cs !== undefined) {
      const cfvo = readCfValues(cs, ns, warn);
      const colors = readColors(cs, ns, warn, readColor);
      if (cfvo === null || colors === null) continue;
      colorScale = Object.freeze({ cfvo, colors });
    } else if (db !== undefined) {
      const cfvo = readCfValues(db, ns, warn);
      const colors = readColors(db, ns, warn, readColor);
      if (cfvo === null || colors === null) continue;
      const color = colors[0];
      if (color === undefined) {
        warn('A data bar has no colour; the rule is ignored.');
        continue;
      }
      dataBar = Object.freeze({
        cfvo,
        color,
        minLength: parseInteger(db.attrs.get('minLength'), 0, 4_294_967_295) ?? 10,
        maxLength: parseInteger(db.attrs.get('maxLength'), 0, 4_294_967_295) ?? 90,
        showValue: parseBool(db.attrs.get('showValue'), true),
      });
    } else if (is !== undefined) {
      const cfvo = readCfValues(is, ns, warn);
      if (cfvo === null) continue;
      const raw = is.attrs.get('iconSet');
      const set = raw === undefined ? '3TrafficLights1' : ICON_SETS.find((s) => s === raw);
      if (set === undefined) {
        warn(`A conditional formatting rule has the unknown icon set "${String(raw).slice(0, 40)}"; it is ignored.`);
        continue;
      }
      iconSet = Object.freeze({ iconSet: set, cfvo, reverse: parseBool(is.attrs.get('reverse'), false), showValue: parseBool(is.attrs.get('showValue'), true), percent: parseBool(is.attrs.get('percent'), true) });
    }
    // Attributes of other rule types are ignored (§18.3.1.10, SR-09): LibreOffice writes, for
    // example, rank="0" and text="" on every rule.
    const top = type === 'top10';
    const avg = type === 'aboveAverage';
    const textType = CF_TEXT_TYPES.includes(type);
    const stdDev = avg ? r.attrs.get('stdDev') : undefined;
    rules.push(
      Object.freeze({
        type,
        priority,
        stopIfTrue: parseBool(r.attrs.get('stopIfTrue'), false),
        style,
        formulas: Object.freeze(childrenNamed(r, ns, 'formula').slice(0, 3).map((f) => decodeXstring(f.text))),
        operator: type === 'cellIs' || textType ? enumAttr('operator', CF_OPERATORS) : null,
        text: textType ? (r.attrs.get('text') ?? null) : null,
        timePeriod: type === 'timePeriod' ? enumAttr('timePeriod', TIME_PERIODS) : null,
        rank: top ? parseInteger(r.attrs.get('rank'), 0, 4_294_967_295) : null,
        percent: top && parseBool(r.attrs.get('percent'), false),
        bottom: top && parseBool(r.attrs.get('bottom'), false),
        aboveAverage: !avg || parseBool(r.attrs.get('aboveAverage'), true),
        equalAverage: avg && parseBool(r.attrs.get('equalAverage'), false),
        stdDev: stdDev === undefined ? null : parseInteger(stdDev, -2_147_483_648, 2_147_483_647),
        colorScale,
        dataBar,
        iconSet,
      }),
    );
  }
  return rules.length === 0 ? null : Object.freeze({ sqref, rules: Object.freeze(rules) });
}

/**
 * Collects the `conditionalFormatting` and `dataValidations` elements of a worksheet while it is
 * streamed through the SAX parser (they follow `sheetData`, so this costs no memory per row).
 * Each element is kept as a small tree until its end tag; elements and attributes count against
 * `maxElements` for the whole sheet, so a hostile part cannot grow it without bound.
 */
export function createRuleCollector(ctx: {
  readonly ns: () => string;
  readonly maxElements: number;
  readonly dxf: (index: number) => CellStyle | undefined;
  readonly readColor: (el: XmlElement) => Color | undefined;
  readonly warn: Warn;
  readonly tooLarge: () => XlvError;
}): {
  start(local: string, depth: number, attrs: ReadonlyMap<string, string>): boolean;
  end(): boolean;
  text(value: string): boolean;
  result(): { conditionalFormats: ConditionalFormat[]; dataValidations: DataValidation[] };
} {
  const conditionalFormats: ConditionalFormat[] = [];
  const dataValidations: DataValidation[] = [];
  let stack: XmlElement[] | null = null;
  let count = 0;
  return {
    start(local, depth, attrs) {
      if (stack === null) {
        if (depth !== 2 || (local !== 'conditionalFormatting' && local !== 'dataValidations')) return false;
        stack = [];
      }
      count += 1 + attrs.size;
      if (count > ctx.maxElements) throw ctx.tooLarge();
      const el: XmlElement = { ns: ctx.ns(), local, attrs, children: [], text: '' };
      stack[stack.length - 1]?.children.push(el);
      stack.push(el);
      return true;
    },
    end() {
      if (stack === null) return false;
      const done = stack.pop();
      if (stack.length === 0 && done !== undefined) {
        stack = null;
        if (done.local === 'dataValidations') {
          const room = MAX_DATA_VALIDATIONS - dataValidations.length;
          dataValidations.push(...parseDataValidations(done, done.ns, ctx.warn).slice(0, Math.max(0, room)));
        } else {
          const cf = parseConditionalFormatting(done, done.ns, ctx.dxf, ctx.readColor, ctx.warn);
          if (cf !== null) conditionalFormats.push(cf);
        }
      }
      return true;
    },
    text(value) {
      if (stack === null) return false;
      const top = stack[stack.length - 1];
      if (top !== undefined) top.text += value;
      return true;
    },
    result() {
      return { conditionalFormats, dataValidations };
    },
  };
}
