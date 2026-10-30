// Core properties of a package: the Core Properties part, `docProps/core.xml` (ECMA-376 Part 2
// Clause 8, Annex E; facts F-OPC-16 … F-OPC-20 in docs/facts/opc.md, decision D-A02).
import { XlvError } from './errors';
import type { ReadLimits } from './limits';
import type { WarningCollector } from './util/warnings';
import { escapeText, firstLoneSurrogate } from './xml/writer';
import { parseXmlTree, type XmlElement } from './xml/tree';

/**
 * Document properties of a workbook (the package's core properties, ECMA-376 Part 2 Clause 8).
 * Every property is optional. Reading gives the properties found in `docProps/core.xml`;
 * writing stores the ones given. `company` and `manager` are not core properties (they belong
 * to `docProps/app.xml`) and are not supported yet.
 *
 * @example
 * ```ts
 * import { createWorkbook, type WorkbookProperties } from 'xlvane';
 *
 * const properties: WorkbookProperties = { title: 'Q3 sales', creator: 'Finance', created: new Date(Date.UTC(2026, 9, 1)) };
 * createWorkbook({ properties }).addSheet('Sales');
 * ```
 */
export interface WorkbookProperties {
  /** `dc:title`: the name of the document. */
  readonly title?: string;
  /** `dc:subject`: the topic of the content. */
  readonly subject?: string;
  /** `dc:creator`: who made the content. */
  readonly creator?: string;
  /** `cp:keywords`: keywords for searching; one string, as stored. */
  readonly keywords?: string;
  /** `dc:description`: an explanation of the content. */
  readonly description?: string;
  /** `cp:lastModifiedBy`: who made the last change. */
  readonly lastModifiedBy?: string;
  /** `cp:category`, for example "Report". */
  readonly category?: string;
  /** `cp:contentStatus`, for example "Draft" or "Final". */
  readonly contentStatus?: string;
  /** `dc:identifier`: an unambiguous reference to the document. */
  readonly identifier?: string;
  /** `dc:language`: the language of the content, for example "en-US". */
  readonly language?: string;
  /** `cp:revision`: the revision number, as text. */
  readonly revision?: string;
  /** `cp:version`: the version number, as text. */
  readonly version?: string;
  /** `dcterms:created`: when the document was created. */
  readonly created?: Date;
  /** `dcterms:modified`: when the document was last changed. */
  readonly modified?: Date;
  /** `cp:lastPrinted`: when the document was last printed. */
  readonly lastPrinted?: Date;
}

/** The core properties namespace (Part 2 Table E.1). */
export const CORE_PROPERTIES_NS = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
/** Namespace of the ISO 15836-1 elements (Part 2 §8.3.4.2). */
const DC_NS = 'http://purl.org/dc/elements/1.1/';
/** Namespace of the ISO 15836-2 elements (Part 2 §8.3.4.3). */
const DCTERMS_NS = 'http://purl.org/dc/terms/';
const XSI_NS = 'http://www.w3.org/2001/XMLSchema-instance';
/** The core properties relationship type (Part 2 Table E.3; [MS-OI29500] §2.1.30). */
export const CORE_PROPERTIES_REL = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
/** The type ECMA-376 Part 4 §13.2.10 lists for Transitional documents; accepted when reading. */
const CORE_PROPERTIES_REL_P4 = 'http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties';
/** Media type of the Core Properties part (Part 2 Table E.2). */
export const CORE_PROPERTIES_CONTENT_TYPE = 'application/vnd.openxmlformats-package.core-properties+xml';
/** Part name used by the writer (a convention; parts are found through relationships, E-08). */
export const CORE_PROPERTIES_PART = 'docProps/core.xml';

/** `true` for a relationship type that points to the Core Properties part. */
export function isCorePropertiesRel(type: string): boolean {
  return type === CORE_PROPERTIES_REL || type === CORE_PROPERTIES_REL_P4;
}

type TextKey = 'title' | 'subject' | 'creator' | 'keywords' | 'description' | 'lastModifiedBy' | 'category' | 'contentStatus' | 'identifier' | 'language' | 'revision' | 'version';
type DateKey = 'created' | 'modified' | 'lastPrinted';

/** Property name, element namespace prefix and local name, in the order of CT_CoreProperties. */
const ELEMENTS: readonly (readonly [TextKey | DateKey, 'cp' | 'dc' | 'dcterms'])[] = [
  ['category', 'cp'],
  ['contentStatus', 'cp'],
  ['created', 'dcterms'],
  ['creator', 'dc'],
  ['description', 'dc'],
  ['identifier', 'dc'],
  ['keywords', 'cp'],
  ['language', 'dc'],
  ['lastModifiedBy', 'cp'],
  ['lastPrinted', 'cp'],
  ['modified', 'dcterms'],
  ['revision', 'cp'],
  ['subject', 'dc'],
  ['title', 'dc'],
  ['version', 'cp'],
];
const DATE_KEYS: readonly string[] = ['created', 'modified', 'lastPrinted'];
const KEYS: readonly string[] = ELEMENTS.map(([key]) => key);
const PREFIX_NS = { cp: CORE_PROPERTIES_NS, dc: DC_NS, dcterms: DCTERMS_NS } as const;

function invalidProperty(key: string, message: string, hint: string): XlvError {
  return new XlvError('XLV_E_INVALID_OPTION', message, { hint, details: { option: `properties.${key}` } });
}

/**
 * Checks document properties given by a caller and returns a copy with only the given ones.
 * Text must be a string that XML can hold; dates must be valid `Date` objects in the years
 * 0–9999 (W3CDTF uses four-digit years).
 */
export function normalizeProperties(value: unknown): WorkbookProperties {
  if (value === undefined) return {};
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new XlvError('XLV_E_INVALID_OPTION', 'The option "properties" must be an object.', {
      hint: 'For example { properties: { title: "Q3 sales", creator: "Finance" } }.',
      details: { option: 'properties' },
    });
  }
  const out: Record<string, string | Date> = {};
  for (const key of Object.keys(value)) {
    if (!KEYS.includes(key)) {
      throw new XlvError('XLV_E_INVALID_OPTION', `Unknown document property "${key.slice(0, 40)}".`, {
        hint: `Known properties: ${KEYS.join(', ')}. company and manager are not supported yet.`,
        details: { option: `properties.${key.slice(0, 40)}`, known: KEYS },
      });
    }
    const v = (value as Record<string, unknown>)[key];
    if (v === undefined) continue;
    if (DATE_KEYS.includes(key)) {
      if (!(v instanceof Date) || Number.isNaN(v.getTime()) || v.getUTCFullYear() < 0 || v.getUTCFullYear() > 9999) {
        throw invalidProperty(key, `The document property "${key}" must be a valid Date between the years 0 and 9999.`, `For example ${key}: new Date().`);
      }
      out[key] = new Date(v.getTime());
    } else {
      if (typeof v !== 'string') throw invalidProperty(key, `The document property "${key}" must be a string.`, `For example ${key}: "text".`);
      if (firstLoneSurrogate(v) !== -1 || /[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]/.test(v)) {
        throw invalidProperty(key, `The document property "${key}" contains a character that XML cannot store.`, 'Remove control characters (other than tab, line feed and carriage return) and unpaired surrogates.');
      }
      out[key] = v;
    }
  }
  return out as WorkbookProperties;
}

/** A date as W3CDTF in UTC with whole seconds, e.g. `2026-10-30T12:00:00Z` (F-OPC-19). */
function w3cdtf(d: Date): string {
  const iso = d.toISOString();
  // toISOString gives "YYYY-MM-DDTHH:mm:ss.sssZ" for the years 0–9999.
  return `${iso.slice(0, 19)}Z`;
}

/** The Core Properties part, or `null` when no property is set. */
export function corePropertiesXml(properties: WorkbookProperties): string | null {
  let body = '';
  for (const [key, prefix] of ELEMENTS) {
    const v = properties[key];
    if (v === undefined) continue;
    const name = `${prefix}:${key}`;
    if (v instanceof Date) {
      // dcterms:created and dcterms:modified need xsi:type="dcterms:W3CDTF" (Part 2 §8.3.4.3);
      // cp:lastPrinted is xsd:dateTime and takes no attribute (§8.3.4.4.6).
      body += `<${name}${prefix === 'dcterms' ? ' xsi:type="dcterms:W3CDTF"' : ''}>${w3cdtf(v)}</${name}>`;
    } else {
      body += `<${name}>${escapeText(v)}</${name}>`;
    }
  }
  if (body === '') return null;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<cp:coreProperties xmlns:cp="${CORE_PROPERTIES_NS}" xmlns:dc="${DC_NS}" xmlns:dcterms="${DCTERMS_NS}" xmlns:xsi="${XSI_NS}">${body}</cp:coreProperties>`
  );
}

const DATE_RE = /^(\d{4})(?:-(\d{2})(?:-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(\.\d+)?)?(Z|[+-]\d{2}:\d{2})?)?)?)?$/;

/**
 * Parses a W3CDTF / xsd:dateTime value: `YYYY`, `YYYY-MM`, `YYYY-MM-DD` or a date with time,
 * optional seconds and fraction, and `Z` or `±hh:mm`. A value without a zone is read as UTC.
 * Returns `null` for anything else.
 */
export function parseW3cdtf(text: string): Date | null {
  const m = DATE_RE.exec(text.trim());
  if (m === null) return null;
  const num = (s: string | undefined, fallback: number): number => (s === undefined ? fallback : Number(s));
  const year = num(m[1], 0);
  const month = num(m[2], 1);
  const day = num(m[3], 1);
  const hour = num(m[4], 0);
  const minute = num(m[5], 0);
  const second = num(m[6], 0);
  const ms = m[7] === undefined ? 0 : Math.floor(Number(`0${m[7]}`) * 1000);
  if (month < 1 || month > 12 || day < 1 || hour > 23 || minute > 59 || second > 59) return null;
  const date = new Date(0);
  date.setUTCFullYear(year, month - 1, day);
  date.setUTCHours(hour, minute, second, ms);
  if (date.getUTCDate() !== day) return null;
  const zone = m[8];
  if (zone !== undefined && zone !== 'Z') {
    const sign = zone.startsWith('-') ? -1 : 1;
    const zh = Number(zone.slice(1, 3));
    const zm = Number(zone.slice(4, 6));
    if (zh > 23 || zm > 59) return null;
    date.setTime(date.getTime() - sign * (zh * 60 + zm) * 60_000);
  }
  return date;
}

/** Text of an element including the text of its children (keywords may wrap words in `value`). */
function allText(el: XmlElement): string {
  if (el.children.length === 0) return el.text;
  return [el.text, ...el.children.map(allText)].map((s) => s.trim()).filter((s) => s !== '').join(' ');
}

/** Context for reading the Core Properties part. */
export interface CorePropertiesContext {
  readonly part: string;
  readonly limits: ReadLimits;
  readonly warnings: WarningCollector;
}

/**
 * Reads the Core Properties part. Unknown elements are ignored; a date that cannot be read is
 * left out with `XLV_W_INVALID_VALUE`; malformed XML or another root element gives a warning and
 * no properties, because the part is optional metadata.
 */
export function parseCoreProperties(xml: string, ctx: CorePropertiesContext): WorkbookProperties {
  let root: XmlElement;
  try {
    root = parseXmlTree(xml, ctx);
  } catch (error) {
    if (error instanceof XlvError && error.code === 'XLV_E_XML_MALFORMED') {
      ctx.warnings.add('XLV_W_INVALID_VALUE', 'The document properties part is not well-formed XML; its properties are ignored.', { part: ctx.part });
      return {};
    }
    throw error;
  }
  if (root.ns !== CORE_PROPERTIES_NS || root.local !== 'coreProperties') {
    ctx.warnings.add('XLV_W_INVALID_VALUE', `The document properties part has the root element <${root.local}>; its properties are ignored.`, { part: ctx.part });
    return {};
  }
  const out: Record<string, string | Date> = {};
  for (const [key, prefix] of ELEMENTS) {
    const el = root.children.find((c) => c.ns === PREFIX_NS[prefix] && c.local === key);
    if (el === undefined) continue;
    const text = allText(el);
    if (DATE_KEYS.includes(key)) {
      if (text.trim() === '') continue;
      const date = parseW3cdtf(text);
      if (date === null) {
        ctx.warnings.add('XLV_W_INVALID_VALUE', `The document property ${prefix}:${key} has the value "${text.trim().slice(0, 40)}", which is not a date; it is ignored.`, { part: ctx.part });
        continue;
      }
      out[key] = date;
    } else {
      out[key] = text;
    }
  }
  return out as WorkbookProperties;
}
