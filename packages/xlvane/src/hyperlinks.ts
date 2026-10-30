// Cell hyperlinks (ECMA-376 Part 1 §18.3.1.47 hyperlink, §18.3.1.48 hyperlinks, §15.3;
// Part 4 §13.3, §15.10.10; [MS-OI29500] §2.1.626; docs/facts/sml-write.md W2-42 … W2-44;
// docs/decisions.md D-A08). A link to a web page or file is stored as a relationship of the
// sheet with TargetMode="External"; a link to a place in the workbook is stored in `location`
// only. Targets are data: xlvane never opens, fetches or resolves them.
import { XlvError } from './errors';
import { firstLoneSurrogate } from './xml/writer';

/**
 * A hyperlink read from a worksheet. `target` is the address of an external resource (a URL
 * or file path, exactly as stored); `location` is a place in this workbook (a sheet and cell or
 * a defined name, e.g. `"'Sales 2026'!A1"`) or, with a `target`, a place inside that resource.
 *
 * @example
 * ```ts
 * import { readWorkbook, type Hyperlink } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const links: readonly Hyperlink[] = readWorkbook(bytes).getSheet('Sales').hyperlinks;
 * for (const link of links) console.log(link.ref, link.target ?? link.location);
 * ```
 */
export interface Hyperlink {
  /** The cell or range the link is attached to, for example `"B4"` or `"B4:C5"`. */
  readonly ref: string;
  /** External target (URL, `mailto:` address, file path), or `null` for a link inside the workbook. */
  readonly target: string | null;
  /** Place in the workbook, or inside the external target; `null` if not given. */
  readonly location: string | null;
  /** Text shown when the pointer rests on the link; `null` if not given. */
  readonly tooltip: string | null;
  /** Display text stored with the link, if it differs from the cell text; `null` if not given. */
  readonly display: string | null;
}

/**
 * Where a written hyperlink points: `url` for an external resource, `location` for a place in
 * the workbook (or inside the resource, together with `url`). At least one of them is needed.
 * Limits follow [MS-OI29500] §2.1.626: `location` and `display` at most 2084 characters,
 * `tooltip` at most 255.
 *
 * @example
 * ```ts
 * import { createWorkbook, type HyperlinkTarget } from 'xlvane';
 *
 * const web: HyperlinkTarget = { url: 'https://example.com/report', tooltip: 'Open the report' };
 * const inside: HyperlinkTarget = { location: "'Sales 2026'!A1" };
 * createWorkbook().addSheet('Links').set('A1', 'Web').link('A1', web).set('A2', 'Sales').link('A2', inside);
 * ```
 */
export interface HyperlinkTarget {
  readonly url?: string;
  readonly location?: string;
  readonly tooltip?: string;
  readonly display?: string;
}

/** A validated hyperlink of the writer: the range is already normalised. */
export interface NormalizedHyperlink {
  readonly ref: string;
  readonly url: string | null;
  readonly location: string | null;
  readonly tooltip: string | null;
  readonly display: string | null;
}

const KEYS = ['url', 'location', 'tooltip', 'display'] as const;
/** [MS-OI29500] §2.1.626 a–c (W2-43). */
const MAX_LENGTH: Readonly<Record<(typeof KEYS)[number], number>> = { url: 65_536, location: 2084, display: 2084, tooltip: 255 };

function invalid(message: string, where: { sheet: string; cell: string }, hint: string, property?: string): XlvError {
  return new XlvError('XLV_E_INVALID_HYPERLINK', message, { location: where, hint, details: { cell: where.cell, ...(property === undefined ? {} : { property }) } });
}

/** Checks a hyperlink given to the writer; throws `XLV_E_INVALID_HYPERLINK`. */
export function normalizeHyperlink(value: unknown, where: { sheet: string; cell: string }): Omit<NormalizedHyperlink, 'ref'> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw invalid(`The hyperlink for ${where.cell} must be an object.`, where, 'Pass { url: "https://…" } or { location: "Sheet1!A1" }.');
  }
  const out: Record<string, string | null> = { url: null, location: null, tooltip: null, display: null };
  for (const key of Object.keys(value)) {
    if (!(KEYS as readonly string[]).includes(key)) {
      throw invalid(`Unknown hyperlink property "${key.slice(0, 40)}" for ${where.cell}.`, where, `Known properties: ${KEYS.join(', ')}.`, key.slice(0, 40));
    }
    const v = (value as Record<string, unknown>)[key];
    if (v === undefined) continue;
    const k = key as (typeof KEYS)[number];
    if (typeof v !== 'string') throw invalid(`The hyperlink property "${k}" for ${where.cell} must be a string.`, where, `For example ${k}: "text".`, k);
    if (v.length > MAX_LENGTH[k]) {
      throw invalid(`The hyperlink property "${k}" for ${where.cell} has ${String(v.length)} characters; at most ${String(MAX_LENGTH[k])} are allowed.`, where, `Shorten "${k}" to ${String(MAX_LENGTH[k])} characters or fewer.`, k);
    }
    if (firstLoneSurrogate(v) !== -1) throw invalid(`The hyperlink property "${k}" for ${where.cell} contains an unpaired surrogate.`, where, 'Remove the broken character.', k);
    // The URL becomes a relationship Target (xsd:anyURI, written as is): control characters
    // including tab, LF and CR, and U+FFFE/U+FFFF, which XML cannot carry, are refused (D-D05).
    if (k === 'url' && (v === '' || /[\x00-\x1F\x7F\uFFFE\uFFFF]/.test(v))) {
      throw invalid(`The hyperlink URL for ${where.cell} is empty or contains control characters or non-characters.`, where, 'Pass the address as text, for example url: "https://example.com"; percent-encode unusual characters.', k);
    }
    out[k] = v;
  }
  if (out['url'] === null && (out['location'] === null || out['location'] === '')) {
    throw invalid(`The hyperlink for ${where.cell} has neither a url nor a location.`, where, 'Pass { url: "https://…" } for a web page or file, or { location: "Sheet1!A1" } for a place in the workbook.');
  }
  return out as unknown as Omit<NormalizedHyperlink, 'ref'>;
}

/** Schemes that are ordinary for links in a spreadsheet; others are reported (D-A08). */
const USUAL_SCHEME = /^(?:https?|mailto):/i;

/** Whether an external target uses a scheme other than http, https or mailto. */
export function isUnusualTarget(target: string): boolean {
  return /^[A-Za-z][A-Za-z0-9+.-]*:/.test(target) ? !USUAL_SCHEME.test(target) : true;
}
