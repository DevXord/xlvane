// A small, non-validating XML 1.0 parser for Office Open XML parts.
//
// Security properties (docs/decisions.md, D-101):
// - Document type declarations are always rejected, so there are no external entities (XXE)
//   and no entity expansion ("billion laughs"). Only the five predefined entities and
//   character references are recognised (XML 1.0 §4.1, §4.6).
// - The parser is iterative: element nesting uses an explicit stack, never recursion.
// - Depth, attribute count, name length and text length are limited.
// - Attributes are delivered in a Map, so names from the file never become object keys.
//
// Conformance notes (docs/facts/zip-xml-web.md X-01 … X-16, N-01 … N-09): line ends are normalised before parsing (§2.11),
// attribute values are normalised as for undeclared (CDATA) attributes (§3.3.3), every
// character must match the Char production (§2.2), names follow the Name production (§2.3)
// and namespaces follow Namespaces in XML 1.0.
import { XlvError, type XlvDetails } from '../errors';
import type { ReadLimits } from '../limits';

/** The limits the XML parser enforces. */
export type XmlLimits = Pick<ReadLimits, 'maxXmlDepth' | 'maxXmlAttributes' | 'maxXmlNameLength' | 'maxXmlTextLength'>;

/** Receives parse events in document order. */
export interface XmlHandler {
  /**
   * A start tag. `attrs` maps attribute names to values: unprefixed attributes by their local
   * name, namespaced ones in Clark notation (`{namespace}local`). Namespace declarations are
   * not included.
   */
  start(ns: string, local: string, attrs: ReadonlyMap<string, string>): void;
  /** An end tag (also reported for empty-element tags). */
  end(ns: string, local: string): void;
  /** Character data between two tags, with references expanded and CDATA sections included. */
  text(value: string): void;
}

/** Options for {@link parseXml}. */
export interface XmlParseOptions {
  /** Part name used in error locations. */
  readonly part: string;
  readonly limits: XmlLimits;
}

/** Namespace bound to the `xml` prefix by definition (Namespaces in XML 1.0, §3). */
export const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const XMLNS_NAMESPACE = 'http://www.w3.org/2000/xmlns/';

// Character codes used by the scanner.
const TAB = 0x09;
const LF = 0x0a;
const CR = 0x0d;
const SPACE = 0x20;
const QUOTE = 0x22;
const HASH = 0x23;
const AMP = 0x26;
const APOS = 0x27;
const MINUS = 0x2d;
const DOT = 0x2e;
const SLASH = 0x2f;
const COLON = 0x3a;
const SEMICOLON = 0x3b;
const LT = 0x3c;
const EQ = 0x3d;
const GT = 0x3e;
const QUESTION = 0x3f;
const UNDERSCORE = 0x5f;

function isWhitespace(c: number): boolean {
  return c === SPACE || c === LF || c === TAB || c === CR;
}

function isNameStartChar(c: number): boolean {
  if (c < 0x80) {
    return (c >= 0x61 && c <= 0x7a) || (c >= 0x41 && c <= 0x5a) || c === UNDERSCORE || c === COLON;
  }
  return (
    (c >= 0xc0 && c <= 0xd6) ||
    (c >= 0xd8 && c <= 0xf6) ||
    (c >= 0xf8 && c <= 0x2ff) ||
    (c >= 0x370 && c <= 0x37d) ||
    (c >= 0x37f && c <= 0x1fff) ||
    (c >= 0x200c && c <= 0x200d) ||
    (c >= 0x2070 && c <= 0x218f) ||
    (c >= 0x2c00 && c <= 0x2fef) ||
    (c >= 0x3001 && c <= 0xd7ff) ||
    (c >= 0xf900 && c <= 0xfdcf) ||
    (c >= 0xfdf0 && c <= 0xfffd) ||
    (c >= 0x10000 && c <= 0xeffff)
  );
}

function isNameChar(c: number): boolean {
  if (c < 0x80) {
    return (
      (c >= 0x61 && c <= 0x7a) ||
      (c >= 0x41 && c <= 0x5a) ||
      (c >= 0x30 && c <= 0x39) ||
      c === UNDERSCORE ||
      c === COLON ||
      c === MINUS ||
      c === DOT
    );
  }
  return isNameStartChar(c) || c === 0xb7 || (c >= 0x300 && c <= 0x36f) || (c >= 0x203f && c <= 0x2040);
}

/** The Char production of XML 1.0 for a code point. */
function isXmlChar(c: number): boolean {
  return (
    c === TAB ||
    c === LF ||
    c === CR ||
    (c >= 0x20 && c <= 0xd7ff) ||
    (c >= 0xe000 && c <= 0xfffd) ||
    (c >= 0x10000 && c <= 0x10ffff)
  );
}

const PREDEFINED_ENTITIES = new Map([
  ['lt', '<'],
  ['gt', '>'],
  ['amp', '&'],
  ['apos', "'"],
  ['quot', '"'],
]);

interface OpenElement {
  readonly qname: string;
  readonly ns: string;
  readonly local: string;
  /** Previous bindings to restore when the element closes, or null if it declared none. */
  readonly undo: [prefix: string, previous: string | undefined][] | null;
}

/**
 * Parses an XML document held in a string and reports events to `handler`.
 * Throws {@link XlvError} with `XLV_E_XML_MALFORMED`, `XLV_E_XML_DTD_FORBIDDEN`,
 * `XLV_E_XML_ENTITY_FORBIDDEN` or `XLV_E_LIMIT_EXCEEDED`.
 */
export function parseXml(xml: string, handler: XmlHandler, options: XmlParseOptions): void {
  new Parser(xml, handler, options).run();
}

class Parser {
  private readonly s: string;
  private pos = 0;
  private readonly stack: OpenElement[] = [];
  /** Current prefix bindings; '' is the default namespace. */
  private readonly bindings = new Map<string, string>([['xml', XML_NAMESPACE]]);
  private textBuffer = '';
  private rootSeen = false;

  constructor(
    xml: string,
    private readonly handler: XmlHandler,
    private readonly options: XmlParseOptions,
  ) {
    // §2.11: translate CR LF and lone CR to LF before parsing.
    this.s = xml.indexOf('\r') === -1 ? xml : xml.replace(/\r\n?/g, '\n');
  }

  run(): void {
    this.checkCharacters();
    const s = this.s;
    if (s.startsWith('<?xml') && (s.length === 5 || isWhitespace(s.charCodeAt(5)) || s.charCodeAt(5) === QUESTION)) {
      this.skipProcessingInstruction(true);
    }
    while (this.pos < s.length) {
      const lt = s.indexOf('<', this.pos);
      if (this.stack.length === 0) {
        const end = lt === -1 ? s.length : lt;
        for (let i = this.pos; i < end; i++) {
          if (!isWhitespace(s.charCodeAt(i))) {
            this.fail(this.rootSeen ? 'Text after the root element.' : 'Text before the root element.', i);
          }
        }
        if (lt === -1) {
          this.pos = s.length;
          break;
        }
        this.pos = lt;
        this.markupOutsideRoot();
        continue;
      }
      if (lt === -1) {
        this.fail(`Element <${this.stack[this.stack.length - 1]?.qname ?? ''}> is not closed.`, s.length);
      }
      if (lt > this.pos) this.characterData(this.pos, lt);
      this.pos = lt;
      this.markupInContent();
    }
    if (this.stack.length > 0) {
      this.fail(`Element <${this.stack[this.stack.length - 1]?.qname ?? ''}> is not closed.`, s.length);
    }
    if (!this.rootSeen) this.fail('The document has no root element.', s.length);
  }

  /** Every character must match the Char production (XML 1.0 §2.2). */
  private checkCharacters(): void {
    const s = this.s;
    for (let i = 0; i < s.length; i++) {
      const c = s.charCodeAt(i);
      if (c >= 0x20 && c < 0xd800) continue;
      if (c === LF || c === TAB || c === CR) continue;
      if (c >= 0xd800 && c <= 0xdbff) {
        const next = s.charCodeAt(i + 1);
        if (next >= 0xdc00 && next <= 0xdfff) {
          i++;
          continue;
        }
      } else if (c >= 0xe000 && c <= 0xfffd) {
        continue;
      }
      this.fail(`Character U+${c.toString(16).toUpperCase().padStart(4, '0')} is not allowed in XML.`, i);
    }
  }

  private markupOutsideRoot(): void {
    const s = this.s;
    if (s.startsWith('<!--', this.pos)) return this.skipComment();
    if (s.startsWith('<?', this.pos)) return this.skipProcessingInstruction(false);
    if (s.startsWith('<!DOCTYPE', this.pos)) return this.failDoctype();
    if (s.startsWith('<!', this.pos)) this.fail('Unexpected markup declaration outside the root element.', this.pos);
    if (s.charCodeAt(this.pos + 1) === SLASH) this.fail('End tag without a matching start tag.', this.pos);
    if (this.rootSeen) this.fail('A second root element is not allowed.', this.pos);
    this.rootSeen = true;
    this.startTag();
  }

  private markupInContent(): void {
    const s = this.s;
    const next = s.charCodeAt(this.pos + 1);
    if (next === SLASH) return this.endTag();
    if (next === 0x21 /* ! */) {
      if (s.startsWith('<!--', this.pos)) return this.skipComment();
      if (s.startsWith('<![CDATA[', this.pos)) return this.cdata();
      if (s.startsWith('<!DOCTYPE', this.pos)) return this.failDoctype();
      this.fail('Unexpected markup declaration in element content.', this.pos);
    }
    if (next === QUESTION) return this.skipProcessingInstruction(false);
    this.startTag();
  }

  // ----- character data -------------------------------------------------------------------

  private characterData(from: number, to: number): void {
    const raw = this.s.slice(from, to);
    const close = raw.indexOf(']]>');
    if (close !== -1) this.fail('The sequence "]]>" is not allowed in character data.', from + close);
    this.appendText(raw.indexOf('&') === -1 ? raw : this.expandReferences(raw, from), from);
  }

  private cdata(): void {
    const start = this.pos + 9;
    const end = this.s.indexOf(']]>', start);
    if (end === -1) this.fail('CDATA section is not terminated.', this.pos);
    this.appendText(this.s.slice(start, end), this.pos);
    this.pos = end + 3;
  }

  private appendText(text: string, at: number): void {
    if (text.length === 0) return;
    if (this.textBuffer.length + text.length > this.options.limits.maxXmlTextLength) {
      this.limit('maxXmlTextLength', 'Text node is longer than the limit.', at);
    }
    this.textBuffer += text;
  }

  private flushText(): void {
    if (this.textBuffer.length > 0) {
      const text = this.textBuffer;
      this.textBuffer = '';
      this.handler.text(text);
    }
  }

  /**
   * Expands entity and character references in `raw`, which starts at offset `base`.
   */
  private expandReferences(raw: string, base: number): string {
    let out = '';
    let last = 0;
    let amp = raw.indexOf('&');
    while (amp !== -1) {
      out += raw.slice(last, amp);
      const semi = raw.indexOf(';', amp + 1);
      if (semi === -1) this.fail('"&" must start an entity or character reference ending with ";".', base + amp);
      const body = raw.slice(amp + 1, semi);
      if (body.charCodeAt(0) === HASH) {
        out += this.characterReference(body, base + amp);
      } else {
        const value = PREDEFINED_ENTITIES.get(body);
        if (value === undefined) {
          if (body.length > this.options.limits.maxXmlNameLength) {
            this.limit('maxXmlNameLength', 'Entity name is longer than the limit.', base + amp);
          }
          if (!this.isName(body)) this.fail('"&" must start an entity or character reference.', base + amp);
          throw new XlvError(
            'XLV_E_XML_ENTITY_FORBIDDEN',
            `Reference to entity "&${body};", which is not one of the five predefined XML entities.`,
            {
              location: this.location(base + amp),
              hint: 'Only &lt; &gt; &amp; &apos; &quot; and character references such as &#160; are allowed, because xlvane does not process DTDs. Re-save the file in a spreadsheet application.',
              details: { entity: body },
            },
          );
        }
        out += value;
      }
      last = semi + 1;
      amp = raw.indexOf('&', last);
    }
    out += raw.slice(last);
    return out;
  }

  /**
   * `&#` [0-9]+ `;` or `&#x` [0-9a-fA-F]+ `;` (XML 1.0 [66], X-02): any number of digits, so
   * leading zeros are allowed; only a lowercase `x` introduces the hexadecimal form.
   */
  private characterReference(body: string, at: number): string {
    const shown = body.length > 20 ? `${body.slice(0, 20)}…` : body;
    const hex = body.charCodeAt(1) === 0x78; /* x */
    const start = hex ? 2 : 1;
    if (body.length === start) this.fail(`Invalid character reference "&${shown};".`, at);
    let code = 0;
    for (let i = start; i < body.length; i++) {
      const c = body.charCodeAt(i);
      let d: number;
      if (c >= 0x30 && c <= 0x39) d = c - 0x30;
      else if (hex && c >= 0x61 && c <= 0x66) d = c - 0x57;
      else if (hex && c >= 0x41 && c <= 0x46) d = c - 0x37;
      else return this.fail(`Invalid character reference "&${shown};".`, at);
      // Stop growing once the value is out of range; the digits are still all validated.
      if (code <= 0x10ffff) code = code * (hex ? 16 : 10) + d;
    }
    if (!isXmlChar(code)) this.fail(`Character reference "&${shown};" refers to a character not allowed in XML.`, at);
    return String.fromCodePoint(code);
  }

  // ----- tags -----------------------------------------------------------------------------

  private startTag(): void {
    const s = this.s;
    const tagStart = this.pos;
    this.flushText();
    this.pos++;
    const qname = this.readName();
    const raw: [name: string, value: string, at: number][] = [];
    let selfClosing = false;
    for (;;) {
      const hadSpace = this.skipWhitespace();
      const c = s.charCodeAt(this.pos);
      if (c === GT) {
        this.pos++;
        break;
      }
      if (c === SLASH) {
        if (s.charCodeAt(this.pos + 1) !== GT) this.fail('Expected ">" after "/".', this.pos);
        this.pos += 2;
        selfClosing = true;
        break;
      }
      if (Number.isNaN(c)) this.fail(`Start tag <${qname}> is not terminated.`, tagStart);
      if (!hadSpace) this.fail('Attributes must be separated by whitespace.', this.pos);
      const at = this.pos;
      const name = this.readName();
      this.skipWhitespace();
      if (s.charCodeAt(this.pos) !== EQ) this.fail(`Attribute "${name}" has no value.`, this.pos);
      this.pos++;
      this.skipWhitespace();
      const quote = s.charCodeAt(this.pos);
      if (quote !== QUOTE && quote !== APOS) this.fail(`Value of attribute "${name}" must be quoted.`, this.pos);
      const valueStart = this.pos + 1;
      const valueEnd = s.indexOf(quote === QUOTE ? '"' : "'", valueStart);
      if (valueEnd === -1) this.fail(`Value of attribute "${name}" is not terminated.`, this.pos);
      if (valueEnd - valueStart > this.options.limits.maxXmlTextLength) {
        this.limit('maxXmlTextLength', `Value of attribute "${name}" is longer than the limit.`, valueStart);
      }
      let value = s.slice(valueStart, valueEnd);
      const lt = value.indexOf('<');
      if (lt !== -1) this.fail('"<" is not allowed in attribute values.', valueStart + lt);
      // §3.3.3: each literal whitespace character becomes a space; CR is already gone (§2.11).
      if (value.indexOf('\n') !== -1 || value.indexOf('\t') !== -1) value = value.replace(/[\n\t]/g, ' ');
      if (value.indexOf('&') !== -1) value = this.expandReferences(value, valueStart);
      for (const [existing] of raw) {
        if (existing === name) this.fail(`Attribute "${name}" appears twice.`, at);
      }
      raw.push([name, value, at]);
      if (raw.length > this.options.limits.maxXmlAttributes) {
        this.limit('maxXmlAttributes', `Element <${qname}> has more attributes than the limit.`, at);
      }
      this.pos = valueEnd + 1;
    }

    if (this.stack.length + 1 > this.options.limits.maxXmlDepth) {
      this.limit('maxXmlDepth', 'Elements are nested deeper than the limit.', tagStart);
    }

    // Namespace declarations first: they apply to the element itself and its attributes.
    let undo: [string, string | undefined][] | null = null;
    for (const [name, value, at] of raw) {
      let prefix: string;
      if (name === 'xmlns') prefix = '';
      else if (name.startsWith('xmlns:')) {
        prefix = name.slice(6);
        // PrefixedAttName ::= 'xmlns:' NCName (Namespaces in XML 1.0 [2], N-01).
        this.checkNcName(prefix, at);
      } else continue;
      this.checkDeclaration(prefix, value, at);
      (undo ??= []).push([prefix, this.bindings.get(prefix)]);
      this.bindings.set(prefix, value);
    }

    const [ns, local] = this.resolve(qname, true, tagStart);
    const attrs = new Map<string, string>();
    for (const [name, value, at] of raw) {
      if (name === 'xmlns' || name.startsWith('xmlns:')) continue;
      const [attrNs, attrLocal] = this.resolve(name, false, at);
      const key = attrNs === '' ? attrLocal : `{${attrNs}}${attrLocal}`;
      if (attrs.has(key)) this.fail(`Attribute "${name}" appears twice (same namespace and local name).`, at);
      attrs.set(key, value);
    }

    this.handler.start(ns, local, attrs);
    if (selfClosing) {
      this.handler.end(ns, local);
      this.restore(undo);
    } else {
      this.stack.push({ qname, ns, local, undo });
    }
  }

  private endTag(): void {
    const tagStart = this.pos;
    this.pos += 2;
    const qname = this.readName();
    this.skipWhitespace();
    if (this.s.charCodeAt(this.pos) !== GT) this.fail(`End tag </${qname}> is not terminated.`, tagStart);
    this.pos++;
    const open = this.stack.pop();
    if (open === undefined || open.qname !== qname) {
      this.fail(
        open === undefined ? `End tag </${qname}> has no start tag.` : `End tag </${qname}> does not match <${open.qname}>.`,
        tagStart,
      );
    }
    this.flushText();
    this.handler.end(open.ns, open.local);
    this.restore(open.undo);
  }

  private restore(undo: [string, string | undefined][] | null): void {
    if (undo === null) return;
    for (let i = undo.length - 1; i >= 0; i--) {
      const entry = undo[i];
      if (entry === undefined) continue;
      const [prefix, previous] = entry;
      if (previous === undefined) this.bindings.delete(prefix);
      else this.bindings.set(prefix, previous);
    }
  }

  /** Namespace constraints on declarations (Namespaces in XML 1.0, §3). */
  private checkDeclaration(prefix: string, uri: string, at: number): void {
    // Namespace names become part of attribute keys ({uri}local), so they are bounded like
    // names; real Office Open XML namespace names are well under 100 characters.
    if (uri.length > this.options.limits.maxXmlNameLength) {
      this.limit('maxXmlNameLength', 'Namespace name is longer than the limit.', at);
    }
    if (prefix === 'xmlns') this.fail('The prefix "xmlns" must not be declared.', at);
    if (prefix === 'xml' && uri !== XML_NAMESPACE) this.fail('The prefix "xml" must not be bound to another namespace.', at);
    if (prefix !== 'xml' && uri === XML_NAMESPACE) this.fail('The XML namespace must not be bound to another prefix.', at);
    if (uri === XMLNS_NAMESPACE) this.fail('The xmlns namespace must not be declared.', at);
    if (prefix !== '' && uri === '') this.fail(`Prefix "${prefix}" cannot be undeclared in XML 1.0.`, at);
    if (prefix !== '') this.checkNcName(prefix, at);
  }

  private resolve(qname: string, element: boolean, at: number): [ns: string, local: string] {
    const colon = qname.indexOf(':');
    if (colon === -1) return [element ? (this.bindings.get('') ?? '') : '', qname];
    const prefix = qname.slice(0, colon);
    const local = qname.slice(colon + 1);
    this.checkNcName(prefix, at);
    this.checkNcName(local, at);
    const ns = this.bindings.get(prefix);
    if (ns === undefined) this.fail(`Namespace prefix "${prefix}" is not declared.`, at);
    return [ns, local];
  }

  private checkNcName(name: string, at: number): void {
    if (name.length === 0 || name.indexOf(':') !== -1) this.fail(`"${name}" is not a valid name.`, at);
  }

  private readName(): string {
    const s = this.s;
    const start = this.pos;
    const first = s.codePointAt(start) ?? -1;
    if (!isNameStartChar(first)) this.fail('Expected a name.', start);
    let i = start + (first > 0xffff ? 2 : 1);
    for (;;) {
      const c = s.charCodeAt(i);
      if (c < 0x80) {
        if (!isNameChar(c)) break;
        i++;
      } else {
        const cp = s.codePointAt(i) ?? -1;
        if (!isNameChar(cp)) break;
        i += cp > 0xffff ? 2 : 1;
      }
      if (i - start > this.options.limits.maxXmlNameLength) {
        this.limit('maxXmlNameLength', 'Name is longer than the limit.', start);
      }
    }
    this.pos = i;
    return s.slice(start, i);
  }

  private isName(text: string): boolean {
    if (text.length === 0) return false;
    const first = text.codePointAt(0) ?? -1;
    if (!isNameStartChar(first)) return false;
    for (let i = first > 0xffff ? 2 : 1; i < text.length; ) {
      const cp = text.codePointAt(i) ?? -1;
      if (!isNameChar(cp)) return false;
      i += cp > 0xffff ? 2 : 1;
    }
    return true;
  }

  private skipWhitespace(): boolean {
    const start = this.pos;
    while (isWhitespace(this.s.charCodeAt(this.pos))) this.pos++;
    return this.pos > start;
  }

  // ----- comments, processing instructions, DOCTYPE ---------------------------------------

  private skipComment(): void {
    const start = this.pos;
    const end = this.s.indexOf('--', start + 4);
    if (end === -1) this.fail('Comment is not terminated.', start);
    if (this.s.charCodeAt(end + 2) !== GT) this.fail('"--" is not allowed inside a comment.', end);
    this.pos = end + 3;
  }

  private skipProcessingInstruction(declaration: boolean): void {
    const start = this.pos;
    this.pos += 2;
    const target = this.readName();
    if (!declaration && target.toLowerCase() === 'xml') {
      this.fail('The XML declaration is allowed only at the very start of the document.', start);
    }
    const end = this.s.indexOf('?>', this.pos);
    if (end === -1) this.fail('Processing instruction is not terminated.', start);
    const c = this.s.charCodeAt(this.pos);
    if (this.pos !== end && !isWhitespace(c)) this.fail('Invalid processing instruction target.', start);
    if (declaration) {
      // XMLDecl ::= '<?xml' VersionInfo …, VersionInfo ::= S 'version' Eq … (XML 1.0 [23], [24], X-13).
      let i = this.pos;
      while (isWhitespace(this.s.charCodeAt(i))) i++;
      if (i === this.pos || !this.s.startsWith('version', i)) {
        this.fail('The XML declaration must start with the version, e.g. <?xml version="1.0"?>.', start);
      }
    }
    this.pos = end + 2;
  }

  private failDoctype(): never {
    throw new XlvError('XLV_E_XML_DTD_FORBIDDEN', 'The part contains a document type declaration (DTD).', {
      location: this.location(this.pos),
      hint: 'xlvane rejects DTDs to protect against external entities (XXE) and entity expansion attacks. Office Open XML parts do not use DTDs; if the file was generated by a program, report the problem to its authors.',
    });
  }

  // ----- errors ---------------------------------------------------------------------------

  private location(offset: number): { part: string; line: number; column: number } {
    let line = 1;
    let lineStart = 0;
    for (let i = this.s.indexOf('\n'); i !== -1 && i < offset; i = this.s.indexOf('\n', i + 1)) {
      line++;
      lineStart = i + 1;
    }
    return { part: this.options.part, line, column: offset - lineStart + 1 };
  }

  private fail(message: string, offset: number): never {
    throw new XlvError('XLV_E_XML_MALFORMED', message, {
      location: this.location(offset),
      hint: 'The part is not well-formed XML. Re-save the file in a spreadsheet application; if it was generated by a program, report the problem to its authors.',
    });
  }

  private limit(limit: keyof XmlLimits, message: string, offset: number): never {
    const details: XlvDetails = { limit, value: this.options.limits[limit] };
    throw new XlvError('XLV_E_LIMIT_EXCEEDED', message, {
      location: this.location(offset),
      hint: `If you trust this file, raise the limit, for example { limits: { ${limit}: ${String(this.options.limits[limit] * 4)} } }.`,
      details,
    });
  }
}

/**
 * Decodes the bytes of an XML part. Office Open XML packaging parts are UTF-8 or UTF-16
 * (docs/facts/ecma-376.md E-55); a byte order mark selects UTF-16, otherwise UTF-8 is used. Invalid byte sequences
 * and declarations of other encodings are rejected.
 */
export function decodeXml(bytes: Uint8Array, part: string): string {
  let encoding: 'utf-8' | 'utf-16le' | 'utf-16be' = 'utf-8';
  if (bytes[0] === 0xff && bytes[1] === 0xfe) encoding = 'utf-16le';
  else if (bytes[0] === 0xfe && bytes[1] === 0xff) encoding = 'utf-16be';
  let text: string;
  try {
    text = new TextDecoder(encoding, { fatal: true }).decode(bytes);
  } catch (cause) {
    throw new XlvError('XLV_E_XML_ENCODING', `The part is not valid ${encoding.toUpperCase()}.`, {
      location: { part },
      hint: 'Office Open XML parts must be encoded in UTF-8 or UTF-16. Re-save the file in a spreadsheet application.',
      cause,
    });
  }
  const declared = declaredEncoding(text);
  if (declared !== null) {
    const normalized = declared.toLowerCase();
    const ok = encoding === 'utf-8' ? normalized === 'utf-8' : normalized === 'utf-16' || normalized === encoding;
    if (!ok) {
      throw new XlvError('XLV_E_XML_ENCODING', `The part declares encoding "${declared}", which is not supported.`, {
        location: { part },
        hint: 'Office Open XML parts must be encoded in UTF-8 or UTF-16. Re-save the file in a spreadsheet application.',
        details: { encoding: declared },
      });
    }
  }
  return text;
}

/** Returns the `encoding` pseudo-attribute of the XML declaration, if present. */
function declaredEncoding(text: string): string | null {
  if (!text.startsWith('<?xml')) return null;
  const end = text.indexOf('?>');
  if (end === -1) return null;
  const decl = text.slice(5, end);
  const at = decl.indexOf('encoding');
  if (at === -1) return null;
  let i = at + 8;
  while (isWhitespace(decl.charCodeAt(i))) i++;
  if (decl.charCodeAt(i) !== EQ) return null;
  i++;
  while (isWhitespace(decl.charCodeAt(i))) i++;
  const quote = decl[i];
  if (quote !== '"' && quote !== "'") return null;
  const close = decl.indexOf(quote, i + 1);
  return close === -1 ? null : decl.slice(i + 1, close);
}
