import { XlvError } from '../errors';
import type { ReadLimits } from '../limits';
import { parseXml, type XmlLimits } from './parser';

/** An element of a small XML part parsed into memory. */
export interface XmlElement {
  readonly ns: string;
  readonly local: string;
  /** Attributes as in {@link XmlHandler.start}: local name, or `{namespace}local`. */
  readonly attrs: ReadonlyMap<string, string>;
  readonly children: XmlElement[];
  /** Concatenated character data that appears directly inside this element. */
  text: string;
}

/** Options for {@link parseXmlTree}. */
export interface XmlTreeOptions {
  readonly part: string;
  readonly limits: XmlLimits & Pick<ReadLimits, 'maxXmlTreeElements'>;
}

/**
 * Parses a whole part into a tree. Use only for small parts (workbook, relationships, content
 * types, styles); worksheets and shared strings are processed with {@link parseXml} directly.
 */
export function parseXmlTree(xml: string, options: XmlTreeOptions): XmlElement {
  const stack: XmlElement[] = [];
  let root: XmlElement | undefined;
  let count = 0;
  parseXml(
    xml,
    {
      start(ns, local, attrs) {
        // Elements and attributes both stay in memory, so both count against the budget.
        count += 1 + attrs.size;
        if (count > options.limits.maxXmlTreeElements) {
          throw new XlvError('XLV_E_LIMIT_EXCEEDED', 'The part has more elements and attributes than the limit.', {
            location: { part: options.part },
            hint: `If you trust this file, raise the limit, for example { limits: { maxXmlTreeElements: ${String(options.limits.maxXmlTreeElements * 4)} } }.`,
            details: { limit: 'maxXmlTreeElements', value: options.limits.maxXmlTreeElements },
          });
        }
        const element: XmlElement = { ns, local, attrs, children: [], text: '' };
        const parent = stack[stack.length - 1];
        if (parent === undefined) root = element;
        else parent.children.push(element);
        stack.push(element);
      },
      end() {
        stack.pop();
      },
      text(value) {
        const current = stack[stack.length - 1];
        if (current !== undefined) current.text += value;
      },
    },
    options,
  );
  // parseXml guarantees exactly one root element.
  return root as XmlElement;
}

/** Returns the direct children of `element` with the given namespace and local name. */
export function childrenNamed(element: XmlElement, ns: string, local: string): XmlElement[] {
  return element.children.filter((c) => c.local === local && c.ns === ns);
}

/** Returns the first direct child of `element` with the given namespace and local name. */
export function childNamed(element: XmlElement, ns: string, local: string): XmlElement | undefined {
  return element.children.find((c) => c.local === local && c.ns === ns);
}
