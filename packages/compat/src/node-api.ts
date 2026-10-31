// The file helpers of `xlvane/node`, loaded only when a file is read or written by name. The
// specifier is not a literal, so bundlers for browsers and Cloudflare Workers do not follow it and
// the bundle stays free of node: imports (docs/decisions.md D-R01).
import { XlvError } from 'xlvane';
import type * as XlvaneNode from 'xlvane/node';

/** Loads `xlvane/node`; outside Node.js, Bun and Deno it throws XLV_E_NOT_SUPPORTED. */
export async function xlvaneNode(method: string): Promise<typeof XlvaneNode> {
  try {
    const specifier = 'xlvane/node';
    return (await import(/* @vite-ignore */ /* webpackIgnore: true */ specifier)) as typeof XlvaneNode;
  } catch (cause) {
    throw new XlvError('XLV_E_NOT_SUPPORTED', `${method} needs a runtime with a file system (Node.js, Bun or Deno).`, {
      hint: 'Read or write the bytes yourself and use xlsx.load() / xlsx.writeBuffer() instead.',
      details: { feature: method },
      cause,
    });
  }
}
