// The file helpers of `xlvane/node`, loaded only when a file is read or written by name, through
// the package's own import `#xlvane-node` (package.json "imports"): under the `node` condition
// (Node.js, Bun, Deno) it is `xlvane/node`; elsewhere (browsers, Workers) it is dist/no-node.js,
// which throws. Bundlers resolve it with their own conditions, so browser and Workers bundles
// contain no node: imports (docs/decisions.md D-R01, D-R07).
import { XlvError } from 'xlvane';
import type * as XlvaneNode from 'xlvane/node';

/** Loads `xlvane/node`; outside Node.js, Bun and Deno it throws XLV_E_NOT_SUPPORTED. */
export async function xlvaneNode(method: string): Promise<typeof XlvaneNode> {
  try {
    return (await import('#xlvane-node')) as unknown as typeof XlvaneNode;
  } catch (cause) {
    throw new XlvError('XLV_E_NOT_SUPPORTED', `${method} needs a runtime with a file system (Node.js, Bun or Deno).`, {
      hint: 'Read or write the bytes yourself and use xlsx.load() / xlsx.writeBuffer() instead.',
      details: { feature: method },
      cause,
    });
  }
}

/**
 * Reads a file through `#node-fs` (node:fs/promises under the `node` condition), for image
 * files given by name; outside Node.js, Bun and Deno it throws XLV_E_NOT_SUPPORTED.
 */
export async function readFileBytes(path: string, method: string): Promise<Uint8Array> {
  let fs: { readFile(path: string): Promise<Uint8Array> };
  try {
    fs = (await import('#node-fs')) as unknown as typeof fs;
  } catch (cause) {
    throw new XlvError('XLV_E_NOT_SUPPORTED', `${method} needs a runtime with a file system (Node.js, Bun or Deno).`, {
      hint: "Pass the image contents instead: workbook.addImage({ buffer, extension: 'png' }).",
      details: { feature: method },
      cause,
    });
  }
  const data = await fs.readFile(path);
  return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
}
