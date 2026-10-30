// Pictures on worksheets: the Drawings part of a sheet (ECMA-376 Part 1 §12.3.8, §20.5), the
// Image parts it points to (§15.2.14) and the sheet background (`picture`, §18.3.1.67). Facts:
// docs/facts/tables-drawings.md TD-30 … TD-39; decisions: docs/decisions.md D-D03. Only pictures
// are modelled; charts, shapes and group shapes in a drawing are counted and reported, and
// `editCells` keeps every drawing byte for byte.
import { XlvError } from './errors';
import type { WarningCollector } from './util/warnings';
import type { ReadLimits } from './limits';
import type { OpcPackage, Relationship } from './opc/package';
import { formatCellRef, parseCellRef, parseRange } from './sml/cell-ref';
import { classifyRelationship, PACKAGE_REL_NS, REL_NS } from './sml/namespaces';
import { childNamed, parseXmlTree, type XmlElement } from './xml/tree';
import { escapeAttr } from './xml/writer';

/**
 * Image formats xlvane writes and recognises (maintainer's decision, D-D03).
 *
 * @example
 * ```ts
 * import { imageFormat, type ImageFormat } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const format: ImageFormat | null = imageFormat(bytes); // 'png', 'jpeg', 'gif' or null
 * ```
 */
export type ImageFormat = 'png' | 'jpeg' | 'gif';

/**
 * An image stored in the workbook (an Image part, Part 1 §15.2.14). Images used several times
 * are one object. `format` is recognised from the first bytes of the image, not from the part
 * name; it is `null` for other formats (EMF, WMF, TIFF, …), which are kept as they are.
 *
 * @example
 * ```ts
 * import { readWorkbook, type WorkbookImage } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const images: readonly WorkbookImage[] = readWorkbook(bytes).images;
 * for (const image of images) console.log(image.part, image.format, image.bytes.length);
 * ```
 */
export interface WorkbookImage {
  /** Part name in the package, for example `/xl/media/image1.png`. */
  readonly part: string;
  readonly format: ImageFormat | null;
  /** The content type the package declares for the part, or `null`. */
  readonly contentType: string | null;
  readonly bytes: Uint8Array;
}

/**
 * A corner of a picture: the cell it is in (0-based `col` and `row`, as stored, and the A1 name
 * in `cell`) and the offset from that cell's top-left corner in EMU (English Metric Units:
 * 914 400 per inch, 12 700 per point; Part 1 §20.1.2.1).
 *
 * @example
 * ```ts
 * import type { AnchorMarker } from 'xlvane';
 *
 * const corner: AnchorMarker = { cell: 'B3', col: 1, row: 2, colOffset: 0, rowOffset: 0 };
 * ```
 */
export interface AnchorMarker {
  readonly cell: string;
  readonly col: number;
  readonly row: number;
  readonly colOffset: number;
  readonly rowOffset: number;
}

/**
 * How a picture is placed (Part 1 §20.5.2.33 twoCellAnchor, §20.5.2.24 oneCellAnchor,
 * §20.5.2.1 absoluteAnchor). Sizes and positions are in EMU. `editAs` says how a two-cell
 * picture follows changes of its cells: `twoCell` (moves and resizes, the default),
 * `oneCell` (moves only) or `absolute` (neither).
 *
 * @example
 * ```ts
 * import type { ImageAnchor } from 'xlvane';
 *
 * function describe(anchor: ImageAnchor): string {
 *   if (anchor.type === 'twoCell') return `${anchor.from.cell} to ${anchor.to.cell}`;
 *   if (anchor.type === 'oneCell') return `at ${anchor.from.cell}`;
 *   return `at ${String(anchor.position.x)}, ${String(anchor.position.y)} EMU`;
 * }
 * ```
 */
export type ImageAnchor =
  | { readonly type: 'twoCell'; readonly from: AnchorMarker; readonly to: AnchorMarker; readonly editAs: 'twoCell' | 'oneCell' | 'absolute' }
  | { readonly type: 'oneCell'; readonly from: AnchorMarker; readonly size: { readonly width: number; readonly height: number } }
  | { readonly type: 'absolute'; readonly position: { readonly x: number; readonly y: number }; readonly size: { readonly width: number; readonly height: number } };

/**
 * A picture on a worksheet. `image` is `null` when the picture links to an image outside the
 * file or its image part is missing. `hyperlink` is the target of a click on the picture, read as
 * text and never opened.
 *
 * @example
 * ```ts
 * import { readWorkbook, type SheetImage } from 'xlvane';
 *
 * declare const bytes: Uint8Array;
 * const pictures: readonly SheetImage[] = readWorkbook(bytes).getSheet('Report').images;
 * for (const p of pictures) console.log(p.name, p.anchor.type, p.image?.format);
 * ```
 */
export interface SheetImage {
  readonly name: string;
  readonly description: string | null;
  readonly anchor: ImageAnchor;
  readonly image: WorkbookImage | null;
  readonly hyperlink: string | null;
}

/**
 * Where a picture goes, for `NewSheet.addImage`:
 * - a range such as `"B2:D6"`: the picture covers these cells;
 * - `{ from, to }`: corners, each a cell name (its top-left corner) or a 0-based `{ col, row }`
 *   position whose fraction is a part of the cell (`{ col: 1.5, row: 2 }` is the middle of the top
 *   edge of B3);
 * - `{ from, size }`: the top-left corner and the size in pixels.
 *
 * `editAs` (two corners only) is written as given; without it the schema default `twoCell`
 * applies (the picture moves and resizes with its cells).
 *
 * @example
 * ```ts
 * import type { ImageAnchorInput } from 'xlvane';
 *
 * const cells: ImageAnchorInput = 'B2:D6';
 * const corners: ImageAnchorInput = { from: { col: 1.5, row: 1 }, to: 'E8', editAs: 'oneCell' };
 * const sized: ImageAnchorInput = { from: 'B2', size: { width: 160, height: 90 } };
 * ```
 */
export type ImageAnchorInput =
  | string
  | { readonly from: ImagePositionInput; readonly to: ImagePositionInput; readonly editAs?: 'twoCell' | 'oneCell' | 'absolute' }
  | { readonly from: ImagePositionInput; readonly size: { readonly width: number; readonly height: number } };

/**
 * A corner of a picture: a cell name (its top-left corner) or a 0-based fractional position.
 *
 * @example
 * ```ts
 * import type { ImagePositionInput } from 'xlvane';
 *
 * const a: ImagePositionInput = 'C4';
 * const b: ImagePositionInput = { col: 2.25, row: 3.5 };
 * ```
 */
export type ImagePositionInput = string | { readonly col: number; readonly row: number };

/**
 * Options of `NewSheet.addImage`: the picture's name and description (alternative text) and a
 * hyperlink that a click on the picture opens.
 *
 * @example
 * ```ts
 * import type { ImageOptions } from 'xlvane';
 *
 * const options: ImageOptions = { description: 'Company logo', hyperlink: 'https://example.com' };
 * ```
 */
export interface ImageOptions {
  readonly name?: string;
  readonly description?: string;
  readonly hyperlink?: string;
}

/** Content type of a Drawings part (Part 1 §12.3.8, TD-30). */
export const DRAWING_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml';
/** Content types of the image formats (Part 1 §15.2.14 note, TD-31). */
export const IMAGE_CONTENT_TYPES: Readonly<Record<ImageFormat, string>> = { png: 'image/png', jpeg: 'image/jpeg', gif: 'image/gif' };
const XDR_NS = { transitional: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', strict: 'http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing' } as const;
const A_NS = { transitional: 'http://schemas.openxmlformats.org/drawingml/2006/main', strict: 'http://purl.oclc.org/ooxml/drawingml/main' } as const;
/** EMU per pixel at 96 pixels per inch: 914 400 / 96 (TD-33; the pixel size is a convention). */
export const EMU_PER_PIXEL = 9525;
/** Last 0-based column and row of a drawing marker ([MS-OI29500] §2.1.1358 ff., TD-32). */
const MAX_MARKER_COL = 16_383;
const MAX_MARKER_ROW = 1_048_575;
/** ST_PositiveCoordinate maximum (dml-main.xsd L204–209). */
const MAX_COORDINATE = 27_273_042_316_900;

/**
 * The format of an image from its first bytes (TD-31): PNG `89 50 4E 47 0D 0A 1A 0A`, JPEG
 * `FF D8 FF`, GIF `GIF87a` or `GIF89a`. These signatures were observed in the image files of the
 * POI corpus and of LibreOffice 24.2.7.2; the format specifications were not available.
 *
 * @example
 * ```ts
 * import { imageFormat } from 'xlvane';
 *
 * console.log(imageFormat(new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61]))); // 'gif'
 * ```
 */
export function imageFormat(bytes: Uint8Array): ImageFormat | null {
  const b = (i: number): number => bytes[i] ?? -1;
  if (b(0) === 0x89 && b(1) === 0x50 && b(2) === 0x4e && b(3) === 0x47 && b(4) === 0x0d && b(5) === 0x0a && b(6) === 0x1a && b(7) === 0x0a) return 'png';
  if (b(0) === 0xff && b(1) === 0xd8 && b(2) === 0xff) return 'jpeg';
  if (b(0) === 0x47 && b(1) === 0x49 && b(2) === 0x46 && b(3) === 0x38 && (b(4) === 0x37 || b(4) === 0x39) && b(5) === 0x61) return 'gif';
  return null;
}

// ---------------------------------------------------------------------------------------------
// Reading

/** The pictures of a worksheet and its background image. */
export interface SheetDrawing {
  readonly images: readonly SheetImage[];
  readonly background: WorkbookImage | null;
}

/** Image parts already read, by part name, so that every image is one object. */
export type MediaCache = Map<string, WorkbookImage | null>;

function mediaOf(pkg: OpcPackage, rel: Relationship | undefined, media: MediaCache): WorkbookImage | null {
  if (rel === undefined || rel.external) return null;
  const cached = media.get(rel.target);
  if (cached !== undefined) return cached;
  const bytes = pkg.readBytes(rel.target);
  const image = bytes === undefined ? null : Object.freeze({ part: rel.target, format: imageFormat(bytes), contentType: pkg.contentType(rel.target) ?? null, bytes });
  media.set(rel.target, image);
  return image;
}

/** Parses an ST_Coordinate: an integer in EMU, or a number with in, cm, mm or pt (TD-33). */
function coordinate(text: string | undefined): number | null {
  const t = (text ?? '').trim();
  if (/^-?\d+$/.test(t)) return Number(t);
  const m = /^(-?\d+(?:\.\d+)?)(in|cm|mm|pt)$/.exec(t);
  if (m === null) return null;
  const per = { in: 914_400, cm: 360_000, mm: 36_000, pt: 12_700 }[m[2] as 'in' | 'cm' | 'mm' | 'pt'];
  return Math.round(Number(m[1]) * per);
}

/**
 * Reads the Drawings part and the background picture of a worksheet. Pictures are read; other
 * objects (charts, shapes, connectors, group shapes, content parts, alternate content) are
 * counted and reported once with `XLV_W_DRAWING_CONTENT_NOT_READ`. Broken markers skip their
 * picture with a warning; a malformed or foreign drawing part gives a warning and no pictures.
 */
export function readSheetDrawing(
  pkg: OpcPackage,
  sheetPart: string,
  sheet: string,
  ids: { readonly drawing: string | null; readonly picture: string | null },
  ctx: { readonly limits: ReadLimits; readonly warnings: WarningCollector },
  media: MediaCache,
): SheetDrawing {
  if (ids.drawing === null && ids.picture === null) return { images: [], background: null };
  const rels = pkg.relationships(sheetPart);
  const byId = (id: string | null, name: string): Relationship | undefined => (id === null ? undefined : rels.find((r) => r.id === id && classifyRelationship(r.type)?.name === name));
  const background = mediaOf(pkg, byId(ids.picture, 'image'), media);
  const drawing = byId(ids.drawing, 'drawing');
  if (drawing === undefined || drawing.external) {
    if (ids.drawing !== null) ctx.warnings.add('XLV_W_PART_MISSING', 'The sheet refers to a drawing that is not a drawing part of the package; its pictures are not read.', { part: sheetPart, sheet });
    return { images: [], background };
  }
  const xml = pkg.readXml(drawing.target);
  if (xml === undefined) {
    ctx.warnings.add('XLV_W_PART_MISSING', 'The drawing part of the sheet does not exist; its pictures are not read.', { part: drawing.target, sheet });
    return { images: [], background };
  }
  const where = { part: drawing.target, sheet };
  const warn = (message: string): void => ctx.warnings.add('XLV_W_INVALID_VALUE', message, where);
  let root: XmlElement;
  try {
    root = parseXmlTree(xml, { part: drawing.target, limits: ctx.limits });
  } catch (error) {
    if (error instanceof XlvError && error.code === 'XLV_E_XML_MALFORMED') {
      warn('The drawing part of the sheet is not well-formed XML; its pictures are not read.');
      return { images: [], background };
    }
    throw error;
  }
  const xdr = root.ns === XDR_NS.strict ? XDR_NS.strict : XDR_NS.transitional;
  const a = xdr === XDR_NS.strict ? A_NS.strict : A_NS.transitional;
  const r = xdr === XDR_NS.strict ? REL_NS.strict : REL_NS.transitional;
  if (root.local !== 'wsDr' || root.ns !== xdr) {
    warn(`The drawing part of the sheet has the root element <${root.local}>; its pictures are not read.`);
    return { images: [], background };
  }
  const drawingRels = pkg.relationships(drawing.target);
  const images: SheetImage[] = [];
  let other = 0;
  const marker = (el: XmlElement | undefined): AnchorMarker | null => {
    if (el === undefined) return null;
    const int = (local: string, max: number): number | null => {
      const t = (childNamed(el, xdr, local)?.text ?? '').trim();
      return /^\d+$/.test(t) && Number(t) <= max ? Number(t) : null;
    };
    const col = int('col', MAX_MARKER_COL);
    const row = int('row', MAX_MARKER_ROW);
    const colOffset = coordinate(childNamed(el, xdr, 'colOff')?.text);
    const rowOffset = coordinate(childNamed(el, xdr, 'rowOff')?.text);
    if (col === null || row === null || colOffset === null || rowOffset === null) return null;
    return Object.freeze({ cell: formatCellRef(row + 1, col + 1), col, row, colOffset, rowOffset });
  };
  const size = (el: XmlElement | undefined): { width: number; height: number } | null => {
    const w = coordinate(el?.attrs.get('cx'));
    const h = coordinate(el?.attrs.get('cy'));
    return w === null || h === null || w < 0 || h < 0 ? null : Object.freeze({ width: w, height: h });
  };
  for (const el of root.children) {
    if (el.ns !== xdr || (el.local !== 'twoCellAnchor' && el.local !== 'oneCellAnchor' && el.local !== 'absoluteAnchor')) {
      if (el.ns !== xdr || el.local !== 'extLst') other++;
      continue;
    }
    const pic = childNamed(el, xdr, 'pic');
    if (pic === undefined) {
      other++;
      continue;
    }
    let anchor: ImageAnchor | null = null;
    if (el.local === 'twoCellAnchor') {
      const from = marker(childNamed(el, xdr, 'from'));
      const to = marker(childNamed(el, xdr, 'to'));
      const editAs = el.attrs.get('editAs') ?? 'twoCell';
      if (from !== null && to !== null) anchor = { type: 'twoCell', from, to, editAs: editAs === 'oneCell' || editAs === 'absolute' ? editAs : 'twoCell' };
    } else if (el.local === 'oneCellAnchor') {
      const from = marker(childNamed(el, xdr, 'from'));
      const ext = size(childNamed(el, xdr, 'ext'));
      if (from !== null && ext !== null) anchor = { type: 'oneCell', from, size: ext };
    } else {
      const pos = childNamed(el, xdr, 'pos');
      const x = coordinate(pos?.attrs.get('x'));
      const y = coordinate(pos?.attrs.get('y'));
      const ext = size(childNamed(el, xdr, 'ext'));
      if (x !== null && y !== null && ext !== null) anchor = { type: 'absolute', position: Object.freeze({ x, y }), size: ext };
    }
    if (anchor === null) {
      warn('A picture of the sheet has an invalid position; it is skipped.');
      continue;
    }
    const nv = childNamed(childNamed(pic, xdr, 'nvPicPr') ?? pic, xdr, 'cNvPr');
    const blip = childNamed(childNamed(pic, xdr, 'blipFill') ?? pic, a, 'blip');
    const embed = blip?.attrs.get(`{${r}}embed`);
    const imageRel = embed === undefined ? undefined : drawingRels.find((x) => x.id === embed && classifyRelationship(x.type)?.name === 'image');
    const linkId = nv === undefined ? undefined : childNamed(nv, a, 'hlinkClick')?.attrs.get(`{${r}}id`);
    const linkRel = linkId === undefined ? undefined : drawingRels.find((x) => x.id === linkId && classifyRelationship(x.type)?.name === 'hyperlink');
    images.push(
      Object.freeze({
        name: nv?.attrs.get('name') ?? '',
        description: nv?.attrs.get('descr') ?? null,
        anchor: Object.freeze(anchor),
        image: mediaOf(pkg, imageRel, media),
        hyperlink: linkRel?.target ?? null,
      }),
    );
  }
  if (other > 0) {
    ctx.warnings.add('XLV_W_DRAWING_CONTENT_NOT_READ', `The drawing of the sheet has ${String(other)} object${other === 1 ? '' : 's'} other than pictures (charts, shapes or groups); xlvane reports only pictures, and editing keeps the drawing unchanged.`, where);
  }
  return { images, background };
}

// ---------------------------------------------------------------------------------------------
// Writing

/** A picture as stored by `NewSheet.addImage`, before its anchor is resolved. */
export interface StoredImage {
  readonly image: number;
  readonly anchor: ImageAnchorInput;
  /** `null`: "Picture N", numbered within the drawing. */
  readonly name: string | null;
  readonly description: string | null;
  readonly hyperlink: string | null;
}

/** A marker in file form: 0-based cell and offsets in EMU. */
interface Marker {
  readonly col: number;
  readonly row: number;
  readonly colOff: number;
  readonly rowOff: number;
}

/** A picture with its anchor resolved to markers (and the size in EMU for one-cell anchors). */
export interface ResolvedImage {
  readonly image: number;
  readonly from: Marker;
  readonly to: Marker | null;
  readonly ext: { readonly cx: number; readonly cy: number } | null;
  readonly editAs: 'twoCell' | 'oneCell' | 'absolute' | null;
  readonly name: string | null;
  readonly description: string | null;
  readonly hyperlink: string | null;
}

function invalid(message: string, hint: string, details?: Record<string, string | number>): XlvError {
  return new XlvError('XLV_E_INVALID_OPTION', message, { hint, ...(details ? { details } : {}) });
}

/** Checks the image bytes given to `NewWorkbook.addImage` and returns their format. */
export function checkImage(bytes: unknown): { bytes: Uint8Array; format: ImageFormat } {
  const data = bytes instanceof Uint8Array ? bytes : bytes instanceof ArrayBuffer ? new Uint8Array(bytes) : null;
  if (data === null) throw new XlvError('XLV_E_INVALID_IMAGE', 'addImage expects the bytes of an image (a Uint8Array).', { hint: 'Pass the file contents, for example new Uint8Array(await file.arrayBuffer()).' });
  const format = imageFormat(data);
  if (format === null) {
    throw new XlvError('XLV_E_INVALID_IMAGE', 'The image is not a PNG, JPEG or GIF file (recognised from its first bytes).', {
      hint: 'Convert the image to PNG or JPEG first; xlvane writes PNG, JPEG and GIF images.',
      details: { size: data.length },
    });
  }
  return { bytes: data, format };
}

/** Validates the options of `NewSheet.addImage` (the anchor is resolved when the sheet is written). */
export function normalizeImage(image: number, count: number, anchor: unknown, options: unknown): StoredImage {
  if (typeof image !== 'number' || !Number.isInteger(image) || image < 0 || image >= count) {
    throw invalid(`Image ${String(image).slice(0, 20)} was not added to this workbook.`, 'Pass the number that workbook.addImage(bytes) returned.', { image: String(image).slice(0, 20) });
  }
  checkAnchor(anchor);
  let name: string | null = null;
  let description: string | null = null;
  let hyperlink: string | null = null;
  if (options !== undefined) {
    if (typeof options !== 'object' || options === null) throw invalid('The image options must be an object.', 'For example { description: "Logo" }.');
    for (const [key, value] of Object.entries(options)) {
      if (key !== 'name' && key !== 'description' && key !== 'hyperlink') throw invalid(`Unknown image option "${key.slice(0, 40)}".`, 'Use name, description or hyperlink.', { option: key.slice(0, 40) });
      if (value === undefined) continue;
      if (typeof value !== 'string' || value.length > 2084 || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(value) || (key !== 'description' && value === '')) {
        throw invalid(`The image option "${key}" must be a text of 1 to 2084 characters without control characters.`, key === 'hyperlink' ? 'For example hyperlink: "https://example.com".' : 'For example name: "Logo".', { option: key });
      }
      if (key === 'name') name = value;
      else if (key === 'description') description = value;
      else hyperlink = value;
    }
  }
  return { image, anchor: anchor as ImageAnchorInput, name, description, hyperlink };
}

function position(p: unknown, what: string): { col: number; row: number } {
  if (typeof p === 'string') {
    const at = parseCellRef(p.trim().toUpperCase());
    if (at === null) throw invalid(`The picture ${what} "${p.slice(0, 20)}" is not a cell.`, 'Use a cell such as "B2" or a position such as { col: 1, row: 1 }.', { option: what });
    return { col: at.col - 1, row: at.row - 1 };
  }
  if (typeof p === 'object' && p !== null) {
    const { col, row } = p as { col?: unknown; row?: unknown };
    const keys = Object.keys(p);
    if (keys.every((k) => k === 'col' || k === 'row') && typeof col === 'number' && typeof row === 'number' && Number.isFinite(col) && Number.isFinite(row) && col >= 0 && row >= 0 && col < MAX_MARKER_COL + 1 && row < MAX_MARKER_ROW + 1) {
      return { col, row };
    }
  }
  throw invalid(`The picture ${what} is not a valid position.`, 'Use a cell such as "B2", or { col, row } with 0-based numbers (col below 16384, row below 1048576).', { option: what });
}

function checkAnchor(anchor: unknown): void {
  if (typeof anchor === 'string') {
    if (parseRange(anchor.trim().toUpperCase()) === null) throw invalid(`The picture range "${anchor.slice(0, 40)}" is not valid.`, 'Use a range such as "B2:D6".', { option: 'anchor' });
    return;
  }
  if (typeof anchor !== 'object' || anchor === null) throw invalid('The picture position must be a range or an object.', 'For example "B2:D6", { from: "B2", to: "E8" } or { from: "B2", size: { width: 160, height: 90 } }.', { option: 'anchor' });
  const o = anchor as Record<string, unknown>;
  for (const key of Object.keys(o)) {
    if (!['from', 'to', 'size', 'editAs'].includes(key)) throw invalid(`Unknown picture position option "${key.slice(0, 40)}".`, 'Use from with to (and editAs), or from with size.', { option: key.slice(0, 40) });
  }
  const from = position(o['from'], 'from');
  if ((o['to'] === undefined) === (o['size'] === undefined)) throw invalid('A picture position needs either "to" or "size".', 'For example { from: "B2", to: "E8" } or { from: "B2", size: { width: 160, height: 90 } }.', { option: 'anchor' });
  if (o['to'] !== undefined) {
    const to = position(o['to'], 'to');
    if (to.col < from.col || to.row < from.row) throw invalid('The picture corner "to" is above or left of "from".', 'Give the top-left corner as from and the bottom-right corner as to.', { option: 'to' });
    const editAs = o['editAs'];
    if (editAs !== undefined && editAs !== 'twoCell' && editAs !== 'oneCell' && editAs !== 'absolute') throw invalid('editAs must be "twoCell", "oneCell" or "absolute".', 'Use editAs: "oneCell" for a picture that moves but does not resize with its cells.', { option: 'editAs' });
  } else {
    if (o['editAs'] !== undefined) throw invalid('editAs applies to pictures with two corners only.', 'Use { from, to, editAs } or leave editAs out.', { option: 'editAs' });
    const s = o['size'] as { width?: unknown; height?: unknown } | null;
    const ok = (v: unknown): boolean => typeof v === 'number' && Number.isFinite(v) && v > 0 && v * EMU_PER_PIXEL <= MAX_COORDINATE;
    if (typeof s !== 'object' || s === null || !ok(s.width) || !ok(s.height)) throw invalid('The picture size must be { width, height } in pixels, both above 0.', 'For example size: { width: 160, height: 90 }.', { option: 'size' });
  }
}

/**
 * Column width and row height in pixels, for fractional positions (TD-33, PROVISIONAL): a column
 * of stored width `w` is `round(w × 7)` pixels (W2-03: the width includes the padding, in
 * characters of the 7-pixel maximum digit width of Calibri 11), a column without a width 64
 * pixels (Excel's default, observed); a row of `h` points is `round(h × 96 / 72)` pixels, a row
 * without a height 20 pixels (15 points, the `defaultRowHeight` xlvane writes).
 */
export interface SheetMetrics {
  colPixels(col: number): number;
  rowPixels(row: number): number;
}

/** Metrics from the column widths and row heights of a sheet being written (1-based maps). */
export function metricsOf(columns: ReadonlyMap<number, { readonly width: number | null; readonly hidden: boolean }>, rows: ReadonlyMap<number, { readonly height: number | null; readonly hidden: boolean }>): SheetMetrics {
  return {
    colPixels(col) {
      const c = columns.get(col + 1);
      if (c?.hidden) return 0;
      return c?.width == null ? 64 : Math.round(c.width * 7);
    },
    rowPixels(row) {
      const r = rows.get(row + 1);
      if (r?.hidden) return 0;
      return r?.height == null ? 20 : Math.round((r.height * 96) / 72);
    },
  };
}

function toMarker(p: { col: number; row: number }, m: SheetMetrics): Marker {
  const col = Math.floor(p.col);
  const row = Math.floor(p.row);
  return { col, row, colOff: Math.round((p.col - col) * m.colPixels(col) * EMU_PER_PIXEL), rowOff: Math.round((p.row - row) * m.rowPixels(row) * EMU_PER_PIXEL) };
}

/** Resolves the anchor of a stored picture with the sheet's column widths and row heights. */
export function resolveImage(s: StoredImage, m: SheetMetrics): ResolvedImage {
  const base = { image: s.image, name: s.name, description: s.description, hyperlink: s.hyperlink };
  const anchor = s.anchor;
  if (typeof anchor === 'string') {
    const range = parseRange(anchor.trim().toUpperCase())!;
    return { ...base, from: { col: range.start.col - 1, row: range.start.row - 1, colOff: 0, rowOff: 0 }, to: { col: range.end.col, row: range.end.row, colOff: 0, rowOff: 0 }, ext: null, editAs: null };
  }
  const from = toMarker(position(anchor.from, 'from'), m);
  if ('to' in anchor && anchor.to !== undefined) return { ...base, from, to: toMarker(position(anchor.to, 'to'), m), ext: null, editAs: anchor.editAs ?? null };
  const size = (anchor as { size: { width: number; height: number } }).size;
  return { ...base, from, to: null, ext: { cx: Math.round(size.width * EMU_PER_PIXEL), cy: Math.round(size.height * EMU_PER_PIXEL) }, editAs: null };
}

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const attr = (name: string, value: string | number): string => ` ${name}="${typeof value === 'number' ? String(value) : escapeAttr(value)}"`;

/**
 * The Image parts of a package: each distinct image (same format and bytes) is stored once, as
 * `xl/media/imageN.<png|jpeg|gif>` (TD-31), in the order the pictures use them.
 */
export class MediaParts {
  private readonly byIndex = new Map<number, number>();
  private readonly byKey = new Map<string, number[]>();
  private readonly list: { readonly bytes: Uint8Array; readonly format: ImageFormat }[] = [];

  constructor(private readonly images: readonly { readonly bytes: Uint8Array; readonly format: ImageFormat }[]) {}

  /** Target of image `index` relative to `xl/drawings/` or `xl/worksheets/`. */
  target(index: number): string {
    let n = this.byIndex.get(index);
    if (n === undefined) {
      const img = this.images[index]!;
      let h = 2166136261;
      for (const b of img.bytes) h = Math.imul(h ^ b, 16777619);
      const key = `${img.format}:${String(img.bytes.length)}:${String(h >>> 0)}`;
      const same = (this.byKey.get(key) ?? []).find((k) => sameBytes(this.list[k - 1]!.bytes, img.bytes));
      if (same !== undefined) n = same;
      else {
        this.list.push(img);
        n = this.list.length;
        this.byKey.set(key, [...(this.byKey.get(key) ?? []), n]);
      }
      this.byIndex.set(index, n);
    }
    return `../media/image${String(n)}.${this.list[n - 1]!.format}`;
  }

  /** The Image parts as ZIP entries. */
  entries(): { name: string; data: Uint8Array }[] {
    return this.list.map((img, i) => ({ name: `xl/media/image${String(i + 1)}.${img.format}`, data: img.bytes }));
  }

  /** The formats used, for the `Default` content types. */
  formats(): ImageFormat[] {
    return [...new Set(this.list.map((img) => img.format))];
  }
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/** The Drawings part of one worksheet and the elements for the sheet part. */
export interface DrawingParts {
  /** `<drawing r:id="…"/>`, written before `legacyDrawing` (CT_Worksheet, TD-34). */
  readonly drawing: string;
  /** `<picture r:id="…"/>`, written after `legacyDrawing` (CT_Worksheet, TD-34). */
  readonly picture: string;
  /** The drawing part and its relationships, or `null` without pictures. */
  readonly part: { readonly name: string; readonly xml: string; readonly relsName: string; readonly rels: string } | null;
}

function markerXml(local: 'from' | 'to', m: Marker): string {
  return `<xdr:${local}><xdr:col>${String(m.col)}</xdr:col><xdr:colOff>${String(m.colOff)}</xdr:colOff><xdr:row>${String(m.row)}</xdr:row><xdr:rowOff>${String(m.rowOff)}</xdr:rowOff></xdr:${local}>`;
}

/**
 * Builds the Drawings part `xl/drawings/drawing<n>.xml` (Part 1 §20.5, TD-35) for the pictures of
 * a sheet, adds the `drawing` and background `image` relationships to `rels`, and returns the
 * sheet elements. Each picture is an `xdr:pic` in a `twoCellAnchor` or `oneCellAnchor`, with a
 * unique `cNvPr` id, `a:blip r:embed` to its image and `a:stretch/a:fillRect`, as in the Part 1
 * Annex L picture example; one image relationship per image and drawing.
 */
export function drawingParts(pictures: readonly ResolvedImage[], background: number | null, media: MediaParts, n: number, rels: { add(type: string, target: string, external?: boolean): string }): DrawingParts {
  const picture = background === null ? '' : `<picture${attr('r:id', rels.add('image', media.target(background)))}/>`;
  if (pictures.length === 0) return { drawing: '', picture, part: null };
  const items: string[] = [];
  const imageIds = new Map<number, string>();
  const add = (type: string, target: string, external = false): string => {
    const id = `rId${String(items.length + 1)}`;
    items.push(`<Relationship${attr('Id', id)}${attr('Type', REL_BASE + type)}${attr('Target', target)}${external ? ' TargetMode="External"' : ''}/>`);
    return id;
  };
  let xml = `${XML_DECL}<xdr:wsDr xmlns:xdr="${XDR_NS.transitional}" xmlns:a="${A_NS.transitional}" xmlns:r="${REL_BASE.slice(0, -1)}">`;
  pictures.forEach((p, i) => {
    let embed = imageIds.get(p.image);
    if (embed === undefined) {
      embed = add('image', media.target(p.image));
      imageIds.set(p.image, embed);
    }
    const link = p.hyperlink === null ? '' : `<a:hlinkClick${attr('r:id', add('hyperlink', p.hyperlink, true))}/>`;
    const body =
      `<xdr:pic><xdr:nvPicPr><xdr:cNvPr${attr('id', i + 1)}${attr('name', p.name ?? `Picture ${String(i + 1)}`)}${p.description === null ? '' : attr('descr', p.description)}>${link}</xdr:cNvPr><xdr:cNvPicPr/></xdr:nvPicPr>` +
      `<xdr:blipFill><a:blip${attr('r:embed', embed)}/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
      '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/>';
    if (p.to !== null) xml += `<xdr:twoCellAnchor${p.editAs === null ? '' : attr('editAs', p.editAs)}>${markerXml('from', p.from)}${markerXml('to', p.to)}${body}</xdr:twoCellAnchor>`;
    else xml += `<xdr:oneCellAnchor>${markerXml('from', p.from)}<xdr:ext${attr('cx', p.ext!.cx)}${attr('cy', p.ext!.cy)}/>${body}</xdr:oneCellAnchor>`;
  });
  xml += '</xdr:wsDr>';
  const name = `xl/drawings/drawing${String(n)}.xml`;
  return {
    drawing: `<drawing${attr('r:id', rels.add('drawing', `../drawings/drawing${String(n)}.xml`))}/>`,
    picture,
    part: { name, xml, relsName: drawingRelsPartName(n), rels: `${XML_DECL}<Relationships xmlns="${PACKAGE_REL_NS}">${items.join('')}</Relationships>` },
  };
}

/** Part name of the relationships part of drawing `n`. */
export function drawingRelsPartName(n: number): string {
  return `xl/drawings/_rels/drawing${String(n)}.xml.rels`;
}
