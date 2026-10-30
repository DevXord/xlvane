import type { XlvLocation, XlvWarning, XlvWarningCode } from '../errors';

const MAX_EXAMPLE_CELLS = 10;

interface Entry {
  code: XlvWarningCode;
  message: string;
  location: XlvLocation;
  count: number;
  cells: string[];
}

/**
 * Collects warnings, aggregating repeated occurrences of the same code in the same sheet or
 * part into one warning with a count and up to ten example cells. Entries are found through
 * nested maps (code, then sheet, then part), so no key is built from file data: a sheet name of
 * millions of characters costs nothing per warning.
 */
export class WarningCollector {
  private readonly byCode = new Map<XlvWarningCode, Map<string, Map<string, Entry>>>();
  private readonly order: Entry[] = [];

  add(code: XlvWarningCode, message: string, location: XlvLocation = {}): void {
    let bySheet = this.byCode.get(code);
    if (bySheet === undefined) {
      bySheet = new Map();
      this.byCode.set(code, bySheet);
    }
    const sheet = location.sheet ?? '';
    let byPart = bySheet.get(sheet);
    if (byPart === undefined) {
      byPart = new Map();
      bySheet.set(sheet, byPart);
    }
    const part = location.part ?? '';
    const entry = byPart.get(part);
    if (entry === undefined) {
      const { cell, ...rest } = location;
      const created: Entry = { code, message, location: rest, count: 1, cells: cell === undefined ? [] : [cell] };
      byPart.set(part, created);
      this.order.push(created);
      return;
    }
    entry.count++;
    if (location.cell !== undefined && entry.cells.length < MAX_EXAMPLE_CELLS) entry.cells.push(location.cell);
  }

  /** Warnings in the order their first occurrence was reported. */
  list(): XlvWarning[] {
    return this.order.map((e) =>
      e.cells.length > 0
        ? { code: e.code, message: e.message, location: e.location, count: e.count, cells: [...e.cells] }
        : { code: e.code, message: e.message, location: e.location, count: e.count },
    );
  }
}
