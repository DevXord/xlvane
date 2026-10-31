// The one text form of dates in xlvane's text outputs: CSV, JSON (`toJSON`, `dates: "iso"`),
// `inspect()` samples and the CLI (docs/decisions.md D-306, D-R02).

function pad2(n: number): string {
  return n < 10 ? `0${String(n)}` : String(n);
}

/**
 * ISO 8601 without a time zone, because cell dates have none: the date alone at midnight
 * (`2024-01-31`), otherwise date and time (`2024-01-31T13:45:00`, with `.sss` only when there are
 * milliseconds). The UTC fields of the `Date` are the cell's wall-clock date and time (V-* facts).
 */
export function isoDate(d: Date): string {
  const date = `${String(d.getUTCFullYear()).padStart(4, '0')}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
  const ms = d.getUTCMilliseconds();
  if (d.getUTCHours() === 0 && d.getUTCMinutes() === 0 && d.getUTCSeconds() === 0 && ms === 0) return date;
  const time = `${pad2(d.getUTCHours())}:${pad2(d.getUTCMinutes())}:${pad2(d.getUTCSeconds())}`;
  return `${date}T${time}${ms === 0 ? '' : `.${String(ms).padStart(3, '0')}`}`;
}
