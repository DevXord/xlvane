// Stands in for `xlvane/node` in the portable smoke bundle (scripts/build-smoke.mjs): file
// helpers exist only in Node.js, and the smoke suite never calls them.
function unavailable(): never {
  throw new Error('xlvane/node is not available in this runtime.');
}
export const readWorkbookFile = unavailable;
export const writeWorkbookFile = unavailable;
export const openWorkbookStreamFile = unavailable;
export const writeWorkbookStreamFile = unavailable;
