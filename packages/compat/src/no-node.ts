// What `#xlvane-node` resolves to outside Node.js, Bun and Deno (package.json "imports"): loading
// it fails, and src/node-api.ts turns the failure into XLV_E_NOT_SUPPORTED.
throw new Error('xlvane/node is not available in this runtime.');
