import { defineConfig } from 'tsdown';

export default defineConfig({
  // cli: the `xlvane-mcp` command (package.json "bin").
  entry: { index: 'src/index.ts', cli: 'src/bin.ts' },
  format: 'esm',
  platform: 'neutral',
  target: 'es2022',
  dts: true,
  clean: true,
  // Dependencies stay imports: xlvane and the MCP SDK are installed next to this package.
  deps: { neverBundle: [/^node:/, /^xlvane/, /^@modelcontextprotocol\//] },
});
