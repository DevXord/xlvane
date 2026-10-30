// Node.js globals used by tests and scripts only. Not visible to package sources, so core code
// cannot accidentally depend on `process` (docs/decisions.md, D-006).

declare const process: {
  readonly env: Record<string, string | undefined>;
  readonly platform: string;
  readonly versions: Readonly<Record<string, string | undefined>>;
};
