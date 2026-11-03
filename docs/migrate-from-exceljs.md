# Migrating from ExcelJS

`@xlvane/compat` offers the commonly used part of the exceljs 4.4.0 API on top of xlvane. It is installed under the name `exceljs` with an npm alias, so code that imports `exceljs` keeps its imports:

```json
"dependencies": {
  "exceljs": "npm:@xlvane/compat@^1"
}
```

npm installs an alias `<alias>@npm:<name>` under the alias name in `node_modules` (npm/cli `docs/lib/content/using-npm/package-spec.md`, "Aliases"; docs/exceljs-compat.md), so `import ExcelJS from 'exceljs'` and `require('exceljs')` load `@xlvane/compat`. What is supported, partly supported and not supported is listed method by method in [docs/exceljs-compat.md](https://github.com/DevXord/xlvane/blob/main/docs/exceljs-compat.md) (also in `node_modules/xlvane/docs/exceljs-compat.md` after installing xlvane).

This guide uses the `xlvane migrate` command of the `xlvane` package. The `xlvane` commands, the report shown below and the TypeScript examples are checked by the tests in `tests/docs/`; the npm commands were run with npm 10.9.9 on 2026-11-03.

## Before you start

- **Node.js.** `@xlvane/compat` is an ES module and declares `engines.node` `>=22.12`. `require('exceljs')` of an ES module works without a flag from Node.js 22.12 (docs/exceljs-compat.md); `import` works too, and so do named imports such as `import { Workbook } from 'exceljs'`.
- **Version control.** Commit `package.json` and the lockfile first, so the change can be reviewed and undone.

## 1. See what would change

In the directory that holds the project's `package.json`:

```sh
npx xlvane migrate
```

`xlvane migrate` does the following, and nothing else:

1. It reads `package.json` of the project directory (the current directory, or the directory given as the first argument). A `package.json` that is a symbolic link is refused.
2. It looks for an `"exceljs"` entry in `dependencies`, `devDependencies`, `optionalDependencies` and `peerDependencies`, and shows how each would change to `"npm:@xlvane/compat@^1"`. Other sections, such as `overrides`, are not changed.
3. It prints the Node.js requirement of `require('exceljs')` next to the project's own `engines.node`.
4. It scans the source files of the project (`.js`, `.mjs`, `.cjs`, `.jsx`, `.ts`, `.mts`, `.cts`, `.tsx`, up to 2 MiB each and 10 000 files) for exceljs API that `@xlvane/compat` does not support: `ModelContainer` and deep imports of `exceljs/lib/...`. Only files that contain the text `exceljs` are checked; `node_modules`, folders whose names start with a dot, and symbolic links are skipped. The matching is textual, so a finding is a place to look at, not proof.
5. It writes nothing unless `--force` is given.

For a project whose `package.json` has `"exceljs": "^4.4.0"` in `dependencies`, with a file `src/report.js` that imports `exceljs` and a file `src/legacy.cjs` that requires `exceljs/lib/xlsx/xform/base-xform`, the report is:

```text
Change in package.json:
  dependencies: - "exceljs": "^4.4.0"
                + "exceljs": "npm:@xlvane/compat@^1"
Node.js: require('exceljs') of @xlvane/compat needs Node.js 22.12 or later (package.json engines.node: not set).

Scanned 2 source files for exceljs API that @xlvane/compat does not fully support:
  not supported src/legacy.cjs:1:27  deep import (exceljs/lib/...)
  Details: docs/exceljs-compat.md. Unsupported calls throw XLV_E_NOT_SUPPORTED at run time.

Nothing was written. Run "xlvane migrate --force" to change package.json.
```

The same report as JSON (schema in [docs/cli-json.md](https://github.com/DevXord/xlvane/blob/main/docs/cli-json.md)):

```sh
npx xlvane migrate --json
```

## 2. Change package.json and install

```sh
npx xlvane migrate --force --dry-run   # show the change without writing it
npx xlvane migrate --force             # change package.json
npm install                            # install @xlvane/compat under the name exceljs
```

`--force` replaces only the value of each `"exceljs"` entry; the rest of `package.json` keeps its formatting and key order. Instead of `xlvane migrate`, the alias can also be installed directly:

```sh
npm install exceljs@npm:@xlvane/compat@^1
```

npm then saves the version it installed as the range (with npm 10.9.9 and `@xlvane/compat` 1.2.2: `"exceljs": "npm:@xlvane/compat@^1.2.2"`).

The code stays as it was. For example, this exceljs code runs unchanged on `@xlvane/compat`:

```ts
import ExcelJS from 'exceljs';

const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet('Sales');
sheet.columns = [
  { header: 'Region', key: 'region', width: 15 },
  { header: 'Revenue', key: 'revenue', width: 12 },
];
sheet.addRow({ region: 'North', revenue: 1200 });
await workbook.xlsx.writeFile('sales.xlsx');

const check = new ExcelJS.Workbook();
await check.xlsx.readFile('sales.xlsx');
console.log(check.getWorksheet('Sales')?.getCell('B2').value); // 1200
```

In CommonJS, `const ExcelJS = require('exceljs');` gives the same object.

## 3. Check the result

- **Run the project's tests.** Unsupported API throws an `XlvError` whose `code` is `XLV_E_NOT_SUPPORTED` and whose message names the method; nothing is dropped silently.
- **Look at each finding of `xlvane migrate`.** Deep imports of `exceljs/lib/...` do not resolve (`exceljs/dist/es5` and the `exceljs/dist/exceljs*.js` files do); `ModelContainer` throws `XLV_E_NOT_SUPPORTED` when it is constructed.
- **Read the differences.** [docs/exceljs-compat.md](https://github.com/DevXord/xlvane/blob/main/docs/exceljs-compat.md) lists every method with its status and the exceljs source of each behaviour, and ends with the differences found in reviews and cross-checks. The ones most likely to matter are summarised in the [@xlvane/compat README](https://github.com/DevXord/xlvane/blob/main/packages/compat/README.md#differences-from-exceljs): CSV output escapes text that starts like a formula (`formulaInjection: "off"` writes it unchanged, as exceljs does), a sheet id outside 1 to 65534 is written as the lowest unused id, only the first view of a worksheet is written, the module object is frozen, the streaming writer compresses in JavaScript, and input files are checked against ZIP and XML limits.
- **TypeScript.** The type names of exceljs's `index.d.ts` are exported and work with `moduleResolution` `bundler` and `nodenext` and with `module: commonjs`. Under `module: node16`, TypeScript refuses `import ExcelJS = require('exceljs')` of an ES module (TS1471); use `nodenext` or `commonjs`.
- **Browsers and Workers.** The methods that read or write files (`xlsx.readFile`, `xlsx.writeFile`, `csv.readFile`, `csv.writeFile`, the `filename` forms of the streaming classes and of `addImage`) work in Node.js, Bun and Deno; in browsers and Workers they throw `XLV_E_NOT_SUPPORTED`.

To handle an unsupported call in code:

```ts
import ExcelJS from 'exceljs';

const workbook = new ExcelJS.Workbook();
try {
  await workbook.xlsx.readFile('report.xlsx');
} catch (error) {
  const code = (error as { code?: unknown }).code;
  if (code === 'XLV_E_NOT_SUPPORTED') console.error('Not supported by @xlvane/compat:', (error as Error).message);
  else throw error;
}
```

## 4. Roll back

Install exceljs again under its own name, with the range the project used before (here `^4.4.0`):

```sh
npm install exceljs@^4.4.0
```

npm replaces the alias in `package.json` and `node_modules`. Restoring `package.json` and the lockfile from version control and running `npm install` has the same effect.

## Reporting a problem

When exceljs code behaves differently on `@xlvane/compat` and the difference is not listed in docs/exceljs-compat.md, please open an issue at https://github.com/DevXord/xlvane/issues with a small script that shows it.
