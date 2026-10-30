// Runs the smoke bundle in the current runtime (Bun, Deno) and prints the results as one JSON
// line; the exit code is 1 when a check failed. Uses only standard ECMAScript plus `import()`.
// Usage: bun run tests/runtimes/run-bundle.mjs <smoke.js>
//        deno run --allow-read=<dir> --no-prompt tests/runtimes/run-bundle.mjs <smoke.js>
const args = typeof Deno !== 'undefined' ? Deno.args : process.argv.slice(2);
const bundle = args[0];
if (!bundle) throw new Error('Pass the path of smoke.js');
const url = new URL(bundle.startsWith('/') ? `file://${bundle}` : bundle, `file://${typeof Deno !== 'undefined' ? Deno.cwd() : process.cwd()}/`);
const { runSmoke } = await import(url.href);
const results = await runSmoke();
console.log(JSON.stringify(results));
const failed = results.some((r) => !r.ok);
if (typeof Deno !== 'undefined') Deno.exit(failed ? 1 : 0);
else process.exitCode = failed ? 1 : 0;
