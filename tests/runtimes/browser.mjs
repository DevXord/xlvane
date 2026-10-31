// Runs the smoke bundle in Chromium through Playwright (D-904). The bundle is served by a small
// static HTTP server on 127.0.0.1, imported by a page as an ES module, and its results are read
// with page.evaluate (Playwright API: chromium.launch, browser.newPage, page.goto, page.evaluate).
import { readFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { chromium } from 'playwright';

const PAGE = '<!doctype html><meta charset="utf-8"><title>xlvane smoke</title>';

/** Returns the smoke results from Chromium. */
export async function runInBrowser(bundle) {
  const code = await readFile(bundle);
  const server = createServer((req, res) => {
    if (req.url === '/smoke.js') {
      res.writeHead(200, { 'content-type': 'text/javascript; charset=utf-8' });
      res.end(code);
    } else {
      res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' });
      res.end(PAGE);
    }
  });
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  const { port } = server.address();
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    const errors = [];
    page.on('pageerror', (error) => errors.push(String(error)));
    await page.goto(`http://127.0.0.1:${port}/`);
    const results = await page.evaluate(async () => {
      const { runSmoke } = await import('/smoke.js');
      return runSmoke();
    });
    if (errors.length > 0) results.push({ name: 'page errors', ok: false, error: errors.join('; ') });
    // The page must not see Node.js globals: the bundle runs on web APIs only.
    const globals = await page.evaluate(() => ({ process: typeof globalThis.process, Buffer: typeof globalThis.Buffer }));
    results.push({ name: 'no Node.js globals in the page', ok: globals.process === 'undefined' && globals.Buffer === 'undefined', error: JSON.stringify(globals) });
    return results.map((r) => (r.ok ? { name: r.name, ok: true } : r));
  } finally {
    await browser.close();
    server.close();
  }
}
