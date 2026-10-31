// Runs the smoke bundle in workerd, the open-source Cloudflare Workers runtime (D-904). A config
// in the Cap'n Proto text format declares one module worker (the bundle plus a small entry
// module that answers fetch with the results) and one HTTP socket; the format follows
// samples/helloworld_esm/config.capnp in github.com/cloudflare/workerd. `workerd serve` runs it.
import { spawn } from 'node:child_process';
import { copyFile, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { createServer } from 'node:net';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../..', import.meta.url));
/** Must not be later than the release date of the workerd binary (1.20260730.1). */
export const COMPATIBILITY_DATE = '2026-07-01';

const ENTRY = `import { runSmoke } from './smoke.js';
export default {
  async fetch() {
    const results = await runSmoke();
    results.push({ name: 'no Node.js globals in the worker', ok: typeof globalThis.process === 'undefined' && typeof globalThis.Buffer === 'undefined' });
    return Response.json(results);
  },
};
`;

function freePort() {
  return new Promise((resolve, reject) => {
    const s = createServer();
    s.once('error', reject);
    s.listen(0, '127.0.0.1', () => {
      const { port } = s.address();
      s.close(() => resolve(port));
    });
  });
}

/** Returns the smoke results from workerd. */
export async function runInWorkerd(bundle) {
  const dir = await mkdtemp(join(tmpdir(), 'xlvane-workerd-'));
  const port = await freePort();
  await copyFile(bundle, join(dir, 'smoke.js'));
  await writeFile(join(dir, 'entry.js'), ENTRY);
  await writeFile(
    join(dir, 'config.capnp'),
    `using Workerd = import "/workerd/workerd.capnp";

const config :Workerd.Config = (
  services = [ (name = "main", worker = .smokeWorker) ],
  sockets = [ (name = "http", address = "127.0.0.1:${port}", http = (), service = "main") ]
);

const smokeWorker :Workerd.Worker = (
  modules = [
    (name = "entry.js", esModule = embed "entry.js"),
    (name = "smoke.js", esModule = embed "smoke.js")
  ],
  compatibilityDate = "${COMPATIBILITY_DATE}",
);
`,
  );
  const bin = join(root, 'node_modules/.bin/workerd');
  const child = spawn(bin, ['serve', join(dir, 'config.capnp'), 'config'], { stdio: ['ignore', 'ignore', 'pipe'] });
  let stderr = '';
  child.stderr.on('data', (d) => (stderr += d));
  try {
    const deadline = Date.now() + 30_000;
    for (;;) {
      if (child.exitCode !== null) throw new Error(`workerd exited with ${child.exitCode}: ${stderr}`);
      try {
        const res = await fetch(`http://127.0.0.1:${port}/`);
        if (!res.ok) throw new Error(`HTTP ${res.status}: ${await res.text()}\n${stderr}`);
        return await res.json();
      } catch (error) {
        if (error instanceof Error && error.message.startsWith('HTTP')) throw error;
        if (Date.now() > deadline) throw new Error(`workerd did not answer: ${String(error)}\n${stderr}`);
        await new Promise((r) => setTimeout(r, 200));
      }
    }
  } finally {
    child.kill();
    await rm(dir, { recursive: true, force: true });
  }
}
