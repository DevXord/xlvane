// Arguments of `xlvane-mcp` (runCli), with the server replaced by a stub.
import { describe, expect, it } from 'vitest';
import { runWith } from '../../packages/mcp/src/server';
import { withData } from './helpers';

async function run(args: string[], cwd: string): Promise<{ code: number; stdout: string; stderr: string; served: boolean }> {
  let stdout = '';
  let stderr = '';
  let served = false;
  const code = await runWith(args, {
    stdout: (t) => (stdout += t),
    stderr: (t) => (stderr += t),
    cwd,
    serve: async (factory) => {
      factory();
      served = true;
    },
  });
  return { code, stdout, stderr, served };
}

describe('xlvane-mcp arguments', () => {
  it('requires at least one --root', async () => {
    const r = await run([], '/');
    expect(r).toMatchObject({ code: 2, served: false, stdout: '' });
    expect(r.stderr).toContain('xlvane-mcp --root ./data');
  });

  it('refuses a missing directory and unknown options', async () => {
    expect(await run(['--root', '/definitely/missing'], '/')).toMatchObject({ code: 3, served: false });
    expect(await run(['--root', '.', '--bogus'], '/')).toMatchObject({ code: 2, served: false });
    expect(await run(['--root', '.', '--max-rows', '0'], '/')).toMatchObject({ code: 2, served: false });
  });

  it('prints help to stdout without serving', async () => {
    const r = await run(['--help'], '/');
    expect(r).toMatchObject({ code: 0, served: false });
    expect(r.stdout).toContain('--allow-write');
  });

  it('serves with relative roots resolved against the working directory; messages go to stderr', async () => {
    await withData(async (dir) => {
      const r = await run(['--root', '.', '--allow-write'], dir);
      expect(r).toMatchObject({ code: 0, served: true, stdout: '' });
      expect(r.stderr).toContain('(writing enabled)');
    });
  });
});
