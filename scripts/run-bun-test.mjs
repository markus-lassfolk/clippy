#!/usr/bin/env node
/**
 * Run the Bun test CLI. Prefer a `bun` on PATH (CI / local installs); fall back to
 * the same Bun minor line as CI (see .github/workflows/ci.yml) via npx.
 */
import { spawnSync } from 'node:child_process';

const args = process.argv.slice(2);
if (args.length === 0) {
  console.error('usage: node scripts/run-bun-test.mjs <bun-args...>');
  process.exit(1);
}

function run(cmd, cmdArgs) {
  const r = spawnSync(cmd, cmdArgs, {
    stdio: 'inherit',
    env: { ...process.env, NODE_ENV: 'test' }
  });
  if (r.error?.code === 'ENOENT') return 'missing';
  if (r.error) {
    console.error(`Error spawning ${cmd}:`, r.error.message);
  }
  process.exit(r.status ?? 1);
}

const miss = run('bun', args);
if (miss === 'missing') {
  const pinned = ['--yes', 'bun@1.3.11', ...args];
  const r2 = spawnSync('npx', pinned, {
    stdio: 'inherit',
    shell: process.platform === 'win32',
    env: { ...process.env, NODE_ENV: 'test' }
  });
  process.exit(r2.status ?? 1);
}
