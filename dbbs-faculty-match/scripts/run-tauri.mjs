#!/usr/bin/env node
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { createRequire } from 'node:module';

const args = process.argv.slice(2);
const isBuildCommand = args[0] === 'build';
const isMac = process.platform === 'darwin';
const forwardedArgs = [...args];

if (
  isBuildCommand &&
  isMac &&
  !forwardedArgs.some((value, index) => value === '--bundles' || value.startsWith('--bundles='))
) {
  // Avoid the default DMG bundler which fails on macOS 13 by delegating to our own
  // post-processing step. We request only the app bundle here and rebuild the DMG later.
  forwardedArgs.push('--bundles', 'app');
}

const require = createRequire(import.meta.url);
const tauriCli = require.resolve('@tauri-apps/cli/tauri.js');

// Launch the JS entry point with Node on every platform. Windows .cmd shims
// cannot be spawned directly, and using a shell would alter argument handling.
const child = spawn(process.execPath, [tauriCli, ...forwardedArgs], {
  stdio: 'inherit',
  env: process.env,
});

const [code, signal] = await once(child, 'exit');

if (typeof code === 'number' ? code !== 0 : signal) {
  process.exit(code ?? 1);
}

if (isBuildCommand && isMac) {
  const { buildDmg } = await import('./create-dmg.mjs');
  await buildDmg({ debug: forwardedArgs.includes('--debug') });
}
