#!/usr/bin/env node
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

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

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const tauriExecutable = path.join(
  __dirname,
  '..',
  'node_modules',
  '.bin',
  process.platform === 'win32' ? 'tauri.cmd' : 'tauri'
);

const child = spawn(tauriExecutable, forwardedArgs, {
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
