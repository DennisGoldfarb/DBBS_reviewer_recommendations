import { mkdtemp, rm, symlink, access, mkdir, unlink, cp, readFile, readdir } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawn } from 'node:child_process';

async function exec(command, args, options = {}) {
  return await new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      stdio: 'inherit',
      ...options,
    });

    child.on('error', reject);
    child.on('exit', (code) => {
      if (code === 0) {
        resolve();
      } else {
        reject(new Error(`${command} exited with code ${code}`));
      }
    });
  });
}

function resolveArch() {
  if (process.env.TAURI_ENV_ARCH) {
    return process.env.TAURI_ENV_ARCH;
  }

  switch (process.arch) {
    case 'arm64':
      return 'aarch64';
    case 'x64':
      return 'x64';
    default:
      return process.arch;
  }
}

async function codesignAppBundle(appBundlePath) {
  const identity =
    process.env.APPLE_CODESIGN_IDENTITY ?? process.env.CODESIGN_IDENTITY;

  if (!identity) {
    console.warn(
      'Skipping codesign because APPLE_CODESIGN_IDENTITY was not provided.'
    );
    return;
  }

  const entitlements = process.env.APPLE_CODESIGN_ENTITLEMENTS;
  const macosDir = path.join(appBundlePath, 'Contents', 'MacOS');
  const binaries = await readdir(macosDir, { withFileTypes: true });

  console.log('Codesigning macOS executables with hardened runtime...');

  for (const entry of binaries) {
    if (!entry.isFile()) {
      continue;
    }

    const target = path.join(macosDir, entry.name);
    await exec('codesign', [
      '--force',
      '--options',
      'runtime',
      '--timestamp',
      '--sign',
      identity,
      target,
    ]);
  }

  const appArgs = [
    '--force',
    '--options',
    'runtime',
    '--timestamp',
    '--sign',
    identity,
  ];

  if (entitlements) {
    appArgs.push('--entitlements', entitlements);
  }

  appArgs.push(appBundlePath);

  await exec('codesign', appArgs);
  await exec('codesign', [
    '--verify',
    '--deep',
    '--strict',
    '--verbose=2',
    appBundlePath,
  ]);
}

export async function buildDmg({ debug }) {
  if (process.platform !== 'darwin') {
    return;
  }

  const __dirname = path.dirname(fileURLToPath(import.meta.url));
  const rootDir = path.resolve(__dirname, '..');
  const configPath = path.join(rootDir, 'src-tauri', 'tauri.conf.json');
  const config = JSON.parse(await readFile(configPath, 'utf8'));
  const productName = config.productName;
  const version = config.version;

  const bundleDir = path.join(
    rootDir,
    'src-tauri',
    'target',
    debug ? 'debug' : 'release',
    'bundle'
  );

  const appBundlePath = path.join(bundleDir, 'macos', `${productName}.app`);
  try {
    await access(appBundlePath);
  } catch (error) {
    console.warn(`Skipping DMG creation because app bundle was not found at ${appBundlePath}`);
    return;
  }

  await codesignAppBundle(appBundlePath);

  const dmgDir = path.join(bundleDir, 'dmg');
  await mkdir(dmgDir, { recursive: true });

  const dmgPath = path.join(dmgDir, `${productName}_${version}_${resolveArch()}.dmg`);
  try {
    await unlink(dmgPath);
  } catch (error) {
    if (error.code !== 'ENOENT') {
      throw error;
    }
  }

  const stagingDir = await mkdtemp(path.join(os.tmpdir(), 'dbbs-faculty-match-dmg-'));
  const stagedAppPath = path.join(stagingDir, `${productName}.app`);
  await cp(appBundlePath, stagedAppPath, { recursive: true });

  const applicationsLink = path.join(stagingDir, 'Applications');
  try {
    await symlink('/Applications', applicationsLink);
  } catch (error) {
    if (error.code !== 'EEXIST') {
      throw error;
    }
  }

  const hdiutilArgs = [
    'create',
    '-volname',
    productName,
    '-fs',
    'HFS+',
    '-srcfolder',
    stagingDir,
    '-ov',
    '-format',
    'UDZO',
    dmgPath,
  ];

  console.log('Creating DMG using hdiutil...');
  try {
    await exec('hdiutil', hdiutilArgs);
    console.log(`Created DMG at ${dmgPath}`);
  } finally {
    await rm(stagingDir, { recursive: true, force: true });
  }
}
