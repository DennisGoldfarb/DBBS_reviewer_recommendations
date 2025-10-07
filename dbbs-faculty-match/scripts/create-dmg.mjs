import { mkdtemp, rm, symlink, access, mkdir, cp, readFile } from 'node:fs/promises';
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

async function ensureApplicationsShortcut(applicationsLink, stagingDir) {
  await rm(applicationsLink, { force: true, recursive: true });

  const finderScriptLines = [
    'tell application "Finder"',
    `  set dmgFolder to POSIX file ${JSON.stringify(stagingDir)}`,
    '  try',
    '    delete every item of dmgFolder whose name is "Applications"',
    '  end try',
    '  set newAlias to make alias file at dmgFolder to POSIX file "/Applications"',
    '  set name of newAlias to "Applications"',
    'end tell',
  ];

  try {
    const finderArgs = finderScriptLines.flatMap((line) => ['-e', line]);
    await exec('osascript', finderArgs);
  } catch (finderError) {
    console.warn('Failed to create Finder alias for Applications shortcut. Falling back to symlink.', finderError);

    try {
      await symlink('/Applications', applicationsLink);
    } catch (symlinkError) {
      if (symlinkError.code !== 'EEXIST') {
        throw symlinkError;
      }
    }
  }
}

export async function buildDmg({ debug = false, dmgPath: providedDmgPath, volumeName } = {}) {
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

  const defaultDmgDir = path.join(bundleDir, 'dmg');
  const dmgPath = providedDmgPath
    ? path.resolve(providedDmgPath)
    : path.join(defaultDmgDir, `${productName}_${version}_${resolveArch()}.dmg`);

  await mkdir(path.dirname(dmgPath), { recursive: true });
  await rm(dmgPath, { force: true });

  const stagingDir = await mkdtemp(path.join(os.tmpdir(), 'dbbs-faculty-match-dmg-'));
  const stagedAppPath = path.join(stagingDir, `${productName}.app`);
  await cp(appBundlePath, stagedAppPath, { recursive: true });

  const applicationsLink = path.join(stagingDir, 'Applications');
  await ensureApplicationsShortcut(applicationsLink, stagingDir);

  const hdiutilArgs = [
    'create',
    '-volname',
    volumeName ?? productName,
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

  return dmgPath;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const cliArgs = process.argv.slice(2);
  let debug = false;
  let dmgPath;
  let volumeName;

  const requireValue = (name, currentIndex) => {
    const nextValue = cliArgs[currentIndex + 1];
    if (typeof nextValue === 'undefined') {
      console.error(`Missing value for ${name}`);
      process.exit(1);
    }

    return nextValue;
  };

  for (let index = 0; index < cliArgs.length; index += 1) {
    const arg = cliArgs[index];

    if (arg === '--debug') {
      debug = true;
      continue;
    }

    if (arg === '--release') {
      debug = false;
      continue;
    }

    if (arg.startsWith('--dmg-path=')) {
      dmgPath = arg.slice('--dmg-path='.length);
      continue;
    }

    if (arg === '--dmg-path') {
      dmgPath = requireValue('--dmg-path', index);
      index += 1;
      continue;
    }

    if (arg.startsWith('--volume-name=')) {
      volumeName = arg.slice('--volume-name='.length);
      continue;
    }

    if (arg === '--volume-name') {
      volumeName = requireValue('--volume-name', index);
      index += 1;
      continue;
    }

    console.error(`Unknown argument: ${arg}`);
    process.exit(1);
  }

  try {
    const resultPath = await buildDmg({ debug, dmgPath, volumeName });
    if (resultPath) {
      console.log(`DMG created at ${resultPath}`);
    }
  } catch (error) {
    console.error('Failed to create DMG:', error);
    process.exit(1);
  }
}
