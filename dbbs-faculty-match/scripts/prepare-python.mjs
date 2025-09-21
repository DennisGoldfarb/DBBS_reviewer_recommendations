import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, '..');
const srcTauriDir = path.join(projectRoot, 'src-tauri');
const resourcesDir = path.join(srcTauriDir, 'resources');
const requirementsPath = path.join(projectRoot, 'python', 'requirements.txt');

const platformMap = {
  win32: 'windows',
  darwin: 'macos',
  linux: 'linux'
};

const archMap = {
  x64: 'x86_64',
  arm64: 'aarch64'
};

function fail(message) {
  console.error(`\u274c ${message}`);
  process.exit(1);
}

if (!fs.existsSync(requirementsPath)) {
  fail(`Missing requirements file: ${requirementsPath}`);
}

const platform = platformMap[process.platform];
if (!platform) {
  fail(`Unsupported platform for bundling Python runtime: ${process.platform}`);
}

const arch = archMap[process.arch];
if (!arch) {
  fail(`Unsupported architecture for bundling Python runtime: ${process.arch}`);
}

const runtimeDirName = `${platform}-${arch}`;
const pythonRootDir = path.join(resourcesDir, 'python');
const runtimeDir = path.join(pythonRootDir, runtimeDirName);
const metadataPath = path.join(runtimeDir, '.bundle-metadata.json');

function ensureDirectory(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function hashFile(filePath) {
  const hash = createHash('sha256');
  hash.update(fs.readFileSync(filePath));
  return hash.digest('hex');
}

const requirementsHash = hashFile(requirementsPath);
let reuseExisting = false;

if (fs.existsSync(metadataPath)) {
  try {
    const metadata = JSON.parse(fs.readFileSync(metadataPath, 'utf8'));
    if (
      metadata.requirementsHash === requirementsHash &&
      typeof metadata.pythonVersion === 'string' &&
      metadata.pythonVersion.startsWith('3.11.')
    ) {
      reuseExisting = true;
      console.log(`\u2705 Bundled Python runtime for ${runtimeDirName} is up to date.`);
    } else {
      console.log(
        `\u2139\ufe0f Runtime metadata changed; rebuilding bundled Python runtime for ${runtimeDirName}.`
      );
    }
  } catch (error) {
    console.warn(`\u26a0\ufe0f Unable to parse ${metadataPath}; rebuilding runtime.`);
  }
}

function runCommand(command, args, options = {}) {
  const result = spawnSync(command, args, {
    stdio: 'inherit',
    env: {
      ...process.env,
      PIP_DISABLE_PIP_VERSION_CHECK: '1',
      ...options.env
    },
    shell: options.shell ?? false
  });
  if (result.error) {
    throw result.error;
  }
  if (result.status !== 0) {
    throw new Error(`${command} ${args.join(' ')} exited with status ${result.status}`);
  }
}

function findPythonCandidate() {
  const candidates = process.platform === 'win32'
    ? [
        ['py', '-3.11'],
        ['py', '-3'],
        ['python3.11'],
        ['python3'],
        ['python']
      ]
    : [
        ['python3.11'],
        ['python3'],
        ['python']
      ];

  for (const parts of candidates) {
    const [command, ...initialArgs] = parts;
    const versionCheck = spawnSync(
      command,
      [
        ...initialArgs,
        '-c',
        'import sys; print(sys.version_info[0] == 3 and sys.version_info[1] == 11)'
      ],
      {
        stdio: ['ignore', 'pipe', 'ignore']
      }
    );
    if (versionCheck.status === 0 && versionCheck.stdout?.toString().trim() === 'True') {
      return { command, args: initialArgs };
    }
  }

  return null;
}

function resolveRuntimePython(runtimePath) {
  const candidates = process.platform === 'win32'
    ? [path.join(runtimePath, 'Scripts', 'python.exe'), path.join(runtimePath, 'Scripts', 'python')]
    : [path.join(runtimePath, 'bin', 'python3'), path.join(runtimePath, 'bin', 'python')];

  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return null;
}

function resolveSitePackages(runtimePython) {
  const result = spawnSync(
    runtimePython,
    ['-c', 'import sysconfig; print(sysconfig.get_paths()["purelib"])'],
    {
      stdio: ['ignore', 'pipe', 'inherit']
    }
  );

  if (result.status !== 0) {
    fail('Unable to resolve the site-packages directory for the bundled runtime.');
  }

  const sitePackages = result.stdout?.toString().trim();
  if (!sitePackages) {
    fail('Bundled runtime did not report a valid site-packages directory.');
  }

  return sitePackages;
}

function pruneTorchPackage(sitePackagesDir) {
  const torchDir = path.join(sitePackagesDir, 'torch');
  if (!fs.existsSync(torchDir)) {
    return;
  }

  const removals = [];
  const dropDirs = [
    path.join(torchDir, 'include'),
    path.join(torchDir, 'share'),
    path.join(torchDir, 'lib', 'cmake'),
    path.join(torchDir, 'lib', 'pkgconfig')
  ];

  for (const dir of dropDirs) {
    if (fs.existsSync(dir)) {
      fs.rmSync(dir, { recursive: true, force: true });
      removals.push(path.relative(torchDir, dir) || path.basename(dir));
    }
  }

  const libDir = path.join(torchDir, 'lib');
  if (fs.existsSync(libDir)) {
    for (const entry of fs.readdirSync(libDir)) {
      const lower = entry.toLowerCase();
      if (lower.endsWith('.pdb') || lower.endsWith('.lib') || lower.endsWith('.exp') || lower.endsWith('.ilk')) {
        fs.rmSync(path.join(libDir, entry), { force: true });
        removals.push(path.join('lib', entry));
      }
    }
  }

  if (removals.length > 0) {
    console.log(
      `\u2139\ufe0f Removed development-only torch assets from bundled runtime: ${removals.join(', ')}.`
    );
  } else {
    console.log('\u2139\ufe0f No torch development assets needed pruning from bundled runtime.');
  }
}

ensureDirectory(pythonRootDir);

if (!reuseExisting && fs.existsSync(runtimeDir)) {
  fs.rmSync(runtimeDir, { recursive: true, force: true });
}

let runtimePython = resolveRuntimePython(runtimeDir);

if (!reuseExisting) {
  const python = findPythonCandidate();
  if (!python) {
    fail(
      'Unable to locate a Python 3.11 interpreter to create the bundled runtime. Install Python 3.11 and ensure it is on your PATH.'
    );
  }

  console.log(`\u2139\ufe0f Using ${python.command} to create bundled Python runtime at ${runtimeDir}.`);

  runCommand(python.command, [...python.args, '-m', 'venv', runtimeDir]);
  runtimePython = resolveRuntimePython(runtimeDir);
  if (!runtimePython) {
    fail(`Virtual environment creation succeeded but no interpreter was found in ${runtimeDir}.`);
  }

  console.log('\u2139\ufe0f Upgrading pip and installing dependencies for bundled Python runtime.');
  runCommand(runtimePython, ['-m', 'pip', 'install', '--upgrade', 'pip', 'setuptools', 'wheel']);
  runCommand(runtimePython, ['-m', 'pip', 'install', '--no-cache-dir', '-r', requirementsPath]);

  const versionResult = spawnSync(
    runtimePython,
    ['-c', 'import sys; print(".".join(map(str, sys.version_info[:3])))'],
    {
      stdio: ['ignore', 'pipe', 'inherit']
    }
  );
  if (versionResult.status !== 0) {
    fail('Unable to determine the Python version for the bundled runtime.');
  }

  const metadata = {
    requirementsHash,
    createdAt: new Date().toISOString(),
    platform,
    arch,
    pythonVersion: versionResult.stdout?.toString().trim()
  };
  fs.writeFileSync(metadataPath, JSON.stringify(metadata, null, 2));
  console.log(`\u2705 Bundled Python runtime prepared for ${runtimeDirName}.`);
} else if (!runtimePython) {
  fail(`Bundled runtime at ${runtimeDir} exists but no interpreter was found.`);
}

const sitePackagesDir = resolveSitePackages(runtimePython);
pruneTorchPackage(sitePackagesDir);
