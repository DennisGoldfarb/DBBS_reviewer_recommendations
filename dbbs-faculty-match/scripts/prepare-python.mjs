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

function removePythonBytecode(rootDir) {
  const entries = fs.readdirSync(rootDir, { withFileTypes: true });
  for (const entry of entries) {
    const entryPath = path.join(rootDir, entry.name);
    if (entry.isDirectory()) {
      if (entry.name === '__pycache__') {
        fs.rmSync(entryPath, { recursive: true, force: true });
      } else {
        removePythonBytecode(entryPath);
      }
    } else if (entry.isFile() && entry.name.endsWith('.pyc')) {
      fs.rmSync(entryPath, { force: true });
    }
  }
}

function removeDirectoryIfExists(targetPath) {
  if (fs.existsSync(targetPath)) {
    fs.rmSync(targetPath, { recursive: true, force: true });
  }
}

function pruneBundledRuntime(rootDir) {
  const sitePackagesCandidates = [
    path.join(rootDir, 'Lib', 'site-packages'),
    path.join(rootDir, 'lib', 'python3.11', 'site-packages')
  ];

  for (const sitePackages of sitePackagesCandidates) {
    if (!fs.existsSync(sitePackages)) {
      continue;
    }

    const torchDir = path.join(sitePackages, 'torch');
    if (fs.existsSync(torchDir)) {
      const includeDir = path.join(torchDir, 'include');
      if (fs.existsSync(includeDir)) {
        console.log(
          '\u2139\ufe0f Removing torch C++ headers from bundled runtime to keep Windows paths short.'
        );
        removeDirectoryIfExists(includeDir);
      }
    }
  }
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

if (!reuseExisting) {
  ensureDirectory(pythonRootDir);
  if (fs.existsSync(runtimeDir)) {
    fs.rmSync(runtimeDir, { recursive: true, force: true });
  }

  const python = findPythonCandidate();
  if (!python) {
    fail(
      'Unable to locate a Python 3.11 interpreter to create the bundled runtime. Install Python 3.11 and ensure it is on your PATH.'
    );
  }

  console.log(`\u2139\ufe0f Using ${python.command} to create bundled Python runtime at ${runtimeDir}.`);

  runCommand(python.command, [...python.args, '-m', 'venv', runtimeDir]);

  const runtimePython = resolveRuntimePython(runtimeDir);
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
}

if (fs.existsSync(runtimeDir)) {
  console.log('\u2139\ufe0f Removing Python bytecode caches from bundled runtime to avoid long Windows paths.');
  removePythonBytecode(runtimeDir);
  pruneBundledRuntime(runtimeDir);
}
