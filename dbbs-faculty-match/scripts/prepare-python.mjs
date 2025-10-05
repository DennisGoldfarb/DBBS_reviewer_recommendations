import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { pipeline } from 'node:stream/promises';
import { Readable } from 'node:stream';
import AdmZip from 'adm-zip';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, '..');
const srcTauriDir = path.join(projectRoot, 'src-tauri');
const resourcesDir = path.join(srcTauriDir, 'resources');
const requirementsPath = path.join(projectRoot, 'python', 'requirements.txt');
const pythonAppSource = path.join(projectRoot, 'python', 'app');

const PYTHON_MAJOR = 3;
const PYTHON_MINOR = 11;
const PYTHON_TAG = `${PYTHON_MAJOR}${PYTHON_MINOR}`;
const EMBED_VERSION = '3.11.9';

const platformMap = {
  win32: 'windows',
  darwin: 'macos',
  linux: 'linux'
};

const archMap = {
  x64: 'x86_64',
  arm64: 'aarch64'
};

const embedDistributions = {
  x86_64: {
    url: `https://www.python.org/ftp/python/${EMBED_VERSION}/python-${EMBED_VERSION}-embed-amd64.zip`,
    archiveName: `python-${EMBED_VERSION}-embed-amd64.zip`,
    sha256: null
  },
  aarch64: {
    url: `https://www.python.org/ftp/python/${EMBED_VERSION}/python-${EMBED_VERSION}-embed-arm64.zip`,
    archiveName: `python-${EMBED_VERSION}-embed-arm64.zip`,
    sha256: null
  }
};

function fail(message) {
  console.error(`\u274c ${message}`);
  process.exit(1);
}

function ensureDirectory(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function removeDirectoryIfExists(targetPath) {
  if (fs.existsSync(targetPath)) {
    fs.rmSync(targetPath, { recursive: true, force: true });
  }
}

function hashFile(filePath) {
  const hash = createHash('sha256');
  hash.update(fs.readFileSync(filePath));
  return hash.digest('hex');
}

function removePythonBytecode(rootDir) {
  if (!fs.existsSync(rootDir)) {
    return;
  }

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

function pruneTorchHeaders(sitePackagesRoots) {
  for (const sitePackages of sitePackagesRoots) {
    if (!sitePackages) {
      continue;
    }

    if (!fs.existsSync(sitePackages)) {
      continue;
    }

    const torchDir = path.join(sitePackages, 'torch');
    if (!fs.existsSync(torchDir)) {
      continue;
    }

    const includeDir = path.join(torchDir, 'include');
    if (fs.existsSync(includeDir)) {
      console.log(
        '\u2139\ufe0f Removing torch C++ headers from bundled runtime to keep Windows paths short.'
      );
      removeDirectoryIfExists(includeDir);
    }
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
  const versionCheckSnippet = 'import sys; print(sys.version_info[0] == 3 and sys.version_info[1] == 11)';
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
      [...initialArgs, '-c', versionCheckSnippet],
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
    ? [
        path.join(runtimePath, 'python.exe'),
        path.join(runtimePath, 'python'),
        path.join(runtimePath, 'Scripts', 'python.exe'),
        path.join(runtimePath, 'Scripts', 'python')
      ]
    : [path.join(runtimePath, 'bin', 'python3'), path.join(runtimePath, 'bin', 'python')];

  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return null;
}

async function downloadFile(url, destination) {
  if (typeof fetch !== 'function') {
    throw new Error('Global fetch API is unavailable. Use Node.js 18 or newer to run prepare-python.');
  }
  console.log(`\u2139\ufe0f Downloading ${url}`);
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Failed to download ${url}: ${response.status} ${response.statusText}`);
  }

  const body = response.body;
  if (!body) {
    throw new Error(`Response body for ${url} was empty.`);
  }

  await pipeline(Readable.fromWeb(body), fs.createWriteStream(destination));
}

function verifySha256(filePath, expected) {
  if (!expected) {
    return;
  }

  const actual = hashFile(filePath);
  if (actual.toLowerCase() !== expected.toLowerCase()) {
    throw new Error(
      `Checksum mismatch for ${path.basename(filePath)}. Expected ${expected}, got ${actual}.`
    );
  }
}

function extractZipArchive(archivePath, destination) {
  const zip = new AdmZip(archivePath);
  zip.extractAllTo(destination, true);
}

function copyDirectoryRecursive(source, destination) {
  const entries = fs.readdirSync(source, { withFileTypes: true });
  for (const entry of entries) {
    const sourcePath = path.join(source, entry.name);
    const destPath = path.join(destination, entry.name);

    if (entry.isDirectory()) {
      ensureDirectory(destPath);
      copyDirectoryRecursive(sourcePath, destPath);
    } else if (entry.isSymbolicLink()) {
      const target = fs.readlinkSync(sourcePath);
      try {
        fs.symlinkSync(target, destPath);
      } catch (error) {
        fs.copyFileSync(sourcePath, destPath);
      }
    } else if (entry.isFile()) {
      fs.copyFileSync(sourcePath, destPath);
    }
  }
}

function syncPythonAppSources(destination) {
  removeDirectoryIfExists(destination);
  ensureDirectory(destination);

  if (!fs.existsSync(pythonAppSource)) {
    console.log(
      '\u2139\ufe0f No python/app directory found; creating an empty application folder in the embedded runtime.'
    );
    return;
  }

  console.log('\u2139\ufe0f Copying python/app sources into the embedded runtime.');
  copyDirectoryRecursive(pythonAppSource, destination);
}

function readMetadata(metadataPath) {
  if (!fs.existsSync(metadataPath)) {
    return null;
  }

  try {
    return JSON.parse(fs.readFileSync(metadataPath, 'utf8'));
  } catch (error) {
    console.warn(`\u26a0\ufe0f Failed to parse runtime metadata at ${metadataPath}:`, error);
    return null;
  }
}

function writeMetadata(metadataPath, metadata) {
  fs.writeFileSync(metadataPath, JSON.stringify(metadata, null, 2));
}

function finalizeRuntime(runtimeDir, sitePackagesRoots) {
  console.log('\u2139\ufe0f Removing Python bytecode caches from bundled runtime.');
  removePythonBytecode(runtimeDir);
  pruneTorchHeaders(sitePackagesRoots);
}

async function prepareWindowsEmbedRuntime(options) {
  const {
    runtimeDir,
    metadataPath,
    requirementsHash,
    arch,
    runtimeDirName,
    pythonRootDir
  } = options;

  const embed = embedDistributions[arch];
  if (!embed) {
    fail(`No embedded CPython distribution configured for Windows ${arch}.`);
  }

  const metadata = readMetadata(metadataPath);
  const pythonExecutable = path.join(runtimeDir, 'python.exe');
  const runtimePresent = fs.existsSync(runtimeDir) && fs.existsSync(pythonExecutable);
  const reuseExisting =
    runtimePresent &&
    metadata?.distribution === 'cpython-embed' &&
    metadata?.requirementsHash === requirementsHash &&
    metadata?.embedVersion === EMBED_VERSION;

  const sitePackagesDir = path.join(runtimeDir, 'site-packages');
  const runtimeAppDir = path.join(runtimeDir, 'app');
  const pthPath = path.join(runtimeDir, `python${PYTHON_TAG}._pth`);

  if (!reuseExisting) {
    console.log(`\u2139\ufe0f Building embedded CPython runtime for ${runtimeDirName}.`);
    removeDirectoryIfExists(runtimeDir);
    ensureDirectory(runtimeDir);

    const downloadPath = path.join(pythonRootDir, embed.archiveName);
    await downloadFile(embed.url, downloadPath);
    verifySha256(downloadPath, embed.sha256);
    extractZipArchive(downloadPath, runtimeDir);
    fs.rmSync(downloadPath, { force: true });

    ensureDirectory(sitePackagesDir);
  } else {
    console.log(`\u2705 Bundled Python runtime for ${runtimeDirName} is up to date.`);
    ensureDirectory(sitePackagesDir);
  }

  const python = findPythonCandidate();
  if (!python) {
    fail('Unable to locate a Python 3.11 interpreter to prepare the embedded runtime.');
  }

  if (!reuseExisting) {
    console.log('\u2139\ufe0f Installing Python packages into the embedded runtime.');
    removeDirectoryIfExists(sitePackagesDir);
    ensureDirectory(sitePackagesDir);
    runCommand(python.command, [
      ...python.args,
      '-m',
      'pip',
      'install',
      '--no-compile',
      '--upgrade',
      '--target',
      sitePackagesDir,
      '-r',
      requirementsPath
    ]);
  }

  const pthContents = [
    `python${PYTHON_TAG}.zip`,
    '.\\Lib',
    '.\\DLLs',
    '.\\site-packages',
    '.\\app',
    'import site'
  ].join('\r\n');
  fs.writeFileSync(pthPath, `${pthContents}\r\n`);

  syncPythonAppSources(runtimeAppDir);
  finalizeRuntime(runtimeDir, [sitePackagesDir]);

  const updatedMetadata = {
    distribution: 'cpython-embed',
    embedVersion: EMBED_VERSION,
    requirementsHash,
    pythonVersion: EMBED_VERSION,
    platform: 'windows',
    arch,
    updatedAt: new Date().toISOString()
  };
  writeMetadata(metadataPath, updatedMetadata);
}

function preparePosixVirtualEnv(options) {
  const {
    runtimeDir,
    metadataPath,
    requirementsHash,
    runtimeDirName,
    platform,
    arch
  } = options;

  const metadata = readMetadata(metadataPath);
  const runtimePythonExisting = resolveRuntimePython(runtimeDir);
  let reuseExisting = false;

  if (
    runtimePythonExisting &&
    metadata?.distribution === 'virtualenv' &&
    metadata?.requirementsHash === requirementsHash &&
    typeof metadata.pythonVersion === 'string' &&
    metadata.pythonVersion.startsWith(`${PYTHON_MAJOR}.${PYTHON_MINOR}.`)
  ) {
    reuseExisting = true;
    console.log(`\u2705 Bundled Python runtime for ${runtimeDirName} is up to date.`);
  } else if (metadata) {
    console.log(
      `\u2139\ufe0f Runtime metadata changed; rebuilding bundled Python runtime for ${runtimeDirName}.`
    );
  }

  if (!reuseExisting) {
    removeDirectoryIfExists(runtimeDir);
    ensureDirectory(runtimeDir);

    const python = findPythonCandidate();
    if (!python) {
      fail('Unable to locate a Python 3.11 interpreter to create the bundled runtime.');
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

    const metadataToWrite = {
      distribution: 'virtualenv',
      requirementsHash,
      createdAt: new Date().toISOString(),
      platform,
      arch,
      pythonVersion: versionResult.stdout?.toString().trim()
    };
    writeMetadata(metadataPath, metadataToWrite);
  }

  const sitePackagesRoots = [
    path.join(runtimeDir, 'Lib', 'site-packages'),
    path.join(runtimeDir, 'lib', `python${PYTHON_MAJOR}.${PYTHON_MINOR}`, 'site-packages')
  ];
  finalizeRuntime(runtimeDir, sitePackagesRoots);
}

async function main() {
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

  ensureDirectory(resourcesDir);
  const pythonRootDir = path.join(resourcesDir, 'python');
  ensureDirectory(pythonRootDir);

  const runtimeDirName = `${platform}-${arch}`;
  const runtimeDir = path.join(pythonRootDir, runtimeDirName);
  const metadataPath = path.join(runtimeDir, '.bundle-metadata.json');

  const requirementsHash = hashFile(requirementsPath);

  if (platform === 'windows') {
    await prepareWindowsEmbedRuntime({
      runtimeDir,
      metadataPath,
      requirementsHash,
      arch,
      runtimeDirName,
      pythonRootDir
    });
  } else {
    preparePosixVirtualEnv({
      runtimeDir,
      metadataPath,
      requirementsHash,
      runtimeDirName,
      platform,
      arch
    });
  }
}

main().catch((error) => {
  console.error('\u274c Failed to prepare embedded Python runtime:', error);
  process.exit(1);
});
