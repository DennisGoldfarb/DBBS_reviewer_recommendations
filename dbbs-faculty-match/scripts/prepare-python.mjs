import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import https from 'node:https';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import AdmZip from 'adm-zip';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, '..');
const srcTauriDir = path.join(projectRoot, 'src-tauri');
const resourcesDir = path.join(srcTauriDir, 'resources');
const pythonRootDir = path.join(resourcesDir, 'python');
const requirementsPath = path.join(projectRoot, 'python', 'requirements.txt');
const vendorRoot = path.join(projectRoot, 'python', 'vendor');
const appSourceDir = path.join(projectRoot, 'python', 'app');

const PYTHON_VERSION = '3.11.9';
const EMBEDDABLE_FILENAME = `python-${PYTHON_VERSION}-embed-amd64.zip`;
const EMBEDDABLE_URL = `https://www.python.org/ftp/python/${PYTHON_VERSION}/${EMBEDDABLE_FILENAME}`;

const platformMap = {
  win32: 'windows',
  darwin: 'macos',
  linux: 'linux'
};

const archMap = {
  x64: 'x86_64',
  arm64: 'aarch64'
};

function fail(message, error) {
  if (error) {
    console.error(error);
  }
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

function copyDirectoryContents(source, destination) {
  if (!fs.existsSync(source)) {
    return false;
  }

  ensureDirectory(destination);
  const entries = fs.readdirSync(source, { withFileTypes: true });
  for (const entry of entries) {
    const sourcePath = path.join(source, entry.name);
    const destinationPath = path.join(destination, entry.name);
    if (entry.isDirectory()) {
      copyDirectoryContents(sourcePath, destinationPath);
    } else if (entry.isFile()) {
      fs.copyFileSync(sourcePath, destinationPath);
    } else if (entry.isSymbolicLink()) {
      const target = fs.readlinkSync(sourcePath);
      try {
        fs.symlinkSync(target, destinationPath);
      } catch (error) {
        if (error.code === 'EEXIST') {
          fs.rmSync(destinationPath);
          fs.symlinkSync(target, destinationPath);
        } else {
          throw error;
        }
      }
    }
  }

  return true;
}

function pruneTorchHeaders(sitePackagesDir) {
  const torchDir = path.join(sitePackagesDir, 'torch');
  if (!fs.existsSync(torchDir)) {
    return;
  }

  const includeDir = path.join(torchDir, 'include');
  if (fs.existsSync(includeDir)) {
    console.log('\u2139\ufe0f Removing torch C++ headers from bundled runtime to keep Windows paths short.');
    removeDirectoryIfExists(includeDir);
  }
}

function writeMetadata(runtimeDir, metadata) {
  const metadataPath = path.join(runtimeDir, '.bundle-metadata.json');
  fs.writeFileSync(metadataPath, JSON.stringify(metadata, null, 2));
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

function pythonIdentifier() {
  const major = PYTHON_VERSION.split('.')[0];
  const minor = PYTHON_VERSION.split('.')[1];
  return `${major}${minor}`;
}

function configureEmbeddablePaths(runtimeDir) {
  const identifier = pythonIdentifier();
  const pthPath = path.join(runtimeDir, `python${identifier}._pth`);
  const lines = [
    `python${identifier}.zip`,
    '.\\Lib',
    '.\\DLLs',
    '.\\site-packages',
    '.\\app',
    'import site'
  ];
  fs.writeFileSync(pthPath, lines.join('\r\n'));
}

function readDirectoryOrNull(dir) {
  if (!fs.existsSync(dir)) {
    return null;
  }
  return fs.readdirSync(dir);
}

async function downloadFile(url, destination) {
  await new Promise((resolve, reject) => {
    const request = https.get(url, response => {
      if (response.statusCode && response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
        const redirectedUrl = new URL(response.headers.location, url).toString();
        downloadFile(redirectedUrl, destination).then(resolve).catch(reject);
        return;
      }

      if (response.statusCode !== 200) {
        reject(new Error(`Download failed with status ${response.statusCode}`));
        return;
      }

      ensureDirectory(path.dirname(destination));
      const fileStream = fs.createWriteStream(destination);
      response.pipe(fileStream);
      fileStream.on('finish', () => {
        fileStream.close(resolve);
      });
      fileStream.on('error', reject);
    });

    request.on('error', reject);
  });
}

async function resolveEmbeddableArchive() {
  const windowsVendorDir = path.join(vendorRoot, 'windows-x86_64');
  const vendoredArchive = path.join(windowsVendorDir, EMBEDDABLE_FILENAME);
  if (fs.existsSync(vendoredArchive)) {
    return vendoredArchive;
  }

  const downloadDir = path.join(projectRoot, 'python', '.downloads');
  const downloadPath = path.join(downloadDir, EMBEDDABLE_FILENAME);
  if (fs.existsSync(downloadPath)) {
    return downloadPath;
  }

  console.log(`\u2139\ufe0f Downloading Windows embeddable Python ${PYTHON_VERSION}...`);
  try {
    await downloadFile(EMBEDDABLE_URL, downloadPath);
    return downloadPath;
  } catch (error) {
    console.warn('\u26a0\ufe0f Failed to download the embeddable distribution automatically.');
    throw error;
  }
}

async function prepareWindowsEmbeddable() {
  console.log('\u2139\ufe0f Preparing Windows embeddable Python runtime.');

  const requirementsHash = hashFile(requirementsPath);
  const sitePackagesSource = path.join(vendorRoot, 'windows-x86_64', 'site-packages');
  if (!fs.existsSync(sitePackagesSource)) {
    fail(
      'Missing vendored Windows site-packages at python/vendor/windows-x86_64/site-packages. Populate this directory with the dependencies built for Python 3.11 x86_64.'
    );
  }

  const appEntries = readDirectoryOrNull(appSourceDir);
  if (!appEntries || appEntries.length === 0) {
    console.log('\u26a0\ufe0f No Python application files found under python/app. A placeholder directory will be bundled.');
  }

  removeDirectoryIfExists(pythonRootDir);
  ensureDirectory(pythonRootDir);

  let archivePath;
  try {
    archivePath = await resolveEmbeddableArchive();
    console.log(`\u2139\ufe0f Using embeddable archive: ${archivePath}`);
  } catch (error) {
    fail(
      `Unable to obtain the embeddable distribution automatically. Download ${EMBEDDABLE_FILENAME} from https://www.python.org/ftp/python/${PYTHON_VERSION}/ and place it in python/vendor/windows-x86_64 before retrying.`,
      error
    );
  }

  try {
    const zip = new AdmZip(archivePath);
    zip.extractAllTo(pythonRootDir, true);
  } catch (error) {
    fail('Unable to extract the embeddable Python archive.', error);
  }

  configureEmbeddablePaths(pythonRootDir);

  const sitePackagesDestination = path.join(pythonRootDir, 'site-packages');
  removeDirectoryIfExists(sitePackagesDestination);
  ensureDirectory(sitePackagesDestination);
  copyDirectoryContents(sitePackagesSource, sitePackagesDestination);

  const appDestination = path.join(pythonRootDir, 'app');
  removeDirectoryIfExists(appDestination);
  if (appEntries && appEntries.length > 0) {
    ensureDirectory(appDestination);
    copyDirectoryContents(appSourceDir, appDestination);
  } else {
    ensureDirectory(appDestination);
  }

  pruneTorchHeaders(sitePackagesDestination);
  removePythonBytecode(pythonRootDir);

  writeMetadata(pythonRootDir, {
    bundler: 'embeddable',
    platform: 'windows',
    arch: 'x86_64',
    pythonVersion: PYTHON_VERSION,
    requirementsHash,
    embeddableArchive: path.basename(archivePath),
    updatedAt: new Date().toISOString()
  });

  console.log('\u2705 Bundled Windows embeddable Python runtime is ready.');
}

function preparePosixVirtualEnv(platform, arch) {
  const requirementsHash = hashFile(requirementsPath);
  const runtimeDir = path.join(pythonRootDir, `${platform}-${arch}`);
  const metadataPath = path.join(runtimeDir, '.bundle-metadata.json');
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
        console.log(`\u2705 Bundled Python runtime for ${platform}-${arch} is up to date.`);
      } else {
        console.log(
          `\u2139\ufe0f Runtime metadata changed; rebuilding bundled Python runtime for ${platform}-${arch}.`
        );
      }
    } catch (error) {
      console.warn(
        '\u26a0\ufe0f Unable to parse existing runtime metadata; rebuilding bundled Python runtime.'
      );
    }
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
    console.log(`\u2705 Bundled Python runtime prepared for ${platform}-${arch}.`);
  }

  if (fs.existsSync(runtimeDir)) {
    console.log('\u2139\ufe0f Removing Python bytecode caches from bundled runtime to avoid long paths.');
    removePythonBytecode(runtimeDir);
    pruneTorchHeaders(path.join(runtimeDir, 'Lib', 'site-packages'));
    pruneTorchHeaders(path.join(runtimeDir, 'lib', 'python3.11', 'site-packages'));
  }
}

async function main() {
  if (!fs.existsSync(requirementsPath)) {
    fail(`Missing requirements file: ${requirementsPath}`);
  }

  ensureDirectory(resourcesDir);

  if (process.platform === 'win32') {
    await prepareWindowsEmbeddable();
    return;
  }

  const platform = platformMap[process.platform];
  if (!platform) {
    fail(`Unsupported platform for bundling Python runtime: ${process.platform}`);
  }

  const arch = archMap[process.arch];
  if (!arch) {
    fail(`Unsupported architecture for bundling Python runtime: ${process.arch}`);
  }

  preparePosixVirtualEnv(platform, arch);
}

main().catch(error => {
  fail('An unexpected error occurred while preparing the Python runtime.', error);
});
