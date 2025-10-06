import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, '..');
const srcTauriDir = path.join(projectRoot, 'src-tauri');
const binariesDir = path.join(srcTauriDir, 'binaries');
const pythonDir = path.join(projectRoot, 'python');
const scriptPath = path.join(pythonDir, 'embedding_helper.py');
const requirementsPath = path.join(pythonDir, 'requirements.txt');
const buildDir = path.join(pythonDir, '.sidecar-build');
const venvDir = path.join(buildDir, 'venv');
const distDir = path.join(buildDir, 'dist');
const workDir = path.join(buildDir, 'work');
const specDir = path.join(buildDir, 'spec');
const metadataPath = path.join(buildDir, 'metadata.json');
const sidecarName = 'embedding-helper';

function fail(message) {
  console.error(`\u274c ${message}`);
  process.exit(1);
}

function ensureExists(filePath, description) {
  if (!fs.existsSync(filePath)) {
    fail(`${description} not found: ${filePath}`);
  }
}

ensureExists(scriptPath, 'Embedding helper script');
ensureExists(requirementsPath, 'Python requirements file');

function hashFile(filePath) {
  const hash = createHash('sha256');
  hash.update(fs.readFileSync(filePath));
  return hash.digest('hex');
}

const metadataSignature = {
  scriptHash: hashFile(scriptPath),
  requirementsHash: hashFile(requirementsPath),
};

function readMetadata() {
  if (!fs.existsSync(metadataPath)) {
    return null;
  }
  try {
    return JSON.parse(fs.readFileSync(metadataPath, 'utf8'));
  } catch (error) {
    console.warn(`\u26a0\ufe0f Unable to parse ${metadataPath}; rebuilding sidecar.`);
    return null;
  }
}

function writeMetadata(data) {
  fs.mkdirSync(buildDir, { recursive: true });
  fs.writeFileSync(metadataPath, JSON.stringify(data, null, 2));
}

function runCommand(command, args, options = {}) {
  const result = spawnSync(command, args, {
    stdio: options.stdio || 'inherit',
    cwd: options.cwd || projectRoot,
    env: {
      ...process.env,
      PIP_DISABLE_PIP_VERSION_CHECK: '1',
      PYTHONUTF8: '1',
      ...(options.env || {}),
    },
    shell: options.shell ?? false,
  });

  if (result.error) {
    throw result.error;
  }
  if (result.status !== 0) {
    throw new Error(`${command} ${args.join(' ')} exited with status ${result.status}`);
  }

  return result.stdout ? result.stdout.toString() : '';
}

function getTargetTriple() {
  const output = runCommand('rustc', ['-vV'], { stdio: 'pipe' });
  const match = output.match(/host:\s*(\S+)/);
  if (!match) {
    fail('Unable to determine the Rust host target triple.');
  }
  return match[1];
}

function getPythonFromVenv(venvPath) {
  if (process.platform === 'win32') {
    return path.join(venvPath, 'Scripts', 'python.exe');
  }
  return path.join(venvPath, 'bin', 'python3');
}

function removeIfExists(targetPath) {
  if (fs.existsSync(targetPath)) {
    fs.rmSync(targetPath, { recursive: true, force: true });
  }
}

function findPythonInterpreter() {
  const candidates = process.platform === 'win32'
    ? [
        { command: 'py', args: ['-3.11'] },
        { command: 'py', args: ['-3'] },
        { command: 'python3.11', args: [] },
        { command: 'python3', args: [] },
        { command: 'python', args: [] },
      ]
    : [
        { command: 'python3.11', args: [] },
        { command: 'python3', args: [] },
        { command: 'python', args: [] },
      ];

  for (const candidate of candidates) {
    const result = spawnSync(candidate.command, [...candidate.args, '--version'], {
      stdio: 'pipe',
      shell: false,
    });
    if (result.error || result.status !== 0) {
      continue;
    }
    const output = result.stdout?.toString() || result.stderr?.toString() || '';
    const match = output.match(/Python\s+(\d+)\.(\d+)\.(\d+)/);
    if (!match) {
      continue;
    }
    const major = Number.parseInt(match[1], 10);
    const minor = Number.parseInt(match[2], 10);
    if (Number.isNaN(major) || Number.isNaN(minor) || major !== 3 || minor < 11) {
      continue;
    }
    return candidate;
  }

  fail('Python 3.11 or newer is required to build the embedding sidecar.');
}

function prepareVirtualEnv() {
  removeIfExists(venvDir);
  const python = findPythonInterpreter();
  runCommand(python.command, [...python.args, '-m', 'venv', venvDir]);
  const pythonExecutable = getPythonFromVenv(venvDir);
  if (!fs.existsSync(pythonExecutable)) {
    fail(`Virtual environment Python executable not found at ${pythonExecutable}`);
  }
  runCommand(pythonExecutable, ['-m', 'pip', 'install', '--upgrade', 'pip']);
  runCommand(pythonExecutable, ['-m', 'pip', 'install', 'pyinstaller']);
  runCommand(pythonExecutable, ['-m', 'pip', 'install', '-r', requirementsPath]);
  return pythonExecutable;
}

function buildSidecar(pythonExecutable, targetTriple) {
  fs.mkdirSync(buildDir, { recursive: true });
  removeIfExists(distDir);
  removeIfExists(workDir);
  removeIfExists(specDir);

  const pyInstallerArgs = [
    '-m',
    'PyInstaller',
    '--onefile',
    '--name',
    sidecarName,
    '--collect-submodules',
    'numpy',
    '--collect-data',
    'numpy',
    '--collect-submodules',
    'torch',
    '--collect-data',
    'torch',
    '--distpath',
    distDir,
    '--workpath',
    workDir,
    '--specpath',
    specDir,
    scriptPath,
  ];

  runCommand(pythonExecutable, pyInstallerArgs);

  const extension = process.platform === 'win32' ? '.exe' : '';
  const builtBinary = path.join(distDir, `${sidecarName}${extension}`);
  if (!fs.existsSync(builtBinary)) {
    fail(`PyInstaller did not produce the expected binary at ${builtBinary}`);
  }

  fs.mkdirSync(binariesDir, { recursive: true });

  for (const entry of fs.readdirSync(binariesDir)) {
    if (entry.startsWith(`${sidecarName}-`)) {
      fs.rmSync(path.join(binariesDir, entry), { force: true });
    }
  }

  const destinationName = `${sidecarName}-${targetTriple}${extension}`;
  const destinationPath = path.join(binariesDir, destinationName);
  fs.renameSync(builtBinary, destinationPath);

  console.log(`\u2705 Built sidecar binary at ${destinationPath}`);
  return destinationName;
}

const existingMetadata = readMetadata();
const targetTriple = getTargetTriple();
const expectedBinaryName = `${sidecarName}-${targetTriple}${process.platform === 'win32' ? '.exe' : ''}`;
const existingBinaryPath = path.join(binariesDir, expectedBinaryName);

if (
  existingMetadata &&
  existingMetadata.scriptHash === metadataSignature.scriptHash &&
  existingMetadata.requirementsHash === metadataSignature.requirementsHash &&
  existingMetadata.targetTriple === targetTriple &&
  fs.existsSync(existingBinaryPath)
) {
  console.log(`\u2705 Sidecar binary ${expectedBinaryName} is up to date.`);
  process.exit(0);
}

const pythonExecutable = prepareVirtualEnv();
const producedBinaryName = buildSidecar(pythonExecutable, targetTriple);
writeMetadata({
  ...metadataSignature,
  targetTriple,
  binary: producedBinaryName,
  generatedAt: new Date().toISOString(),
});
