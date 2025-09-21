#!/usr/bin/env node

const { spawnSync } = require('node:child_process');
const { existsSync, mkdirSync, createWriteStream } = require('node:fs');
const { dirname, resolve } = require('node:path');
const { Console } = require('node:console');

const ExitStatus = {
  SUCCESS: 0,
  XATTR_FAILURE: 1,
};

const LOG_PATH = process.env.SETFILE_SHIM_LOG
  ? resolve(process.env.SETFILE_SHIM_LOG)
  : null;

function ensureLogStream(path) {
  if (!path) {
    return null;
  }

  const dir = dirname(path);
  if (!existsSync(dir)) {
    mkdirSync(dir, { recursive: true });
  }

  return createWriteStream(path, { flags: 'a' });
}

const logStream = ensureLogStream(LOG_PATH);
const shimConsole = logStream
  ? new Console(logStream, logStream)
  : new Console(process.stdout, process.stderr);

function formatArgs(args) {
  return args
    .map((part) => (part.includes(' ') ? `"${part}"` : part))
    .join(' ');
}

function warn(message) {
  shimConsole.warn(`[setfile-shim] ${message}`);
}

function info(message) {
  shimConsole.log(`[setfile-shim] ${message}`);
}

function runXattr(args, options = {}) {
  const result = spawnSync('xattr', args, {
    encoding: 'utf8',
    stdio: 'pipe',
    ...options,
  });

  if (result.error) {
    throw Object.assign(
      new Error(`Failed to invoke xattr: ${result.error.message}`),
      { cause: result.error }
    );
  }

  if (typeof result.status === 'number' && result.status !== ExitStatus.SUCCESS) {
    if (result.status === ExitStatus.XATTR_FAILURE) {
      warn(
        `xattr ${formatArgs(args)} exited with status ${result.status}.` +
          (result.stderr ? ` stderr: ${result.stderr.trim()}` : '')
      );
      return null;
    }

    const stderr = result.stderr ? result.stderr.trim() : '';
    const stdout = result.stdout ? result.stdout.trim() : '';
    const diagnostic = [stdout, stderr].filter(Boolean).join('\n');

    const error = new Error(
      `xattr ${formatArgs(args)} failed with status ${result.status}` +
        (diagnostic ? `\n${diagnostic}` : '')
    );
    error.result = result;
    throw error;
  }

  return result;
}

function readFinderInfo(targetPath) {
  const result = runXattr(['-px', 'com.apple.FinderInfo', targetPath]);
  if (!result) {
    warn(`FinderInfo read skipped for ${targetPath}.`);
    return null;
  }

  const output = (result.stdout || '').replace(/\s+/g, '');
  if (!output) {
    return Buffer.alloc(32);
  }

  try {
    const buffer = Buffer.from(output, 'hex');
    if (buffer.length === 32) {
      return buffer;
    }
    if (buffer.length > 32) {
      return buffer.subarray(0, 32);
    }
    const expanded = Buffer.alloc(32);
    buffer.copy(expanded, 0, 0, buffer.length);
    return expanded;
  } catch (error) {
    warn(`Unable to parse FinderInfo for ${targetPath}: ${error.message}`);
    return Buffer.alloc(32);
  }
}

function writeFinderInfo(targetPath, finderInfo) {
  let buffer;

  if (Buffer.isBuffer(finderInfo)) {
    buffer = Buffer.from(finderInfo);
  } else {
    const sanitized = String(finderInfo || '').replace(/\s+/g, '');
    buffer = sanitized ? Buffer.from(sanitized, 'hex') : Buffer.alloc(32);
  }

  if (buffer.length !== 32) {
    if (buffer.length > 32) {
      buffer = buffer.subarray(0, 32);
    } else {
      const expanded = Buffer.alloc(32);
      buffer.copy(expanded, 0, 0, buffer.length);
      buffer = expanded;
    }
  }

  const result = runXattr([
    '-wx',
    'com.apple.FinderInfo',
    buffer.toString('hex'),
    targetPath,
  ]);

  if (!result) {
    warn(`FinderInfo update skipped for ${targetPath}. Continuing without Finder metadata.`);
    return null;
  }

  return result;
}

function setCreatorCode(targetPath, creatorCode) {
  const info = readFinderInfo(targetPath);
  if (info === null) {
    warn(`Creator code update skipped for ${targetPath}.`);
    return null;
  }

  const buffer = Buffer.isBuffer(info) ? Buffer.from(info) : Buffer.alloc(32);
  const code = Buffer.from(String(creatorCode).padEnd(4, ' ').slice(0, 4), 'ascii');
  code.copy(buffer, 4);

  const result = writeFinderInfo(targetPath, buffer);
  if (!result) {
    warn(`Creator code update skipped for ${targetPath}. Continuing without Finder metadata.`);
    return null;
  }

  return result;
}

function applyAttributeFlags(targetPath, { set = 0, clear = 0 } = {}) {
  const info = readFinderInfo(targetPath);
  if (info === null) {
    warn(`Attribute flag update skipped for ${targetPath}.`);
    return null;
  }

  const buffer = Buffer.isBuffer(info) ? Buffer.from(info) : Buffer.alloc(32);
  const currentFlags = buffer.readUInt16BE(8);
  const updatedFlags = (currentFlags | set) & ~clear;
  buffer.writeUInt16BE(updatedFlags, 8);

  const result = writeFinderInfo(targetPath, buffer);
  if (!result) {
    warn(`Finder attribute flags skipped for ${targetPath}. Continuing without Finder metadata.`);
    return null;
  }

  return result;
}

module.exports = {
  ExitStatus,
  runXattr,
  readFinderInfo,
  writeFinderInfo,
  setCreatorCode,
  applyAttributeFlags,
};

if (require.main === module) {
  const [, , ...argv] = process.argv;
  if (argv.length === 0) {
    info('setfile-shim invoked without arguments; no-op.');
    process.exit(0);
  }

  try {
    runXattr(argv);
  } catch (error) {
    warn(error.message);
    process.exitCode = 1;
  }
}
