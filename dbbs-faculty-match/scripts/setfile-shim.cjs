#!/usr/bin/env node

const { spawnSync } = require('node:child_process');

const ExitStatus = {
  SUCCESS: 0,
  XATTR_FAILURE: 1,
};

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
      return null;
    }

    const stderr = result.stderr ? result.stderr.trim() : '';
    const stdout = result.stdout ? result.stdout.trim() : '';
    const diagnostic = [stdout, stderr].filter(Boolean).join('\n');

    const error = new Error(
      `xattr ${args.join(' ')} failed with status ${result.status}` +
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
    return null;
  }

  return result;
}

function setCreatorCode(targetPath, creatorCode) {
  const info = readFinderInfo(targetPath);
  if (info === null) {
    return null;
  }

  const buffer = Buffer.isBuffer(info) ? Buffer.from(info) : Buffer.alloc(32);
  const code = Buffer.from(String(creatorCode).padEnd(4, ' ').slice(0, 4), 'ascii');
  code.copy(buffer, 4);

  const result = writeFinderInfo(targetPath, buffer);
  if (!result) {
    return null;
  }

  return result;
}

function applyAttributeFlags(targetPath, { set = 0, clear = 0 } = {}) {
  const info = readFinderInfo(targetPath);
  if (info === null) {
    return null;
  }

  const buffer = Buffer.isBuffer(info) ? Buffer.from(info) : Buffer.alloc(32);
  const currentFlags = buffer.readUInt16BE(8);
  const updatedFlags = (currentFlags | set) & ~clear;
  buffer.writeUInt16BE(updatedFlags, 8);

  const result = writeFinderInfo(targetPath, buffer);
  if (!result) {
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
    process.exit(0);
  }

  try {
    runXattr(argv);
  } catch (error) {
    console.error(error.message);
    process.exitCode = 1;
  }
}
