#!/usr/bin/env node

const { spawnSync } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

function createShimLogger() {
  const logPath = (process.env.SETFILE_SHIM_LOG || '').trim();
  const debugEnabled = /^(1|true|on|yes)$/i.test(
    (process.env.SETFILE_SHIM_DEBUG || '').trim()
  );

  let resolvedLogPath = null;
  if (logPath) {
    try {
      const directory = path.dirname(logPath);
      if (directory && directory !== '.') {
        fs.mkdirSync(directory, { recursive: true });
      }
      resolvedLogPath = logPath;
    } catch (error) {
      const fallbackMessage =
        `Unable to create shim log directory for ${logPath}: ${error.message}`;
      console.warn(fallbackMessage);
    }
  }

  const writeEntry = (level, message) => {
    const timestamp = new Date().toISOString();
    const entry = `[${timestamp}] [${level}] ${message}`;

    if (level === 'ERROR') {
      console.error(entry);
    } else if (level === 'WARN') {
      console.warn(entry);
    } else if (debugEnabled) {
      console.log(entry);
    }

    if (resolvedLogPath) {
      try {
        fs.appendFileSync(resolvedLogPath, entry + os.EOL, 'utf8');
      } catch (error) {
        console.warn(
          `Unable to append to shim log ${resolvedLogPath}: ${error.message}`
        );
        resolvedLogPath = null;
      }
    }
  };

  return {
    debug(message) {
      writeEntry('DEBUG', message);
    },
    info(message) {
      writeEntry('INFO', message);
    },
    warn(message) {
      writeEntry('WARN', message);
    },
    error(message) {
      writeEntry('ERROR', message);
    },
  };
}

const shimLog = createShimLogger();

const ExitStatus = {
  SUCCESS: 0,
  XATTR_FAILURE: 1,
};

function describeCommand(args) {
  return `xattr ${args.map((arg) => JSON.stringify(String(arg))).join(' ')}`;
}

function collectDiagnostics(result) {
  const output = [];
  const stderr = result.stderr ? String(result.stderr).trim() : '';
  const stdout = result.stdout ? String(result.stdout).trim() : '';

  if (stdout) {
    output.push(`stdout: ${stdout}`);
  }
  if (stderr) {
    output.push(`stderr: ${stderr}`);
  }

  return output.join('\n');
}

function runXattr(args, options = {}) {
  const command = describeCommand(args);
  shimLog.debug(`Invoking ${command}`);

  const result = spawnSync('xattr', args, {
    encoding: 'utf8',
    stdio: 'pipe',
    ...options,
  });

  if (result.error) {
    const errorMessage = `Failed to invoke ${command}: ${result.error.message}`;
    shimLog.error(errorMessage);
    throw Object.assign(new Error(errorMessage), { cause: result.error });
  }

  if (typeof result.status === 'number' && result.status !== ExitStatus.SUCCESS) {
    if (result.status === ExitStatus.XATTR_FAILURE) {
      const diagnostic = collectDiagnostics(result);
      const warning =
        `${command} exited with status ${result.status}; continuing in best-effort mode.` +
        (diagnostic ? `\n${diagnostic}` : '');
      shimLog.warn(warning);
      return null;
    }

    const diagnostic = collectDiagnostics(result);

    const errorMessage =
      `${command} failed with status ${result.status}` +
      (diagnostic ? `\n${diagnostic}` : '');
    shimLog.error(errorMessage);
    const error = new Error(errorMessage);
    error.result = result;
    throw error;
  }

  if (result.signal) {
    const signalMessage = `${command} terminated due to signal ${result.signal}`;
    shimLog.error(signalMessage);
    const error = new Error(signalMessage);
    error.result = result;
    throw error;
  }

  shimLog.debug(`${command} completed successfully.`);

  return result;
}

function readFinderInfo(targetPath) {
  const result = runXattr(['-px', 'com.apple.FinderInfo', targetPath]);
  if (!result) {
    shimLog.warn(
      `FinderInfo read skipped for ${targetPath}; shim is running in best-effort mode.`
    );
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
    shimLog.warn(
      `FinderInfo write skipped for ${targetPath}; shim is running in best-effort mode.`
    );
    return null;
  }

  return result;
}

function setCreatorCode(targetPath, creatorCode) {
  const info = readFinderInfo(targetPath);
  if (info === null) {
    shimLog.warn(
      `Creator code update skipped for ${targetPath}; FinderInfo could not be read.`
    );
    return null;
  }

  const buffer = Buffer.isBuffer(info) ? Buffer.from(info) : Buffer.alloc(32);
  const code = Buffer.from(String(creatorCode).padEnd(4, ' ').slice(0, 4), 'ascii');
  code.copy(buffer, 4);

  const result = writeFinderInfo(targetPath, buffer);
  if (!result) {
    shimLog.warn(
      `Creator code update skipped for ${targetPath}; FinderInfo write failed in best-effort mode.`
    );
    return null;
  }

  return result;
}

function applyAttributeFlags(targetPath, { set = 0, clear = 0 } = {}) {
  const info = readFinderInfo(targetPath);
  if (info === null) {
    shimLog.warn(
      `Flag update skipped for ${targetPath}; FinderInfo could not be read.`
    );
    return null;
  }

  const buffer = Buffer.isBuffer(info) ? Buffer.from(info) : Buffer.alloc(32);
  const currentFlags = buffer.readUInt16BE(8);
  const updatedFlags = (currentFlags | set) & ~clear;
  buffer.writeUInt16BE(updatedFlags, 8);

  const result = writeFinderInfo(targetPath, buffer);
  if (!result) {
    shimLog.warn(
      `Flag update skipped for ${targetPath}; FinderInfo write failed in best-effort mode.`
    );
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
  createShimLogger,
  shimLog,
};

if (require.main === module) {
  const [, , ...argv] = process.argv;
  if (argv.length === 0) {
    process.exit(0);
  }

  try {
    runXattr(argv);
  } catch (error) {
    shimLog.error(error.message);
    process.exitCode = 1;
  }
}
