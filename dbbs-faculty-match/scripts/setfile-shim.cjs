#!/usr/bin/env node

const { spawnSync } = require('node:child_process');

const LOG_PREFIX = '[setfile-shim]';

function bindConsoleMethod(method, fallback = null) {
  if (typeof console[method] === 'function') {
    return console[method].bind(console, LOG_PREFIX);
  }

  if (fallback && typeof console[fallback] === 'function') {
    return console[fallback].bind(console, LOG_PREFIX);
  }

  return () => {};
}

function createConsoleLogger() {
  return {
    debug: bindConsoleMethod('debug', 'log'),
    info: bindConsoleMethod('info', 'log'),
    warn: bindConsoleMethod('warn', 'error'),
    error: bindConsoleMethod('error', 'warn'),
  };
}

const logger = createConsoleLogger();

function setLogger(customLogger) {
  const fallbackLogger = createConsoleLogger();
  if (!customLogger || typeof customLogger !== 'object') {
    Object.assign(logger, fallbackLogger);
    return logger;
  }

  logger.debug = typeof customLogger.debug === 'function'
    ? customLogger.debug.bind(customLogger)
    : fallbackLogger.debug;
  logger.info = typeof customLogger.info === 'function'
    ? customLogger.info.bind(customLogger)
    : fallbackLogger.info;
  logger.warn = typeof customLogger.warn === 'function'
    ? customLogger.warn.bind(customLogger)
    : fallbackLogger.warn;
  logger.error = typeof customLogger.error === 'function'
    ? customLogger.error.bind(customLogger)
    : fallbackLogger.error;

  return logger;
}

function getLogger() {
  return logger;
}

function formatCommand(command, args = []) {
  return [command, ...(Array.isArray(args) ? args : [])].join(' ');
}

function formatSpawnResultDiagnostic(result) {
  if (!result || typeof result !== 'object') {
    return '';
  }

  const details = [];

  if (result.stdout) {
    const stdout = result.stdout.toString().trim();
    if (stdout) {
      details.push(stdout);
    }
  }

  if (result.stderr) {
    const stderr = result.stderr.toString().trim();
    if (stderr) {
      details.push(stderr);
    }
  }

  if (typeof result.signal === 'string' && result.signal) {
    details.push(`signal: ${result.signal}`);
  }

  return details.join('\n');
}

function logXattrSoftFailure(args, result, context = {}) {
  const activeLogger = getLogger();
  const { operation, targetPath } = context;
  const command = formatCommand('xattr', args);
  const baseMessage = operation
    ? `${operation} failed with exit status ${ExitStatus.XATTR_FAILURE}; continuing without modifying Finder metadata.`
    : `xattr exited with status ${ExitStatus.XATTR_FAILURE}; continuing without modifying Finder metadata.`;

  const location = targetPath ? ` (${targetPath})` : '';
  const diagnostic = formatSpawnResultDiagnostic(result);
  const message = `${command}${location ? ` ${location}` : ''}`;

  if (diagnostic) {
    activeLogger.warn(`${baseMessage}\n${message}\n${diagnostic}`);
  } else {
    activeLogger.warn(`${baseMessage}\n${message}`);
  }
}

const ExitStatus = {
  SUCCESS: 0,
  XATTR_FAILURE: 1,
};

const ATTRIBUTE_FLAGS = {
  A: 0x8000, // Alias file
  B: 0x2000, // Bundle bit (package)
  C: 0x0400, // Custom icon
};

function runXattr(args, options = {}) {
  const { context, ...spawnOptions } = options;
  const result = spawnSync('xattr', args, {
    encoding: 'utf8',
    stdio: 'pipe',
    ...spawnOptions,
  });

  if (result.error) {
    throw Object.assign(
      new Error(`Failed to invoke xattr: ${result.error.message}`),
      { cause: result.error }
    );
  }

  if (typeof result.status === 'number' && result.status !== ExitStatus.SUCCESS) {
    if (result.status === ExitStatus.XATTR_FAILURE) {
      logXattrSoftFailure(args, result, context);
      return null;
    }

    const diagnostic = formatSpawnResultDiagnostic(result);

    const error = new Error(
      `${formatCommand('xattr', args)} failed with status ${result.status}` +
        (diagnostic ? `\n${diagnostic}` : '')
    );
    error.result = result;
    throw error;
  }

  return result;
}

function readFinderInfo(targetPath) {
  const result = runXattr(['-px', 'com.apple.FinderInfo', targetPath], {
    context: { operation: 'readFinderInfo', targetPath },
  });
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
  ], {
    context: { operation: 'writeFinderInfo', targetPath },
  });

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

function setTypeCode(targetPath, typeCode) {
  const info = readFinderInfo(targetPath);
  if (info === null) {
    return null;
  }

  const buffer = Buffer.isBuffer(info) ? Buffer.from(info) : Buffer.alloc(32);
  const code = Buffer.from(String(typeCode).padEnd(4, ' ').slice(0, 4), 'ascii');
  code.copy(buffer, 0);

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
  setTypeCode,
  applyAttributeFlags,
  logger,
  setLogger,
  getLogger,
  createConsoleLogger,
  formatCommand,
  formatSpawnResultDiagnostic,
};

function parseAttributeSpec(spec, state) {
  if (!spec) {
    throw new Error('Missing attribute specification for -a option.');
  }

  for (const symbol of spec.split('')) {
    const upper = symbol.toUpperCase();
    const flag = ATTRIBUTE_FLAGS[upper];

    if (!flag) {
      throw new Error(`Unsupported Finder attribute flag: ${symbol}`);
    }

    if (symbol === upper) {
      state.set |= flag;
      state.clear &= ~flag;
    } else {
      state.clear |= flag;
      state.set &= ~flag;
    }
  }
}

function parseCliArguments(argv) {
  const attributeState = { set: 0, clear: 0 };
  const options = {
    attributes: attributeState,
    creatorCode: null,
    typeCode: null,
    targets: [],
  };

  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];

    if (argument === '--') {
      const remaining = argv.slice(index + 1);
      if (remaining.length === 0) {
        throw new Error('Expected at least one target path after --.');
      }
      options.targets.push(...remaining);
      break;
    }

    if (argument === '-a' || argument.startsWith('-a')) {
      let spec;

      if (argument.length > 2) {
        spec = argument.slice(2);
      } else {
        index += 1;
        if (index >= argv.length) {
          throw new Error('Missing value for -a option.');
        }
        spec = argv[index];
      }

      parseAttributeSpec(spec, attributeState);
      continue;
    }

    if (argument === '-c' || argument.startsWith('-c')) {
      let code;

      if (argument.length > 2) {
        code = argument.slice(2);
      } else {
        index += 1;
        if (index >= argv.length) {
          throw new Error('Missing value for -c option.');
        }
        code = argv[index];
      }

      if (!code) {
        throw new Error('Creator code cannot be empty.');
      }

      options.creatorCode = code;
      continue;
    }

    if (argument === '-t' || argument.startsWith('-t')) {
      let code;

      if (argument.length > 2) {
        code = argument.slice(2);
      } else {
        index += 1;
        if (index >= argv.length) {
          throw new Error('Missing value for -t option.');
        }
        code = argv[index];
      }

      if (!code) {
        throw new Error('Type code cannot be empty.');
      }

      options.typeCode = code;
      continue;
    }

    if (argument.startsWith('-')) {
      throw new Error(`Unsupported option: ${argument}`);
    }

    options.targets.push(argument);
  }

  if (options.targets.length === 0) {
    throw new Error('Expected at least one target path.');
  }

  return options;
}

function applyOperationsToTarget(targetPath, options) {
  const { attributes, creatorCode, typeCode } = options;

  if (attributes.set || attributes.clear) {
    applyAttributeFlags(targetPath, attributes);
  }

  if (creatorCode) {
    setCreatorCode(targetPath, creatorCode);
  }

  if (typeCode) {
    setTypeCode(targetPath, typeCode);
  }
}

if (require.main === module) {
  const [, , ...argv] = process.argv;
  if (argv.length === 0) {
    process.exit(ExitStatus.SUCCESS);
  }

  try {
    const options = parseCliArguments(argv);
    for (const target of options.targets) {
      applyOperationsToTarget(target, options);
    }
  } catch (error) {
    console.error(error.message);
    process.exitCode = 1;
  }
}
