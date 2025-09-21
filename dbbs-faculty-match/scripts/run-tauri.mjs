import { spawn } from "node:child_process";
import {
  accessSync,
  chmodSync,
  constants,
  existsSync,
  mkdirSync,
  writeFileSync,
} from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";

const args = process.argv.slice(2);
const env = { ...process.env };
let shimLogPath;

if (process.platform === "darwin") {
  const pathEntries = (env.PATH ?? "").split(path.delimiter).filter(Boolean);
  const hasSetFile = pathEntries.some((entry) => {
    try {
      accessSync(path.join(entry, "SetFile"), constants.X_OK);
      return true;
    } catch {
      return false;
    }
  });

  if (!hasSetFile) {
    const shimDir = path.join(tmpdir(), "dbbs-faculty-match", "setfile-shim");
    mkdirSync(shimDir, { recursive: true });
    shimLogPath = path.join(shimDir, "setfile-shim.log");
    const shimPath = path.join(shimDir, "SetFile");

    const shimSource = String.raw`#!/usr/bin/env node
"use strict";

const { spawnSync } = require("node:child_process");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

const FINDER_INFO_ATTR = "com.apple.FinderInfo";
const FLAG_HAS_CUSTOM_ICON = 0x0400;
const FLAG_HAS_BEEN_INITED = 0x0100;
const LOG_ENV = "DBBS_SETFILE_LOG";
const MODE_ENV = "DBBS_SETFILE_MODE";
const TRACE_ENV = "DBBS_SETFILE_TRACE";
const DEFAULT_LOG = path.join(
  os.tmpdir(),
  "dbbs-faculty-match",
  "setfile-shim",
  "setfile.log",
);

function ensureLogPath(target) {
  try {
    fs.mkdirSync(path.dirname(target), { recursive: true });
  } catch (error) {
    if (error && error.code !== "EEXIST") {
      console.warn(
        `SetFile shim: unable to prepare log directory ${path.dirname(target)} (${error.message})`,
      );
    }
  }
}

function resolveLogPath() {
  const explicit = process.env[LOG_ENV];
  if (explicit && typeof explicit === "string" && explicit.length > 0) {
    ensureLogPath(explicit);
    return explicit;
  }
  ensureLogPath(DEFAULT_LOG);
  return DEFAULT_LOG;
}

const logPath = resolveLogPath();
const traceValue = String(process.env[TRACE_ENV] || "").toLowerCase();
const traceEnabled = traceValue === "1" || traceValue === "true";
const modeValue = String(process.env[MODE_ENV] || "apply").toLowerCase();
const dryRun = modeValue === "observe" || modeValue === "dry-run";

function appendLog(level, message) {
  const timestamp = new Date().toISOString();
  const line = `[${timestamp}] [${level}] ${message}\n`;
  try {
    fs.appendFileSync(logPath, line, { encoding: "utf8" });
  } catch (error) {
    console.warn(`SetFile shim: failed to append to log (${error.message})`);
  }
}

function log(level, message) {
  appendLog(level, message);
  if (level === "ERROR") {
    console.error(`SetFile shim: ${message}`);
  } else if (level === "WARN") {
    console.warn(`SetFile shim: ${message}`);
  } else if (traceEnabled || level !== "DEBUG") {
    console.log(`SetFile shim: ${message}`);
  }
}

function runXattr(args, context) {
  const result = spawnSync("xattr", args, { encoding: "utf8" });
  if (result.error) {
    const err = new Error(`${context}: ${result.error.message}`);
    err.cause = result.error;
    throw err;
  }
  if (typeof result.status === "number" && result.status !== 0) {
    const stderr = result.stderr ? String(result.stderr).trim() : "";
    const stdout = result.stdout ? String(result.stdout).trim() : "";
    const messageParts = [`exit code ${result.status}`];
    if (stderr) {
      messageParts.push(`stderr: ${stderr}`);
    }
    if (stdout) {
      messageParts.push(`stdout: ${stdout}`);
    }
    const err = new Error(`${context}: ${messageParts.join("; ")}`);
    err.code = "XATTR_FAILURE";
    err.stderr = stderr;
    err.stdout = stdout;
    err.status = result.status;
    throw err;
  }
  return result.stdout ? String(result.stdout) : "";
}

function readFinderInfo(target) {
  try {
    const output = runXattr(
      ["-px", FINDER_INFO_ATTR, target],
      `read Finder info for ${target}`,
    );
    const hex = output.replace(/\s+/g, "");
    if (!hex) {
      log("DEBUG", `Finder info empty for ${target}`);
      return Buffer.alloc(32);
    }
    const buffer = Buffer.from(hex, "hex");
    if (buffer.length === 32) {
      return buffer;
    }
    const info = Buffer.alloc(32);
    buffer.copy(info, 0, 0, Math.min(buffer.length, 32));
    return info;
  } catch (error) {
    const message = error && error.message ? error.message : String(error);
    if (
      error &&
      error.code === "XATTR_FAILURE" &&
      typeof error.stderr === "string" &&
      /No such xattr/i.test(error.stderr)
    ) {
      log("DEBUG", `Finder info missing for ${target}; using zero-filled buffer (${message})`);
      return Buffer.alloc(32);
    }
    throw error;
  }
}

function writeFinderInfo(target, info) {
  if (!Buffer.isBuffer(info) || info.length !== 32) {
    throw new Error("Finder info must be a 32-byte buffer");
  }
  if (dryRun) {
    log("INFO", `[dry-run] Skipping Finder info write for ${target}`);
    return;
  }
  const hex = info.toString("hex");
  runXattr(["-wx", FINDER_INFO_ATTR, hex, target], `write Finder info for ${target}`);
}

function setCreatorCode(target, code) {
  if (typeof code !== "string" || code.length !== 4) {
    throw new Error("Creator code must be exactly four ASCII characters");
  }
  if (Buffer.byteLength(code, "ascii") !== 4) {
    throw new Error("Creator code must use ASCII characters");
  }
  const info = readFinderInfo(target);
  const codeBuffer = Buffer.alloc(4);
  codeBuffer.write(code, 0, 4, "ascii");
  codeBuffer.copy(info, 4);
  writeFinderInfo(target, info);
  log("INFO", `Updated creator code to ${code} for ${target}`);
}

function applyAttributeFlags(target, flagSpec) {
  if (typeof flagSpec !== "string" || flagSpec.length === 0) {
    log("WARN", `Empty flag specification provided for ${target}`);
    return;
  }
  const info = readFinderInfo(target);
  let flags = info.readUInt16BE(8);
  for (const flag of flagSpec) {
    if (flag === "C") {
      flags |= FLAG_HAS_CUSTOM_ICON;
      flags |= FLAG_HAS_BEEN_INITED;
    } else if (flag === "c") {
      flags &= ~FLAG_HAS_CUSTOM_ICON;
    } else {
      log("WARN", `Unsupported flag '${flag}' for ${target}`);
    }
  }
  info.writeUInt16BE(flags, 8);
  writeFinderInfo(target, info);
  log("INFO", `Applied attribute flags '${flagSpec}' to ${target}`);
}

function parseInvocation(argv) {
  const parsed = {
    flags: [],
    creator: undefined,
    target: undefined,
    extras: [],
  };
  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];
    if (token === "-a" || token === "-c") {
      if (i + 1 >= argv.length) {
        throw new Error(`Missing argument for ${token}`);
      }
      const value = argv[i + 1];
      i += 1;
      if (token === "-a") {
        parsed.flags.push(value);
      } else {
        parsed.creator = value;
      }
      continue;
    }
    if (!parsed.target && typeof token === "string" && !token.startsWith("-")) {
      parsed.target = token;
      continue;
    }
    parsed.extras.push(token);
  }
  return parsed;
}

function main() {
  const argv = process.argv.slice(2);
  if (argv.length === 0) {
    log("DEBUG", "No arguments provided; exiting");
    return;
  }

  log("INFO", `Invocation arguments: ${JSON.stringify(argv)}`);

  const parsed = parseInvocation(argv);
  if (!parsed.target) {
    throw new Error("Target path argument is required");
  }
  if (!fs.existsSync(parsed.target)) {
    throw new Error(`Target not found: ${parsed.target}`);
  }

  if (dryRun) {
    log("WARN", "Running in observe mode; no modifications will be applied");
  }

  if (parsed.extras.length > 0) {
    log(
      "WARN",
      `Unsupported or additional arguments encountered: ${parsed.extras.join(" ")}`,
    );
  }

  if (parsed.creator) {
    setCreatorCode(parsed.target, parsed.creator);
  }
  if (parsed.flags.length > 0) {
    for (const spec of parsed.flags) {
      applyAttributeFlags(parsed.target, spec);
    }
  }
  if (!parsed.creator && parsed.flags.length === 0) {
    log("WARN", "No supported operations requested; treated as no-op");
  }
  log("INFO", `Completed processing for ${parsed.target}`);
}

try {
  main();
} catch (error) {
  const message = error && error.message ? error.message : String(error);
  log("ERROR", message);
  if (error && error.stack) {
    appendLog("ERROR", error.stack);
  }
  process.exitCode = 1;
}
`;

    writeFileSync(shimPath, `${shimSource}\n`, { encoding: "utf8" });
    chmodSync(shimPath, 0o755);
    env.PATH = `${shimDir}${path.delimiter}${env.PATH ?? ""}`;
    env.DBBS_SETFILE_LOG = shimLogPath;
    env.DBBS_SETFILE_MODE = env.DBBS_SETFILE_MODE ?? "apply";
    env.DBBS_SETFILE_TRACE = env.DBBS_SETFILE_TRACE ?? "0";
    console.warn(
      "SetFile command not found on PATH. Using fallback shim to emulate required behavior.",
    );
    console.warn(`SetFile shim logging to: ${shimLogPath}`);
  }
}

const isWindows = process.platform === "win32";
const command = "tauri";
const spawnOptions = {
  stdio: "inherit",
  env,
  shell: isWindows,
};

let child;

try {
  child = spawn(command, args, spawnOptions);
} catch (error) {
  console.error("Failed to start Tauri CLI:", error);
  process.exitCode = 1;
  process.exit();
}

child.on("error", (error) => {
  console.error("Failed to start Tauri CLI:", error);
  process.exitCode = 1;
});

child.on("exit", (code, signal) => {
  if (signal) {
    process.kill(process.pid, signal);
    return;
  }
  if (code && code !== 0 && shimLogPath) {
    console.error(`Tauri CLI exited with code ${code}. Review shim log at ${shimLogPath}`);
  }
  process.exit(code ?? 0);
});
