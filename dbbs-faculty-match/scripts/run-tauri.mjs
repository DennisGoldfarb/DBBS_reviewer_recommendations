import { spawn } from "node:child_process";
import { accessSync, constants, existsSync, mkdirSync, writeFileSync, chmodSync } from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";

const args = process.argv.slice(2);
const env = { ...process.env };

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
    if (!existsSync(shimDir)) {
      mkdirSync(shimDir, { recursive: true });
    }
    const shimPath = path.join(shimDir, "SetFile");
    const scriptLines = [
      "#!/usr/bin/env node",
      '"use strict";',
      "",
      'const { spawnSync } = require("node:child_process");',
      'const { existsSync } = require("node:fs");',
      "",
      'const FINDER_INFO_ATTR = "com.apple.FinderInfo";',
      "const FLAG_HAS_CUSTOM_ICON = 0x0400;",
      "const FLAG_HAS_BEEN_INITED = 0x0100;",
      "",
      "function readFinderInfo(target) {",
      '  const result = spawnSync("xattr", ["-px", FINDER_INFO_ATTR, target], { encoding: "utf8" });',
      '  if (result.status === 0 && typeof result.stdout === "string") {',
      '    const hex = result.stdout.replace(/\\s+/g, "");',
      '    if (hex.length > 0) {',
      '      const buffer = Buffer.from(hex, "hex");',
      '      if (buffer.length === 32) {',
      '        return Buffer.from(buffer);',
      '      }',
      '      const info = Buffer.alloc(32);',
      '      buffer.copy(info, 0, 0, Math.min(buffer.length, 32));',
      '      return info;',
      "    }",
      "  }",
      "  return Buffer.alloc(32);",
      "}",
      "",
      "function writeFinderInfo(target, info) {",
      '  if (!Buffer.isBuffer(info) || info.length !== 32) {',
      '    throw new Error("Finder info must be 32 bytes");',
      "  }",
      '  const hex = info.toString("hex");',
      '  const result = spawnSync("xattr", ["-wx", FINDER_INFO_ATTR, hex, target], { encoding: "utf8" });',
      '  if (result.status !== 0) {',
      '    const message = result.stderr ? String(result.stderr).trim() : "";',
      '    throw new Error(message || "xattr write failed");',
      "  }",
      "}",
      "",
      "function ensureTarget(target) {",
      '  if (!existsSync(target)) {',
      '    throw new Error(`Target not found: ${target}`);',
      "  }",
      "}",
      "",
      "function setCreatorCode(target, code) {",
      '  if (typeof code !== "string" || code.length !== 4) {',
      '    throw new Error("Creator code must be exactly four characters");',
      "  }",
      '  if (Buffer.byteLength(code, "ascii") !== 4) {',
      '    throw new Error("Creator code must use ASCII characters");',
      "  }",
      "  const info = readFinderInfo(target);",
      "  const codeBuffer = Buffer.alloc(4);",
      '  codeBuffer.write(code, 0, 4, "ascii");',
      "  codeBuffer.copy(info, 4);",
      "  writeFinderInfo(target, info);",
      "}",
      "",
      "function applyAttributeFlags(target, flagSpec) {",
      "  if (!flagSpec) {",
      "    return;",
      "  }",
      "  const info = readFinderInfo(target);",
      "  let flags = info.readUInt16BE(8);",
      "  for (const flag of flagSpec) {",
      '    if (flag === "C") {',
      "      flags |= FLAG_HAS_CUSTOM_ICON;",
      "      flags |= FLAG_HAS_BEEN_INITED;",
      '    } else if (flag === "c") {',
      "      flags &= ~FLAG_HAS_CUSTOM_ICON;",
      "    }",
      "  }",
      "  info.writeUInt16BE(flags, 8);",
      "  writeFinderInfo(target, info);",
      "}",
      "",
      "function main() {",
      "  const argv = process.argv.slice(2);",
      "  if (argv.length === 0) {",
      "    return;",
      "  }",
      "  try {",
      '    if (argv[0] === "-c") {',
      "      if (argv.length < 3) {",
      '        throw new Error("Missing argument for -c");',
      "      }",
      "      const code = argv[1];",
      "      const target = argv[2];",
      "      ensureTarget(target);",
      "      setCreatorCode(target, code);",
      "      return;",
      "    }",
      '    if (argv[0] === "-a") {',
      "      if (argv.length < 3) {",
      '        throw new Error("Missing argument for -a");',
      "      }",
      "      const flags = argv[1];",
      "      const target = argv[2];",
      "      ensureTarget(target);",
      "      applyAttributeFlags(target, flags);",
      "      return;",
      "    }",
      '    console.warn(`SetFile shim: unsupported arguments ${argv.join(" ")}`);',
      "  } catch (error) {",
      "    const message = error && error.message ? error.message : String(error);",
      '    console.error(`SetFile shim error: ${message}`);',
      "    process.exitCode = 1;",
      "  }",
      "}",
      "",
      "main();",
    ];
    const script = `${scriptLines.join("\n")}\n`;
    writeFileSync(shimPath, script, { encoding: "utf8" });
    chmodSync(shimPath, 0o755);
    env.PATH = `${shimDir}${path.delimiter}${env.PATH ?? ""}`;
    console.warn("SetFile command not found on PATH. Using fallback shim to emulate required behavior.");
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
  process.exit(code ?? 0);
});
