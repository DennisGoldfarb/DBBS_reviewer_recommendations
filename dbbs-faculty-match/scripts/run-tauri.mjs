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
    if (!existsSync(shimPath)) {
      const script = `#!/usr/bin/env bash
# Auto-generated shim to satisfy Tauri's DMG bundler when macOS \`SetFile\` is unavailable.
# The shim performs a no-op but reports success so that the bundler can proceed.
exit 0
`;
      writeFileSync(shimPath, script, { encoding: "utf8" });
      chmodSync(shimPath, 0o755);
    }
    env.PATH = `${shimDir}${path.delimiter}${env.PATH ?? ""}`;
    console.warn("SetFile command not found on PATH. Using no-op shim to allow DMG bundling to continue.");
  }
}

const command = process.platform === "win32" ? "tauri.cmd" : "tauri";
const child = spawn(command, args, {
  stdio: "inherit",
  env,
});

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
