import { spawn } from "node:child_process";
import {
  accessSync,
  chmodSync,
  constants,
  mkdirSync,
  readFileSync,
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

    const shimTemplateUrl = new URL("./setfile-shim.cjs", import.meta.url);
    let shimSource;
    try {
      shimSource = readFileSync(shimTemplateUrl, { encoding: "utf8" });
    } catch (error) {
      console.error("Unable to load SetFile shim template:", error);
      throw error;
    }

    writeFileSync(shimPath, shimSource, { encoding: "utf8" });
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
