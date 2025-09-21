import { spawn, spawnSync } from "node:child_process";
import {
  accessSync,
  chmodSync,
  constants,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";

const rawArgs = process.argv.slice(2);
const forwardedArgs = [];
let diagnosticsRequested = false;

for (const arg of rawArgs) {
  if (arg === "--diagnostics" || arg === "--sanity-checks") {
    diagnosticsRequested = true;
    continue;
  }
  forwardedArgs.push(arg);
}

const args = forwardedArgs;
const env = { ...process.env };
let shimLogPath;
let shimPath;
let setFilePath;

const diagnosticsEnabled =
  diagnosticsRequested ||
  isTruthy(env.DBBS_TAURI_DIAGNOSTICS) ||
  isTruthy(env.DBBS_TAURI_DEBUG);

const diagLog = diagnosticsEnabled
  ? (message) => console.log(`[tauri-runner][diagnostic] ${message}`)
  : () => {};
const diagWarn = diagnosticsEnabled
  ? (message) => console.warn(`[tauri-runner][diagnostic] ${message}`)
  : () => {};

if (diagnosticsEnabled) {
  diagLog(`Diagnostics enabled. Forwarded args: ${JSON.stringify(args)}`);
  diagLog(`Working directory: ${process.cwd()}`);
  diagLog(`Node version: ${process.version}`);
  diagLog(`Platform: ${process.platform} (${process.arch})`);
}

if (process.platform === "darwin") {
  const pathEntries = (env.PATH ?? "").split(path.delimiter).filter(Boolean);

  if (diagnosticsEnabled) {
    diagLog(
      `PATH entries (${pathEntries.length}): ${
        pathEntries.length > 0 ? pathEntries.join(", ") : "(empty)"
      }`,
    );
  }

  for (const entry of pathEntries) {
    const candidate = path.join(entry, "SetFile");
    try {
      accessSync(candidate, constants.X_OK);
      setFilePath = candidate;
      break;
    } catch {
      // continue searching
    }
  }

  const hasSetFile = Boolean(setFilePath);

  if (diagnosticsEnabled) {
    if (hasSetFile) {
      diagLog(`Found native SetFile at ${setFilePath}`);
    } else {
      diagWarn("Native SetFile not found on PATH; fallback shim will be used");
    }
  }

  if (!hasSetFile) {
    const shimDir = path.join(tmpdir(), "dbbs-faculty-match", "setfile-shim");
    mkdirSync(shimDir, { recursive: true });
    shimLogPath = path.join(shimDir, "setfile-shim.log");
    shimPath = path.join(shimDir, "SetFile");

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
    if (!env.DBBS_SETFILE_TRACE && diagnosticsEnabled) {
      env.DBBS_SETFILE_TRACE = "1";
      diagLog("Enabled DBBS_SETFILE_TRACE=1 for verbose shim logging");
    } else {
      env.DBBS_SETFILE_TRACE = env.DBBS_SETFILE_TRACE ?? "0";
    }
    console.warn(
      "SetFile command not found on PATH. Using fallback shim to emulate required behavior.",
    );
    console.warn(`SetFile shim logging to: ${shimLogPath}`);
    if (diagnosticsEnabled) {
      diagLog(`SetFile shim prepared at ${shimPath}`);
    }
  }

  if (diagnosticsEnabled) {
    runMacDiagnostics({
      env,
      usingShim: !hasSetFile,
      shimPath,
      shimLogPath,
      setFilePath,
    });
  }
}

const isWindows = process.platform === "win32";
const command = "tauri";
const spawnOptions = {
  stdio: "inherit",
  env,
  shell: isWindows,
};

if (diagnosticsEnabled) {
  diagLog(`Invoking Tauri CLI: ${command} ${args.join(" ")}`);
}

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
    emitShimLog(shimLogPath, diagnosticsEnabled ? 160 : 80);
  }
  if (code && code !== 0 && diagnosticsEnabled && process.platform === "darwin") {
    summarizeBundleArtifacts();
  }
  process.exit(code ?? 0);
});

function isTruthy(value) {
  if (value === undefined || value === null) {
    return false;
  }
  const normalized = String(value).trim().toLowerCase();
  return ["1", "true", "yes", "on", "enable", "enabled"].includes(normalized);
}

function formatOutput(output, limit = 2000) {
  if (!output) {
    return "";
  }
  const text = String(output).trim();
  if (!text) {
    return "";
  }
  if (text.length > limit) {
    return `${text.slice(0, limit)}… (truncated)`;
  }
  return text;
}

function runCommandForDiagnostics(label, commandName, commandArgs = [], options = {}) {
  diagLog(`Running diagnostic command: ${label}`);
  try {
    const result = spawnSync(commandName, commandArgs, {
      encoding: "utf8",
      ...options,
    });
    if (result.error) {
      diagWarn(`${label} spawn error: ${result.error.message}`);
      return;
    }
    diagLog(`${label} exit code: ${result.status ?? "(null)"}`);
    const stdout = formatOutput(result.stdout);
    if (stdout) {
      diagLog(`${label} stdout: ${stdout}`);
    }
    const stderr = formatOutput(result.stderr);
    if (stderr) {
      diagLog(`${label} stderr: ${stderr}`);
    }
  } catch (error) {
    diagWarn(`${label} threw error: ${error.message}`);
  }
}

function runShimSelfTest(shimPathValue, shimEnv) {
  if (!shimPathValue) {
    diagWarn("Shim self-test skipped; shim path unavailable");
    return;
  }

  diagLog("Running SetFile shim self-test");
  let tempDir;
  try {
    tempDir = mkdtempSync(path.join(tmpdir(), "dbbs-setfile-selftest-"));
  } catch (error) {
    diagWarn(`Unable to create temporary directory for shim self-test: ${error.message}`);
    return;
  }

  const testFile = path.join(tempDir, "shim-test.txt");
  try {
    writeFileSync(testFile, "diagnostic", { encoding: "utf8" });
  } catch (error) {
    diagWarn(`Unable to create shim test file: ${error.message}`);
    cleanupTemp(tempDir);
    return;
  }

  try {
    const shimResult = spawnSync(shimPathValue, ["-a", "C", testFile], {
      encoding: "utf8",
      env: shimEnv,
    });
    if (shimResult.error) {
      diagWarn(`Shim invocation failed: ${shimResult.error.message}`);
    } else {
      diagLog(`Shim exit code: ${shimResult.status ?? "(null)"}`);
      const stdout = formatOutput(shimResult.stdout);
      if (stdout) {
        diagLog(`Shim stdout: ${stdout}`);
      }
      const stderr = formatOutput(shimResult.stderr);
      if (stderr) {
        diagLog(`Shim stderr: ${stderr}`);
      }
    }

    const inspectResult = spawnSync("xattr", ["-p", "com.apple.FinderInfo", testFile], {
      encoding: "utf8",
    });
    if (inspectResult.error) {
      diagWarn(`xattr inspection failed: ${inspectResult.error.message}`);
    } else {
      diagLog(`xattr exit code: ${inspectResult.status ?? "(null)"}`);
      const stdout = formatOutput(inspectResult.stdout);
      if (stdout) {
        diagLog(`xattr stdout: ${stdout}`);
      }
      const stderr = formatOutput(inspectResult.stderr);
      if (stderr) {
        diagLog(`xattr stderr: ${stderr}`);
      }
    }
  } catch (error) {
    diagWarn(`Shim self-test threw error: ${error.message}`);
  } finally {
    cleanupTemp(tempDir);
  }
}

function cleanupTemp(directory) {
  if (!directory) {
    return;
  }
  try {
    rmSync(directory, { recursive: true, force: true });
  } catch (error) {
    diagWarn(`Failed to remove temporary directory ${directory}: ${error.message}`);
  }
}

function runMacDiagnostics({ env: diagEnv, usingShim, shimPath: shimPathValue, shimLogPath: shimLog, setFilePath: nativeSetFilePath, }) {
  diagLog("--- macOS diagnostics ---");
  diagLog(`SetFile in use: ${usingShim ? shimPathValue ?? "(shim unavailable)" : nativeSetFilePath ?? "(not detected)"}`);
  diagLog(`DBBS_SETFILE_LOG: ${diagEnv.DBBS_SETFILE_LOG ?? "(not set)"}`);
  diagLog(`DBBS_SETFILE_MODE: ${diagEnv.DBBS_SETFILE_MODE ?? "(not set)"}`);
  diagLog(`DBBS_SETFILE_TRACE: ${diagEnv.DBBS_SETFILE_TRACE ?? "(not set)"}`);
  diagLog(`TMPDIR: ${tmpdir()}`);

  runCommandForDiagnostics("macOS version", "sw_vers");
  runCommandForDiagnostics("CPU brand", "sysctl", ["-n", "machdep.cpu.brand_string"]);
  runCommandForDiagnostics("which xattr", "which", ["xattr"]);
  runCommandForDiagnostics("which hdiutil", "which", ["hdiutil"]);
  runCommandForDiagnostics("hdiutil version", "hdiutil", ["help"]);
  runCommandForDiagnostics("python3 --version", "python3", ["--version"]);

  if (usingShim) {
    runShimSelfTest(shimPathValue, diagEnv);
    if (shimLog) {
      diagLog(`Shim log path: ${shimLog}`);
    }
  }

  diagLog("--- end macOS diagnostics ---");
}

function emitShimLog(logPathValue, lineLimit = 120) {
  try {
    if (!existsSync(logPathValue)) {
      console.error(`SetFile shim log not found at ${logPathValue}`);
      return;
    }
    const content = readFileSync(logPathValue, { encoding: "utf8" });
    if (!content) {
      console.error(`SetFile shim log is empty at ${logPathValue}`);
      return;
    }
    const lines = content.replace(/\r\n/g, "\n").split("\n");
    const tail = lineLimit > 0 ? lines.slice(-lineLimit) : lines;
    console.error("----- SetFile shim log (tail) -----");
    for (const line of tail) {
      if (line) {
        console.error(line);
      }
    }
    console.error("----- End SetFile shim log -----");
  } catch (error) {
    console.error(`Unable to read SetFile shim log at ${logPathValue}:`, error);
  }
}

function summarizeBundleArtifacts() {
  const bundleDir = path.join(
    process.cwd(),
    "src-tauri",
    "target",
    "release",
    "bundle",
    "dmg",
  );
  diagLog(`Inspecting DMG bundle directory: ${bundleDir}`);
  try {
    const entries = readdirSync(bundleDir, { withFileTypes: true });
    if (entries.length === 0) {
      diagWarn("DMG bundle directory exists but is empty");
    } else {
      diagLog(
        `DMG bundle entries: ${entries
          .map((entry) => `${entry.name}${entry.isDirectory() ? "/" : ""}`)
          .join(", ")}`,
      );
    }
  } catch (error) {
    diagWarn(`Unable to list DMG bundle directory: ${error.message}`);
    return;
  }

  const scriptPath = path.join(bundleDir, "bundle_dmg.sh");
  if (existsSync(scriptPath)) {
    try {
      const stats = statSync(scriptPath);
      diagLog(
        `bundle_dmg.sh mode ${(stats.mode & 0o777).toString(8)}, size ${stats.size} bytes`,
      );
    } catch (error) {
      diagWarn(`Unable to stat bundle_dmg.sh: ${error.message}`);
    }
    try {
      const content = readFileSync(scriptPath, { encoding: "utf8" });
      const lines = content.replace(/\r\n/g, "\n").split("\n");
      const head = lines.slice(0, 40).join("\n");
      diagLog(`bundle_dmg.sh head:\n${head}`);
    } catch (error) {
      diagWarn(`Unable to read bundle_dmg.sh: ${error.message}`);
    }
  } else {
    diagWarn("bundle_dmg.sh not found after failure");
  }

  const dmgLog = path.join(bundleDir, "bundle.log");
  if (existsSync(dmgLog)) {
    try {
      const content = readFileSync(dmgLog, { encoding: "utf8" });
      const lines = content.replace(/\r\n/g, "\n").split("\n");
      const tail = lines.slice(-80).join("\n");
      diagLog(`bundle.log tail:\n${tail}`);
    } catch (error) {
      diagWarn(`Unable to read bundle.log: ${error.message}`);
    }
  } else {
    diagWarn("bundle.log not present in DMG bundle directory");
  }
}
