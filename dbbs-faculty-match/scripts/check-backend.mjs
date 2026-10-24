import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

// Tests and static checks use the real Rust code without requiring a packaged ML sidecar.
// Installer builds continue to use tauri.conf.json and include the sidecar.
const config = JSON.parse(process.env.TAURI_CONFIG || "{}");
const [mode = "test", ...extraArgs] = process.argv.slice(2);
const args =
  mode === "lint"
    ? [
        "clippy",
        "--locked",
        "--lib",
        "--tests",
        ...extraArgs,
        "--",
        "-D",
        "warnings",
      ]
    : ["test", "--locked", "--lib", ...extraArgs];
const result = spawnSync("cargo", args, {
  cwd: fileURLToPath(new URL("../src-tauri", import.meta.url)),
  env: {
    ...process.env,
    TAURI_CONFIG: JSON.stringify({
      ...config,
      bundle: { ...config.bundle, externalBin: [] },
    }),
  },
  stdio: "inherit",
});
if (result.error) console.error(result.error.message);
process.exit(result.status ?? 1);
