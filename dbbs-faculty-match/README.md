# Reviewer Recommendation Console

A Tauri + React desktop interface for configuring student-to-faculty matching
runs. This build focuses on validating configuration choices and verifying that
referenced files and directories exist before the matching backend is wired in.

## Available tasks

- **Prompt validation** – enter a single student description for matching.
- **Document validation** – confirm a text, PDF, or Word document exists.
- **Spreadsheet validation** – check a tabular file containing prompts.
- **Directory validation** – verify a folder of student documents is available.
- **Faculty embedding refresh** – trigger a placeholder confirmation that an
  embedding refresh would run.

## Getting started

1. Install dependencies:

   ```bash
   npm install
   ```

2. Run the desktop application in development mode:

   ```bash
   npm run tauri dev
   ```

3. Build the frontend for production:

   ```bash
   npm run build
   ```

4. (Optional) Prepare the embedded Python runtime that ships with the desktop
   installers. This step requires Python **3.11** to be installed and available
   on your `PATH`, runs automatically during `tauri build`, and can be invoked
   manually to verify dependency installation:

   ```bash
   npm run prepare-python
   ```

   The `prepare-python` script creates an isolated virtual environment under
   `src-tauri/resources/python/<platform>-<arch>` and installs the packages needed
   to generate embeddings (`torch`, `transformers`, and their dependencies). The
   runtime currently pins `torch==2.2.2` and `transformers==4.56.2`, the newest
   versions that ship wheels for every platform we target with Python 3.11. The
   Tauri bundler copies these resources into the platform-specific installer so
   users do not need a system-wide Python installation.

## Troubleshooting desktop builds

Pass the `--diagnostics` flag or set `DBBS_TAURI_DEBUG=1` when invoking the Tauri
wrapper to collect additional environment data and SetFile shim logs during macOS
builds:

```bash
DBBS_TAURI_DEBUG=1 npm run tauri build
# or
npm run tauri -- --diagnostics build
```

When diagnostics are enabled the wrapper:

- Prints the detected platform, Node.js version, and `PATH` entries before
  invoking the Tauri CLI.
- Records which `SetFile` implementation (native or shim) will be used and where
  the shim log is written.
- Runs macOS sanity checks (`sw_vers`, `sysctl`, `which xattr`, `python3 --version`)
  and executes a SetFile shim self-test when the shim is active.
- Dumps the tail of the shim log and inspects the generated `bundle_dmg.sh` and
  `bundle.log` artifacts if the Tauri CLI exits with a non-zero status.

The shim trace level (`DBBS_SETFILE_TRACE`) is automatically elevated when
diagnostics are enabled. You can override this behaviour by exporting
`DBBS_SETFILE_TRACE=0` before running the command if the log volume becomes
excessive.

The form displays a confirmation payload after validation so that you can review
which settings will be submitted to the backend service. File pickers fall back
silently if the operating system denies access; you can always paste a path into
the accompanying text field.
