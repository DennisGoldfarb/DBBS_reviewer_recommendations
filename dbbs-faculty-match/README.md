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

4. (Optional) Compile the Python embedding helper into the sidecar binary that
   ships with the desktop installers. This step requires Python **3.11** to be
   installed and available on your `PATH`, runs automatically during
   `tauri build`, and can be invoked manually to verify dependency installation:

   ```bash
   npm run build-sidecars
   ```

The `build-sidecars` script creates an isolated virtual environment under
`python/.sidecar-build`, installs the Python dependencies required to generate
embeddings (`torch`, `transformers`, and their dependencies), and bundles the
`embedding_helper.py` script with PyInstaller. The resulting executable is
placed in `src-tauri/binaries` and embedded as a Tauri sidecar so users do not
need a system-wide Python installation.

The form displays a confirmation payload after validation so that you can review
which settings will be submitted to the backend service. File pickers fall back
silently if the operating system denies access; you can always paste a path into
the accompanying text field.

