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
to generate embeddings (`torch`, `transformers`, and their dependencies). After
installation it rewrites the generated `pyvenv.cfg` so the environment's `home`
entry points back to the bundled runtime, keeping the interpreter relocatable in
the packaged installer. The runtime currently pins `torch==2.2.2` and
`transformers==4.56.2`, the newest versions that ship wheels for every platform
we target with Python 3.11. The Tauri bundler copies these resources into the
platform-specific installer so users do not need a system-wide Python
installation. When bundled, the runtime lives under
`<install dir>/resources/python/<platform>-<arch>`, allowing the Windows build to
expose `<install dir>/resources/python/<platform>-<arch>/Scripts/python.exe` for the
Rust backend.

The form displays a confirmation payload after validation so that you can review
which settings will be submitted to the backend service. File pickers fall back
silently if the operating system denies access; you can always paste a path into
the accompanying text field.

