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
   installers. The build now vendors the official CPython **3.11** embeddable
   distribution on Windows and reuses virtual environments on macOS/Linux. The
   script runs automatically during `tauri build` and can be invoked manually to
   verify that the dependencies and Python payload are staged correctly:

   ```bash
   npm run prepare-python
   ```

On Windows the script:

1. Looks for vendored dependencies under
   `python/vendor/windows-x86_64/site-packages` (populate this directory from a
   Windows virtual environment that matches Python 3.11 x86_64).
2. Downloads (or reuses) the official `python-3.11.9-embed-amd64.zip` archive
   and extracts it into `src-tauri/resources/python/`.
3. Enables `site-packages` via `python311._pth` and copies the contents of
   `python/app` next to the interpreter as your application payload.

macOS and Linux builds continue to create an isolated virtual environment under
`src-tauri/resources/python/<platform>-<arch>` using the Python 3.11 interpreter
available on your machine. In all cases the packaged runtime is self-contained,
bundles the dependencies required for embedding generation (currently
`torch==2.2.2` and `transformers==4.56.2`), and ships with the installer so end
users do not need a system-wide Python installation.

The form displays a confirmation payload after validation so that you can review
which settings will be submitted to the backend service. File pickers fall back
silently if the operating system denies access; you can always paste a path into
the accompanying text field.

