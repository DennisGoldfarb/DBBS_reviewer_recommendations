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

## macOS signing and notarization

Apple's notarization service requires both the primary Tauri executable and the
`embedding-helper` sidecar to be signed with the hardened runtime and a secure
timestamp before you submit a DMG. The `scripts/create-dmg.mjs` helper now signs
every binary inside `DBBS Faculty Match.app/Contents/MacOS` automatically when
an identity is provided via environment variables, ensuring the notarization log
no longer reports missing hardened runtime or timestamp metadata.

1. Build the release bundle on macOS:

   ```bash
   npm run tauri -- build
   ```

2. Export your Developer ID Application identity and (optionally) an
   entitlements file path so the DMG helper can codesign on your behalf:

   ```bash
   export APPLE_CODESIGN_IDENTITY="Developer ID Application: Jane Doe (TEAMID1234)"
   export APPLE_CODESIGN_ENTITLEMENTS="/absolute/path/to/entitlements.plist" # optional
   ```

3. Create the DMG, which will sign the app bundle, the `dbbs-faculty-match`
   binary, and the `embedding-helper` sidecar with `--options runtime` and
   `--timestamp` before packaging:

   ```bash
   node ./scripts/create-dmg.mjs
   ```

4. Submit the resulting DMG for notarization using the keychain profile created
   with `xcrun notarytool store-credentials`, and staple the ticket after
   approval:

   ```bash
   xcrun notarytool submit src-tauri/target/release/bundle/dmg/DBBS\ Faculty\ Match_0.1.0_macos_arm64.dmg \
     --keychain-profile "notary-profile" \
     --wait
   xcrun stapler staple src-tauri/target/release/bundle/dmg/DBBS\ Faculty\ Match_0.1.0_macos_arm64.dmg
   ```

If you prefer to run `codesign` manually, sign each executable in
`Contents/MacOS` with `--options runtime --timestamp` before signing the app
bundle itself. Verifying with `codesign --verify --deep --strict --verbose=2` on
the `.app` bundle should complete without errors before submitting to Apple.

