# DBBS Reviewer Recommendations

The DBBS Reviewer Recommendations desktop app pairs students with faculty based on
shared research interests, for admissions reviews, qualifying exam committees,
and rotations.

The React interface calls a Rust backend; a bundled Python helper generates
PubMedBERT embeddings locally on CPU. The model may need to be downloaded the
first time it runs.

## Using the app

Dark mode is the default for the interface and native window. The theme toggle
can switch to light mode, and explicit choices are saved.

1. Supply a research prompt, a document, a spreadsheet of students, or a folder of
   documents. Browse, paste a path, or drop an item onto its labeled upload area.
2. Choose all faculty, filter by program, or supply a faculty roster and map its
   identifier columns to those in the active dataset.
3. Optionally upload an **Exclusion spreadsheet** to omit faculty for each student.
   Map its student ID and faculty name columns to the student inputs and faculty
   dataset. This can exclude letter writers, PIs, or other faculty.
4. Enter the number of faculty recommendations per student and select **Run matching**.
5. Review the similarities, faculty ranks, and student ranks. Spreadsheet and folder
   runs can be saved as an Excel workbook with **Matches**, **Student Summary**, and
   **Faculty Summary** worksheets.

Documents support TXT, PDF, DOCX, RTF, and RTF/text-compatible DOC files. Legacy
binary DOC support is limited; convert those documents to DOCX or PDF if extraction
fails. Student spreadsheets and faculty filter lists support CSV, TSV/TXT, XLSX,
and XLS. Faculty dataset replacements support TSV/TXT and XLSX/XLS.

Drop one file or folder at a time onto the corresponding area. Dropped spreadsheets
load their column previews automatically. A folder run processes the files directly
inside that folder; it does not recurse into subfolders. File browsing and pasted
paths remain available alongside drag and drop.

Long student interests and documents are matched using overlapping sections:
510 content tokens per section with a 50-token overlap. PubMedBERT's 512-token
input window includes two special tokens. Every section is embedded
locally, and each faculty member's score is their best section match. This covers
pasted interests, individual documents, student spreadsheet rows, and folder runs.
Short inputs keep their existing behavior. Longer inputs take more processing time.

The **Faculty dataset** area lets you preview the packaged dataset, configure a
replacement, or restore the default. Replacements validate the selected columns
before changing the active file. Importing a dataset refreshes its embeddings;
if that refresh fails, use **Update faculty embeddings** before matching again.
During import or **Update faculty embeddings**, "Lab Website" labels are removed
and descriptions exceeding the active model's context are pruned with the
small BERT-based LLMLingua-2 model on CPU. It selects existing words from the
whole description for topic matching; it does not generate a summary. Shorter
descriptions receive no pruning. PubMedBERT's budget is 510 content tokens plus
two special tokens. The uploaded spreadsheet retains its
original text. First use downloads the pruning model (about 710 MB); subsequent
use can run offline alongside the embedding model. **Update faculty embeddings**
regenerates the index with PubMedBERT. The application uses this one embedding
model, pinned to a fixed revision for reproducible results. PubMedBERT was retained
for fast local CPU inference after comparing alternative models.

## Development

Work in `dbbs-faculty-match/`. Requirements are Node.js 22.12+ (22.x, 24.x, or 26+),
a stable Rust toolchain, the [Tauri platform prerequisites](https://v2.tauri.app/start/prerequisites/),
and Python 3.11+ for building the embedding helper. Python 3.11 is used in CI and
is the recommended sidecar build environment.

```sh
cd dbbs-faculty-match
npm ci
npm run build-sidecars
npm run tauri dev
```

Building the sidecar installs the Python requirements in an isolated environment
and bundles the helper with PyInstaller. It can take time and requires network
access. `npm run dev` starts a browser preview of the interface; matching, native
dialogs, and filesystem drag and drop require the desktop application.
Set `DDBS_SIDECAR_PYTHON` to an existing Python interpreter with the required
dependencies and PyInstaller installed to reuse a prepared build environment.

```sh
npm run check                         # formatting, TypeScript, production frontend, UI tests
npm run test:backend                  # real Rust unit tests, no packaged ML runtime needed
npm run lint:backend                  # strict Rust static analysis
python3 -m unittest discover -s python/tests -v
cargo fmt --manifest-path src-tauri/Cargo.toml --all --check
```

Use `npm run test:watch` while working on the interface and `npm run format` before
review. Commit both `package-lock.json` and `src-tauri/Cargo.lock` when dependencies
change. The code-check workflow runs on macOS and Windows; the packaging workflow
continues to build the installers separately.

```sh
npm run tauri build
```

Installer builds prepare the real embedding sidecar automatically. Unit tests
exclude that bundled executable through a process-local Tauri configuration;
they do not change installer configuration or download the embedding model.

Release builds derive the app and installer version from the release tag. Both
`v1.2.3` and `1.2.3` produce version `1.2.3`; prerelease tags are supported too.
Manual releases use the supplied tag and check out that tag's code. CI validates
the tag before creating the GitHub release, then stamps the Tauri, npm, and Cargo
manifests and lockfiles before compiling. These generated version changes are
not committed back to the repository. To check this logic locally, run
`node --test scripts/release-version.test.mjs` from `dbbs-faculty-match/`.

For architecture, maintenance guidance, and the manual desktop smoke test, see the
[application developer guide](dbbs-faculty-match/README.md).
