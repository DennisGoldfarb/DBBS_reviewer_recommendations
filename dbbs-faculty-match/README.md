# Application developer guide

The [repository README](../README.md) covers setup, supported inputs, commands,
and packaging.

## Code map

```text
src/
  App.tsx                    Screen composition and shared feedback
  components/                Student inputs, faculty filters, dataset dialogs, results
    FileDropZone.tsx         Shared upload area and feedback
    PreviewTable.tsx         Shared tabular preview
  hooks/                     State and actions for each workflow
    useNativeFileDrop.ts     Native event lifecycle, hit testing, path validation
    useStudentInput.ts       Student input modes and spreadsheet column selections
    useFacultyFilters.ts     Program selection and roster identifier mappings
    useFacultyExclusions.ts Exclusion upload and column selections
    useFacultyDataset.ts     Dataset loading, import configuration, restore
    useEmbeddingHelper.ts    Helper readiness, retry, progress, refresh
    useMatching.ts           Matching requests and saving generated results
  services/
    backend.ts               Typed desktop command boundary
    fileDialogs.ts           Native file-picker boundary
  lib/                       Pure formatting, column, theme, and upload helpers
  styles/                    Theme, forms, previews, dataset, results, upload styles
  types.ts                   Frontend command payloads and response contracts
  test/                      Synthetic fixtures and test cleanup
src-tauri/src/
  lib.rs                     Plugin setup, startup, command registration
  commands.rs                Tauri commands and upload validation
  config.rs                  Resource names, model and event constants
  models.rs                  Serializable command, embedding, and result models
  matching.rs                Request validation and matching orchestration
  scoring.rs                 Similarity, faculty ordering, student rankings
  exclusions.rs              First/last name normalization and faculty exclusions
  directory_matching.rs      Folder extraction and batch matching
  spreadsheet_matching.rs    Student spreadsheet batch matching
  spreadsheets.rs            Tabular readers, column suggestions, normalization
  documents.rs               TXT, PDF, Word, and RTF extraction
  exports.rs                 Excel worksheets, formulas, and references
  dataset.rs                 Dataset analysis, metadata, replacement, persistence
  embeddings.rs              Refresh, progress, and Python sidecar management
  tests/                     Domain regression tests
python/
  embedding_helper.py        Persistent JSON-lines embedding worker
  faculty_preparation.py     Local LLMLingua-2 word pruning for faculty topic matching
  student_chunking.py        Complete student token coverage with overlapping windows
  tests/                     Dependency-free text and protocol tests
scripts/                     Development, sidecar, installer, and test entry points
```

Keep view markup in components, state transitions in workflow hooks, pure functions
in `lib/`, and desktop calls in `services/backend.ts`. Keep Rust domain code free
of command registration. If a command contract changes, update both `models.rs`
and `types.ts`/`backend.ts`, together with a regression test.

## Upload event flow

Tauri handles operating-system file drops before ordinary HTML drop handlers on
some platforms. `useNativeFileDrop` subscribes to `getCurrentWindow().onDragDropEvent`
and releases its subscription on unmount, including late registration during
React StrictMode cleanup. Windows physical coordinates are converted to viewport coordinates; macOS and
GTK already supply logical points. Scaling macOS coordinates again would miss
the upload area on Retina displays.

A drop must contain exactly one path. The frontend checks its extension; the Rust
`validate_upload_path` command checks existence, file versus folder, and extension.
A valid path enters the same state/preview flow as a file selected with Browse.
Drops outside an area, into disabled controls, or after unmount do not change input.

Changing a path clears its previous preview and column mappings. Request counters
prevent older preview responses from replacing newer selections. Replacing the
faculty dataset invalidates roster mappings and old matching results.

## Faculty exclusions

Use exclusions to omit letter writers, PIs, or any other faculty for each student.
The optional exclusion upload accepts the same spreadsheet formats as student
inputs. `Person Reference ID` and `Ref N Name` columns are recognized automatically;
users can change the column selections. Only faculty names are used. Select the
student spreadsheet ID column separately from its output identifier columns. For
a single prompt/document, enter the student ID; for folders, document filenames
without their extension must equal student IDs. IDs are compared as trimmed
text, preserving leading zeros. Students absent from the file have no exclusions.

`exclusions.rs` normalizes both faculty and exclusion names to lowercase `first last`,
moving the text before a comma to the end and collapsing whitespace. It does not
infer nicknames or missing name parts. Every faculty row sharing a normalized name
is excluded for that student. Repeated student records combine their faculty
exclusions. These exclusions intersect existing faculty filters before scoring
and selecting the top results, across all four input modes. No embedding refresh
is needed. Results and the existing workbook sheets contain eligible matches only.

## Dataset and embedding lifecycle

Packaged resources are copied to the application data directory on first use.
Faculty dataset metadata records its configured columns and program memberships.
A replacement is analyzed without altering the active dataset, then staged beside
the destination and renamed into place. Old faculty embeddings are invalidated
before replacement so a failed refresh cannot silently match the new dataset
against an old index. Missing embeddings produce an actionable error.

The Python helper retains loaded models across JSON-lines requests. Optional ML
dependencies load only when embeddings are requested, allowing protocol and text
normalization tests to run without torch, transformers, or a downloaded model.
The packaged faculty source remains unchanged.

## Local embeddings

The application uses only `NeuML/pubmedbert-base-embeddings`, pinned to revision
`b79526d6ef3645e0df4530322e266f24c829f5ef`. `python/embedding_config.py` defines
the model, revision, and 512-token input limit. Tokenizer and model load the same
revision once per helper process. Inference uses attention-mask mean pooling on
CPU, without retrieval prefixes or an extra normalization pass. Cosine similarity
handles vector normalization during ranking.

Faculty refreshes, imports, prompt matching, spreadsheet rows, and folder runs all
use PubMedBERT. **Update faculty embeddings** replaces an index with freshly
prepared PubMedBERT vectors. An experimental model index produces an actionable
refresh error before matching; its vectors are never mixed with PubMedBERT.
Legacy indexes with an empty model identifier continue to use PubMedBERT.

Faculty refresh requests enable `facultyPreparation`, also used automatically
on dataset import. "Lab Website" labels are removed. Remaining descriptions are
kept intact if they fit the 510-content-token budget plus two special tokens. Only overlong descriptions load the small
`microsoft/llmlingua-2-bert-base-multilingual-cased-meetingbank` classifier, pinned
to revision `5f0c82792b7ea14c6484e015b6a072009496b7f2`, in float32 on CPU.
The large XLM-RoBERTa model is not used. The small model has about 177 million
parameters and a 710 MB weights download, cached after first use.

The classifier scores the entire description through overlapping 512-token
windows with a 64-token overlap. Subword scores are averaged within source words.
Lowest-scoring whole words are deleted until the embedding tokenizer's budget
is met. There are no manual term-protection rules or generated summaries. The
pruning model is released before embedding. Failures stop the refresh instead
of silently truncating faculty text. Original spreadsheet text stays intact;
optional index preparation metadata records retained text and token counts.
Old `faculty_text_cache.json` summaries are never used.

## Student interests and long documents

`student_chunking.py` preserves all student text through 510-content-token
sections with 50-token overlap. Each section adds PubMedBERT's CLS and SEP tokens.
Token IDs are used directly, avoiding decode/re-encode losses at boundaries.
Inputs fitting one section retain the original short-text tokenizer path.

The helper returns the first vector in `embedding` and additional sections in
`additionalEmbeddings` on the same row ID. Rust scores each faculty member by the
maximum cosine similarity across that student's sections, then sorts and limits
the faculty list. Each faculty member appears once. Filters and reciprocal
student rankings use those final scores. This applies to pasted interests,
documents, spreadsheet rows, and folder runs. Student text is never pruned.

## Model decision

PubMedBERT was retained for its fast CPU inference and competitive topic matching.
Experimental pooling, task prefixes, alternative token limits, and model selection
are not part of the application or its packaged Python helper. Local input data,
model-comparison reports, and experiment caches are ignored by Git.

Use `scripts/rebuild-bundled-embeddings.py` with the embedding helper's Python
requirements to regenerate the packaged faculty index directly from the bundled
TSV. It uses the same production preparation and embedding pipeline as uploads;
it does not depend on evaluation caches or a user's active dataset.

Sidecar build signatures cover every top-level Python module. Builds clear their
project cache and verify that frozen bytecode matches source before success.

## Verification

Frontend tests use the real Tauri window adapter with a simulated IPC boundary
for macOS Retina and Windows payloads, and cover upload targets, native
listener cleanup, position routing, validation failures, Browse/drop preview
behavior, stale requests, form submission, locking, counts, and themes. Rust tests
exercise real parsing and export libraries: CSV/TSV input, DOCX extraction, ranking,
program filtering, dataset replacement, and reopening generated Excel workbooks.
Python tests cover text cleanup and worker protocol recovery without ML downloads.
They also verify faculty token boundaries, unchanged short inputs, footer cleanup,
whole-source window scoring, whole-word deletion, topic retention with controlled
scores, model failure cleanup, and ignoring old summary caches. Rust
tests cover preparation metadata compatibility and reporting helper failures.
Student regressions cover exact token coverage and overlap, limit boundaries,
unchanged short-input tokenization, grouped batch IDs, failing later chunks,
maximum similarity (including negative scores), faculty filters/limits,
deduplication and student-rank counts.
Exclusion tests cover name order/case/spacing, duplicate faculty, leading-zero IDs,
per-student filtering before the result limit, XLSX parsing, column mapping,
native exclusion uploads, and stale-preview removal.

Before a release, also smoke-test the packaged app on macOS and Windows:

- Wait for helper readiness; run a prompt against the packaged faculty dataset.
- Drop a valid document and a folder containing several documents into their areas.
- Drop a student spreadsheet, verify column suggestions, and run matching.
- Match a long input with a different topic near the end. Verify all sections
  are embedded, the later topic can affect matches, and each faculty appears once.
  Repeat using a spreadsheet and folder, checking that sections stay with the
  correct student/document and short inputs keep their previous results.
- Drop a faculty roster, check its identifier mapping, and apply a program filter.
- Drop a replacement faculty dataset, confirm columns, refresh embeddings, and restore the default.
- Import faculty descriptions exceeding the active model's context. Check local
  LLMLingua-2 loading/pruning progress and preserved originals. Verify the real
  input fits the model's total token budget and metadata records
  `truncated: false`. Verify a 1,000-token description is pruned and
  PubMedBERT still respects its 512-token input window.
  Check important topic terms
  from the ends of long descriptions, and test a missing-model download failure.
- Try an incompatible file and multiple items; confirm that errors are visible.
- Exercise Browse and pasted paths; save a workbook and open its three worksheets.
- Repeat a drop on a scaled/Retina display, after scrolling, and after switching themes.

The automated native-event tests do not replace a real Finder/Explorer drag gesture
or a packaged model inference test. The code-check workflow runs formatting, builds,
and regression suites on macOS and Windows; installer verification remains separate.
