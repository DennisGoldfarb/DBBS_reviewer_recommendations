import { FormEvent, useMemo, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { open } from "@tauri-apps/plugin-dialog";
import "./App.css";

type TaskType = "prompt" | "document" | "spreadsheet" | "directory";
type FacultyScope = "all" | "program" | "custom";

interface FacultyFormEntry {
  firstName: string;
  lastName: string;
  identifier: string;
}

interface FacultyEntry {
  firstName: string;
  lastName: string;
  identifier?: string;
}

interface PathConfirmation {
  label: string;
  path: string;
}

interface SubmissionDetails {
  taskType: TaskType;
  facultyScope: FacultyScope;
  validatedPaths: PathConfirmation[];
  programFilters: string[];
  customFaculty: FacultyEntry[];
  recommendationsPerStudent: number;
  recommendationsPerFaculty: number;
  updateEmbeddings: boolean;
  promptPreview?: string;
}

interface SubmissionResponse {
  summary: string;
  warnings: string[];
  details: SubmissionDetails;
}

interface StatusMessage {
  variant: "success" | "error" | "info";
  message: string;
}

const createFacultyFormEntry = (): FacultyFormEntry => ({
  firstName: "",
  lastName: "",
  identifier: "",
});

const parsePrograms = (raw: string): string[] =>
  raw
    .split(/[\n,]/)
    .map((value) => value.trim())
    .filter((value) => value.length > 0);

const sanitizeFacultyEntries = (entries: FacultyFormEntry[]): FacultyEntry[] =>
  entries
    .map((entry) => ({
      firstName: entry.firstName.trim(),
      lastName: entry.lastName.trim(),
      identifier: entry.identifier.trim() || undefined,
    }))
    .filter((entry) => entry.firstName.length > 0 && entry.lastName.length > 0);

function App() {
  const [taskType, setTaskType] = useState<TaskType>("prompt");
  const [promptText, setPromptText] = useState("");
  const [documentPath, setDocumentPath] = useState("");
  const [spreadsheetPath, setSpreadsheetPath] = useState("");
  const [directoryPath, setDirectoryPath] = useState("");
  const [facultyScope, setFacultyScope] = useState<FacultyScope>("all");
  const [programInput, setProgramInput] = useState("");
  const [customFaculty, setCustomFaculty] = useState<FacultyFormEntry[]>([]);
  const [facultyRecCount, setFacultyRecCount] = useState("3");
  const [studentRecCount, setStudentRecCount] = useState("0");
  const [updateEmbeddings, setUpdateEmbeddings] = useState(false);

  const [result, setResult] = useState<SubmissionResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isUpdatingEmbeddings, setIsUpdatingEmbeddings] = useState(false);
  const [embeddingStatus, setEmbeddingStatus] = useState<StatusMessage | null>(null);

  const sanitizedCustomFaculty = useMemo(
    () => sanitizeFacultyEntries(customFaculty),
    [customFaculty],
  );

  const handleTaskTypeChange = (value: TaskType) => {
    setTaskType(value);
    setError(null);
    setResult(null);
  };

  const handleFacultyScopeChange = (value: FacultyScope) => {
    setFacultyScope(value);
    setError(null);
    setResult(null);

    if (value === "custom" && customFaculty.length === 0) {
      setCustomFaculty([createFacultyFormEntry()]);
    }
  };

  const handleFileSelection = async (
    setter: (value: string) => void,
    options: Parameters<typeof open>[0],
  ) => {
    try {
      const selection = await open(options);
      if (typeof selection === "string") {
        setter(selection);
      }
    } catch (selectionError) {
      const message =
        selectionError instanceof Error
          ? selectionError.message
          : String(selectionError);
      setError(`Unable to open a file dialog: ${message}`);
    }
  };

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setIsSubmitting(true);
    setError(null);
    setResult(null);

    const programFilters = parsePrograms(programInput);
    const facultyRecommendations = Math.max(
      1,
      Number.parseInt(facultyRecCount, 10) || 0,
    );
    const studentRecommendations = Math.max(
      0,
      Number.parseInt(studentRecCount, 10) || 0,
    );

    try {
      const response = await invoke<SubmissionResponse>(
        "submit_matching_request",
        {
          payload: {
            taskType,
            promptText: taskType === "prompt" ? promptText : undefined,
            documentPath:
              taskType === "document" && documentPath.trim().length > 0
                ? documentPath.trim()
                : undefined,
            spreadsheetPath:
              taskType === "spreadsheet" && spreadsheetPath.trim().length > 0
                ? spreadsheetPath.trim()
                : undefined,
            directoryPath:
              taskType === "directory" && directoryPath.trim().length > 0
                ? directoryPath.trim()
                : undefined,
            facultyScope,
            programFilters:
              facultyScope === "program" && programFilters.length > 0
                ? programFilters
                : undefined,
            customFaculty:
              facultyScope === "custom" && sanitizedCustomFaculty.length > 0
                ? sanitizedCustomFaculty
                : undefined,
            facultyRecsPerStudent: facultyRecommendations,
            studentRecsPerFaculty: studentRecommendations,
            updateEmbeddings,
          },
        },
      );

      setResult(response);
    } catch (submissionError) {
      const message =
        submissionError instanceof Error
          ? submissionError.message
          : String(submissionError);
      setError(message);
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleEmbeddingsUpdate = async () => {
    setIsUpdatingEmbeddings(true);
    setEmbeddingStatus(null);

    try {
      const message = await invoke<string>("update_faculty_embeddings");
      setEmbeddingStatus({ variant: "success", message });
    } catch (updateError) {
      const message =
        updateError instanceof Error ? updateError.message : String(updateError);
      setEmbeddingStatus({ variant: "error", message });
    } finally {
      setIsUpdatingEmbeddings(false);
    }
  };

  return (
    <div className="app-shell">
      <main>
        <header className="page-header">
          <h1>Reviewer Recommendation Console</h1>
          <p className="description">
            Configure matching runs, narrow the faculty roster, and track the
            options that will be sent to the matching backend.
          </p>
        </header>

        <form className="matching-form" onSubmit={handleSubmit}>
          <fieldset>
            <legend>Student inputs</legend>
            <p className="section-description">
              Select how student research interests will be supplied. The
              current build validates paths and settings without launching
              matching jobs.
            </p>

            <div className="radio-group">
              <label className="radio-option">
                <input
                  type="radio"
                  name="input-type"
                  value="prompt"
                  checked={taskType === "prompt"}
                  onChange={() => handleTaskTypeChange("prompt")}
                />
                <span>Single prompt</span>
              </label>
              <label className="radio-option">
                <input
                  type="radio"
                  name="input-type"
                  value="document"
                  checked={taskType === "document"}
                  onChange={() => handleTaskTypeChange("document")}
                />
                <span>Document file</span>
              </label>
              <label className="radio-option">
                <input
                  type="radio"
                  name="input-type"
                  value="spreadsheet"
                  checked={taskType === "spreadsheet"}
                  onChange={() => handleTaskTypeChange("spreadsheet")}
                />
                <span>Spreadsheet</span>
              </label>
              <label className="radio-option">
                <input
                  type="radio"
                  name="input-type"
                  value="directory"
                  checked={taskType === "directory"}
                  onChange={() => handleTaskTypeChange("directory")}
                />
                <span>Directory of files</span>
              </label>
            </div>

            {taskType === "prompt" && (
              <div className="input-stack">
                <label htmlFor="prompt-text">Prompt text</label>
                <textarea
                  id="prompt-text"
                  value={promptText}
                  onChange={(event) => setPromptText(event.target.value)}
                  placeholder="Describe the student's research interests..."
                />
                <p className="small-note">
                  Provide a single free-form description that will be embedded
                  and compared against faculty profiles.
                </p>
              </div>
            )}

            {taskType === "document" && (
              <div className="input-stack">
                <label>Research interest document</label>
                <div className="button-row inline">
                  <button
                    type="button"
                    className="secondary"
                    onClick={() =>
                      handleFileSelection(setDocumentPath, {
                        multiple: false,
                        filters: [
                          {
                            name: "Supported documents",
                            extensions: ["txt", "pdf", "doc", "docx"],
                          },
                        ],
                      })
                    }
                  >
                    Browse…
                  </button>
                  <input
                    type="text"
                    value={documentPath}
                    onChange={(event) => setDocumentPath(event.target.value)}
                    placeholder="Paste or confirm the document path"
                  />
                </div>
                {documentPath && (
                  <div className="path-preview">{documentPath}</div>
                )}
              </div>
            )}

            {taskType === "spreadsheet" && (
              <div className="input-stack">
                <label>Spreadsheet with prompts</label>
                <div className="button-row inline">
                  <button
                    type="button"
                    className="secondary"
                    onClick={() =>
                      handleFileSelection(setSpreadsheetPath, {
                        multiple: false,
                        filters: [
                          {
                            name: "Spreadsheets",
                            extensions: ["csv", "tsv", "xlsx", "xls", "ods"],
                          },
                        ],
                      })
                    }
                  >
                    Browse…
                  </button>
                  <input
                    type="text"
                    value={spreadsheetPath}
                    onChange={(event) => setSpreadsheetPath(event.target.value)}
                    placeholder="Paste or confirm the spreadsheet path"
                  />
                </div>
                {spreadsheetPath && (
                  <div className="path-preview">{spreadsheetPath}</div>
                )}
                <p className="small-note">
                  Each row should include an identifier column (for example,
                  name or ID) and a column containing the student's prompt.
                </p>
              </div>
            )}

            {taskType === "directory" && (
              <div className="input-stack">
                <label>Directory containing documents</label>
                <div className="button-row inline">
                  <button
                    type="button"
                    className="secondary"
                    onClick={() =>
                      handleFileSelection(setDirectoryPath, {
                        directory: true,
                      })
                    }
                  >
                    Browse…
                  </button>
                  <input
                    type="text"
                    value={directoryPath}
                    onChange={(event) => setDirectoryPath(event.target.value)}
                    placeholder="Paste or confirm the directory path"
                  />
                </div>
                {directoryPath && (
                  <div className="path-preview">{directoryPath}</div>
                )}
                <p className="small-note">
                  Each document in the folder will be treated as a separate
                  student submission. The filename will be used as the
                  identifier.
                </p>
              </div>
            )}
          </fieldset>

          <fieldset>
            <legend>Faculty filters</legend>
            <p className="section-description">
              Limit the available faculty before matching to respect program
              boundaries or user-specified lists.
            </p>

            <div className="radio-group">
              <label className="radio-option">
                <input
                  type="radio"
                  name="faculty-scope"
                  value="all"
                  checked={facultyScope === "all"}
                  onChange={() => handleFacultyScopeChange("all")}
                />
                <span>All faculty</span>
              </label>
              <label className="radio-option">
                <input
                  type="radio"
                  name="faculty-scope"
                  value="program"
                  checked={facultyScope === "program"}
                  onChange={() => handleFacultyScopeChange("program")}
                />
                <span>Limit to programs</span>
              </label>
              <label className="radio-option">
                <input
                  type="radio"
                  name="faculty-scope"
                  value="custom"
                  checked={facultyScope === "custom"}
                  onChange={() => handleFacultyScopeChange("custom")}
                />
                <span>Provide a faculty list</span>
              </label>
            </div>

            {facultyScope === "program" && (
              <div className="input-stack">
                <label htmlFor="programs">Programs or tracks</label>
                <textarea
                  id="programs"
                  value={programInput}
                  onChange={(event) => setProgramInput(event.target.value)}
                  placeholder="One program per line, or separate entries with commas"
                />
                <p className="small-note">
                  You can specify up to four programs per faculty member.
                  Duplicate names are removed automatically.
                </p>
              </div>
            )}

            {facultyScope === "custom" && (
              <div className="input-stack">
                <div className="faculty-entry-list">
                  {customFaculty.map((entry, index) => (
                    <div className="faculty-entry-row" key={index}>
                      <input
                        type="text"
                        placeholder="First name"
                        value={entry.firstName}
                        onChange={(event) => {
                          const value = event.target.value;
                          setCustomFaculty((previous) =>
                            previous.map((item, itemIndex) =>
                              itemIndex === index
                                ? { ...item, firstName: value }
                                : item,
                            ),
                          );
                        }}
                      />
                      <input
                        type="text"
                        placeholder="Last name"
                        value={entry.lastName}
                        onChange={(event) => {
                          const value = event.target.value;
                          setCustomFaculty((previous) =>
                            previous.map((item, itemIndex) =>
                              itemIndex === index
                                ? { ...item, lastName: value }
                                : item,
                            ),
                          );
                        }}
                      />
                      <input
                        type="text"
                        placeholder="Identifier (optional)"
                        value={entry.identifier}
                        onChange={(event) => {
                          const value = event.target.value;
                          setCustomFaculty((previous) =>
                            previous.map((item, itemIndex) =>
                              itemIndex === index
                                ? { ...item, identifier: value }
                                : item,
                            ),
                          );
                        }}
                      />
                      <button
                        type="button"
                        className="remove-button"
                        onClick={() =>
                          setCustomFaculty((previous) =>
                            previous.filter((_, itemIndex) => itemIndex !== index),
                          )
                        }
                      >
                        Remove
                      </button>
                    </div>
                  ))}
                </div>
                <button
                  type="button"
                  className="add-entry-button"
                  onClick={() =>
                    setCustomFaculty((previous) => [
                      ...previous,
                      createFacultyFormEntry(),
                    ])
                  }
                >
                  Add faculty entry
                </button>
                <p className="small-note">
                  Provide first and last names. Include an internal identifier
                  when available to avoid name collisions.
                </p>
              </div>
            )}
          </fieldset>

          <fieldset>
            <legend>Recommendation settings</legend>
            <div className="number-row">
              <label>
                Faculty recommendations per student
                <input
                  type="number"
                  min={1}
                  value={facultyRecCount}
                  onChange={(event) => setFacultyRecCount(event.target.value)}
                />
              </label>
              <label>
                Student recommendations per faculty
                <input
                  type="number"
                  min={0}
                  value={studentRecCount}
                  onChange={(event) => setStudentRecCount(event.target.value)}
                />
              </label>
            </div>
            <div className="checkbox-row">
              <input
                id="update-embeddings"
                type="checkbox"
                checked={updateEmbeddings}
                onChange={(event) => setUpdateEmbeddings(event.target.checked)}
              />
              <label htmlFor="update-embeddings">
                Refresh faculty embeddings before running this match
              </label>
            </div>
          </fieldset>

          <div className="button-row">
            <button type="submit" disabled={isSubmitting}>
              {isSubmitting ? "Validating…" : "Validate matching request"}
            </button>
            <button
              type="button"
              className="secondary"
              onClick={handleEmbeddingsUpdate}
              disabled={isUpdatingEmbeddings}
            >
              {isUpdatingEmbeddings
                ? "Checking embeddings…"
                : "Update faculty embeddings"}
            </button>
          </div>
        </form>

        {error && (
          <div className="status-banner status-error">{error}</div>
        )}

        {result && (
          <section className="result-card">
            <h2>Submission ready</h2>
            <p>{result.summary}</p>

            {result.warnings.length > 0 && (
              <ul className="warning-list">
                {result.warnings.map((warning, index) => (
                  <li key={index}>{warning}</li>
                ))}
              </ul>
            )}

            <div className="detail-grid">
              <div className="detail-card">
                <h3>Input configuration</h3>
                <dl>
                  <dt>Type</dt>
                  <dd>{result.details.taskType}</dd>
                  {result.details.promptPreview && (
                    <>
                      <dt>Prompt preview</dt>
                      <dd>
                        <pre className="prompt-preview">
                          {result.details.promptPreview}
                        </pre>
                      </dd>
                    </>
                  )}
                  {result.details.validatedPaths.length > 0 && (
                    <>
                      <dt>Validated paths</dt>
                      <dd>
                        <ul className="path-list">
                          {result.details.validatedPaths.map((path) => (
                            <li key={`${path.label}-${path.path}`}>
                              <strong>{path.label}:</strong> {path.path}
                            </li>
                          ))}
                        </ul>
                      </dd>
                    </>
                  )}
                </dl>
              </div>

              <div className="detail-card">
                <h3>Faculty scope</h3>
                <dl>
                  <dt>Scope</dt>
                  <dd>{result.details.facultyScope}</dd>
                  {result.details.programFilters.length > 0 && (
                    <>
                      <dt>Programs</dt>
                      <dd>
                        <ul className="path-list">
                          {result.details.programFilters.map((program) => (
                            <li key={program}>{program}</li>
                          ))}
                        </ul>
                      </dd>
                    </>
                  )}
                  {result.details.customFaculty.length > 0 && (
                    <>
                      <dt>Faculty list</dt>
                      <dd>
                        <table className="faculty-table">
                          <thead>
                            <tr>
                              <th>Name</th>
                              <th>Identifier</th>
                            </tr>
                          </thead>
                          <tbody>
                            {result.details.customFaculty.map((entry, index) => (
                              <tr key={`${entry.firstName}-${entry.lastName}-${index}`}>
                                <td>
                                  {entry.firstName} {entry.lastName}
                                </td>
                                <td>{entry.identifier ?? "—"}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </dd>
                    </>
                  )}
                </dl>
              </div>

              <div className="detail-card">
                <h3>Recommendation limits</h3>
                <dl>
                  <dt>Faculty per student</dt>
                  <dd>{result.details.recommendationsPerStudent}</dd>
                  <dt>Students per faculty</dt>
                  <dd>{result.details.recommendationsPerFaculty}</dd>
                  <dt>Refresh embeddings</dt>
                  <dd>{result.details.updateEmbeddings ? "Yes" : "No"}</dd>
                </dl>
              </div>
            </div>
          </section>
        )}

        {embeddingStatus && (
          <div
            className={`status-banner ${
              embeddingStatus.variant === "success"
                ? "status-success"
                : embeddingStatus.variant === "error"
                  ? "status-error"
                  : "status-info"
            }`}
          >
            {embeddingStatus.message}
          </div>
        )}
      </main>
    </div>
  );
}

export default App;
