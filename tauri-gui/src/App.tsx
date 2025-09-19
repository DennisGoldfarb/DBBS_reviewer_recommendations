import { FormEvent, useCallback, useEffect, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { open } from "@tauri-apps/plugin-dialog";
import "./App.css";

type TaskType = "prompt" | "document" | "spreadsheet" | "directory";
type FacultyScope = "all" | "program" | "custom";

type ProgramName = string;

interface PathConfirmation {
  label: string;
  path: string;
}

interface SpreadsheetPreview {
  headers: string[];
  rows: string[][];
  suggestedPromptColumns: number[];
  suggestedIdentifierColumns: number[];
}

interface FacultyDatasetAnalysis {
  embeddingColumns: string[];
  identifierColumns: string[];
  programColumns: string[];
  availablePrograms: string[];
}

interface FacultyDatasetPreviewResult {
  preview: SpreadsheetPreview;
  suggestedEmbeddingColumns: number[];
  suggestedIdentifierColumns: number[];
  suggestedProgramColumns: number[];
}

interface FacultyDatasetStatus {
  path: string | null;
  canonicalPath: string | null;
  sourcePath: string | null;
  lastModified: string | null;
  rowCount: number | null;
  columnCount: number | null;
  isValid: boolean;
  isDefault: boolean;
  message: string | null;
  messageVariant: StatusMessage["variant"] | null;
  preview: SpreadsheetPreview | null;
  analysis: FacultyDatasetAnalysis | null;
}

interface SubmissionDetails {
  taskType: TaskType;
  facultyScope: FacultyScope;
  validatedPaths: PathConfirmation[];
  programFilters: string[];
  customFacultyPath: string | null;
  recommendationsPerStudent: number;
  recommendationsPerFaculty: number;
  promptPreview?: string;
  spreadsheetPromptColumns: string[];
  spreadsheetIdentifierColumns: string[];
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

function App() {
  const [taskType, setTaskType] = useState<TaskType>("prompt");
  const [promptText, setPromptText] = useState("");
  const [documentPath, setDocumentPath] = useState("");
  const [spreadsheetPath, setSpreadsheetPath] = useState("");
  const [directoryPath, setDirectoryPath] = useState("");
  const [facultyScope, setFacultyScope] = useState<FacultyScope>("all");
  const [availablePrograms, setAvailablePrograms] = useState<ProgramName[]>([]);
  const [selectedPrograms, setSelectedPrograms] = useState<ProgramName[]>([]);
  const [customFacultyPath, setCustomFacultyPath] = useState("");
  const [facultyRecCount, setFacultyRecCount] = useState("10");
  const [studentRecCount, setStudentRecCount] = useState("0");

  const [spreadsheetPreview, setSpreadsheetPreview] =
    useState<SpreadsheetPreview | null>(null);
  const [spreadsheetPreviewError, setSpreadsheetPreviewError] =
    useState<string | null>(null);
  const [isLoadingSpreadsheetPreview, setIsLoadingSpreadsheetPreview] =
    useState(false);
  const [selectedIdentifierColumns, setSelectedIdentifierColumns] = useState<
    number[]
  >([]);
  const [selectedPromptColumns, setSelectedPromptColumns] = useState<number[]>(
    [],
  );

  const [result, setResult] = useState<SubmissionResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isUpdatingEmbeddings, setIsUpdatingEmbeddings] = useState(false);
  const [embeddingStatus, setEmbeddingStatus] = useState<StatusMessage | null>(null);
  const [datasetStatus, setDatasetStatus] =
    useState<FacultyDatasetStatus | null>(null);
  const [isDatasetLoading, setIsDatasetLoading] = useState(true);
  const [isDatasetBusy, setIsDatasetBusy] = useState(false);
  const [datasetBanner, setDatasetBanner] = useState<StatusMessage | null>(null);
  const [isDatasetPreviewOpen, setIsDatasetPreviewOpen] = useState(false);
  const [isDatasetConfigurationOpen, setIsDatasetConfigurationOpen] =
    useState(false);
  const [datasetConfigurationPath, setDatasetConfigurationPath] =
    useState("");
  const [datasetConfigurationPreview, setDatasetConfigurationPreview] =
    useState<SpreadsheetPreview | null>(null);
  const [isLoadingDatasetConfiguration, setIsLoadingDatasetConfiguration] =
    useState(false);
  const [
    datasetConfigurationEmbeddingColumns,
    setDatasetConfigurationEmbeddingColumns,
  ] = useState<number[]>([]);
  const [
    datasetConfigurationIdentifierColumns,
    setDatasetConfigurationIdentifierColumns,
  ] = useState<number[]>([]);
  const [
    datasetConfigurationProgramColumns,
    setDatasetConfigurationProgramColumns,
  ] = useState<number[]>([]);
  const [datasetConfigurationError, setDatasetConfigurationError] = useState<
    string | null
  >(null);

  const applyDatasetStatus = useCallback(
    (
      status: FacultyDatasetStatus,
      fallbackVariant?: StatusMessage["variant"],
    ) => {
      setDatasetStatus(status);
      if (status.message) {
        const variant = status.messageVariant ?? fallbackVariant;
        if (variant === "success" || variant === "error") {
          setDatasetBanner({ variant, message: status.message });
        } else {
          setDatasetBanner(null);
        }
      } else {
        setDatasetBanner(null);
      }
    },
    [],
  );

  useEffect(() => {
    const loadStatus = async () => {
      try {
        const status = await invoke<FacultyDatasetStatus>(
          "get_faculty_dataset_status",
        );
        applyDatasetStatus(status);
      } catch (statusError) {
        const message =
          statusError instanceof Error
            ? statusError.message
            : String(statusError);
        setDatasetStatus(null);
        setDatasetBanner({
          variant: "error",
          message: `Unable to load the faculty dataset status: ${message}`,
        });
      } finally {
        setIsDatasetLoading(false);
      }
    };

    loadStatus();
  }, [applyDatasetStatus]);

  useEffect(() => {
    if (!datasetStatus?.analysis) {
      setAvailablePrograms([]);
      setSelectedPrograms([]);
      if (facultyScope === "program") {
        setFacultyScope("all");
      }
      return;
    }

    const programs = [...datasetStatus.analysis.availablePrograms];
    setAvailablePrograms(programs);
    setSelectedPrograms((previous) => {
      const filtered = previous.filter((program) => programs.includes(program));
      if (
        filtered.length === previous.length &&
        filtered.every((value, index) => value === previous[index])
      ) {
        return previous;
      }
      return filtered;
    });

    if (programs.length === 0 && facultyScope === "program") {
      setFacultyScope("all");
    }
  }, [datasetStatus, facultyScope]);

  const runEmbeddingRefresh = useCallback(
    async (statusForValidation: FacultyDatasetStatus | null) => {
      const currentStatus = statusForValidation ?? datasetStatus;
      if (!currentStatus || !currentStatus.isValid) {
        const message =
          currentStatus?.message ??
          "Provide a valid faculty dataset before refreshing embeddings.";
        setEmbeddingStatus({ variant: "error", message });
        return;
      }

      setIsUpdatingEmbeddings(true);
      setEmbeddingStatus(null);

      try {
        const message = await invoke<string>("update_faculty_embeddings");
        setEmbeddingStatus({ variant: "success", message });
      } catch (updateError) {
        const message =
          updateError instanceof Error
            ? updateError.message
            : String(updateError);
        setEmbeddingStatus({ variant: "error", message });
      } finally {
        setIsUpdatingEmbeddings(false);
      }
    },
    [datasetStatus],
  );

  const selectDatasetFile = async () => {
    try {
      const selection = await open({
        multiple: false,
        filters: [
          {
            name: "Faculty dataset",
            extensions: ["tsv", "txt", "xlsx", "xls"],
          },
        ],
      });

      if (typeof selection === "string") {
        return selection;
      }
      return null;
    } catch (selectionError) {
      const message =
        selectionError instanceof Error
          ? selectionError.message
          : String(selectionError);
      setDatasetBanner({
        variant: "error",
        message: `Unable to open a file dialog: ${message}`,
      });
      return null;
    }
  };

  const openDatasetConfiguration = async (path: string) => {
    resetDatasetConfiguration();
    setDatasetConfigurationPath(path);
    setIsDatasetConfigurationOpen(true);
    setIsLoadingDatasetConfiguration(true);
    try {
      const preview = await invoke<FacultyDatasetPreviewResult>(
        "preview_faculty_dataset_replacement",
        { path },
      );
      setDatasetConfigurationPreview(preview.preview);
      setDatasetConfigurationEmbeddingColumns(
        [...preview.suggestedEmbeddingColumns].sort((a, b) => a - b),
      );
      setDatasetConfigurationIdentifierColumns(
        [...preview.suggestedIdentifierColumns].sort((a, b) => a - b),
      );
      setDatasetConfigurationProgramColumns(
        [...preview.suggestedProgramColumns].sort((a, b) => a - b),
      );
    } catch (previewError) {
      const message =
        previewError instanceof Error
          ? previewError.message
          : String(previewError);
      setDatasetConfigurationError(message);
    } finally {
      setIsLoadingDatasetConfiguration(false);
    }
  };

  const handleDatasetReplacement = async () => {
    setDatasetBanner(null);
    const selection = await selectDatasetFile();
    if (!selection) {
      return;
    }

    void openDatasetConfiguration(selection);
  };

  const closeDatasetConfiguration = () => {
    if (isDatasetBusy) {
      return;
    }
    setIsDatasetConfigurationOpen(false);
    resetDatasetConfiguration();
  };

  const handleConfirmDatasetReplacement = async () => {
    if (!datasetConfigurationPreview) {
      setDatasetConfigurationError(
        "Load the dataset preview before importing the faculty roster.",
      );
      return;
    }

    if (datasetConfigurationEmbeddingColumns.length === 0) {
      setDatasetConfigurationError(
        "Select at least one column containing faculty research interests or similar embedding content.",
      );
      return;
    }

    if (datasetConfigurationIdentifierColumns.length === 0) {
      setDatasetConfigurationError(
        "Select at least one column that uniquely identifies each faculty member.",
      );
      return;
    }

    setDatasetConfigurationError(null);
    setDatasetBanner(null);
    setIsDatasetBusy(true);

    try {
      const status = await invoke<FacultyDatasetStatus>(
        "replace_faculty_dataset",
        {
          path: datasetConfigurationPath,
          configuration: {
            embeddingColumns: datasetConfigurationEmbeddingColumns,
            identifierColumns: datasetConfigurationIdentifierColumns,
            programColumns: datasetConfigurationProgramColumns,
          },
        },
      );
      applyDatasetStatus(status, "success");

      if (!status.isValid) {
        setDatasetConfigurationError(
          status.message ??
            "Review the selected columns and try importing the dataset again.",
        );
        return;
      }

      setIsDatasetConfigurationOpen(false);
      resetDatasetConfiguration();

      await runEmbeddingRefresh(status);
    } catch (replacementError) {
      const message =
        replacementError instanceof Error
          ? replacementError.message
          : String(replacementError);
      setDatasetConfigurationError(message);
      setDatasetBanner({
        variant: "error",
        message: `Unable to replace the faculty dataset: ${message}`,
      });
    } finally {
      setIsDatasetBusy(false);
    }
  };

  const handleDatasetRestore = async () => {
    setDatasetBanner(null);
    setIsDatasetBusy(true);
    try {
      const status = await invoke<FacultyDatasetStatus>(
        "restore_default_faculty_dataset",
      );
      applyDatasetStatus(status, "success");
    } catch (restoreError) {
      const message =
        restoreError instanceof Error
          ? restoreError.message
          : String(restoreError);
      setDatasetBanner({
        variant: "error",
        message: `Unable to restore the packaged dataset: ${message}`,
      });
    } finally {
      setIsDatasetBusy(false);
    }
  };

  const formatDatasetTimestamp = (value: string | null) => {
    if (!value) {
      return "Not available";
    }
    const parsed = new Date(value);
    if (Number.isNaN(parsed.getTime())) {
      return value;
    }
    return parsed.toLocaleString();
  };

  const resetSpreadsheetConfiguration = () => {
    setSpreadsheetPreview(null);
    setSpreadsheetPreviewError(null);
    setSelectedIdentifierColumns([]);
    setSelectedPromptColumns([]);
    setIsLoadingSpreadsheetPreview(false);
  };

  const resetDatasetConfiguration = () => {
    setDatasetConfigurationPreview(null);
    setDatasetConfigurationEmbeddingColumns([]);
    setDatasetConfigurationIdentifierColumns([]);
    setDatasetConfigurationProgramColumns([]);
    setDatasetConfigurationError(null);
    setDatasetConfigurationPath("");
    setIsLoadingDatasetConfiguration(false);
  };

  const handleTaskTypeChange = (value: TaskType) => {
    setTaskType(value);
    setError(null);
    setResult(null);

    if (value !== "spreadsheet") {
      resetSpreadsheetConfiguration();
    }
  };

  const handleFacultyScopeChange = (value: FacultyScope) => {
    if (value === "program" && availablePrograms.length === 0) {
      return;
    }

    setFacultyScope(value);
    setError(null);
    setResult(null);

    if (value !== "custom") {
      setCustomFacultyPath("");
    }

    if (value !== "program") {
      setSelectedPrograms([]);
    }
  };

  const toggleProgramSelection = (program: ProgramName) => {
    if (!availablePrograms.includes(program)) {
      return;
    }

    setSelectedPrograms((current) => {
      if (current.includes(program)) {
        return current.filter((entry) => entry !== program);
      }

      return [...current, program];
    });
    setError(null);
    setResult(null);
  };

  const handleFileSelection = async (
    setter: (value: string) => void,
    options: Parameters<typeof open>[0],
  ) => {
    try {
      const selection = await open(options);
      if (typeof selection === "string") {
        setter(selection);
        return selection;
      }
      return null;
    } catch (selectionError) {
      const message =
        selectionError instanceof Error
          ? selectionError.message
          : String(selectionError);
      setError(`Unable to open a file dialog: ${message}`);
      return null;
    }
  };

  const handleSpreadsheetPathInput = (value: string) => {
    setSpreadsheetPath(value);
    setError(null);
    setResult(null);
    resetSpreadsheetConfiguration();
  };

  const loadSpreadsheetPreview = async (path: string) => {
    const trimmed = path.trim();
    if (trimmed.length === 0) {
      resetSpreadsheetConfiguration();
      return;
    }

    setIsLoadingSpreadsheetPreview(true);
    setSpreadsheetPreviewError(null);
    setError(null);
    setResult(null);

    try {
      const preview = await invoke<SpreadsheetPreview>("analyze_spreadsheet", {
        path: trimmed,
      });

      const identifierSuggestions = Array.from(
        new Set(preview.suggestedIdentifierColumns),
      ).sort((a, b) => a - b);
      const promptSuggestions = Array.from(
        new Set(preview.suggestedPromptColumns),
      ).sort((a, b) => a - b);

      setSpreadsheetPreview(preview);
      setSelectedIdentifierColumns(identifierSuggestions);
      setSelectedPromptColumns(promptSuggestions);
    } catch (analysisError) {
      const message =
        analysisError instanceof Error
          ? analysisError.message
          : String(analysisError);
      setSpreadsheetPreview(null);
      setSpreadsheetPreviewError(message);
      setSelectedIdentifierColumns([]);
      setSelectedPromptColumns([]);
    } finally {
      setIsLoadingSpreadsheetPreview(false);
    }
  };

  const handleSpreadsheetSelection = async () => {
    const selection = await handleFileSelection(setSpreadsheetPath, {
      multiple: false,
      filters: [
        {
          name: "Spreadsheets",
          extensions: ["tsv", "txt", "xlsx", "xls"],
        },
      ],
    });

    if (selection) {
      await loadSpreadsheetPreview(selection);
    }
  };

  const toggleIdentifierColumn = (index: number) => {
    setSelectedIdentifierColumns((current) => {
      const updated = current.includes(index)
        ? current.filter((entry) => entry !== index)
        : [...current, index];
      updated.sort((a, b) => a - b);
      return updated;
    });
    setError(null);
    setResult(null);
  };

  const togglePromptColumn = (index: number) => {
    setSelectedPromptColumns((current) => {
      const updated = current.includes(index)
        ? current.filter((entry) => entry !== index)
        : [...current, index];
      updated.sort((a, b) => a - b);
      return updated;
    });
    setError(null);
    setResult(null);
  };

  const getColumnLabel = (index: number) => {
    if (!spreadsheetPreview) {
      return `Column ${index + 1}`;
    }

    const header = spreadsheetPreview.headers[index];
    if (!header) {
      return `Column ${index + 1}`;
    }

    const trimmed = header.trim();
    return trimmed.length > 0 ? trimmed : `Column ${index + 1}`;
  };

  const mapSelectedColumns = (indexes: number[]) => {
    const unique = Array.from(new Set(indexes)).filter((index) => index >= 0);

    if (!spreadsheetPreview) {
      return unique.map((index) => `Column ${index + 1}`);
    }

    return unique
      .filter((index) => index < spreadsheetPreview.headers.length)
      .map((index) => getColumnLabel(index));
  };

  const getDatasetConfigurationColumnLabel = (index: number) => {
    if (!datasetConfigurationPreview) {
      return `Column ${index + 1}`;
    }

    const header = datasetConfigurationPreview.headers[index];
    if (!header) {
      return `Column ${index + 1}`;
    }

    const trimmed = header.trim();
    return trimmed.length > 0 ? trimmed : `Column ${index + 1}`;
  };

  const toggleDatasetEmbeddingColumn = (index: number) => {
    setDatasetConfigurationEmbeddingColumns((current) => {
      const updated = current.includes(index)
        ? current.filter((entry) => entry !== index)
        : [...current, index];
      updated.sort((a, b) => a - b);
      return updated;
    });
    setDatasetConfigurationError(null);
  };

  const toggleDatasetIdentifierColumn = (index: number) => {
    setDatasetConfigurationIdentifierColumns((current) => {
      const updated = current.includes(index)
        ? current.filter((entry) => entry !== index)
        : [...current, index];
      updated.sort((a, b) => a - b);
      return updated;
    });
    setDatasetConfigurationError(null);
  };

  const toggleDatasetProgramColumn = (index: number) => {
    setDatasetConfigurationProgramColumns((current) => {
      const updated = current.includes(index)
        ? current.filter((entry) => entry !== index)
        : [...current, index];
      updated.sort((a, b) => a - b);
      return updated;
    });
    setDatasetConfigurationError(null);
  };

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setIsSubmitting(true);
    setError(null);
    setResult(null);

    const programFilters = selectedPrograms;
    const facultyRecommendations = Math.max(
      1,
      Number.parseInt(facultyRecCount, 10) || 0,
    );
    const studentRecommendations = Math.max(
      0,
      Number.parseInt(studentRecCount, 10) || 0,
    );

    if (taskType === "spreadsheet") {
      const trimmedPath = spreadsheetPath.trim();
      if (trimmedPath.length === 0) {
        setError("Provide a spreadsheet containing student prompts.");
        setIsSubmitting(false);
        return;
      }

      if (!spreadsheetPreview) {
        setError("Load the spreadsheet preview to choose columns before submitting.");
        setIsSubmitting(false);
        return;
      }

      if (selectedIdentifierColumns.length === 0) {
        setError("Select at least one column that uniquely identifies each student.");
        setIsSubmitting(false);
        return;
      }

      if (selectedPromptColumns.length === 0) {
        setError("Select at least one column containing the student prompts.");
        setIsSubmitting(false);
        return;
      }
    }

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
            customFacultyPath:
              facultyScope === "custom" && customFacultyPath.trim().length > 0
                ? customFacultyPath.trim()
                : undefined,
            facultyRecsPerStudent: facultyRecommendations,
            studentRecsPerFaculty: studentRecommendations,
            spreadsheetPromptColumns:
              taskType === "spreadsheet"
                ? mapSelectedColumns(selectedPromptColumns)
                : undefined,
            spreadsheetIdentifierColumns:
              taskType === "spreadsheet"
                ? mapSelectedColumns(selectedIdentifierColumns)
                : undefined,
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
    await runEmbeddingRefresh(datasetStatus);
  };

  return (
    <div className="app-shell">
      <main>
        <header className="page-header">
          <h1>DBBS Faculty Recommendation Console</h1>
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
                <span>Single document</span>
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
              <div className="input-stack narrow-column">
                <label htmlFor="prompt-text">Prompt text</label>
                <textarea
                  id="prompt-text"
                  className="prompt-textarea"
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
                    onClick={handleSpreadsheetSelection}
                  >
                    Browse…
                  </button>
                  <input
                    type="text"
                    value={spreadsheetPath}
                    onChange={(event) =>
                      handleSpreadsheetPathInput(event.target.value)
                    }
                    placeholder="Paste or confirm the spreadsheet path"
                  />
                  <button
                    type="button"
                    onClick={() => void loadSpreadsheetPreview(spreadsheetPath)}
                    disabled={
                      spreadsheetPath.trim().length === 0 ||
                      isLoadingSpreadsheetPreview
                    }
                  >
                    {isLoadingSpreadsheetPreview
                      ? "Loading preview…"
                      : "Load preview"}
                  </button>
                </div>
                {spreadsheetPath && (
                  <div className="path-preview">{spreadsheetPath}</div>
                )}
                {spreadsheetPreviewError && (
                  <div className="preview-error">{spreadsheetPreviewError}</div>
                )}
                {isLoadingSpreadsheetPreview && !spreadsheetPreviewError && (
                  <div className="preview-status">
                    Analyzing the spreadsheet…
                  </div>
                )}
                {spreadsheetPreview && (
                  <div className="spreadsheet-preview-card">
                    <div className="column-selector-group">
                      <div className="column-selector">
                        <h4>Identifier columns</h4>
                        <p className="small-note">
                          Combine these columns to create a unique identifier
                          for each student.
                        </p>
                        <div className="column-checkbox-list">
                          {spreadsheetPreview.headers.map((_, index) => {
                            const label = getColumnLabel(index);
                            const isChecked =
                              selectedIdentifierColumns.includes(index);
                            return (
                              <label
                                key={`identifier-${index}`}
                                className="column-checkbox-option"
                              >
                                <input
                                  type="checkbox"
                                  checked={isChecked}
                                  onChange={() => toggleIdentifierColumn(index)}
                                />
                                <span>{label}</span>
                              </label>
                            );
                          })}
                        </div>
                      </div>
                      <div className="column-selector">
                        <h4>Prompt columns</h4>
                        <p className="small-note">
                          Select the columns that contain student research
                          interests. Their text will be merged before
                          embedding.
                        </p>
                        <div className="column-checkbox-list">
                          {spreadsheetPreview.headers.map((_, index) => {
                            const label = getColumnLabel(index);
                            const isChecked =
                              selectedPromptColumns.includes(index);
                            return (
                              <label
                                key={`prompt-${index}`}
                                className="column-checkbox-option"
                              >
                                <input
                                  type="checkbox"
                                  checked={isChecked}
                                  onChange={() => togglePromptColumn(index)}
                                />
                                <span>{label}</span>
                              </label>
                            );
                          })}
                        </div>
                      </div>
                    </div>
                    <div className="preview-table-wrapper">
                      <table className="preview-table">
                        <thead>
                          <tr>
                            {spreadsheetPreview.headers.map((_, index) => (
                              <th key={`header-${index}`}>
                                {getColumnLabel(index)}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {spreadsheetPreview.rows.map((row, rowIndex) => (
                            <tr key={`row-${rowIndex}`}>
                              {row.map((value, columnIndex) => (
                                <td
                                  key={`cell-${rowIndex}-${columnIndex}`}
                                  title={value}
                                >
                                  {value}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
                <p className="small-note">
                  Each row should include an identifier column (for example,
                  name or ID) and a column containing the student's prompt.
                  Tab-delimited TSV/TXT or Excel formats are supported.
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
                  disabled={
                    isDatasetLoading || availablePrograms.length === 0
                  }
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
            {availablePrograms.length === 0 &&
              !isDatasetLoading &&
              datasetStatus?.analysis && (
                <p className="small-note">
                  No program affiliations were detected in the active faculty
                  dataset.
                </p>
              )}

            {facultyScope === "program" && (
              <div className="input-stack narrow-column">
                <span className="input-heading">Programs or tracks</span>
                <div className="program-checkbox-grid">
                  {availablePrograms.length === 0 ? (
                    <p className="small-note">
                      No program columns were detected in the active dataset.
                    </p>
                  ) : (
                    availablePrograms.map((program) => {
                      const isSelected = selectedPrograms.includes(program);

                      return (
                        <label key={program} className="checkbox-option">
                          <input
                            type="checkbox"
                            value={program}
                            checked={isSelected}
                            onChange={() => toggleProgramSelection(program)}
                          />
                          <span>{program}</span>
                        </label>
                      );
                    })
                  )}
                </div>
                <p className="small-note">
                  Select the programs that should be included in the faculty
                  roster.
                </p>
              </div>
            )}

            {facultyScope === "custom" && (
              <div className="input-stack">
                <label>Faculty roster spreadsheet</label>
                <div className="button-row inline">
                  <button
                    type="button"
                    className="secondary"
                    onClick={() =>
                      handleFileSelection(setCustomFacultyPath, {
                        multiple: false,
                        filters: [
                          {
                            name: "Faculty rosters",
                            extensions: ["tsv", "txt", "xlsx", "xls"],
                          },
                        ],
                      })
                    }
                  >
                    Browse…
                  </button>
                  <input
                    type="text"
                    value={customFacultyPath}
                    onChange={(event) => setCustomFacultyPath(event.target.value)}
                    placeholder="Paste or confirm the faculty roster path"
                  />
                </div>
                {customFacultyPath && (
                  <div className="path-preview">{customFacultyPath}</div>
                )}
                <p className="small-note">
                  Upload a tab-delimited TSV/TXT or Excel file listing the
                  available faculty members.
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
          </fieldset>

          <section className="dataset-card">
            <div className="dataset-card-header">
              <h2>Faculty dataset</h2>
              <span
                className={`dataset-status-pill ${
                  isDatasetLoading
                    ? "loading"
                    : datasetStatus?.isValid
                      ? "ready"
                      : "warning"
                }`}
              >
                {isDatasetLoading
                  ? "Loading…"
                  : datasetStatus?.isValid
                    ? "Ready"
                    : "Action required"}
              </span>
            </div>
            <p className="dataset-card-description">
              Manage the dataset file that seeds the faculty embedding index
              before running updates.
            </p>
            <dl className="dataset-meta-grid">
              <div>
                <dt>Active file</dt>
                <dd className="dataset-path">
                  {isDatasetLoading
                    ? "Loading…"
                    : datasetStatus?.sourcePath ??
                        datasetStatus?.canonicalPath ??
                        datasetStatus?.path ??
                        "Not configured"}
                </dd>
              </div>
              <div>
                <dt>Last updated</dt>
                <dd>
                  {isDatasetLoading
                    ? "Loading…"
                    : formatDatasetTimestamp(datasetStatus?.lastModified ?? null)}
                </dd>
              </div>
              <div>
                <dt>Rows × columns</dt>
                <dd>
                  {isDatasetLoading
                    ? "Loading…"
                    : datasetStatus?.rowCount != null &&
                        datasetStatus.columnCount != null
                      ? `${datasetStatus.rowCount} × ${datasetStatus.columnCount}`
                      : "Unavailable"}
                </dd>
              </div>
              <div>
                <dt>Source</dt>
                <dd>
                  {isDatasetLoading
                    ? "Loading…"
                    : datasetStatus
                      ? datasetStatus.isDefault
                        ? "Packaged default"
                        : "Custom upload"
                      : "Not configured"}
                </dd>
              </div>
            </dl>
            {datasetStatus?.message &&
              (datasetStatus.messageVariant ??
                (datasetStatus.isValid ? "success" : "error")) !==
                "success" && (
              <p
                className={`dataset-message ${
                  (datasetStatus.messageVariant ??
                    (datasetStatus.isValid ? "success" : "error")) ===
                  "success"
                    ? "dataset-message-success"
                    : (datasetStatus.messageVariant ??
                        (datasetStatus.isValid ? "success" : "error")) ===
                        "error"
                      ? "dataset-message-error"
                      : "dataset-message-info"
                }`}
              >
                {datasetStatus.message}
              </p>
            )}
            <div className="dataset-actions">
              <button
                type="button"
                onClick={handleDatasetReplacement}
                disabled={
                  isDatasetBusy || isDatasetLoading || isDatasetConfigurationOpen
                }
              >
                Replace dataset
              </button>
              <button
                type="button"
                className="secondary"
                onClick={handleDatasetRestore}
                disabled={
                  isDatasetBusy || isDatasetLoading || isDatasetConfigurationOpen
                }
              >
                Restore default
              </button>
              <button
                type="button"
                className="ghost"
                onClick={() => setIsDatasetPreviewOpen(true)}
                disabled={
                  isDatasetBusy ||
                  isDatasetLoading ||
                  isDatasetConfigurationOpen ||
                  !datasetStatus?.preview
                }
              >
                Preview dataset
              </button>
            </div>
          </section>

          {datasetBanner && (
            <div
              className={`status-banner ${
                datasetBanner.variant === "success"
                  ? "status-success"
                  : datasetBanner.variant === "error"
                    ? "status-error"
                    : "status-info"
              } dataset-status-banner`}
            >
              {datasetBanner.message}
            </div>
          )}

          <div className="button-row">
            <button type="submit" disabled={isSubmitting}>
              {isSubmitting ? "Validating…" : "Validate matching request"}
            </button>
            <button
              type="button"
              className="secondary"
              onClick={handleEmbeddingsUpdate}
              disabled={
                isUpdatingEmbeddings ||
                isDatasetLoading ||
                !datasetStatus?.isValid
              }
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
                  {result.details.spreadsheetIdentifierColumns.length > 0 && (
                    <>
                      <dt>Identifier columns</dt>
                      <dd>
                        <ul className="path-list">
                          {result.details.spreadsheetIdentifierColumns.map(
                            (column) => (
                              <li key={`identifier-${column}`}>{column}</li>
                            ),
                          )}
                        </ul>
                      </dd>
                    </>
                  )}
                  {result.details.spreadsheetPromptColumns.length > 0 && (
                    <>
                      <dt>Prompt columns</dt>
                      <dd>
                        <ul className="path-list">
                          {result.details.spreadsheetPromptColumns.map(
                            (column) => (
                              <li key={`prompt-${column}`}>{column}</li>
                            ),
                          )}
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
                  {result.details.customFacultyPath && (
                    <>
                      <dt>Faculty roster</dt>
                      <dd>
                        <div className="path-preview">
                          {result.details.customFacultyPath}
                        </div>
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

        {isDatasetConfigurationOpen && (
          <div className="dataset-preview-overlay">
            <div className="dataset-preview-dialog">
              <div className="dataset-preview-header">
                <h3>Import faculty dataset</h3>
                <button
                  type="button"
                  className="ghost close-button"
                  onClick={closeDatasetConfiguration}
                  disabled={isDatasetBusy}
                >
                  Close
                </button>
              </div>
              {datasetConfigurationPath && (
                <div className="path-preview">{datasetConfigurationPath}</div>
              )}
              {datasetConfigurationError && (
                <div className="preview-error">{datasetConfigurationError}</div>
              )}
              {isLoadingDatasetConfiguration && (
                <div className="preview-status">Loading preview…</div>
              )}
              {!isLoadingDatasetConfiguration && datasetConfigurationPreview && (
                <>
                  <div className="spreadsheet-preview-card">
                    <div className="column-selector-group">
                      <div className="column-selector">
                        <h4>Embedding columns</h4>
                        <p className="small-note">
                          Choose the columns describing faculty research interests
                          or other content that should be embedded.
                        </p>
                        <div className="column-checkbox-list">
                          {datasetConfigurationPreview.headers.map((_, index) => {
                            const isChecked =
                              datasetConfigurationEmbeddingColumns.includes(index);
                            return (
                              <label
                                key={`dataset-embedding-${index}`}
                                className="column-checkbox-option"
                              >
                                <input
                                  type="checkbox"
                                  checked={isChecked}
                                  onChange={() => toggleDatasetEmbeddingColumn(index)}
                                  disabled={isDatasetBusy}
                                />
                                <span>{getDatasetConfigurationColumnLabel(index)}</span>
                              </label>
                            );
                          })}
                        </div>
                      </div>
                      <div className="column-selector">
                        <h4>Identifier columns</h4>
                        <p className="small-note">
                          Select the columns that uniquely identify each faculty
                          member.
                        </p>
                        <div className="column-checkbox-list">
                          {datasetConfigurationPreview.headers.map((_, index) => {
                            const isChecked =
                              datasetConfigurationIdentifierColumns.includes(index);
                            return (
                              <label
                                key={`dataset-identifier-${index}`}
                                className="column-checkbox-option"
                              >
                                <input
                                  type="checkbox"
                                  checked={isChecked}
                                  onChange={() => toggleDatasetIdentifierColumn(index)}
                                  disabled={isDatasetBusy}
                                />
                                <span>{getDatasetConfigurationColumnLabel(index)}</span>
                              </label>
                            );
                          })}
                        </div>
                      </div>
                      <div className="column-selector">
                        <h4>Program columns</h4>
                        <p className="small-note">
                          Optional: include the columns that list program or track
                          affiliations to enable program filtering.
                        </p>
                        <div className="column-checkbox-list">
                          {datasetConfigurationPreview.headers.map((_, index) => {
                            const isChecked =
                              datasetConfigurationProgramColumns.includes(index);
                            return (
                              <label
                                key={`dataset-program-${index}`}
                                className="column-checkbox-option"
                              >
                                <input
                                  type="checkbox"
                                  checked={isChecked}
                                  onChange={() => toggleDatasetProgramColumn(index)}
                                  disabled={isDatasetBusy}
                                />
                                <span>{getDatasetConfigurationColumnLabel(index)}</span>
                              </label>
                            );
                          })}
                        </div>
                      </div>
                    </div>
                  </div>
                  <div className="spreadsheet-preview-card dataset-preview-card">
                    <p className="small-note">
                      Showing the first {datasetConfigurationPreview.rows.length} rows
                      of the selected dataset.
                    </p>
                    <div className="preview-table-wrapper">
                      <table className="preview-table">
                        <thead>
                          <tr>
                            {datasetConfigurationPreview.headers.map((header, index) => (
                              <th key={`dataset-config-header-${index}`}>
                                {header && header.trim().length > 0
                                  ? header
                                  : `Column ${index + 1}`}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {datasetConfigurationPreview.rows.map((row, rowIndex) => (
                            <tr key={`dataset-config-row-${rowIndex}`}>
                              {row.map((value, columnIndex) => (
                                <td
                                  key={`dataset-config-cell-${rowIndex}-${columnIndex}`}
                                  title={value}
                                >
                                  {value}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </>
              )}
              <div className="dataset-config-actions">
                <button
                  type="button"
                  onClick={handleConfirmDatasetReplacement}
                  disabled={
                    isDatasetBusy ||
                    isLoadingDatasetConfiguration ||
                    !datasetConfigurationPreview ||
                    datasetConfigurationEmbeddingColumns.length === 0 ||
                    datasetConfigurationIdentifierColumns.length === 0
                  }
                >
                  {isDatasetBusy ? "Importing…" : "Import dataset"}
                </button>
                <button
                  type="button"
                  className="ghost"
                  onClick={closeDatasetConfiguration}
                  disabled={isDatasetBusy}
                >
                  Cancel
                </button>
              </div>
            </div>
          </div>
        )}

        {isDatasetPreviewOpen && datasetStatus?.preview && (
          <div className="dataset-preview-overlay">
            <div className="dataset-preview-dialog">
              <div className="dataset-preview-header">
                <h3>Faculty dataset preview</h3>
                <button
                  type="button"
                  className="ghost close-button"
                  onClick={() => setIsDatasetPreviewOpen(false)}
                >
                  Close
                </button>
              </div>
              <div className="spreadsheet-preview-card dataset-preview-card">
                <p className="small-note">
                  Showing the first {datasetStatus.preview.rows.length} rows of
                  the active dataset file.
                </p>
                <div className="preview-table-wrapper">
                  <table className="preview-table">
                    <thead>
                      <tr>
                        {datasetStatus.preview.headers.map((header, index) => (
                          <th key={`dataset-header-${index}`}>
                            {header && header.trim().length > 0
                              ? header
                              : `Column ${index + 1}`}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {datasetStatus.preview.rows.map((row, rowIndex) => (
                        <tr key={`dataset-row-${rowIndex}`}>
                          {row.map((value, columnIndex) => (
                            <td
                              key={`dataset-cell-${rowIndex}-${columnIndex}`}
                              title={value}
                            >
                              {value}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

export default App;
