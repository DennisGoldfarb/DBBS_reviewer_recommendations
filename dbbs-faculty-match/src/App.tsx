import { useEffect, useState } from "react";
import { FacultyExclusions } from "./components/FacultyExclusions";
import { useFacultyExclusions } from "./hooks/useFacultyExclusions";
import { StudentInputs } from "./components/StudentInputs";
import { FacultyFilters } from "./components/FacultyFilters";
import { FacultyDatasetCard } from "./components/FacultyDatasetCard";
import { DatasetConfigurationDialog } from "./components/DatasetConfigurationDialog";
import { DatasetPreviewDialog } from "./components/DatasetPreviewDialog";
import { MatchingResults } from "./components/MatchingResults";
import { useEmbeddingHelper } from "./hooks/useEmbeddingHelper";
import { useFacultyDataset } from "./hooks/useFacultyDataset";
import { useFacultyFilters } from "./hooks/useFacultyFilters";
import { useStudentInput } from "./hooks/useStudentInput";
import { useMatching } from "./hooks/useMatching";
import { useFeedback } from "./hooks/useFeedback";
import { useTheme } from "./hooks/useTheme";
import { useFileSelection } from "./services/fileDialogs";
import { describeEmbeddingProgress } from "./lib/formatting";
import "./App.css";

function App() {
  const { theme, toggleTheme, themeToggleLabel } = useTheme();
  const feedback = useFeedback();
  const selectFile = useFileSelection(feedback.setError);
  const embeddings = useEmbeddingHelper();
  const dataset = useFacultyDataset(embeddings.runEmbeddingRefresh);
  useEffect(feedback.clear, [dataset.datasetStatus, feedback.clear]);
  const student = useStudentInput(feedback.clear, selectFile);
  const faculty = useFacultyFilters(
    dataset.datasetStatus,
    feedback.clear,
    selectFile,
  );
  const [facultyRecCount, setFacultyRecCount] = useState("10");
  const exclusions = useFacultyExclusions(feedback.clear);
  const facultyHeaders =
    dataset.datasetStatus?.analysis?.identifierColumns ?? [];
  const studentHeaders = student.spreadsheetPreview?.headers ?? [];
  const facultyNameColumn = facultyHeaders.includes(
    exclusions.facultyNameColumn,
  )
    ? exclusions.facultyNameColumn
    : (facultyHeaders.find((header) =>
        /^(faculty )?name$/i.test(header.trim()),
      ) ?? "");
  const studentIdColumn = studentHeaders.includes(exclusions.studentIdColumn)
    ? exclusions.studentIdColumn
    : (studentHeaders.find((header) =>
        /^person reference id$/i.test(header.trim()),
      ) ??
      studentHeaders.find((header) => /^id$/i.test(header.trim())) ??
      "");
  const matching = useMatching(
    student,
    faculty,
    facultyRecCount,
    feedback,
    () =>
      exclusions.configuration(
        student.taskType,
        facultyNameColumn,
        studentIdColumn,
      ),
  );
  const inputsDisabled =
    matching.isSubmitting ||
    embeddings.isUpdatingEmbeddings ||
    dataset.isDatasetBusy ||
    dataset.isDatasetConfigurationOpen ||
    dataset.isDatasetPreviewOpen;
  const areControlsDisabled =
    inputsDisabled ||
    !embeddings.isEmbeddingHelperReady ||
    dataset.isDatasetConfigurationOpen;

  return (
    <div
      className="app-shell"
      onDragOver={(event) => event.preventDefault()}
      onDrop={(event) => event.preventDefault()}
    >
      <main>
        <header className="page-header">
          <h1>DBBS Faculty Recommendation Console</h1>
          <button
            type="button"
            className="theme-toggle"
            onClick={toggleTheme}
            aria-pressed={theme === "dark"}
            aria-label={themeToggleLabel}
            title={themeToggleLabel}
          >
            <span className="theme-toggle-icon" aria-hidden="true">
              {theme === "dark" ? "🌙" : "☀️"}
            </span>
            <span className="theme-toggle-text">
              <span className="theme-toggle-name">
                {theme === "dark" ? "Dark WashU" : "Light WashU"}
              </span>
              <span className="theme-toggle-action">Toggle theme</span>
            </span>
          </button>
        </header>

        <form className="matching-form" onSubmit={matching.handleSubmit}>
          <StudentInputs
            {...student}
            handleFileSelection={selectFile}
            disabled={inputsDisabled}
          />
          <FacultyFilters
            {...faculty}
            datasetStatus={dataset.datasetStatus}
            isDatasetLoading={dataset.isDatasetLoading}
            disabled={inputsDisabled}
          />
          <FacultyExclusions
            rules={exclusions}
            disabled={inputsDisabled}
            taskType={student.taskType}
            studentSource={
              student.taskType === "spreadsheet"
                ? student.spreadsheetPath
                : student.taskType === "document"
                  ? student.documentPath
                  : student.taskType === "directory"
                    ? student.directoryPath
                    : ""
            }
            facultySource={
              dataset.datasetStatus?.sourcePath ??
              dataset.datasetStatus?.path ??
              "Bundled faculty dataset"
            }
            studentHeaders={studentHeaders}
            facultyHeaders={facultyHeaders}
            facultyColumn={facultyNameColumn}
            studentColumn={studentIdColumn}
            selectFile={selectFile}
            clear={feedback.clear}
          />
          <fieldset disabled={inputsDisabled}>
            <legend>Recommendation limits</legend>
            <label className="input-stack narrow-column">
              Faculty recommendations per student
              <input
                type="number"
                min="1"
                step="1"
                required
                value={facultyRecCount}
                onChange={(event) => {
                  setFacultyRecCount(event.target.value);
                  feedback.clear();
                }}
              />
            </label>
          </fieldset>
          <FacultyDatasetCard
            {...dataset}
            areControlsDisabled={areControlsDisabled}
          />
          {dataset.datasetBanner && (
            <div
              role="status"
              className={`status-banner status-${dataset.datasetBanner.variant} dataset-status-banner`}
            >
              {dataset.datasetBanner.message}
            </div>
          )}
          <div className="button-row">
            <button
              type="submit"
              disabled={
                areControlsDisabled ||
                dataset.isDatasetLoading ||
                !dataset.datasetStatus?.isValid
              }
            >
              {matching.isSubmitting ? "Matching…" : "Run matching"}
            </button>
            <button
              type="button"
              className="secondary"
              onClick={() =>
                void embeddings.runEmbeddingRefresh(dataset.datasetStatus)
              }
              disabled={
                areControlsDisabled ||
                dataset.isDatasetLoading ||
                !dataset.datasetStatus?.isValid
              }
            >
              {embeddings.isUpdatingEmbeddings
                ? "Checking embeddings…"
                : "Update faculty embeddings"}
            </button>
          </div>
        </form>

        {feedback.error && (
          <div role="alert" className="status-banner status-error">
            {feedback.error}
          </div>
        )}
        <MatchingResults {...matching} datasetStatus={dataset.datasetStatus} />
        {!embeddings.isEmbeddingHelperReady &&
          !embeddings.embeddingProgress &&
          !embeddings.embeddingStatus && (
            <div role="status" className="status-banner status-info">
              Preparing the embedding helper…
            </div>
          )}
        {embeddings.embeddingProgress && (
          <div role="status" className="status-banner status-info">
            {describeEmbeddingProgress(embeddings.embeddingProgress)}
          </div>
        )}
        {embeddings.embeddingStatus && (
          <div
            role="status"
            className={`status-banner status-${embeddings.embeddingStatus.variant}`}
          >
            {embeddings.embeddingStatus.message}
          </div>
        )}
        {!embeddings.isEmbeddingHelperReady &&
          embeddings.embeddingStatus?.variant === "error" && (
            <button
              type="button"
              className="secondary"
              onClick={embeddings.retryInitialization}
            >
              Retry embedding helper
            </button>
          )}
        {dataset.isDatasetConfigurationOpen && (
          <DatasetConfigurationDialog {...dataset} />
        )}
        {dataset.isDatasetPreviewOpen && <DatasetPreviewDialog {...dataset} />}
      </main>
    </div>
  );
}

export default App;
