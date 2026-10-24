import { UPLOAD_EXTENSIONS } from "../lib/uploads";
import { PreviewTable } from "./PreviewTable";
import { FileDropZone } from "./FileDropZone";
import type { useStudentInput } from "../hooks/useStudentInput";
import type { FileSelectionHandler } from "../services/fileDialogs";

export function StudentInputs({
  acceptSpreadsheetPath,
  disabled,
  taskType,
  handleTaskTypeChange,
  promptText,
  setPromptText,
  handleFileSelection,
  setDocumentPath,
  documentPath,
  handleSpreadsheetSelection,
  spreadsheetPath,
  handleSpreadsheetPathInput,
  loadSpreadsheetPreview,
  isLoadingSpreadsheetPreview,
  spreadsheetPreviewError,
  spreadsheetPreview,
  getColumnLabel,
  selectedIdentifierColumns,
  toggleIdentifierColumn,
  selectedPromptColumns,
  togglePromptColumn,
  setDirectoryPath,
  directoryPath,
}: ReturnType<typeof useStudentInput> & {
  handleFileSelection: FileSelectionHandler;
} & { disabled: boolean }) {
  return (
    <fieldset disabled={disabled}>
      <legend>Student inputs</legend>
      <p className="section-description">
        Select how student research interests will be supplied. The app compares
        each submission against faculty research profiles.
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
            Provide a single free-form description that will be embedded and
            compared against faculty profiles.
          </p>
        </div>
      )}

      {taskType === "document" && (
        <FileDropZone
          className="input-stack"
          kind="document"
          label="Research interest document upload"
          onPath={setDocumentPath}
          disabled={disabled}
        >
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
                      extensions: [...UPLOAD_EXTENSIONS.document],
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
          {documentPath && <div className="path-preview">{documentPath}</div>}
        </FileDropZone>
      )}

      {taskType === "spreadsheet" && (
        <FileDropZone
          className="input-stack"
          kind="spreadsheet"
          label="Student spreadsheet upload"
          onPath={acceptSpreadsheetPath}
          disabled={disabled}
        >
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
            <div className="preview-status">Analyzing the spreadsheet…</div>
          )}
          {spreadsheetPreview && (
            <div className="spreadsheet-preview-card">
              <div className="column-selector-group">
                <div className="column-selector">
                  <h4>Identifier columns</h4>
                  <p className="small-note">
                    Combine these columns to create a unique identifier for each
                    student. If you leave them all unchecked, the spreadsheet
                    row number will be used instead.
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
                    Select the columns that contain student research interests.
                    Their text will be merged before embedding.
                  </p>
                  <div className="column-checkbox-list">
                    {spreadsheetPreview.headers.map((_, index) => {
                      const label = getColumnLabel(index);
                      const isChecked = selectedPromptColumns.includes(index);
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
                <PreviewTable preview={spreadsheetPreview} />
              </div>
            </div>
          )}
          <p className="small-note">
            Each row should include a prompt column and, ideally, an identifier
            column such as a name or ID. If you do not select an identifier
            column, the spreadsheet row number will be used in the results.
            Tab-delimited TSV/TXT or Excel formats are supported.
          </p>
        </FileDropZone>
      )}

      {taskType === "directory" && (
        <FileDropZone
          className="input-stack"
          kind="directory"
          label="Student folder upload"
          onPath={setDirectoryPath}
          disabled={disabled}
        >
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
          {directoryPath && <div className="path-preview">{directoryPath}</div>}
          <p className="small-note">
            Each document in the folder will be treated as a separate student
            submission. The filename will be used as the identifier.
          </p>
        </FileDropZone>
      )}
    </fieldset>
  );
}
