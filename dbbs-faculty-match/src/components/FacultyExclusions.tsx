import { FileDropZone } from "./FileDropZone";
import { UPLOAD_EXTENSIONS } from "../lib/uploads";
import type { useFacultyExclusions } from "../hooks/useFacultyExclusions";
import type { FileSelectionHandler } from "../services/fileDialogs";
import type { TaskType } from "../types";

export function FacultyExclusions({
  rules,
  disabled,
  taskType,
  studentSource,
  facultySource,
  studentHeaders,
  facultyHeaders,
  facultyColumn,
  studentColumn,
  selectFile,
  clear,
}: {
  rules: ReturnType<typeof useFacultyExclusions>;
  disabled: boolean;
  taskType: TaskType;
  studentSource: string;
  facultySource: string;
  studentHeaders: string[];
  facultyHeaders: string[];
  facultyColumn: string;
  studentColumn: string;
  selectFile: FileSelectionHandler;
  clear: () => void;
}) {
  const filename = (path: string) => path.split(/[\\/]/).pop() || path;
  const columnSelect = (
    label: string,
    value: string,
    headers: string[],
    setter: (value: string) => void,
  ) => (
    <label className="input-stack narrow-column">
      {label}
      <select
        value={value}
        onChange={(event) => {
          setter(event.target.value);
          clear();
        }}
      >
        <option value="">Select a column</option>
        {headers.map((header) => (
          <option key={header} value={header}>
            {header}
          </option>
        ))}
      </select>
    </label>
  );
  return (
    <fieldset disabled={disabled}>
      <legend>Exclude faculty matches (optional)</legend>
      <p className="section-description">
        Upload faculty to avoid for each student, such as letter writers or PIs.
      </p>
      <FileDropZone
        kind="spreadsheet"
        label="Exclusion spreadsheet upload"
        onPath={(path) => void rules.loadPreview(path)}
        disabled={disabled}
      >
        <label className="input-stack">
          Exclusion spreadsheet
          <input
            type="text"
            value={rules.path}
            placeholder="Choose or drop a spreadsheet"
            onChange={(event) => rules.changePath(event.target.value)}
          />
        </label>
        <div className="button-row inline">
          <button
            type="button"
            className="secondary"
            onClick={async () => {
              const path = await selectFile(rules.changePath, {
                multiple: false,
                filters: [
                  {
                    name: "Spreadsheets",
                    extensions: [...UPLOAD_EXTENSIONS.spreadsheet],
                  },
                ],
              });
              if (path) await rules.loadPreview(path);
            }}
          >
            Browse…
          </button>
          {rules.path && (
            <>
              <button
                type="button"
                disabled={rules.isLoading}
                onClick={() => void rules.loadPreview()}
              >
                {rules.isLoading ? "Loading…" : "Load exclusion columns"}
              </button>
              <button
                type="button"
                className="secondary"
                onClick={() => rules.changePath("")}
              >
                Remove exclusion file
              </button>
            </>
          )}
        </div>
      </FileDropZone>
      {rules.error && (
        <p role="alert" className="preview-error">
          {rules.error}
        </p>
      )}
      {rules.preview && (
        <div className="exclusion-columns">
          <section
            className="exclusion-file-panel"
            aria-labelledby="exclusion-upload-heading"
          >
            <header className="exclusion-file-heading">
              <h3 id="exclusion-upload-heading">Exclusion file</h3>
              <p title={rules.path}>{filename(rules.path)}</p>
            </header>
            {columnSelect(
              "Student ID column in exclusion file",
              rules.exclusionStudentIdColumn,
              rules.preview.headers,
              rules.setExclusionStudentIdColumn,
            )}
            <div className="input-stack">
              <span className="input-heading">
                Faculty name columns to exclude
              </span>
              <div className="column-checkbox-list">
                {rules.preview.headers.map((header) => (
                  <label key={header} className="column-checkbox-option">
                    <input
                      type="checkbox"
                      checked={rules.excludedNameColumns.includes(header)}
                      onChange={() => {
                        rules.setExcludedNameColumns((columns) =>
                          columns.includes(header)
                            ? columns.filter((column) => column !== header)
                            : [...columns, header],
                        );
                        clear();
                      }}
                    />
                    <span>{header}</span>
                  </label>
                ))}
              </div>
            </div>
          </section>
          <div className="exclusion-other-files">
            <section
              className="exclusion-file-panel"
              aria-labelledby="exclusion-student-heading"
            >
              <header className="exclusion-file-heading">
                <h3 id="exclusion-student-heading">
                  {taskType === "spreadsheet"
                    ? "Student file"
                    : taskType === "directory"
                      ? "Student documents"
                      : taskType === "document"
                        ? "Student document"
                        : "Single prompt"}
                </h3>
                {studentSource && (
                  <p title={studentSource}>{filename(studentSource)}</p>
                )}
              </header>
              {taskType === "spreadsheet" &&
                columnSelect(
                  "Student ID column",
                  studentColumn,
                  studentHeaders,
                  rules.setStudentIdColumn,
                )}
              {(taskType === "prompt" || taskType === "document") && (
                <label className="input-stack narrow-column">
                  Student ID
                  <input
                    type="text"
                    value={rules.studentId}
                    aria-describedby="exclusion-student-note"
                    onChange={(event) => {
                      rules.setStudentId(event.target.value);
                      clear();
                    }}
                  />
                </label>
              )}
              <p
                id="exclusion-student-note"
                className="small-note exclusion-file-note"
              >
                {taskType === "directory"
                  ? "Each filename must match a student ID in the exclusion file, for example 021074478.pdf."
                  : taskType === "spreadsheet"
                    ? "Match this column to the student ID column on the left. IDs must match exactly, including leading zeros."
                    : "No student name is needed. Enter the student ID from the exclusion file. Without an ID, remove the exclusion file to match without exclusions."}
              </p>
            </section>
            <section
              className="exclusion-file-panel"
              aria-labelledby="exclusion-faculty-heading"
            >
              <header className="exclusion-file-heading">
                <h3 id="exclusion-faculty-heading">Faculty dataset</h3>
                <p title={facultySource}>{filename(facultySource)}</p>
              </header>
              {columnSelect(
                "Faculty name column",
                facultyColumn,
                facultyHeaders,
                rules.setFacultyNameColumn,
              )}
              <p className="small-note exclusion-file-note">
                Names in this column are compared with the selected excluded
                faculty names on the left.
              </p>
            </section>
          </div>
        </div>
      )}
    </fieldset>
  );
}
