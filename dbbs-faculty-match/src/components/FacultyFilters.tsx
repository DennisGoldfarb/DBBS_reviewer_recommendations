import { PreviewTable } from "./PreviewTable";
import { FileDropZone } from "./FileDropZone";
import type { useFacultyFilters } from "../hooks/useFacultyFilters";
import type { FacultyDatasetStatus } from "../types";

export function FacultyFilters({
  acceptRosterPath,
  disabled,
  facultyScope,
  handleFacultyScopeChange,
  isDatasetLoading,
  availablePrograms,
  datasetStatus,
  selectedPrograms,
  toggleProgramSelection,
  handleRosterSelection,
  customFacultyPath,
  handleRosterPathInput,
  loadRosterPreview,
  isLoadingRosterPreview,
  rosterPreviewError,
  rosterWarnings,
  rosterPreview,
  rosterIdentifierMapping,
  handleRosterIdentifierSelection,
  getRosterColumnLabel,
}: ReturnType<typeof useFacultyFilters> & {
  datasetStatus: FacultyDatasetStatus | null;
  isDatasetLoading: boolean;
} & { disabled: boolean }) {
  return (
    <fieldset disabled={disabled}>
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
            disabled={isDatasetLoading || availablePrograms.length === 0}
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
            No program affiliations were detected in the active faculty dataset.
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
            Select the programs that should be included in the faculty roster.
          </p>
        </div>
      )}

      {facultyScope === "custom" && (
        <FileDropZone
          className="input-stack"
          kind="spreadsheet"
          label="Faculty roster upload"
          onPath={acceptRosterPath}
          disabled={disabled}
        >
          <label>Faculty roster spreadsheet</label>
          <div className="button-row inline">
            <button
              type="button"
              className="secondary"
              onClick={() => void handleRosterSelection()}
            >
              Browse…
            </button>
            <input
              type="text"
              value={customFacultyPath}
              onChange={(event) => handleRosterPathInput(event.target.value)}
              placeholder="Paste or confirm the faculty roster path"
            />
            <button
              type="button"
              onClick={() => void loadRosterPreview(customFacultyPath)}
              disabled={
                customFacultyPath.trim().length === 0 || isLoadingRosterPreview
              }
            >
              {isLoadingRosterPreview
                ? "Loading preview…"
                : "Load roster preview"}
            </button>
          </div>
          {customFacultyPath && (
            <div className="path-preview">{customFacultyPath}</div>
          )}
          {rosterPreviewError && (
            <div className="preview-error">{rosterPreviewError}</div>
          )}
          {isLoadingRosterPreview && !rosterPreviewError && (
            <div className="preview-status">Analyzing the roster…</div>
          )}
          {rosterWarnings.length > 0 && (
            <ul className="warning-list">
              {rosterWarnings.map((warning, index) => (
                <li key={`roster-warning-${index}`}>{warning}</li>
              ))}
            </ul>
          )}
          {rosterPreview && (
            <div className="spreadsheet-preview-card">
              <div className="column-selector-group">
                <div className="column-selector">
                  <h4>Roster identifier mapping</h4>
                  <p className="small-note">
                    Match each dataset identifier column to the corresponding
                    column in the roster. Only matched identifiers will be used
                    to limit the faculty list.
                  </p>
                  {datasetStatus?.analysis?.identifierColumns &&
                  datasetStatus.analysis.identifierColumns.length > 0 ? (
                    <div className="column-checkbox-list">
                      {datasetStatus.analysis.identifierColumns.map(
                        (identifier) => {
                          const selected =
                            rosterIdentifierMapping[identifier] ?? null;
                          return (
                            <label
                              key={`roster-identifier-${identifier}`}
                              className="column-checkbox-option"
                            >
                              <span>{identifier}</span>
                              <select
                                value={selected == null ? "" : String(selected)}
                                onChange={(event) =>
                                  handleRosterIdentifierSelection(
                                    identifier,
                                    event.target.value,
                                  )
                                }
                              >
                                <option value="">No match</option>
                                {rosterPreview.headers.map((_, index) => (
                                  <option
                                    key={`roster-column-${identifier}-${index}`}
                                    value={index}
                                  >
                                    {getRosterColumnLabel(index)}
                                  </option>
                                ))}
                              </select>
                            </label>
                          );
                        },
                      )}
                    </div>
                  ) : (
                    <p className="small-note">
                      No identifier columns were detected in the active faculty
                      dataset.
                    </p>
                  )}
                </div>
              </div>
              <div className="preview-table-wrapper">
                <PreviewTable preview={rosterPreview} />
              </div>
            </div>
          )}
          <p className="small-note">
            Upload a tab-delimited TSV/TXT or Excel file listing the available
            faculty members.
          </p>
        </FileDropZone>
      )}
    </fieldset>
  );
}
