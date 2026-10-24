import { PreviewTable } from "./PreviewTable";
import type { useFacultyDataset } from "../hooks/useFacultyDataset";

export function DatasetConfigurationDialog({
  closeDatasetConfiguration,
  isDatasetBusy,
  datasetConfigurationPath,
  datasetConfigurationError,
  isLoadingDatasetConfiguration,
  datasetConfigurationPreview,
  datasetConfigurationEmbeddingColumns,
  toggleDatasetEmbeddingColumn,
  getDatasetConfigurationColumnLabel,
  datasetConfigurationIdentifierColumns,
  toggleDatasetIdentifierColumn,
  datasetConfigurationProgramColumns,
  toggleDatasetProgramColumn,
  handleConfirmDatasetReplacement,
}: ReturnType<typeof useFacultyDataset>) {
  return (
    <div className="dataset-preview-overlay">
      <div
        className="dataset-preview-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="dataset-import-heading"
      >
        <div className="dataset-preview-header">
          <h3 id="dataset-import-heading">Import faculty dataset</h3>
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
                    Choose the columns describing faculty research interests or
                    other content that should be embedded. Long descriptions are
                    automatically pruned on this computer with the small
                    LLMLingua-2 model to fit the matching limit. Your original
                    descriptions are preserved. First use may download the
                    model.
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
                          <span>
                            {getDatasetConfigurationColumnLabel(index)}
                          </span>
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
                            onChange={() =>
                              toggleDatasetIdentifierColumn(index)
                            }
                            disabled={isDatasetBusy}
                          />
                          <span>
                            {getDatasetConfigurationColumnLabel(index)}
                          </span>
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
                          <span>
                            {getDatasetConfigurationColumnLabel(index)}
                          </span>
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
                <PreviewTable preview={datasetConfigurationPreview} />
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
  );
}
