import { PreviewTable } from "./PreviewTable";
import type { useFacultyDataset } from "../hooks/useFacultyDataset";

export function DatasetPreviewDialog({
  setIsDatasetPreviewOpen,
  datasetStatus,
}: ReturnType<typeof useFacultyDataset>) {
  if (!datasetStatus?.preview) return null;
  return (
    <div className="dataset-preview-overlay">
      <div
        className="dataset-preview-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="dataset-preview-heading"
      >
        <div className="dataset-preview-header">
          <h3 id="dataset-preview-heading">Faculty dataset preview</h3>
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
            Showing the first {datasetStatus.preview.rows.length} rows of the
            active dataset file.
          </p>
          <div className="preview-table-wrapper">
            <PreviewTable preview={datasetStatus.preview} />
          </div>
        </div>
      </div>
    </div>
  );
}
