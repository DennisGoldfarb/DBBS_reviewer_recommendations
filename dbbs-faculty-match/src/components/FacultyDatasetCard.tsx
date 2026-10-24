import { FileDropZone } from "./FileDropZone";
import type { useFacultyDataset } from "../hooks/useFacultyDataset";
import { formatDatasetTimestamp } from "../lib/formatting";

export function FacultyDatasetCard({
  openDatasetConfiguration,
  isDatasetLoading,
  datasetStatus,
  handleDatasetReplacement,
  areControlsDisabled,
  isDatasetBusy,
  isDatasetConfigurationOpen,
  handleDatasetRestore,
  setIsDatasetPreviewOpen,
}: ReturnType<typeof useFacultyDataset> & { areControlsDisabled: boolean }) {
  return (
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
        Manage the dataset file that seeds the faculty embedding index before
        running updates.
      </p>
      <dl className="dataset-meta-grid">
        <div>
          <dt>Active file</dt>
          <dd className="dataset-path">
            {isDatasetLoading
              ? "Loading…"
              : (datasetStatus?.sourcePath ??
                datasetStatus?.canonicalPath ??
                datasetStatus?.path ??
                "Not configured")}
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
          (datasetStatus.isValid ? "success" : "error")) !== "success" && (
          <p
            className={`dataset-message ${
              (datasetStatus.messageVariant ??
                (datasetStatus.isValid ? "success" : "error")) === "success"
                ? "dataset-message-success"
                : (datasetStatus.messageVariant ??
                      (datasetStatus.isValid ? "success" : "error")) === "error"
                  ? "dataset-message-error"
                  : "dataset-message-info"
            }`}
          >
            {datasetStatus.message}
          </p>
        )}
      <FileDropZone
        kind="dataset"
        label="Faculty dataset upload"
        onPath={openDatasetConfiguration}
        disabled={
          areControlsDisabled ||
          isDatasetBusy ||
          isDatasetLoading ||
          isDatasetConfigurationOpen
        }
      >
        <div className="dataset-actions">
          <button
            type="button"
            onClick={handleDatasetReplacement}
            disabled={
              areControlsDisabled ||
              isDatasetBusy ||
              isDatasetLoading ||
              isDatasetConfigurationOpen
            }
          >
            Replace dataset
          </button>
          <button
            type="button"
            className="secondary"
            onClick={handleDatasetRestore}
            disabled={
              areControlsDisabled ||
              isDatasetBusy ||
              isDatasetLoading ||
              isDatasetConfigurationOpen
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
      </FileDropZone>
    </section>
  );
}
