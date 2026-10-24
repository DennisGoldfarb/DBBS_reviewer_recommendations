import { UPLOAD_EXTENSIONS } from "../lib/uploads";
import { columnLabel, toggleColumn } from "../lib/columns";
import { useCallback, useEffect, useRef, useState } from "react";
import { backend } from "../services/backend";
import { open } from "@tauri-apps/plugin-dialog";
import type {
  SpreadsheetPreview,
  FacultyDatasetStatus,
  StatusMessage,
} from "../types";

export function useFacultyDataset(
  runEmbeddingRefresh: (status: FacultyDatasetStatus | null) => Promise<void>,
) {
  const previewRequest = useRef(0);
  useEffect(
    () => () => {
      previewRequest.current += 1;
    },
    [],
  );

  const [datasetStatus, setDatasetStatus] =
    useState<FacultyDatasetStatus | null>(null);

  const [isDatasetLoading, setIsDatasetLoading] = useState(true);

  const [isDatasetBusy, setIsDatasetBusy] = useState(false);

  const [datasetBanner, setDatasetBanner] = useState<StatusMessage | null>(
    null,
  );

  const [isDatasetPreviewOpen, setIsDatasetPreviewOpen] = useState(false);

  const [isDatasetConfigurationOpen, setIsDatasetConfigurationOpen] =
    useState(false);

  const [datasetConfigurationPath, setDatasetConfigurationPath] = useState("");

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
    let active = true;
    const loadStatus = async () => {
      try {
        const status = await backend.getFacultyDatasetStatus();
        if (active) applyDatasetStatus(status);
      } catch (statusError) {
        if (!active) return;
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
        if (active) setIsDatasetLoading(false);
      }
    };

    void loadStatus();
    return () => {
      active = false;
    };
  }, [applyDatasetStatus]);

  const selectDatasetFile = async () => {
    try {
      const selection = await open({
        multiple: false,
        filters: [
          {
            name: "Faculty dataset",
            extensions: [...UPLOAD_EXTENSIONS.dataset],
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
    const request = ++previewRequest.current;
    setDatasetConfigurationPath(path);
    setIsDatasetConfigurationOpen(true);
    setIsLoadingDatasetConfiguration(true);
    try {
      const preview = await backend.previewFacultyDatasetReplacement({ path });
      if (request !== previewRequest.current) return;
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
      if (request !== previewRequest.current) return;
      const message =
        previewError instanceof Error
          ? previewError.message
          : String(previewError);
      setDatasetConfigurationError(message);
    } finally {
      if (request === previewRequest.current)
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
      const status = await backend.replaceFacultyDataset({
        path: datasetConfigurationPath,
        configuration: {
          embeddingColumns: datasetConfigurationEmbeddingColumns,
          identifierColumns: datasetConfigurationIdentifierColumns,
          programColumns: datasetConfigurationProgramColumns,
        },
      });
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
      const status = await backend.restoreDefaultFacultyDataset();
      applyDatasetStatus(status, "success");
      if (status.isValid) {
        await runEmbeddingRefresh(status);
      }
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

  const resetDatasetConfiguration = () => {
    previewRequest.current += 1;
    setDatasetConfigurationPreview(null);
    setDatasetConfigurationEmbeddingColumns([]);
    setDatasetConfigurationIdentifierColumns([]);
    setDatasetConfigurationProgramColumns([]);
    setDatasetConfigurationError(null);
    setDatasetConfigurationPath("");
    setIsLoadingDatasetConfiguration(false);
  };

  const getDatasetConfigurationColumnLabel = (index: number) =>
    columnLabel(datasetConfigurationPreview?.headers, index);

  const toggleDatasetEmbeddingColumn = (index: number) => {
    setDatasetConfigurationEmbeddingColumns((current) =>
      toggleColumn(current, index),
    );
    setDatasetConfigurationError(null);
  };

  const toggleDatasetIdentifierColumn = (index: number) => {
    setDatasetConfigurationIdentifierColumns((current) =>
      toggleColumn(current, index),
    );
    setDatasetConfigurationError(null);
  };

  const toggleDatasetProgramColumn = (index: number) => {
    setDatasetConfigurationProgramColumns((current) =>
      toggleColumn(current, index),
    );
    setDatasetConfigurationError(null);
  };

  return {
    datasetStatus,
    isDatasetLoading,
    isDatasetBusy,
    datasetBanner,
    isDatasetPreviewOpen,
    setIsDatasetPreviewOpen,
    isDatasetConfigurationOpen,
    datasetConfigurationPath,
    datasetConfigurationPreview,
    isLoadingDatasetConfiguration,
    datasetConfigurationEmbeddingColumns,
    datasetConfigurationIdentifierColumns,
    datasetConfigurationProgramColumns,
    datasetConfigurationError,
    openDatasetConfiguration,
    handleDatasetReplacement,
    closeDatasetConfiguration,
    handleConfirmDatasetReplacement,
    handleDatasetRestore,
    getDatasetConfigurationColumnLabel,
    toggleDatasetEmbeddingColumn,
    toggleDatasetIdentifierColumn,
    toggleDatasetProgramColumn,
  };
}
