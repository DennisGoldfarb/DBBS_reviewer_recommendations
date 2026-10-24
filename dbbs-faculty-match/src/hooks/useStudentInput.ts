import { UPLOAD_EXTENSIONS } from "../lib/uploads";
import { columnLabel, toggleColumn } from "../lib/columns";
import { useEffect, useRef, useState } from "react";
import { backend } from "../services/backend";
import type { FileSelectionHandler } from "../services/fileDialogs";
import type { TaskType, SpreadsheetPreview } from "../types";

export function useStudentInput(
  onChange: () => void,
  handleFileSelection: FileSelectionHandler,
) {
  const previewRequest = useRef(0);
  useEffect(
    () => () => {
      previewRequest.current += 1;
    },
    [],
  );
  const [taskType, setTaskType] = useState<TaskType>("prompt");

  const [promptText, updatePromptText] = useState("");
  const setPromptText = (value: string) => {
    updatePromptText(value);
    onChange();
  };

  const [documentPath, updateDocumentPath] = useState("");
  const setDocumentPath = (value: string) => {
    updateDocumentPath(value);
    onChange();
  };

  const [spreadsheetPath, setSpreadsheetPath] = useState("");

  const [directoryPath, updateDirectoryPath] = useState("");
  const setDirectoryPath = (value: string) => {
    updateDirectoryPath(value);
    onChange();
  };

  const [spreadsheetPreview, setSpreadsheetPreview] =
    useState<SpreadsheetPreview | null>(null);

  const [spreadsheetPreviewError, setSpreadsheetPreviewError] = useState<
    string | null
  >(null);

  const [isLoadingSpreadsheetPreview, setIsLoadingSpreadsheetPreview] =
    useState(false);

  const [selectedIdentifierColumns, setSelectedIdentifierColumns] = useState<
    number[]
  >([]);

  const [selectedPromptColumns, setSelectedPromptColumns] = useState<number[]>(
    [],
  );

  const resetSpreadsheetConfiguration = () => {
    previewRequest.current += 1;
    setSpreadsheetPreview(null);
    setSpreadsheetPreviewError(null);
    setSelectedIdentifierColumns([]);
    setSelectedPromptColumns([]);
    setIsLoadingSpreadsheetPreview(false);
  };

  const handleTaskTypeChange = (value: TaskType) => {
    setTaskType(value);
    onChange();

    if (value !== "spreadsheet") {
      resetSpreadsheetConfiguration();
    }
  };

  const handleSpreadsheetPathInput = (value: string) => {
    setSpreadsheetPath(value);
    onChange();
    resetSpreadsheetConfiguration();
  };

  const loadSpreadsheetPreview = async (path: string) => {
    const trimmed = path.trim();
    if (trimmed.length === 0) {
      resetSpreadsheetConfiguration();
      return;
    }

    const request = ++previewRequest.current;
    setSpreadsheetPreview(null);
    setSelectedIdentifierColumns([]);
    setSelectedPromptColumns([]);
    setIsLoadingSpreadsheetPreview(true);
    setSpreadsheetPreviewError(null);
    onChange();

    try {
      const preview = await backend.analyzeSpreadsheet({
        path: trimmed,
      });
      if (request !== previewRequest.current) return;

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
      if (request !== previewRequest.current) return;
      const message =
        analysisError instanceof Error
          ? analysisError.message
          : String(analysisError);
      setSpreadsheetPreview(null);
      setSpreadsheetPreviewError(message);
      setSelectedIdentifierColumns([]);
      setSelectedPromptColumns([]);
    } finally {
      if (request === previewRequest.current)
        setIsLoadingSpreadsheetPreview(false);
    }
  };

  const acceptSpreadsheetPath = async (path: string) => {
    handleSpreadsheetPathInput(path);
    await loadSpreadsheetPreview(path);
  };

  const handleSpreadsheetSelection = async () => {
    const selection = await handleFileSelection(handleSpreadsheetPathInput, {
      multiple: false,
      filters: [
        {
          name: "Spreadsheets",
          extensions: [...UPLOAD_EXTENSIONS.spreadsheet],
        },
      ],
    });

    if (selection) {
      await loadSpreadsheetPreview(selection);
    }
  };

  const toggleIdentifierColumn = (index: number) => {
    setSelectedIdentifierColumns((current) => toggleColumn(current, index));
    onChange();
  };

  const togglePromptColumn = (index: number) => {
    setSelectedPromptColumns((current) => toggleColumn(current, index));
    onChange();
  };

  const getColumnLabel = (index: number) =>
    columnLabel(spreadsheetPreview?.headers, index);

  const mapSelectedColumns = (indexes: number[]) => {
    const unique = Array.from(new Set(indexes)).filter((index) => index >= 0);

    if (!spreadsheetPreview) {
      return unique.map((index) => `Column ${index + 1}`);
    }

    return unique
      .filter((index) => index < spreadsheetPreview.headers.length)
      .map((index) => getColumnLabel(index));
  };

  return {
    taskType,
    promptText,
    setPromptText,
    documentPath,
    setDocumentPath,
    spreadsheetPath,
    directoryPath,
    setDirectoryPath,
    spreadsheetPreview,
    spreadsheetPreviewError,
    isLoadingSpreadsheetPreview,
    selectedIdentifierColumns,
    selectedPromptColumns,
    handleTaskTypeChange,
    handleSpreadsheetPathInput,
    loadSpreadsheetPreview,
    acceptSpreadsheetPath,
    handleSpreadsheetSelection,
    toggleIdentifierColumn,
    togglePromptColumn,
    getColumnLabel,
    mapSelectedColumns,
  };
}
