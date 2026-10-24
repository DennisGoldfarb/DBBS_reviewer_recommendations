import { useCallback, useState, type FormEvent } from "react";
import { backend } from "../services/backend";
import { save } from "@tauri-apps/plugin-dialog";
import type { FacultyExclusionConfig, GeneratedSpreadsheet } from "../types";
import type { useFeedback } from "./useFeedback";
import type { useStudentInput } from "./useStudentInput";
import type { useFacultyFilters } from "./useFacultyFilters";

export function useMatching(
  student: ReturnType<typeof useStudentInput>,
  faculty: ReturnType<typeof useFacultyFilters>,
  facultyRecCount: string,
  feedback: ReturnType<typeof useFeedback>,
  getExclusionConfiguration: () => FacultyExclusionConfig | undefined,
) {
  const { result, setResult, error, setError } = feedback;
  const [isSubmitting, setIsSubmitting] = useState(false);
  const {
    taskType,
    promptText,
    documentPath,
    spreadsheetPath,
    directoryPath,
    spreadsheetPreview,
    selectedIdentifierColumns,
    selectedPromptColumns,
    mapSelectedColumns,
  } = student;
  const {
    facultyScope,
    selectedPrograms,
    customFacultyPath,
    rosterPreview,
    rosterIdentifierMapping,
    rosterWarnings,
    mapRosterColumnSelection,
  } = faculty;
  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setIsSubmitting(true);
    setError(null);
    setResult(null);

    const programFilters = selectedPrograms;
    const facultyRecommendations = Number(facultyRecCount);
    if (
      !Number.isSafeInteger(facultyRecommendations) ||
      facultyRecommendations < 1 ||
      facultyRecommendations > 4_294_967_295
    ) {
      setError(
        "Enter a positive whole number of faculty recommendations per student.",
      );
      setIsSubmitting(false);
      return;
    }
    let rosterColumnMap: Record<string, string> = {};

    if (taskType === "spreadsheet") {
      const trimmedPath = spreadsheetPath.trim();
      if (trimmedPath.length === 0) {
        setError("Provide a spreadsheet containing student prompts.");
        setIsSubmitting(false);
        return;
      }

      if (!spreadsheetPreview) {
        setError(
          "Load the spreadsheet preview to choose columns before submitting.",
        );
        setIsSubmitting(false);
        return;
      }

      if (selectedPromptColumns.length === 0) {
        setError("Select at least one column containing the student prompts.");
        setIsSubmitting(false);
        return;
      }
    }

    if (facultyScope === "custom") {
      const trimmedRosterPath = customFacultyPath.trim();
      if (trimmedRosterPath.length === 0) {
        setError(
          "Provide a faculty roster spreadsheet to limit the faculty list.",
        );
        setIsSubmitting(false);
        return;
      }

      if (!rosterPreview) {
        setError(
          "Load the roster preview to map faculty identifiers before submitting.",
        );
        setIsSubmitting(false);
        return;
      }

      rosterColumnMap = mapRosterColumnSelection(rosterIdentifierMapping);
      if (Object.keys(rosterColumnMap).length === 0) {
        setError(
          "Select at least one roster column that corresponds to a faculty identifier.",
        );
        setIsSubmitting(false);
        return;
      }
    }

    try {
      const facultyExclusions = getExclusionConfiguration();
      const response = await backend.submitMatchingRequest({
        payload: {
          taskType,
          facultyExclusions,
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
          spreadsheetPromptColumns:
            taskType === "spreadsheet"
              ? mapSelectedColumns(selectedPromptColumns)
              : undefined,
          spreadsheetIdentifierColumns:
            taskType === "spreadsheet"
              ? mapSelectedColumns(selectedIdentifierColumns)
              : undefined,
          facultyRosterColumnMap:
            facultyScope === "custom" ? rosterColumnMap : undefined,
          facultyRosterWarnings:
            facultyScope === "custom" && rosterWarnings.length > 0
              ? rosterWarnings
              : undefined,
        },
      });

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

  const downloadGeneratedSpreadsheet = useCallback(
    async (
      spreadsheet: GeneratedSpreadsheet | null | undefined,
      context: "directory" | "spreadsheet",
    ) => {
      if (!spreadsheet?.content) {
        setError(
          context === "directory"
            ? "The directory results are not available for download."
            : "The spreadsheet results are not available for download.",
        );
        return;
      }

      try {
        const defaultName = (() => {
          if (spreadsheet.filename && spreadsheet.filename.trim().length > 0) {
            return spreadsheet.filename.trim();
          }

          const timestamp = new Date()
            .toISOString()
            .replace(/[-:]/g, "")
            .replace("T", "-")
            .replace(/\..+/, "");
          return `DBBS_matches_${timestamp}.tsv`;
        })();
        const extension = defaultName.includes(".")
          ? (defaultName.split(".").pop()?.toLowerCase() ?? "tsv")
          : "tsv";

        const saveOptions =
          extension && extension.length > 0
            ? {
                defaultPath: defaultName,
                filters: [
                  {
                    name: `${extension.toUpperCase()} files`,
                    extensions: [extension],
                  },
                ],
              }
            : { defaultPath: defaultName };

        const selectedPath = await save(saveOptions);
        if (!selectedPath) {
          return;
        }

        await backend.saveGeneratedSpreadsheet({
          path: selectedPath,
          content: spreadsheet.content,
          encoding: spreadsheet.encoding ?? null,
        });

        setError(null);
      } catch (downloadError) {
        const message =
          downloadError instanceof Error
            ? downloadError.message
            : String(downloadError);
        setError(`Unable to save the ${context} results: ${message}`);
      }
    },
    [setError],
  );

  const handleDownloadDirectoryResults = useCallback(async () => {
    await downloadGeneratedSpreadsheet(
      result?.directoryResults?.spreadsheet ?? null,
      "directory",
    );
  }, [downloadGeneratedSpreadsheet, result]);

  const handleDownloadSpreadsheetResults = useCallback(async () => {
    await downloadGeneratedSpreadsheet(
      result?.spreadsheetResults?.spreadsheet ?? null,
      "spreadsheet",
    );
  }, [downloadGeneratedSpreadsheet, result]);
  return {
    result,
    error,
    isSubmitting,
    handleSubmit,
    handleDownloadDirectoryResults,
    handleDownloadSpreadsheetResults,
  };
}
