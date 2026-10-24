import { UPLOAD_EXTENSIONS } from "../lib/uploads";
import { columnLabel } from "../lib/columns";
import { useCallback, useEffect, useRef, useState } from "react";
import { backend } from "../services/backend";
import type { FileSelectionHandler } from "../services/fileDialogs";
import type {
  FacultyScope,
  ProgramName,
  SpreadsheetPreview,
  FacultyDatasetStatus,
} from "../types";

export function useFacultyFilters(
  datasetStatus: FacultyDatasetStatus | null,
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

  const [facultyScope, setFacultyScope] = useState<FacultyScope>("all");

  const [availablePrograms, setAvailablePrograms] = useState<ProgramName[]>([]);

  const [selectedPrograms, setSelectedPrograms] = useState<ProgramName[]>([]);

  const [customFacultyPath, setCustomFacultyPath] = useState("");

  const [rosterPreview, setRosterPreview] = useState<SpreadsheetPreview | null>(
    null,
  );

  const [rosterPreviewError, setRosterPreviewError] = useState<string | null>(
    null,
  );

  const [isLoadingRosterPreview, setIsLoadingRosterPreview] = useState(false);

  const [rosterIdentifierMapping, setRosterIdentifierMapping] = useState<
    Record<string, number | null>
  >({});

  const [rosterWarnings, setRosterWarnings] = useState<string[]>([]);

  const resetRosterConfiguration = useCallback(() => {
    previewRequest.current += 1;
    setRosterPreview(null);
    setRosterPreviewError(null);
    setIsLoadingRosterPreview(false);
    setRosterIdentifierMapping({});
    setRosterWarnings([]);
  }, []);

  useEffect(() => {
    if (!datasetStatus?.analysis) {
      setAvailablePrograms([]);
      setSelectedPrograms([]);
      if (facultyScope === "program") {
        setFacultyScope("all");
      }
      return;
    }

    const programs = [...datasetStatus.analysis.availablePrograms];
    setAvailablePrograms(programs);
    setSelectedPrograms((previous) => {
      const filtered = previous.filter((program) => programs.includes(program));
      if (
        filtered.length === previous.length &&
        filtered.every((value, index) => value === previous[index])
      ) {
        return previous;
      }
      return filtered;
    });

    if (programs.length === 0 && facultyScope === "program") {
      setFacultyScope("all");
    }
  }, [datasetStatus, facultyScope]);

  const handleFacultyScopeChange = (value: FacultyScope) => {
    if (value === "program" && availablePrograms.length === 0) {
      return;
    }

    setFacultyScope(value);
    onChange();

    if (value !== "custom") {
      setCustomFacultyPath("");
      resetRosterConfiguration();
    }

    if (value !== "program") {
      setSelectedPrograms([]);
    }
  };

  const toggleProgramSelection = (program: ProgramName) => {
    if (!availablePrograms.includes(program)) {
      return;
    }

    setSelectedPrograms((current) => {
      if (current.includes(program)) {
        return current.filter((entry) => entry !== program);
      }

      return [...current, program];
    });
    onChange();
  };

  const handleRosterPathInput = (value: string) => {
    setCustomFacultyPath(value);
    onChange();
    resetRosterConfiguration();
  };

  const loadRosterPreview = async (path: string) => {
    const trimmed = path.trim();
    if (trimmed.length === 0) {
      resetRosterConfiguration();
      return;
    }

    const request = ++previewRequest.current;
    setRosterPreview(null);
    setRosterIdentifierMapping({});
    setRosterWarnings([]);
    setIsLoadingRosterPreview(true);
    setRosterPreviewError(null);
    onChange();

    try {
      const response = await backend.previewFacultyRoster({
        path: trimmed,
      });

      if (request !== previewRequest.current) return;
      const mapping: Record<string, number | null> = {};
      for (const [identifier, index] of Object.entries(
        response.suggestedIdentifierMatches ?? {},
      )) {
        mapping[identifier] =
          typeof index === "number" && Number.isInteger(index) ? index : null;
      }

      const datasetIdentifiers =
        datasetStatus?.analysis?.identifierColumns ?? [];
      for (const identifier of datasetIdentifiers) {
        if (!(identifier in mapping)) {
          mapping[identifier] = null;
        }
      }

      setRosterPreview(response.preview);
      setRosterWarnings(response.warnings ?? []);
      setRosterIdentifierMapping(mapping);
    } catch (analysisError) {
      if (request !== previewRequest.current) return;
      const message =
        analysisError instanceof Error
          ? analysisError.message
          : String(analysisError);
      setRosterPreview(null);
      setRosterWarnings([]);
      setRosterIdentifierMapping({});
      setRosterPreviewError(message);
    } finally {
      if (request === previewRequest.current) setIsLoadingRosterPreview(false);
    }
  };

  useEffect(() => {
    resetRosterConfiguration();
  }, [datasetStatus, resetRosterConfiguration]);

  const acceptRosterPath = async (path: string) => {
    handleRosterPathInput(path);
    await loadRosterPreview(path);
  };

  const handleRosterSelection = async () => {
    const selection = await handleFileSelection(handleRosterPathInput, {
      multiple: false,
      filters: [
        {
          name: "Faculty rosters",
          extensions: [...UPLOAD_EXTENSIONS.spreadsheet],
        },
      ],
    });

    if (selection) {
      await loadRosterPreview(selection);
    }
  };

  const handleRosterIdentifierSelection = (
    identifier: string,
    rawIndex: string,
  ) => {
    setRosterIdentifierMapping((current) => {
      const updated: Record<string, number | null> = { ...current };
      if (rawIndex === "") {
        updated[identifier] = null;
      } else {
        const parsed = Number.parseInt(rawIndex, 10);
        updated[identifier] = Number.isNaN(parsed) ? null : parsed;
      }
      return updated;
    });
    onChange();
  };

  const getRosterColumnLabel = (index: number) =>
    columnLabel(rosterPreview?.headers, index);

  const mapRosterColumnSelection = (
    mapping: Record<string, number | null>,
  ): Record<string, string> => {
    const entries: [string, string][] = [];

    for (const [identifier, index] of Object.entries(mapping)) {
      if (typeof index !== "number" || Number.isNaN(index)) {
        continue;
      }

      if (!rosterPreview) {
        entries.push([identifier, `Column ${index + 1}`]);
        continue;
      }

      if (index < 0 || index >= rosterPreview.headers.length) {
        continue;
      }

      entries.push([identifier, getRosterColumnLabel(index)]);
    }

    return Object.fromEntries(entries);
  };

  return {
    facultyScope,
    availablePrograms,
    selectedPrograms,
    customFacultyPath,
    rosterPreview,
    rosterPreviewError,
    isLoadingRosterPreview,
    rosterIdentifierMapping,
    rosterWarnings,
    handleFacultyScopeChange,
    toggleProgramSelection,
    handleRosterPathInput,
    loadRosterPreview,
    acceptRosterPath,
    handleRosterSelection,
    handleRosterIdentifierSelection,
    getRosterColumnLabel,
    mapRosterColumnSelection,
  };
}
