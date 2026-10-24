import { useRef, useState } from "react";
import { backend } from "../services/backend";
import type { FacultyExclusionConfig, SpreadsheetPreview } from "../types";

export function useFacultyExclusions(clear: () => void) {
  const [path, setPath] = useState("");
  const [preview, setPreview] = useState<SpreadsheetPreview | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [exclusionStudentIdColumn, setExclusionStudentIdColumn] = useState("");
  const [excludedNameColumns, setExcludedNameColumns] = useState<string[]>([]);
  const [facultyNameColumn, setFacultyNameColumn] = useState("");
  const [studentIdColumn, setStudentIdColumn] = useState("");
  const [studentId, setStudentId] = useState("");
  const request = useRef(0);

  function changePath(value: string) {
    request.current += 1;
    setPath(value);
    setPreview(null);
    setExclusionStudentIdColumn("");
    setExcludedNameColumns([]);
    setError(null);
    setIsLoading(false);
    clear();
  }

  async function loadPreview(value = path) {
    changePath(value);
    if (!value.trim()) return;
    const current = request.current;
    setIsLoading(true);
    try {
      const result = await backend.analyzeSpreadsheet({ path: value.trim() });
      if (current !== request.current) return;
      setPreview(result);
      setExclusionStudentIdColumn(
        result.headers.find((header) =>
          /^person reference id$/i.test(header.trim()),
        ) ??
          result.headers.find((header) => /^id$/i.test(header.trim())) ??
          "",
      );
      setExcludedNameColumns(
        result.headers.filter((header) =>
          /^ref\s*\d+\s+name$/i.test(header.trim()),
        ),
      );
    } catch (cause) {
      if (current === request.current) setError(String(cause));
    } finally {
      if (current === request.current) setIsLoading(false);
    }
  }

  function configuration(
    taskType: string,
    effectiveFacultyColumn: string,
    effectiveStudentColumn: string,
  ): FacultyExclusionConfig | undefined {
    if (!path.trim()) return undefined;
    if (!preview || isLoading)
      throw new Error("Load the exclusion spreadsheet before matching.");
    if (
      !exclusionStudentIdColumn ||
      !excludedNameColumns.length ||
      !effectiveFacultyColumn
    )
      throw new Error(
        "Select the student ID, excluded faculty names, and faculty name columns.",
      );
    if (taskType === "spreadsheet" && !effectiveStudentColumn)
      throw new Error("Select the student ID column for faculty exclusions.");
    if ((taskType === "prompt" || taskType === "document") && !studentId.trim())
      throw new Error("Enter the student ID for faculty exclusions.");
    return {
      path: path.trim(),
      exclusionStudentIdColumn,
      excludedNameColumns,
      facultyNameColumn: effectiveFacultyColumn,
      studentIdColumn: effectiveStudentColumn,
      studentId: studentId.trim(),
    };
  }

  return {
    path,
    changePath,
    preview,
    isLoading,
    error,
    loadPreview,
    exclusionStudentIdColumn,
    setExclusionStudentIdColumn,
    excludedNameColumns,
    setExcludedNameColumns,
    facultyNameColumn,
    setFacultyNameColumn,
    studentIdColumn,
    setStudentIdColumn,
    studentId,
    setStudentId,
    configuration,
  };
}
