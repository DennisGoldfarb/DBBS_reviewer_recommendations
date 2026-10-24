import { act, renderHook } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { useStudentInput } from "./useStudentInput";
import { useFacultyFilters } from "./useFacultyFilters";
import { useFacultyDataset } from "./useFacultyDataset";
import { backend } from "../services/backend";
import { datasetStatus, deferred, preview } from "../test/fixtures";
import type { FacultyDatasetPreviewResult, SpreadsheetPreview } from "../types";

beforeEach(() => {
  vi.spyOn(backend, "getFacultyDatasetStatus").mockResolvedValue(datasetStatus);
});

describe("student spreadsheet inputs", () => {
  it("loads suggestions automatically for dropped spreadsheets", async () => {
    vi.spyOn(backend, "analyzeSpreadsheet").mockResolvedValue(preview);
    const changed = vi.fn();
    const { result } = renderHook(() => useStudentInput(changed, vi.fn()));
    await act(async () => {
      await result.current.acceptSpreadsheetPath("/students.tsv");
    });
    expect(result.current.spreadsheetPath).toBe("/students.tsv");
    expect(result.current.selectedPromptColumns).toEqual([1]);
    expect(result.current.mapSelectedColumns([1, 1, -1, 99])).toEqual([
      "Research interests",
    ]);
    act(() => result.current.handleSpreadsheetPathInput("/different.tsv"));
    expect(result.current.spreadsheetPreview).toBeNull();
    expect(result.current.selectedPromptColumns).toEqual([]);
    expect(changed).toHaveBeenCalled();
  });
  it("shares the same preview flow with Browse", async () => {
    vi.spyOn(backend, "analyzeSpreadsheet").mockResolvedValue(preview);
    const selectFile = vi.fn(async (setter: (path: string) => void) => {
      setter("/browse.tsv");
      return "/browse.tsv";
    });
    const { result } = renderHook(() => useStudentInput(vi.fn(), selectFile));
    await act(async () => {
      await result.current.handleSpreadsheetSelection();
    });
    expect(result.current.spreadsheetPath).toBe("/browse.tsv");
    expect(result.current.spreadsheetPreview).toEqual(preview);
  });
  it("ignores slow results when a newer file has already loaded", async () => {
    const slow = deferred<SpreadsheetPreview>();
    vi.spyOn(backend, "analyzeSpreadsheet")
      .mockReturnValueOnce(slow.promise)
      .mockResolvedValueOnce({
        ...preview,
        headers: ["New name", "New interests"],
      });
    const { result } = renderHook(() => useStudentInput(vi.fn(), vi.fn()));
    let oldRequest!: Promise<void>;
    act(() => {
      oldRequest = result.current.acceptSpreadsheetPath("/old.tsv");
    });
    await act(async () => {
      await result.current.acceptSpreadsheetPath("/new.tsv");
    });
    await act(async () => {
      slow.resolve(preview);
      await oldRequest;
    });
    expect(result.current.spreadsheetPreview?.headers[0]).toBe("New name");
    expect(result.current.spreadsheetPath).toBe("/new.tsv");
  });
  it("does not resurrect a preview after switching input modes", async () => {
    const pending = deferred<SpreadsheetPreview>();
    vi.spyOn(backend, "analyzeSpreadsheet").mockReturnValue(pending.promise);
    const { result } = renderHook(() => useStudentInput(vi.fn(), vi.fn()));
    let request!: Promise<void>;
    act(() => {
      result.current.handleTaskTypeChange("spreadsheet");
      request = result.current.acceptSpreadsheetPath("/old.tsv");
    });
    act(() => result.current.handleTaskTypeChange("prompt"));
    await act(async () => {
      pending.resolve(preview);
      await request;
    });
    expect(result.current.spreadsheetPreview).toBeNull();
  });
});

describe("faculty roster inputs", () => {
  it("clears identifier mappings when the active faculty dataset changes", async () => {
    vi.spyOn(backend, "previewFacultyRoster").mockResolvedValue({
      preview,
      suggestedIdentifierMatches: { Name: 0 },
      warnings: [],
    });
    const onChange = vi.fn();
    const selectFile = vi.fn();
    const { result, rerender } = renderHook(
      ({ status }) => useFacultyFilters(status, onChange, selectFile),
      { initialProps: { status: datasetStatus } },
    );
    await act(async () => {
      await result.current.acceptRosterPath("/roster.tsv");
    });
    expect(
      result.current.mapRosterColumnSelection(
        result.current.rosterIdentifierMapping,
      ),
    ).toEqual({ Name: "Name" });
    rerender({ status: { ...datasetStatus, path: "/other.tsv" } });
    expect(result.current.rosterPreview).toBeNull();
    expect(result.current.rosterIdentifierMapping).toEqual({});
  });
});

describe("dataset import previews", () => {
  it("automatically prepares embeddings after importing a faculty dataset", async () => {
    vi.spyOn(backend, "previewFacultyDatasetReplacement").mockResolvedValue({
      preview,
      suggestedEmbeddingColumns: [1],
      suggestedIdentifierColumns: [0],
      suggestedProgramColumns: [],
    });
    vi.spyOn(backend, "replaceFacultyDataset").mockResolvedValue(datasetStatus);
    const refresh = vi.fn(async () => {});
    const { result } = renderHook(() => useFacultyDataset(refresh));
    await act(async () => {
      await result.current.openDatasetConfiguration("/faculty.tsv");
    });
    await act(async () => {
      await result.current.handleConfirmDatasetReplacement();
    });
    expect(backend.replaceFacultyDataset).toHaveBeenCalledWith({
      path: "/faculty.tsv",
      configuration: {
        embeddingColumns: [1],
        identifierColumns: [0],
        programColumns: [],
      },
    });
    expect(refresh).toHaveBeenCalledExactlyOnceWith(datasetStatus);
    expect(result.current.isDatasetConfigurationOpen).toBe(false);
    expect(result.current.isDatasetBusy).toBe(false);
  });

  it("ignores previews resolved after the import dialog has been closed", async () => {
    const pending = deferred<FacultyDatasetPreviewResult>();
    vi.spyOn(backend, "previewFacultyDatasetReplacement").mockReturnValue(
      pending.promise,
    );
    const refresh = vi.fn();
    const { result } = renderHook(() => useFacultyDataset(refresh));
    let request!: Promise<void>;
    act(() => {
      request = result.current.openDatasetConfiguration("/faculty.tsv");
    });
    act(() => result.current.closeDatasetConfiguration());
    await act(async () => {
      pending.resolve({
        preview,
        suggestedEmbeddingColumns: [1],
        suggestedIdentifierColumns: [0],
        suggestedProgramColumns: [],
      });
      await request;
    });
    expect(result.current.isDatasetConfigurationOpen).toBe(false);
    expect(result.current.datasetConfigurationPreview).toBeNull();
  });
});
