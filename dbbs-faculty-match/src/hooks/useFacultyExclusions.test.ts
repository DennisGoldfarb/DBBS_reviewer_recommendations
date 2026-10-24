import { act, renderHook } from "@testing-library/react";
import { expect, it, vi } from "vitest";
import { backend } from "../services/backend";
import { deferred, preview } from "../test/fixtures";
import type { SpreadsheetPreview } from "../types";
import { useFacultyExclusions } from "./useFacultyExclusions";

it("does not restore a removed exclusion file when an older preview finishes", async () => {
  const pending = deferred<SpreadsheetPreview>();
  vi.spyOn(backend, "analyzeSpreadsheet").mockReturnValue(pending.promise);
  const { result } = renderHook(() => useFacultyExclusions(vi.fn()));
  let load: Promise<void>;
  act(() => {
    load = result.current.loadPreview("/writers.xlsx");
  });
  act(() => result.current.changePath(""));
  await act(async () => {
    pending.resolve(preview);
    await load;
  });
  expect(result.current.preview).toBeNull();
  expect(result.current.path).toBe("");
  expect(result.current.configuration("prompt", "Name", "ID")).toBeUndefined();
});
