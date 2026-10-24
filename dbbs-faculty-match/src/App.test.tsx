import {
  act,
  fireEvent,
  render,
  screen,
  waitFor,
} from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, expect, it, vi } from "vitest";
import App from "./App";
import { THEME_STORAGE_KEY } from "./lib/theme";
import { backend } from "./services/backend";
import { datasetStatus, deferred, matchResponse } from "./test/fixtures";
import type { SubmissionResponse } from "./types";

vi.mock("@tauri-apps/api/event", () => ({
  listen: vi.fn(async () => vi.fn()),
}));
vi.mock("@tauri-apps/plugin-dialog", () => ({ open: vi.fn(), save: vi.fn() }));

beforeEach(() => {
  vi.spyOn(backend, "getFacultyDatasetStatus").mockResolvedValue(datasetStatus);
  vi.spyOn(backend, "ensureEmbeddingHelperReady").mockResolvedValue(true);
  vi.spyOn(backend, "submitMatchingRequest").mockResolvedValue(matchResponse);
});

it("runs prompt matching and clears old results when the input changes", async () => {
  const user = userEvent.setup();
  render(<App />);
  await waitFor(() =>
    expect(
      screen.getByRole<HTMLButtonElement>("button", { name: "Run matching" })
        .disabled,
    ).toBe(false),
  );
  await user.type(screen.getByLabelText("Prompt text"), "Protein structure");
  await user.click(screen.getByRole("button", { name: "Run matching" }));
  expect(await screen.findByText("Matching completed.")).toBeTruthy();
  expect(screen.getByText("Example faculty")).toBeTruthy();
  expect(backend.submitMatchingRequest).toHaveBeenCalledWith(
    expect.objectContaining({
      payload: expect.objectContaining({
        taskType: "prompt",
        facultyRecsPerStudent: 10,
        promptText: "Protein structure",
      }),
    }),
  );
  await user.type(screen.getByLabelText("Prompt text"), " and signaling");
  expect(screen.queryByText("Matching completed.")).toBeNull();
});

it("locks the inputs while a matching job is running", async () => {
  const pending = deferred<SubmissionResponse>();
  vi.mocked(backend.submitMatchingRequest).mockReturnValue(pending.promise);
  const user = userEvent.setup();
  render(<App />);
  await waitFor(() =>
    expect(
      screen.getByRole<HTMLButtonElement>("button", { name: "Run matching" })
        .disabled,
    ).toBe(false),
  );
  await user.type(screen.getByLabelText("Prompt text"), "Protein structure");
  await user.click(screen.getByRole("button", { name: "Run matching" }));
  expect(
    screen.getByRole<HTMLButtonElement>("button", { name: "Matching…" })
      .disabled,
  ).toBe(true);
  expect(
    screen
      .getByLabelText<HTMLInputElement>("Single document")
      .closest("fieldset")?.disabled,
  ).toBe(true);
  await act(async () => {
    pending.resolve(matchResponse);
  });
  expect(await screen.findByText("Matching completed.")).toBeTruthy();
});

it("rejects fractional recommendation counts even when the form is submitted directly", async () => {
  render(<App />);
  await screen.findByText("Ready");
  fireEvent.change(
    screen.getByLabelText("Faculty recommendations per student"),
    { target: { value: "1.5" } },
  );
  fireEvent.submit(
    screen.getByRole("button", { name: "Run matching" }).closest("form")!,
  );
  expect(await screen.findByRole("alert")).toHaveProperty(
    "textContent",
    "Enter a positive whole number of faculty recommendations per student.",
  );
  expect(backend.submitMatchingRequest).not.toHaveBeenCalled();
});

it("persists theme changes", async () => {
  const user = userEvent.setup();
  render(<App />);
  await user.click(
    screen.getByRole("button", { name: "Switch to light WashU theme" }),
  );
  expect(document.documentElement.dataset.theme).toBe("light");
  expect(window.localStorage.getItem(THEME_STORAGE_KEY)).toBe("light");
});

it("can retry initialization after the embedding helper fails to start", async () => {
  vi.mocked(backend.ensureEmbeddingHelperReady)
    .mockRejectedValueOnce("Helper unavailable")
    .mockResolvedValueOnce(true);
  const user = userEvent.setup();
  render(<App />);
  await user.click(
    await screen.findByRole("button", { name: "Retry embedding helper" }),
  );
  await waitFor(() =>
    expect(
      screen.getByRole<HTMLButtonElement>("button", { name: "Run matching" })
        .disabled,
    ).toBe(false),
  );
  expect(backend.ensureEmbeddingHelperReady).toHaveBeenCalledTimes(2);
});
