import { act, render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, expect, it, vi } from "vitest";
import type { DragDropEvent } from "@tauri-apps/api/webview";
import App from "./App";
import { backend } from "./services/backend";
import { datasetStatus, matchResponse, preview } from "./test/fixtures";

const native = vi.hoisted(() => ({
  listeners: new Set<(event: { payload: DragDropEvent }) => Promise<void>>(),
}));
vi.mock("@tauri-apps/api/core", () => ({
  isTauri: () => true,
  invoke: vi.fn(),
}));
vi.mock("@tauri-apps/api/event", () => ({
  listen: vi.fn(async () => vi.fn()),
}));
vi.mock("@tauri-apps/plugin-dialog", () => ({ open: vi.fn(), save: vi.fn() }));
vi.mock("@tauri-apps/api/window", () => ({
  getCurrentWindow: () => ({
    setTheme: vi.fn(async () => {}),
    onDragDropEvent: vi.fn(async (handler) => {
      native.listeners.add(handler);
      return () => native.listeners.delete(handler);
    }),
  }),
}));

beforeEach(() => {
  native.listeners.clear();
  vi.spyOn(backend, "getFacultyDatasetStatus").mockResolvedValue(datasetStatus);
  vi.spyOn(backend, "ensureEmbeddingHelperReady").mockResolvedValue(true);
  vi.spyOn(backend, "submitMatchingRequest").mockResolvedValue(matchResponse);
  vi.spyOn(backend, "analyzeSpreadsheet").mockResolvedValue(preview);
  vi.spyOn(backend, "validateUploadPath").mockImplementation(
    async ({ path }) => path,
  );
  vi.spyOn(backend, "previewFacultyRoster").mockResolvedValue({
    preview,
    warnings: [],
    suggestedIdentifierMatches: { Name: 0 },
  });
  vi.spyOn(backend, "previewFacultyDatasetReplacement").mockResolvedValue({
    preview,
    suggestedEmbeddingColumns: [1],
    suggestedIdentifierColumns: [0],
    suggestedProgramColumns: [],
  });
  vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockImplementation(
    function (this: HTMLElement) {
      const top =
        this.getAttribute("aria-label") === "Faculty dataset upload"
          ? 400
          : this.getAttribute("aria-label") === "Exclusion spreadsheet upload"
            ? 200
            : 0;
      return { left: 0, top, right: 100, bottom: top + 100 } as DOMRect;
    },
  );
});

async function drop(path: string, y = 20) {
  await act(async () => {
    await Promise.all(
      [...native.listeners].map((handler) =>
        handler({
          payload: {
            type: "drop",
            paths: [path],
            position: { x: 20, y },
          } as DragDropEvent,
        }),
      ),
    );
  });
}

it("uses a dropped exclusion file and removes its rules when cleared", async () => {
  vi.mocked(backend.analyzeSpreadsheet).mockResolvedValue({
    headers: ["Person Reference ID", "Ref 1 Name", "Ref 1 Email", "Ref 2 Name"],
    rows: [
      ["00123", "Sabine Dietmann", "ignored@example.org", "Gabor Egervari"],
    ],
    suggestedIdentifierColumns: [0],
    suggestedPromptColumns: [],
  });
  const user = userEvent.setup();
  render(<App />);
  await screen.findByText("Ready");
  await user.type(screen.getByLabelText("Prompt text"), "Protein structure");
  await waitFor(() => expect(native.listeners.size).toBe(2));
  await drop("/writers.xlsx", 220);
  await screen.findByLabelText("Student ID");
  await user.type(screen.getByLabelText("Student ID"), "00123");
  await user.click(screen.getByRole("button", { name: "Run matching" }));
  expect(backend.submitMatchingRequest).toHaveBeenLastCalledWith({
    payload: expect.objectContaining({
      facultyExclusions: {
        path: "/writers.xlsx",
        exclusionStudentIdColumn: "Person Reference ID",
        excludedNameColumns: ["Ref 1 Name", "Ref 2 Name"],
        facultyNameColumn: "Name",
        studentIdColumn: "",
        studentId: "00123",
      },
    }),
  });
  await user.click(
    screen.getByRole("button", { name: "Remove exclusion file" }),
  );
  expect(screen.queryByText("Matching completed.")).toBeNull();
  await user.click(screen.getByRole("button", { name: "Run matching" }));
  expect(backend.submitMatchingRequest).toHaveBeenLastCalledWith({
    payload: expect.objectContaining({ facultyExclusions: undefined }),
  });
});

it("requires a student identity when exclusions are enabled for a prompt", async () => {
  vi.mocked(backend.analyzeSpreadsheet).mockResolvedValue({
    headers: ["Person Reference ID", "Ref 1 Name"],
    rows: [],
    suggestedIdentifierColumns: [0],
    suggestedPromptColumns: [],
  });
  const user = userEvent.setup();
  render(<App />);
  await screen.findByText("Ready");
  await user.type(screen.getByLabelText("Prompt text"), "Protein structure");
  await waitFor(() => expect(native.listeners.size).toBe(2));
  await drop("/writers.xlsx", 220);
  await screen.findByLabelText("Student ID");
  await user.click(screen.getByRole("button", { name: "Run matching" }));
  expect(
    await screen.findByText("Enter the student ID for faculty exclusions."),
  ).toBeTruthy();
  expect(backend.submitMatchingRequest).not.toHaveBeenCalled();
});

it("links student spreadsheet IDs to the exclusion file independently of output columns", async () => {
  vi.mocked(backend.analyzeSpreadsheet).mockImplementation(async ({ path }) =>
    path === "/writers.xlsx"
      ? {
          headers: ["Person Reference ID", "Ref 1 Name"],
          rows: [["00123", "Gabor Egervari"]],
          suggestedIdentifierColumns: [0],
          suggestedPromptColumns: [],
        }
      : {
          headers: ["ID", "Name", "Interests"],
          rows: [["00123", "Student", "Protein structure"]],
          suggestedIdentifierColumns: [1],
          suggestedPromptColumns: [2],
        },
  );
  const user = userEvent.setup();
  render(<App />);
  await screen.findByText("Ready");
  await user.click(screen.getByRole("radio", { name: "Spreadsheet" }));
  await waitFor(() => expect(native.listeners.size).toBe(3));
  await drop("/students.tsv");
  await drop("/writers.xlsx", 220);
  expect(
    (await screen.findByLabelText<HTMLSelectElement>("Student ID column"))
      .value,
  ).toBe("ID");
  await user.click(screen.getByRole("button", { name: "Run matching" }));
  expect(backend.submitMatchingRequest).toHaveBeenCalledWith({
    payload: expect.objectContaining({
      taskType: "spreadsheet",
      spreadsheetIdentifierColumns: ["Name"],
      facultyExclusions: expect.objectContaining({
        studentIdColumn: "ID",
        exclusionStudentIdColumn: "Person Reference ID",
      }),
    }),
  });
});

it.each([
  ["Single document", "/student.pdf", "document", "documentPath"],
  ["Spreadsheet", "/students.tsv", "spreadsheet", "spreadsheetPath"],
  ["Directory of files", "/Student submissions", "directory", "directoryPath"],
])(
  "uses a dropped %s in the matching request",
  async (mode, path, taskType, pathKey) => {
    const user = userEvent.setup();
    render(<App />);
    await screen.findByText("Ready");
    await user.click(screen.getByRole("radio", { name: mode }));
    await waitFor(() => expect(native.listeners.size).toBe(3));
    await drop(path);
    await user.click(screen.getByRole("button", { name: "Run matching" }));
    expect(backend.submitMatchingRequest).toHaveBeenCalledWith(
      expect.objectContaining({
        payload: expect.objectContaining({ taskType, [pathKey]: path }),
      }),
    );
    expect(backend.previewFacultyDatasetReplacement).not.toHaveBeenCalled();
  },
);

it("maps identifiers automatically when a faculty roster is dropped", async () => {
  const user = userEvent.setup();
  render(<App />);
  await screen.findByText("Ready");
  await user.type(screen.getByLabelText("Prompt text"), "Protein structure");
  await user.click(
    screen.getByRole("radio", { name: "Provide a faculty list" }),
  );
  await waitFor(() => expect(native.listeners.size).toBe(3));
  await drop("/faculty roster.tsv");
  await user.click(screen.getByRole("button", { name: "Run matching" }));
  expect(backend.submitMatchingRequest).toHaveBeenCalledWith(
    expect.objectContaining({
      payload: expect.objectContaining({
        facultyScope: "custom",
        customFacultyPath: "/faculty roster.tsv",
        facultyRosterColumnMap: { Name: "Name" },
      }),
    }),
  );
});

it("opens the configuration dialog when a replacement dataset is dropped", async () => {
  render(<App />);
  await screen.findByText("Ready");
  await drop("/new faculty.tsv", 420);
  expect(await screen.findByText("Import faculty dataset")).toBeTruthy();
  expect(backend.previewFacultyDatasetReplacement).toHaveBeenCalledWith({
    path: "/new faculty.tsv",
  });
  expect(
    screen.getByRole<HTMLButtonElement>("button", { name: "Run matching" })
      .disabled,
  ).toBe(true);
});
