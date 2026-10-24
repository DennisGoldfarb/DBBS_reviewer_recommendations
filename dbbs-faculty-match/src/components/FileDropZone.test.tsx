import { StrictMode } from "react";
import { act, render, screen, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { DragDropEvent } from "@tauri-apps/api/webview";
import { FileDropZone } from "./FileDropZone";
import { backend } from "../services/backend";
import { deferred } from "../test/fixtures";

const native = vi.hoisted(() => ({
  listeners: new Set<
    (event: { payload: DragDropEvent }) => void | Promise<void>
  >(),
  stops: [] as ReturnType<typeof vi.fn>[],
}));
vi.mock("@tauri-apps/api/core", () => ({
  isTauri: () => true,
  invoke: vi.fn(),
}));
vi.mock("@tauri-apps/api/window", () => ({
  getCurrentWindow: () => ({
    setTheme: vi.fn(async () => {}),
    onDragDropEvent: vi.fn(async (handler) => {
      native.listeners.add(handler);
      const stop = vi.fn(() => native.listeners.delete(handler));
      native.stops.push(stop);
      return stop;
    }),
  }),
}));

async function emit(
  type: "enter" | "drop" | "leave",
  paths = ["/student.pdf"],
  x = 20,
  y = 20,
) {
  await act(async () => {
    const payload =
      type === "leave" ? { type } : { type, paths, position: { x, y } };
    await Promise.all(
      [...native.listeners].map((handler) =>
        handler({ payload: payload as DragDropEvent }),
      ),
    );
  });
}

beforeEach(() => {
  native.listeners.clear();
  native.stops.length = 0;
  vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue({
    left: 0,
    top: 0,
    right: 100,
    bottom: 100,
  } as DOMRect);
  vi.spyOn(backend, "validateUploadPath").mockImplementation(
    async ({ path }) => path,
  );
});

describe("native file drop zones", () => {
  it("highlights a hovered zone, validates a real path, and clears on drop", async () => {
    const onPath = vi.fn();
    render(<FileDropZone kind="document" label="Upload" onPath={onPath} />);
    await waitFor(() => expect(native.listeners.size).toBe(1));
    await emit("enter");
    expect(screen.getByRole("group").className).toContain(
      "file-drop-zone-active",
    );
    await emit("drop");
    expect(backend.validateUploadPath).toHaveBeenCalledWith({
      path: "/student.pdf",
      kind: "document",
    });
    expect(onPath).toHaveBeenCalledWith("/student.pdf");
    expect(screen.getByRole("group").className).not.toContain(
      "file-drop-zone-active",
    );
  });
  it("routes drops by position so unrelated zones do not consume them", async () => {
    const onPath = vi.fn();
    render(<FileDropZone kind="document" label="Upload" onPath={onPath} />);
    await emit("drop", undefined, 200, 200);
    expect(onPath).not.toHaveBeenCalled();
    expect(backend.validateUploadPath).not.toHaveBeenCalled();
  });
  it("reports unsupported files, multiple items, and filesystem validation errors", async () => {
    const onPath = vi.fn();
    render(<FileDropZone kind="document" label="Upload" onPath={onPath} />);
    await emit("drop", ["/student.exe"]);
    expect(screen.getByRole("alert").textContent).toContain("Choose a");
    await emit("drop", ["/a.pdf", "/b.pdf"]);
    expect(screen.getByRole("alert").textContent).toContain(
      "one file or folder",
    );
    vi.mocked(backend.validateUploadPath).mockRejectedValueOnce(
      "Item was not found",
    );
    await emit("drop");
    expect(screen.getByRole("alert").textContent).toContain("not found");
    expect(onPath).not.toHaveBeenCalled();
  });
  it("accepts folders and ignores disabled controls", async () => {
    const onPath = vi.fn();
    const { rerender } = render(
      <FileDropZone kind="directory" label="Upload" onPath={onPath} />,
    );
    await emit("drop", ["/students"]);
    expect(onPath).toHaveBeenCalledWith("/students");
    rerender(
      <FileDropZone kind="directory" label="Upload" onPath={onPath} disabled />,
    );
    onPath.mockClear();
    await emit("drop", ["/other students"]);
    expect(onPath).not.toHaveBeenCalled();
  });
  it("does not deliver validation results after unmount", async () => {
    const pending = deferred<string>();
    vi.mocked(backend.validateUploadPath).mockReturnValueOnce(pending.promise);
    const onPath = vi.fn();
    const { unmount } = render(
      <FileDropZone kind="document" label="Upload" onPath={onPath} />,
    );
    let emission!: Promise<void>;
    act(() => {
      emission = emit("drop");
    });
    unmount();
    await act(async () => {
      pending.resolve("/student.pdf");
      await emission;
    });
    expect(onPath).not.toHaveBeenCalled();
    expect(native.listeners.size).toBe(0);
  });
  it("cleans up asynchronously registered listeners under React StrictMode", async () => {
    const { unmount } = render(
      <StrictMode>
        <FileDropZone kind="document" label="Upload" onPath={vi.fn()} />
      </StrictMode>,
    );
    await waitFor(() => expect(native.listeners.size).toBe(1));
    unmount();
    expect(native.listeners.size).toBe(0);
    expect(native.stops.every((stop) => stop.mock.calls.length === 1)).toBe(
      true,
    );
  });
});
