import { act, cleanup, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { clearMocks, mockIPC, mockWindows } from "@tauri-apps/api/mocks";
import { FileDropZone } from "../components/FileDropZone";

interface NativeListener {
  event: string;
  target: { kind: string; label: string };
  callback: number;
}

const listeners = new Map<number, NativeListener>();
const validation = vi.fn();
let nextId = 0;

beforeEach(() => {
  listeners.clear();
  nextId = 0;
  vi.stubGlobal("isTauri", true);
  vi.stubGlobal("devicePixelRatio", 2);
  vi.spyOn(navigator, "userAgent", "get").mockReturnValue(
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 15_7)",
  );
  mockWindows("main");
  mockIPC((command, args) => {
    const input = args as Record<string, unknown>;
    if (command === "plugin:event|listen") {
      const id = ++nextId;
      listeners.set(id, {
        event: input.event as string,
        target: input.target as NativeListener["target"],
        callback: input.handler as number,
      });
      return id;
    }
    if (command === "plugin:event|unlisten") {
      listeners.delete(input.eventId as number);
      return;
    }
    if (command === "validate_upload_path") {
      validation(input);
      return input.path;
    }
    throw new Error(`Unexpected desktop command: ${command}`);
  });
  vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue({
    left: 100,
    right: 800,
    top: 300,
    bottom: 500,
  } as DOMRect);
});

afterEach(async () => {
  cleanup();
  await waitFor(() => expect(listeners.size).toBe(0));
  clearMocks();
});

async function nativeEvent(event: string, payload: unknown) {
  const host = window as unknown as {
    __TAURI_INTERNALS__: { runCallback: (id: number, value: unknown) => void };
  };
  await act(async () => {
    for (const [id, listener] of listeners) {
      if (
        listener.event === event &&
        listener.target.kind === "Window" &&
        listener.target.label === "main"
      ) {
        host.__TAURI_INTERNALS__.runCallback(listener.callback, {
          event,
          id,
          payload,
        });
      }
    }
  });
}

describe("real Tauri window event adapter", () => {
  it.each([
    ["document", "/Student research.pdf"],
    ["directory", "/Student submissions"],
    ["spreadsheet", "/students.tsv"],
    ["dataset", "/faculty.tsv"],
  ] as const)(
    "routes a macOS Retina %s drop without scaling logical points twice",
    async (kind, path) => {
      const onPath = vi.fn();
      const { unmount } = render(
        <FileDropZone kind={kind} label="Upload" onPath={onPath} />,
      );
      await waitFor(() => expect(listeners.size).toBe(4));
      expect(
        [...listeners.values()].every(
          (listener) => listener.target.kind === "Window",
        ),
      ).toBe(true);
      await nativeEvent("tauri://drag-enter", {
        paths: [path],
        position: { x: 400, y: 400 },
      });
      expect(screen.getByRole("group").className).toContain(
        "file-drop-zone-active",
      );
      await nativeEvent("tauri://drag-drop", {
        paths: [path],
        position: { x: 400, y: 400 },
      });
      await waitFor(() => expect(onPath).toHaveBeenCalledWith(path));
      expect(validation).toHaveBeenCalledWith({ path, kind });
      unmount();
      await waitFor(() => expect(listeners.size).toBe(0));
    },
  );

  it("converts the actual Windows physical-pixel payload before routing", async () => {
    vi.spyOn(navigator, "userAgent", "get").mockReturnValue(
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    );
    const onPath = vi.fn();
    render(<FileDropZone kind="document" label="Upload" onPath={onPath} />);
    await waitFor(() => expect(listeners.size).toBe(4));
    await nativeEvent("tauri://drag-drop", {
      paths: ["C:\\Research.pdf"],
      position: { x: 800, y: 800 },
    });
    await waitFor(() =>
      expect(onPath).toHaveBeenCalledWith("C:\\Research.pdf"),
    );
  });
});
