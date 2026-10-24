import type { UploadKind } from "../types";

export const UPLOAD_EXTENSIONS = {
  document: ["txt", "pdf", "doc", "docx", "rtf"],
  spreadsheet: ["csv", "tsv", "txt", "xlsx", "xls"],
  dataset: ["tsv", "txt", "xlsx", "xls"],
} as const;

export function singleDroppedPath(paths: string[]): string {
  if (paths.length !== 1) throw new Error("Drop one file or folder at a time.");
  if (!paths[0]?.trim())
    throw new Error("The dropped item has no readable path.");
  return paths[0];
}

export function validateUploadExtension(path: string, kind: UploadKind): void {
  if (kind === "directory") return;
  const filename = path.split(/[\\/]/).pop() ?? "";
  const extension = filename.split(".").pop()?.toLowerCase() ?? "";
  if (!(UPLOAD_EXTENSIONS[kind] as readonly string[]).includes(extension)) {
    throw new Error(
      `Choose a ${UPLOAD_EXTENSIONS[kind].map((value) => `.${value}`).join(", ")} file.`,
    );
  }
}

export function isInsideDropZone(
  point: { x: number; y: number },
  rect: Pick<DOMRect, "left" | "right" | "top" | "bottom">,
  scaleFactor: number,
  userAgent = navigator.userAgent,
): boolean {
  // Wry 0.57's WKWebView and GTK backends already report logical coordinates.
  // WebView2 reports physical pixels. Tauri wraps both as PhysicalPosition,
  // so applying devicePixelRatio on macOS/GTK scales the drop a second time.
  const scale =
    /Windows/i.test(userAgent) &&
    Number.isFinite(scaleFactor) &&
    scaleFactor > 0
      ? scaleFactor
      : 1;
  const x = point.x / scale;
  const y = point.y / scale;
  return x >= rect.left && x < rect.right && y >= rect.top && y < rect.bottom;
}
