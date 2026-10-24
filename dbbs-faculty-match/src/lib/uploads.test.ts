import { describe, expect, it } from "vitest";
import {
  isInsideDropZone,
  singleDroppedPath,
  validateUploadExtension,
} from "./uploads";

describe("dropped paths", () => {
  it("preserves spaces, Unicode, and Windows paths", () => {
    for (const path of [
      "/tmp/René Student/research.pdf",
      "C:\\Student documents\\research.PDF",
    ]) {
      expect(singleDroppedPath([path])).toBe(path);
      expect(() => validateUploadExtension(path, "document")).not.toThrow();
    }
  });
  it("rejects empty and multiple selections instead of silently ignoring files", () => {
    expect(() => singleDroppedPath([])).toThrow("one file or folder");
    expect(() => singleDroppedPath(["a", "b"])).toThrow("one file or folder");
    expect(() => singleDroppedPath([" "])).toThrow("no readable path");
  });
  it("accepts supported formats and extensionless folders", () => {
    for (const path of [
      "student.txt",
      "student.pdf",
      "student.docx",
      "student.rtf",
    ])
      expect(() => validateUploadExtension(path, "document")).not.toThrow();
    for (const path of ["students.csv", "students.TSV", "students.xlsx"])
      expect(() => validateUploadExtension(path, "spreadsheet")).not.toThrow();
    expect(() =>
      validateUploadExtension("/students", "directory"),
    ).not.toThrow();
    expect(() => validateUploadExtension("students.exe", "document")).toThrow(
      "Choose a",
    );
    expect(() => validateUploadExtension("faculty.csv", "dataset")).toThrow(
      "Choose a",
    );
  });
});

describe("native drop hit testing", () => {
  const rect = { left: 50, top: 100, right: 200, bottom: 300 };
  it("converts Windows physical pixels to viewport coordinates", () => {
    expect(isInsideDropZone({ x: 300, y: 240 }, rect, 2, "Windows NT")).toBe(
      true,
    );
    expect(isInsideDropZone({ x: 300, y: 240 }, rect, 1, "Windows NT")).toBe(
      false,
    );
    expect(isInsideDropZone({ x: 40, y: 120 }, rect, 1, "Windows NT")).toBe(
      false,
    );
  });
  it("avoids targeting adjacent zones at their shared edge", () => {
    expect(isInsideDropZone({ x: 200, y: 120 }, rect, 1, "Windows NT")).toBe(
      false,
    );
    expect(isInsideDropZone({ x: 60, y: 300 }, rect, 1, "Windows NT")).toBe(
      false,
    );
    expect(
      isInsideDropZone({ x: 60, y: 120 }, rect, Number.NaN, "Windows NT"),
    ).toBe(true);
  });
  it.each(["Macintosh; Intel Mac OS X", "X11; Linux x86_64"])(
    "keeps %s logical coordinates on a scaled display",
    (userAgent) => {
      const area = { left: 100, right: 800, top: 300, bottom: 500 };
      expect(isInsideDropZone({ x: 400, y: 400 }, area, 2, userAgent)).toBe(
        true,
      );
      expect(isInsideDropZone({ x: 400, y: 200 }, area, 2, userAgent)).toBe(
        false,
      );
    },
  );
});
