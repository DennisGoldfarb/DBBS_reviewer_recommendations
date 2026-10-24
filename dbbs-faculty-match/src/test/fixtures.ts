import type {
  FacultyDatasetStatus,
  SpreadsheetPreview,
  SubmissionResponse,
} from "../types";

export const preview: SpreadsheetPreview = {
  headers: ["Name", "Research interests"],
  rows: [["Example student", "Protein structure"]],
  suggestedIdentifierColumns: [0],
  suggestedPromptColumns: [1],
};

export const datasetStatus: FacultyDatasetStatus = {
  path: "/faculty.tsv",
  canonicalPath: "/faculty.tsv",
  sourcePath: null,
  lastModified: "2025-01-01T00:00:00Z",
  rowCount: 1,
  columnCount: 2,
  isValid: true,
  isDefault: true,
  message: null,
  messageVariant: null,
  preview,
  analysis: {
    embeddingColumns: ["Research interests"],
    identifierColumns: ["Name"],
    programColumns: ["Program"],
    availablePrograms: ["Biochemistry"],
  },
};

export const matchResponse: SubmissionResponse = {
  summary: "Matching completed.",
  warnings: [],
  details: {
    taskType: "prompt",
    facultyScope: "all",
    validatedPaths: [],
    programFilters: [],
    customFacultyPath: null,
    recommendationsPerStudent: 10,
    spreadsheetPromptColumns: [],
    spreadsheetIdentifierColumns: [],
    facultyRosterColumnMap: {},
    facultyRosterWarnings: [],
  },
  promptMatches: [
    {
      prompt: "Protein structure",
      facultyMatches: [
        {
          rowIndex: 0,
          similarity: 0.9,
          identifiers: { Name: "Example faculty" },
        },
      ],
    },
  ],
};

export function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason: unknown) => void;
  const promise = new Promise<T>((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
}
