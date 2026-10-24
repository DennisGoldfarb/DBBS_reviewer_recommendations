export type TaskType = "prompt" | "document" | "spreadsheet" | "directory";
export type UploadKind = "document" | "spreadsheet" | "directory" | "dataset";

export interface SubmissionPayload {
  taskType: TaskType;
  facultyScope: FacultyScope;
  promptText?: string;
  documentPath?: string;
  spreadsheetPath?: string;
  directoryPath?: string;
  programFilters?: string[];
  customFacultyPath?: string;
  facultyRecsPerStudent: number;
  spreadsheetPromptColumns?: string[];
  spreadsheetIdentifierColumns?: string[];
  facultyRosterColumnMap?: Record<string, string>;
  facultyRosterWarnings?: string[];
  facultyExclusions?: FacultyExclusionConfig;
}
export interface FacultyExclusionConfig {
  path: string;
  exclusionStudentIdColumn: string;
  excludedNameColumns: string[];
  facultyNameColumn: string;
  studentIdColumn: string;
  studentId: string;
}
export type FacultyScope = "all" | "program" | "custom";
export type ThemePreference = "light" | "dark";

export type ProgramName = string;

export interface PathConfirmation {
  label: string;
  path: string;
}

export interface SpreadsheetPreview {
  headers: string[];
  rows: string[][];
  suggestedPromptColumns: number[];
  suggestedIdentifierColumns: number[];
}

export interface GeneratedSpreadsheet {
  filename: string;
  mimeType: string;
  content: string;
  encoding?: "utf8" | "base64";
}

export interface DirectoryMatchResults {
  processedDocuments: number;
  matchedDocuments: number;
  skippedDocuments: number;
  totalRows: number;
  preview: SpreadsheetPreview;
  spreadsheet: GeneratedSpreadsheet;
}

export interface SpreadsheetMatchResults {
  processedRows: number;
  matchedRows: number;
  skippedRows: number;
  totalRows: number;
  preview: SpreadsheetPreview;
  spreadsheet: GeneratedSpreadsheet;
}

export interface FacultyDatasetAnalysis {
  embeddingColumns: string[];
  identifierColumns: string[];
  programColumns: string[];
  availablePrograms: string[];
}

export interface FacultyDatasetPreviewResult {
  preview: SpreadsheetPreview;
  suggestedEmbeddingColumns: number[];
  suggestedIdentifierColumns: number[];
  suggestedProgramColumns: number[];
}

export interface FacultyRosterPreviewResult {
  preview: SpreadsheetPreview;
  suggestedIdentifierMatches: Record<string, number | null>;
  warnings: string[];
}

export interface FacultyDatasetStatus {
  path: string | null;
  canonicalPath: string | null;
  sourcePath: string | null;
  lastModified: string | null;
  rowCount: number | null;
  columnCount: number | null;
  isValid: boolean;
  isDefault: boolean;
  message: string | null;
  messageVariant: StatusMessage["variant"] | null;
  preview: SpreadsheetPreview | null;
  analysis: FacultyDatasetAnalysis | null;
}

export interface SubmissionDetails {
  taskType: TaskType;
  facultyScope: FacultyScope;
  validatedPaths: PathConfirmation[];
  programFilters: string[];
  customFacultyPath: string | null;
  recommendationsPerStudent: number;
  promptPreview?: string;
  spreadsheetPromptColumns: string[];
  spreadsheetIdentifierColumns: string[];
  facultyRosterColumnMap: Record<string, string>;
  facultyRosterWarnings: string[];
}

export interface FacultyMatchResult {
  rowIndex: number;
  similarity: number;
  identifiers: Record<string, string>;
  facultyText?: string;
  studentRankForFaculty?: number;
  studentRankTotal?: number;
}

export interface PromptMatchResult {
  prompt: string;
  facultyMatches: FacultyMatchResult[];
}

export interface SubmissionResponse {
  summary: string;
  warnings: string[];
  details: SubmissionDetails;
  promptMatches: PromptMatchResult[];
  directoryResults?: DirectoryMatchResults;
  spreadsheetResults?: SpreadsheetMatchResults;
}

export interface StatusMessage {
  variant: "success" | "error" | "info";
  message: string;
}

export interface EmbeddingProgressPayload {
  phase: string;
  message?: string | null;
  processedRows: number;
  totalRows: number;
  elapsedSeconds?: number | null;
  estimatedRemainingSeconds?: number | null;
}
