import { invoke } from "@tauri-apps/api/core";
import type {
  FacultyDatasetPreviewResult,
  FacultyDatasetStatus,
  FacultyRosterPreviewResult,
  SpreadsheetPreview,
  SubmissionPayload,
  SubmissionResponse,
  UploadKind,
} from "../types";

// Keep command names, argument shapes, and responses together at the desktop boundary.
export const backend = {
  getFacultyDatasetStatus: () =>
    invoke<FacultyDatasetStatus>("get_faculty_dataset_status"),
  ensureEmbeddingHelperReady: () =>
    invoke<boolean>("ensure_embedding_helper_ready"),
  updateFacultyEmbeddings: () => invoke<string>("update_faculty_embeddings"),
  previewFacultyDatasetReplacement: (args: { path: string }) =>
    invoke<FacultyDatasetPreviewResult>(
      "preview_faculty_dataset_replacement",
      args,
    ),
  replaceFacultyDataset: (args: {
    path: string;
    configuration: {
      embeddingColumns: number[];
      identifierColumns: number[];
      programColumns: number[];
    };
  }) => invoke<FacultyDatasetStatus>("replace_faculty_dataset", args),
  restoreDefaultFacultyDataset: () =>
    invoke<FacultyDatasetStatus>("restore_default_faculty_dataset"),
  previewFacultyRoster: (args: { path: string }) =>
    invoke<FacultyRosterPreviewResult>("preview_faculty_roster", args),
  analyzeSpreadsheet: (args: { path: string }) =>
    invoke<SpreadsheetPreview>("analyze_spreadsheet", args),
  submitMatchingRequest: (args: { payload: SubmissionPayload }) =>
    invoke<SubmissionResponse>("submit_matching_request", args),
  saveGeneratedSpreadsheet: (args: {
    path: string;
    content: string;
    encoding: string | null;
  }) => invoke<void>("save_generated_spreadsheet", args),
  validateUploadPath: (args: { path: string; kind: UploadKind }) =>
    invoke<string>("validate_upload_path", args),
};
