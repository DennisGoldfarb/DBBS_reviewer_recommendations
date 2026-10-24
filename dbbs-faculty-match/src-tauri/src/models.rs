use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::path::Path;

#[derive(Debug, Deserialize, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) enum TaskType {
    Prompt,
    Document,
    Spreadsheet,
    Directory,
}

#[derive(Debug, Deserialize, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) enum FacultyScope {
    All,
    Program,
    Custom,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SubmissionPayload {
    pub(crate) task_type: TaskType,
    #[serde(default)]
    pub(crate) prompt_text: Option<String>,
    #[serde(default)]
    pub(crate) document_path: Option<String>,
    #[serde(default)]
    pub(crate) spreadsheet_path: Option<String>,
    #[serde(default)]
    pub(crate) directory_path: Option<String>,
    pub(crate) faculty_scope: FacultyScope,
    #[serde(default)]
    pub(crate) program_filters: Vec<String>,
    #[serde(default)]
    pub(crate) custom_faculty_path: Option<String>,
    pub(crate) faculty_recs_per_student: u32,
    #[serde(default)]
    pub(crate) spreadsheet_prompt_columns: Vec<String>,
    #[serde(default)]
    pub(crate) spreadsheet_identifier_columns: Vec<String>,
    #[serde(default)]
    pub(crate) faculty_roster_column_map: HashMap<String, String>,
    #[serde(default)]
    pub(crate) faculty_roster_warnings: Vec<String>,
    #[serde(default)]
    pub(crate) faculty_exclusions: Option<FacultyExclusionConfig>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct FacultyExclusionConfig {
    pub(crate) path: String,
    pub(crate) exclusion_student_id_column: String,
    pub(crate) excluded_name_columns: Vec<String>,
    pub(crate) faculty_name_column: String,
    #[serde(default)]
    pub(crate) student_id_column: String,
    #[serde(default)]
    pub(crate) student_id: String,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct PathConfirmation {
    pub(crate) label: String,
    pub(crate) path: String,
}

impl PathConfirmation {
    pub(crate) fn new(label: &str, path: &Path) -> Self {
        Self {
            label: label.to_string(),
            path: path.to_string_lossy().into_owned(),
        }
    }
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SubmissionDetails {
    pub(crate) task_type: TaskType,
    pub(crate) faculty_scope: FacultyScope,
    pub(crate) validated_paths: Vec<PathConfirmation>,
    pub(crate) program_filters: Vec<String>,
    pub(crate) custom_faculty_path: Option<String>,
    pub(crate) recommendations_per_student: u32,
    pub(crate) prompt_preview: Option<String>,
    pub(crate) spreadsheet_prompt_columns: Vec<String>,
    pub(crate) spreadsheet_identifier_columns: Vec<String>,
    pub(crate) faculty_roster_column_map: HashMap<String, String>,
    pub(crate) faculty_roster_warnings: Vec<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SubmissionResponse {
    pub(crate) summary: String,
    pub(crate) warnings: Vec<String>,
    pub(crate) details: SubmissionDetails,
    pub(crate) prompt_matches: Vec<PromptMatchResult>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) directory_results: Option<DirectoryMatchResults>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) spreadsheet_results: Option<SpreadsheetMatchResults>,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SpreadsheetPreview {
    pub(crate) headers: Vec<String>,
    pub(crate) rows: Vec<Vec<String>>,
    pub(crate) suggested_prompt_columns: Vec<usize>,
    pub(crate) suggested_identifier_columns: Vec<usize>,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct FacultyDatasetPreviewResponse {
    pub(crate) preview: SpreadsheetPreview,
    pub(crate) suggested_embedding_columns: Vec<usize>,
    pub(crate) suggested_identifier_columns: Vec<usize>,
    pub(crate) suggested_program_columns: Vec<usize>,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct FacultyRosterPreviewResponse {
    pub(crate) preview: SpreadsheetPreview,
    pub(crate) suggested_identifier_matches: HashMap<String, Option<usize>>,
    pub(crate) warnings: Vec<String>,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct FacultyDatasetStatus {
    pub(crate) path: Option<String>,
    pub(crate) canonical_path: Option<String>,
    pub(crate) source_path: Option<String>,
    pub(crate) last_modified: Option<String>,
    pub(crate) row_count: Option<usize>,
    pub(crate) column_count: Option<usize>,
    pub(crate) is_valid: bool,
    pub(crate) is_default: bool,
    pub(crate) message: Option<String>,
    pub(crate) message_variant: Option<String>,
    pub(crate) preview: Option<SpreadsheetPreview>,
    pub(crate) analysis: Option<FacultyDatasetAnalysis>,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct FacultyDatasetAnalysis {
    pub(crate) embedding_columns: Vec<String>,
    pub(crate) identifier_columns: Vec<String>,
    pub(crate) program_columns: Vec<String>,
    pub(crate) available_programs: Vec<String>,
}

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct FacultyDatasetColumnConfiguration {
    #[serde(default)]
    pub(crate) embedding_columns: Vec<usize>,
    #[serde(default)]
    pub(crate) identifier_columns: Vec<usize>,
    #[serde(default)]
    pub(crate) program_columns: Vec<usize>,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct FacultyDatasetMetadata {
    pub(crate) analysis: FacultyDatasetAnalysis,
    pub(crate) memberships: Vec<FacultyProgramMembership>,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct FacultyProgramMembership {
    pub(crate) row_index: usize,
    pub(crate) identifiers: HashMap<String, String>,
    pub(crate) programs: Vec<String>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct EmbeddingRequestPayload {
    pub(crate) model: String,
    pub(crate) texts: Vec<EmbeddingRequestRow>,
    #[serde(default)]
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) item_label: Option<String>,
    #[serde(default)]
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) item_label_plural: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) faculty_preparation: Option<FacultyPreparationOptions>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct FacultyPreparationOptions {}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct EmbeddingRequestRow {
    pub(crate) id: usize,
    pub(crate) text: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct EmbeddingResponsePayload {
    pub(crate) model: String,
    pub(crate) dimension: usize,
    pub(crate) rows: Vec<EmbeddingResponseRow>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct EmbeddingResponseRow {
    pub(crate) id: usize,
    pub(crate) embedding: Vec<f32>,
    #[serde(default)]
    pub(crate) additional_embeddings: Vec<Vec<f32>>,
    #[serde(default)]
    pub(crate) preparation: Option<FacultyTextPreparation>,
}

impl EmbeddingResponseRow {
    pub(crate) fn into_embeddings(self) -> Vec<Vec<f32>> {
        let mut embeddings = Vec::with_capacity(1 + self.additional_embeddings.len());
        embeddings.push(self.embedding);
        embeddings.extend(self.additional_embeddings);
        embeddings
    }
}

#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct FacultyTextPreparation {
    #[serde(default)]
    pub(crate) truncated: bool,
    pub(crate) text: String,
    pub(crate) original_tokens: usize,
    pub(crate) tokens: usize,
    pub(crate) model: String,
    pub(crate) revision: String,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct FacultyEmbeddingEntry {
    pub(crate) row_index: usize,
    pub(crate) identifiers: HashMap<String, String>,
    pub(crate) embedding: Vec<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub(crate) preparation: Option<FacultyTextPreparation>,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct FacultyEmbeddingIndex {
    pub(crate) model: String,
    #[serde(default)]
    pub(crate) generated_at: Option<String>,
    pub(crate) dimension: usize,
    #[serde(default)]
    pub(crate) total_rows: Option<usize>,
    #[serde(default)]
    pub(crate) embedded_rows: Option<usize>,
    #[serde(default)]
    pub(crate) skipped_rows: Option<usize>,
    #[serde(default)]
    pub(crate) embedding_columns: Vec<String>,
    #[serde(default)]
    pub(crate) identifier_columns: Vec<String>,
    pub(crate) entries: Vec<FacultyEmbeddingEntry>,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct PromptMatchResult {
    pub(crate) prompt: String,
    pub(crate) faculty_matches: Vec<FacultyMatchResult>,
}

pub(crate) struct DocumentExtractionResult {
    pub(crate) text: String,
    pub(crate) warnings: Vec<String>,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct FacultyMatchResult {
    pub(crate) row_index: usize,
    pub(crate) similarity: f32,
    pub(crate) identifiers: HashMap<String, String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) faculty_text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) student_rank_for_faculty: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) student_rank_total: Option<usize>,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct GeneratedSpreadsheet {
    pub(crate) filename: String,
    pub(crate) mime_type: String,
    pub(crate) content: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) encoding: Option<String>,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct DirectoryMatchResults {
    pub(crate) processed_documents: usize,
    pub(crate) matched_documents: usize,
    pub(crate) skipped_documents: usize,
    pub(crate) total_rows: usize,
    pub(crate) preview: SpreadsheetPreview,
    pub(crate) spreadsheet: GeneratedSpreadsheet,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SpreadsheetMatchResults {
    pub(crate) processed_rows: usize,
    pub(crate) matched_rows: usize,
    pub(crate) skipped_rows: usize,
    pub(crate) total_rows: usize,
    pub(crate) preview: SpreadsheetPreview,
    pub(crate) spreadsheet: GeneratedSpreadsheet,
}

#[derive(Debug, Clone)]
pub(crate) struct MatchEntry {
    pub(crate) student_values: Vec<String>,
    pub(crate) faculty_values: Vec<String>,
    pub(crate) similarity: Option<f32>,
    pub(crate) student_rank: Option<(usize, Option<usize>)>,
    pub(crate) faculty_rank: Option<usize>,
}

#[derive(Debug)]
pub(crate) struct DirectoryProcessingOutcome {
    pub(crate) warnings: Vec<String>,
    pub(crate) prompt_matches: Vec<PromptMatchResult>,
    pub(crate) results: DirectoryMatchResults,
}

#[derive(Debug)]
pub(crate) struct SpreadsheetProcessingOutcome {
    pub(crate) warnings: Vec<String>,
    pub(crate) prompt_matches: Vec<PromptMatchResult>,
    pub(crate) results: SpreadsheetMatchResults,
}
