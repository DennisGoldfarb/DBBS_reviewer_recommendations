use base64::{engine::general_purpose::STANDARD as Base64Engine, Engine as _};
use calamine::{open_workbook_auto, DataType, Reader};
use chrono::{DateTime, Utc};
use docx_rs::{
    read_docx, DocumentChild, Insert, InsertChild, Paragraph, ParagraphChild, Run, RunChild,
    StructuredDataTag, StructuredDataTagChild, Table, TableCellContent, TableChild, TableRowChild,
};
use pdf_extract::extract_text_from_mem;
use rtf_parser::RtfDocument;
use rust_xlsxwriter::{Format, Workbook};
use serde::{Deserialize, Serialize};
use std::char;
use std::cmp::Ordering;
use std::collections::{BTreeSet, HashMap, HashSet};
use std::convert::TryFrom;
use std::fs;
use std::fs::File;
use std::io::{BufRead, BufReader, BufWriter, Cursor, Write};
use std::path::{Path, PathBuf};
use std::process::{Child, Command, Stdio};
use std::sync::mpsc::{self, Receiver};
use std::sync::{Arc, Mutex};
use std::time::{Instant, SystemTime};
use tauri::{Emitter, Manager};

const FACULTY_DATASET_BASENAME: &str = "faculty_dataset";
const FACULTY_DATASET_DEFAULT_EXTENSION: &str = "tsv";
const FACULTY_DATASET_EXTENSIONS: &[&str] = &["tsv", "txt", "xlsx", "xls"];
const DEFAULT_FACULTY_DATASET: &[u8] = include_bytes!("../assets/default_faculty_dataset.tsv");
const FACULTY_DATASET_METADATA_NAME: &str = "faculty_dataset_metadata.json";
const FACULTY_DATASET_SOURCE_NAME: &str = "faculty_dataset_source.txt";
const FACULTY_EMBEDDINGS_NAME: &str = "faculty_embeddings.json";
const DEFAULT_FACULTY_EMBEDDINGS: &[u8] =
    include_bytes!("../assets/default_faculty_embeddings.json");
const DEFAULT_EMBEDDING_MODEL: &str = "NeuML/pubmedbert-base-embeddings";
const FACULTY_EMBEDDING_PROGRESS_EVENT: &str = "faculty-embedding-progress";

#[derive(Debug, Deserialize, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
enum TaskType {
    Prompt,
    Document,
    Spreadsheet,
    Directory,
}

#[derive(Debug, Deserialize, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
enum FacultyScope {
    All,
    Program,
    Custom,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct SubmissionPayload {
    task_type: TaskType,
    #[serde(default)]
    prompt_text: Option<String>,
    #[serde(default)]
    document_path: Option<String>,
    #[serde(default)]
    spreadsheet_path: Option<String>,
    #[serde(default)]
    directory_path: Option<String>,
    faculty_scope: FacultyScope,
    #[serde(default)]
    program_filters: Vec<String>,
    #[serde(default)]
    custom_faculty_path: Option<String>,
    faculty_recs_per_student: u32,
    #[serde(default)]
    spreadsheet_prompt_columns: Vec<String>,
    #[serde(default)]
    spreadsheet_identifier_columns: Vec<String>,
    #[serde(default)]
    faculty_roster_column_map: HashMap<String, String>,
    #[serde(default)]
    faculty_roster_warnings: Vec<String>,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct PathConfirmation {
    label: String,
    path: String,
}

impl PathConfirmation {
    fn new(label: &str, path: &Path) -> Self {
        Self {
            label: label.to_string(),
            path: path.to_string_lossy().into_owned(),
        }
    }
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct SubmissionDetails {
    task_type: TaskType,
    faculty_scope: FacultyScope,
    validated_paths: Vec<PathConfirmation>,
    program_filters: Vec<String>,
    custom_faculty_path: Option<String>,
    recommendations_per_student: u32,
    prompt_preview: Option<String>,
    spreadsheet_prompt_columns: Vec<String>,
    spreadsheet_identifier_columns: Vec<String>,
    faculty_roster_column_map: HashMap<String, String>,
    faculty_roster_warnings: Vec<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct SubmissionResponse {
    summary: String,
    warnings: Vec<String>,
    details: SubmissionDetails,
    prompt_matches: Vec<PromptMatchResult>,
    #[serde(skip_serializing_if = "Option::is_none")]
    directory_results: Option<DirectoryMatchResults>,
    #[serde(skip_serializing_if = "Option::is_none")]
    spreadsheet_results: Option<SpreadsheetMatchResults>,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct SpreadsheetPreview {
    headers: Vec<String>,
    rows: Vec<Vec<String>>,
    suggested_prompt_columns: Vec<usize>,
    suggested_identifier_columns: Vec<usize>,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct FacultyDatasetPreviewResponse {
    preview: SpreadsheetPreview,
    suggested_embedding_columns: Vec<usize>,
    suggested_identifier_columns: Vec<usize>,
    suggested_program_columns: Vec<usize>,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct FacultyRosterPreviewResponse {
    preview: SpreadsheetPreview,
    suggested_identifier_matches: HashMap<String, Option<usize>>,
    warnings: Vec<String>,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct FacultyDatasetStatus {
    path: Option<String>,
    canonical_path: Option<String>,
    source_path: Option<String>,
    last_modified: Option<String>,
    row_count: Option<usize>,
    column_count: Option<usize>,
    is_valid: bool,
    is_default: bool,
    message: Option<String>,
    message_variant: Option<String>,
    preview: Option<SpreadsheetPreview>,
    analysis: Option<FacultyDatasetAnalysis>,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct FacultyDatasetAnalysis {
    embedding_columns: Vec<String>,
    identifier_columns: Vec<String>,
    program_columns: Vec<String>,
    available_programs: Vec<String>,
}

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct FacultyDatasetColumnConfiguration {
    #[serde(default)]
    embedding_columns: Vec<usize>,
    #[serde(default)]
    identifier_columns: Vec<usize>,
    #[serde(default)]
    program_columns: Vec<usize>,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct FacultyDatasetMetadata {
    analysis: FacultyDatasetAnalysis,
    memberships: Vec<FacultyProgramMembership>,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct FacultyProgramMembership {
    row_index: usize,
    identifiers: HashMap<String, String>,
    programs: Vec<String>,
}

#[tauri::command]
async fn submit_matching_request(
    app_handle: tauri::AppHandle,
    payload: SubmissionPayload,
) -> Result<SubmissionResponse, String> {
    tauri::async_runtime::spawn_blocking(move || perform_matching_request(app_handle, payload))
        .await
        .map_err(|err| format!("Matching task failed: {err}"))?
}

fn perform_matching_request(
    app_handle: tauri::AppHandle,
    payload: SubmissionPayload,
) -> Result<SubmissionResponse, String> {
    let SubmissionPayload {
        task_type,
        prompt_text,
        document_path,
        spreadsheet_path,
        directory_path,
        faculty_scope,
        program_filters,
        custom_faculty_path,
        faculty_recs_per_student,
        spreadsheet_prompt_columns,
        spreadsheet_identifier_columns,
        faculty_roster_column_map,
        faculty_roster_warnings,
    } = payload;

    if faculty_recs_per_student == 0 {
        return Err("Specify at least one faculty recommendation per student.".into());
    }

    let mut warnings = faculty_roster_warnings.clone();
    let mut validated_paths = Vec::new();
    let mut prompt_preview = None;
    let mut selected_prompt_columns = Vec::new();
    let mut selected_identifier_columns = Vec::new();
    let mut detail_identifier_columns = Vec::new();
    let mut prepared_prompt_text: Option<String> = None;
    let mut directory_source: Option<PathBuf> = None;
    let mut spreadsheet_source: Option<PathBuf> = None;
    let mut detail_roster_column_map: HashMap<String, String> = HashMap::new();
    let mut roster_warning_messages = faculty_roster_warnings;

    match task_type {
        TaskType::Prompt => {
            let text = prompt_text.as_deref().map(str::trim).unwrap_or_default();
            if text.is_empty() {
                return Err("Provide a prompt describing the student's interests.".into());
            }
            prompt_preview = Some(build_prompt_preview(text));
            prepared_prompt_text = Some(text.to_string());
        }
        TaskType::Document => {
            let document = resolve_existing_path(document_path, false, "Single document")?;
            if let Some(message) =
                validate_extension(&document, &["txt", "pdf", "doc", "docx"], "document")
            {
                warnings.push(message);
            }
            let extraction = extract_document_prompt(&document)?;
            if extraction.text.trim().is_empty() {
                return Err(
                    "The selected document did not contain any readable text to embed.".into(),
                );
            }
            warnings.extend(extraction.warnings);
            prompt_preview = Some(build_prompt_preview(&extraction.text));
            prepared_prompt_text = Some(extraction.text);
            validated_paths.push(PathConfirmation::new("Document", &document));
        }
        TaskType::Spreadsheet => {
            let spreadsheet = resolve_existing_path(spreadsheet_path, false, "Spreadsheet file")?;
            if let Some(message) =
                validate_extension(&spreadsheet, &["tsv", "txt", "xlsx", "xls"], "spreadsheet")
            {
                warnings.push(message);
            }
            validated_paths.push(PathConfirmation::new("Spreadsheet", &spreadsheet));
            spreadsheet_source = Some(spreadsheet.clone());

            selected_prompt_columns = normalize_columns(spreadsheet_prompt_columns);
            selected_identifier_columns = normalize_columns(spreadsheet_identifier_columns);
            detail_identifier_columns = if selected_identifier_columns.is_empty() {
                vec!["Row number".into()]
            } else {
                selected_identifier_columns.clone()
            };

            if selected_prompt_columns.is_empty() {
                return Err("Select at least one column containing student prompts.".into());
            }
        }
        TaskType::Directory => {
            let directory = resolve_existing_path(directory_path, true, "Directory")?;
            if let Ok(mut entries) = fs::read_dir(&directory) {
                if entries.next().is_none() {
                    warnings.push("The selected directory appears to be empty.".into());
                }
            }
            validated_paths.push(PathConfirmation::new("Directory", &directory));
            directory_source = Some(directory);
        }
    }

    let normalized_programs = normalize_programs(program_filters);
    let mut allowed_faculty_rows: Option<HashSet<usize>> = None;
    let mut faculty_roster_path = None;

    if matches!(faculty_scope, FacultyScope::Custom) {
        let roster = resolve_existing_path(custom_faculty_path, false, "Faculty list")?;
        if let Some(message) =
            validate_extension(&roster, &["tsv", "txt", "xlsx", "xls"], "faculty list")
        {
            warnings.push(message);
        }
        faculty_roster_path = Some(roster.to_string_lossy().into_owned());
        validated_paths.push(PathConfirmation::new("Faculty list", &roster));

        let metadata = load_faculty_dataset_metadata(&app_handle)?.ok_or_else(|| {
            "The faculty dataset metadata is unavailable. Refresh the dataset analysis before limiting faculty by roster.".to_string()
        })?;

        let mut identifier_lookup: HashMap<String, String> = HashMap::new();
        for identifier in &metadata.analysis.identifier_columns {
            identifier_lookup.insert(identifier.trim().to_lowercase(), identifier.clone());
        }

        let mut resolved_map: HashMap<String, String> = HashMap::new();
        for (raw_identifier, roster_label) in faculty_roster_column_map.iter() {
            let normalized_identifier = raw_identifier.trim().to_lowercase();
            if normalized_identifier.is_empty() {
                continue;
            }

            let trimmed_label = roster_label.trim();
            if trimmed_label.is_empty() {
                continue;
            }

            if let Some(original_identifier) = identifier_lookup.get(&normalized_identifier) {
                resolved_map
                    .entry(original_identifier.clone())
                    .or_insert_with(|| trimmed_label.to_string());
            } else {
                let message = format!(
                    "The roster mapping includes an unknown identifier '{raw_identifier}'.",
                );
                warnings.push(message.clone());
                roster_warning_messages.push(message);
            }
        }

        if resolved_map.is_empty() {
            return Err("Map at least one roster column to a faculty identifier.".into());
        }

        let (mut headers, mut rows) = read_full_spreadsheet(&roster)?;
        align_row_lengths(&mut headers, &mut rows);

        detail_roster_column_map = resolved_map.clone();

        let header_map = build_header_index_map(&headers);
        let mut roster_column_indexes: HashMap<String, usize> = HashMap::new();

        for (identifier, roster_label) in resolved_map.iter() {
            let normalized_label = roster_label.trim().to_lowercase();
            let mut column_index = header_map.get(&normalized_label).copied();

            if column_index.is_none() {
                let normalized_target = normalize_identifier_label(roster_label);
                if !normalized_target.is_empty() {
                    for (candidate_index, header) in headers.iter().enumerate() {
                        if normalize_identifier_label(header) == normalized_target {
                            column_index = Some(candidate_index);
                            break;
                        }
                    }
                }
            }

            if let Some(found_index) = column_index {
                roster_column_indexes.insert(identifier.clone(), found_index);
            } else {
                let message = format!(
                    "The roster does not contain a column named '{roster_label}' for identifier '{identifier}'.",
                );
                warnings.push(message.clone());
                roster_warning_messages.push(message);
            }
        }

        if roster_column_indexes.is_empty() {
            return Err(
                "None of the mapped roster columns were found in the roster spreadsheet.".into(),
            );
        }

        let identifier_order: Vec<String> = metadata
            .analysis
            .identifier_columns
            .iter()
            .filter(|identifier| roster_column_indexes.contains_key(*identifier))
            .cloned()
            .collect();

        if identifier_order.is_empty() {
            return Err("Map at least one roster column to a faculty identifier.".into());
        }

        let mut dataset_index: HashMap<String, HashSet<usize>> = HashMap::new();
        for membership in &metadata.memberships {
            let mut parts = Vec::new();
            for identifier in &identifier_order {
                if let Some(value) = membership.identifiers.get(identifier) {
                    let normalized = normalize_identifier_value(value);
                    if normalized.is_empty() {
                        parts.clear();
                        break;
                    }
                    parts.push(normalized);
                } else {
                    parts.clear();
                    break;
                }
            }

            if parts.is_empty() {
                continue;
            }

            let key = parts.join("|");
            dataset_index
                .entry(key)
                .or_default()
                .insert(membership.row_index);
        }

        if dataset_index.is_empty() {
            let message =
                "No faculty dataset identifiers were available for the selected roster columns."
                    .to_string();
            warnings.push(message.clone());
            roster_warning_messages.push(message);
        }

        let mut matched_rows: HashSet<usize> = HashSet::new();
        let mut unmatched_roster_rows = 0usize;

        for row in &rows {
            let mut parts = Vec::new();
            let mut row_missing = false;

            for identifier in &identifier_order {
                let column_index = match roster_column_indexes.get(identifier) {
                    Some(index) => *index,
                    None => {
                        row_missing = true;
                        break;
                    }
                };

                let value = row
                    .get(column_index)
                    .map(|value| normalize_identifier_value(value));
                match value {
                    Some(normalized) if !normalized.is_empty() => parts.push(normalized),
                    _ => {
                        row_missing = true;
                        break;
                    }
                }
            }

            if row_missing || parts.is_empty() {
                unmatched_roster_rows += 1;
                continue;
            }

            let key = parts.join("|");
            if let Some(rows) = dataset_index.get(&key) {
                matched_rows.extend(rows.iter().copied());
            } else {
                unmatched_roster_rows += 1;
            }
        }

        if unmatched_roster_rows > 0 {
            let message = format!(
                "{unmatched_roster_rows} roster row{plural} did not match any faculty dataset entries.",
                plural = if unmatched_roster_rows == 1 { "" } else { "s" }
            );
            warnings.push(message.clone());
            roster_warning_messages.push(message);
        }

        if matched_rows.is_empty() {
            let message =
                "No faculty in the dataset matched the provided roster identifiers.".to_string();
            warnings.push(message.clone());
            roster_warning_messages.push(message);
        }

        for (identifier, &index) in &roster_column_indexes {
            detail_roster_column_map.insert(identifier.clone(), header_label(&headers, index));
        }

        allowed_faculty_rows = Some(matched_rows);
    }

    if matches!(faculty_scope, FacultyScope::Program) && normalized_programs.is_empty() {
        return Err("Provide at least one program to limit the faculty list.".into());
    }

    if matches!(faculty_scope, FacultyScope::Program) {
        let metadata = load_faculty_dataset_metadata(&app_handle)?
            .ok_or_else(|| {
                "The faculty dataset metadata is unavailable. Refresh the dataset analysis before filtering by program.".to_string()
            })?;
        let filtered_rows =
            filter_faculty_rows_by_program(&metadata.memberships, &normalized_programs);
        if filtered_rows.is_empty() {
            warnings
                .push("No faculty members in the dataset matched the selected programs.".into());
        }
        allowed_faculty_rows = Some(filtered_rows);
    }

    if matches!(faculty_scope, FacultyScope::Custom) && faculty_roster_path.is_none() {
        return Err("Provide a faculty roster spreadsheet to limit the faculty list.".into());
    }

    let details = SubmissionDetails {
        task_type: task_type.clone(),
        faculty_scope: faculty_scope.clone(),
        validated_paths,
        program_filters: match faculty_scope {
            FacultyScope::Program => normalized_programs.clone(),
            _ => Vec::new(),
        },
        custom_faculty_path: faculty_roster_path.clone(),
        recommendations_per_student: faculty_recs_per_student,
        prompt_preview,
        spreadsheet_prompt_columns: selected_prompt_columns.clone(),
        spreadsheet_identifier_columns: detail_identifier_columns.clone(),
        faculty_roster_column_map: detail_roster_column_map.clone(),
        faculty_roster_warnings: roster_warning_messages.clone(),
    };

    let summary = build_summary(
        &task_type,
        &faculty_scope,
        faculty_recs_per_student,
        details.program_filters.len(),
        faculty_roster_path.is_some(),
    );

    let mut prompt_matches = Vec::new();
    let mut directory_results = None;
    let mut spreadsheet_results = None;

    let needs_prompt_embedding = prepared_prompt_text.is_some()
        || matches!(task_type, TaskType::Directory | TaskType::Spreadsheet);
    let mut faculty_embedding_index: Option<FacultyEmbeddingIndex> = None;

    if needs_prompt_embedding {
        let index = load_faculty_embedding_index(&app_handle)?;
        if index.entries.is_empty() {
            return Err(
                "No faculty embeddings are available. Generate embeddings before matching.".into(),
            );
        }
        faculty_embedding_index = Some(index);
    }

    if let Some(prompt_text) = prepared_prompt_text {
        let limit = faculty_recs_per_student.max(1) as usize;
        let embedding_index = faculty_embedding_index
            .as_ref()
            .ok_or_else(|| "The faculty embedding index was not loaded.".to_string())?;
        let prompt_embedding = embed_prompt(&app_handle, embedding_index, &prompt_text)?;
        let mut matches = find_best_faculty_matches(
            embedding_index,
            &prompt_embedding,
            limit,
            allowed_faculty_rows.as_ref(),
        );

        if matches!(task_type, TaskType::Prompt | TaskType::Document) {
            if let Err(err) = enrich_matches_with_faculty_text(
                &app_handle,
                &embedding_index.embedding_columns,
                &mut matches,
            ) {
                warnings.push(format!(
                    "Unable to include faculty text in the match results: {err}"
                ));
            }
        }

        prompt_matches.push(PromptMatchResult {
            prompt: match task_type {
                TaskType::Document => build_prompt_preview(&prompt_text),
                _ => prompt_text.clone(),
            },
            faculty_matches: matches,
        });
    }

    if matches!(task_type, TaskType::Directory) {
        let directory_path = directory_source
            .as_ref()
            .ok_or_else(|| "The directory path was not preserved during processing.".to_string())?;
        let embedding_index = faculty_embedding_index
            .as_ref()
            .ok_or_else(|| "The faculty embedding index was not loaded.".to_string())?;
        let limit = faculty_recs_per_student.max(1) as usize;

        let outcome = process_directory_documents(
            &app_handle,
            directory_path,
            embedding_index,
            limit,
            allowed_faculty_rows.as_ref(),
        )?;

        warnings.extend(outcome.warnings);
        prompt_matches.extend(outcome.prompt_matches);
        directory_results = Some(outcome.results);
    }

    if matches!(task_type, TaskType::Spreadsheet) {
        let spreadsheet_path = spreadsheet_source.as_ref().ok_or_else(|| {
            "The spreadsheet path was not preserved during processing.".to_string()
        })?;
        let embedding_index = faculty_embedding_index
            .as_ref()
            .ok_or_else(|| "The faculty embedding index was not loaded.".to_string())?;
        let limit = faculty_recs_per_student.max(1) as usize;

        let outcome = process_prompt_spreadsheet(
            &app_handle,
            spreadsheet_path,
            embedding_index,
            &selected_prompt_columns,
            &selected_identifier_columns,
            limit,
            allowed_faculty_rows.as_ref(),
        )?;

        warnings.extend(outcome.warnings);
        prompt_matches.extend(outcome.prompt_matches);
        spreadsheet_results = Some(outcome.results);
    }

    {
        let mut match_refs: Vec<&mut Vec<FacultyMatchResult>> = prompt_matches
            .iter_mut()
            .map(|result| &mut result.faculty_matches)
            .collect();
        assign_student_rankings(&mut match_refs);
    }

    Ok(SubmissionResponse {
        summary,
        warnings,
        details,
        prompt_matches,
        directory_results,
        spreadsheet_results,
    })
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct EmbeddingRequestPayload {
    model: String,
    texts: Vec<EmbeddingRequestRow>,
    #[serde(default)]
    #[serde(skip_serializing_if = "Option::is_none")]
    item_label: Option<String>,
    #[serde(default)]
    #[serde(skip_serializing_if = "Option::is_none")]
    item_label_plural: Option<String>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct EmbeddingRequestRow {
    id: usize,
    text: String,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct EmbeddingResponsePayload {
    model: String,
    dimension: usize,
    rows: Vec<EmbeddingResponseRow>,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct EmbeddingResponseRow {
    id: usize,
    embedding: Vec<f32>,
}

#[derive(Deserialize)]
#[serde(tag = "type", rename_all = "lowercase")]
enum EmbeddingHelperEnvelope {
    #[serde(rename_all = "camelCase")]
    Result { payload: EmbeddingResponsePayload },
    #[serde(rename_all = "camelCase")]
    Error { message: String },
}

enum EmbeddingHelperMessage {
    Response(EmbeddingResponsePayload),
    Error(String),
    Terminated(Option<String>),
}

struct EmbeddingHelperProcess {
    child: Child,
    stdin: BufWriter<std::process::ChildStdin>,
    receiver: Receiver<EmbeddingHelperMessage>,
    progress_total: Arc<Mutex<Option<usize>>>,
    stderr_buffer: Arc<Mutex<Vec<u8>>>,
    stdout_handle: Option<std::thread::JoinHandle<()>>,
    stderr_handle: Option<std::thread::JoinHandle<()>>,
}

#[derive(Default)]
struct EmbeddingHelperHandle {
    process: Mutex<Option<EmbeddingHelperProcess>>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct EmbeddingHelperCommand<'a> {
    #[serde(rename = "type")]
    command_type: &'static str,
    #[serde(flatten)]
    payload: &'a EmbeddingRequestPayload,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct FacultyEmbeddingEntry {
    row_index: usize,
    identifiers: HashMap<String, String>,
    embedding: Vec<f32>,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct FacultyEmbeddingIndex {
    model: String,
    #[serde(default)]
    generated_at: Option<String>,
    dimension: usize,
    #[serde(default)]
    total_rows: Option<usize>,
    #[serde(default)]
    embedded_rows: Option<usize>,
    #[serde(default)]
    skipped_rows: Option<usize>,
    #[serde(default)]
    embedding_columns: Vec<String>,
    #[serde(default)]
    identifier_columns: Vec<String>,
    entries: Vec<FacultyEmbeddingEntry>,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct PromptMatchResult {
    prompt: String,
    faculty_matches: Vec<FacultyMatchResult>,
}

struct DocumentExtractionResult {
    text: String,
    warnings: Vec<String>,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct FacultyMatchResult {
    row_index: usize,
    similarity: f32,
    identifiers: HashMap<String, String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    faculty_text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    student_rank_for_faculty: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    student_rank_total: Option<usize>,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct GeneratedSpreadsheet {
    filename: String,
    mime_type: String,
    content: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    encoding: Option<String>,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct DirectoryMatchResults {
    processed_documents: usize,
    matched_documents: usize,
    skipped_documents: usize,
    total_rows: usize,
    preview: SpreadsheetPreview,
    spreadsheet: GeneratedSpreadsheet,
}

#[derive(Debug, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct SpreadsheetMatchResults {
    processed_rows: usize,
    matched_rows: usize,
    skipped_rows: usize,
    total_rows: usize,
    preview: SpreadsheetPreview,
    spreadsheet: GeneratedSpreadsheet,
}

#[derive(Debug, Clone)]
struct MatchEntry {
    student_values: Vec<String>,
    faculty_values: Vec<String>,
    similarity: Option<f32>,
    student_rank: Option<(usize, Option<usize>)>,
    faculty_rank: Option<usize>,
}

#[derive(Debug)]
struct DirectoryProcessingOutcome {
    warnings: Vec<String>,
    prompt_matches: Vec<PromptMatchResult>,
    results: DirectoryMatchResults,
}

#[derive(Debug)]
struct SpreadsheetProcessingOutcome {
    warnings: Vec<String>,
    prompt_matches: Vec<PromptMatchResult>,
    results: SpreadsheetMatchResults,
}

fn load_faculty_embedding_index(
    app_handle: &tauri::AppHandle,
) -> Result<FacultyEmbeddingIndex, String> {
    let embeddings_path = dataset_directory(app_handle)?.join(FACULTY_EMBEDDINGS_NAME);
    let data = if embeddings_path.exists() {
        fs::read(&embeddings_path)
            .map_err(|err| format!("Unable to read faculty embeddings: {err}"))?
    } else {
        DEFAULT_FACULTY_EMBEDDINGS.to_vec()
    };

    serde_json::from_slice(&data)
        .map_err(|err| format!("Unable to parse faculty embeddings: {err}"))
}

fn embed_prompt(
    app_handle: &tauri::AppHandle,
    index: &FacultyEmbeddingIndex,
    prompt: &str,
) -> Result<Vec<f32>, String> {
    let model = if index.model.trim().is_empty() {
        DEFAULT_EMBEDDING_MODEL.to_string()
    } else {
        index.model.clone()
    };

    let payload = EmbeddingRequestPayload {
        model,
        texts: vec![EmbeddingRequestRow {
            id: 0,
            text: prompt.to_string(),
        }],
        item_label: Some("text query".into()),
        item_label_plural: Some("text queries".into()),
    };

    let response = run_embedding_helper(app_handle, &payload)?;
    if response.rows.is_empty() {
        return Err("The embedding helper did not return an embedding for the prompt.".into());
    }

    let embedding = response.rows.into_iter().next().unwrap().embedding;
    if embedding.len() != index.dimension {
        return Err(format!(
            "The prompt embedding dimension ({}) does not match the faculty embedding dimension ({}).",
            embedding.len(),
            index.dimension
        ));
    }

    if response.dimension != index.dimension {
        return Err(format!(
            "The embedding helper reported dimension {} but the faculty index uses {}.",
            response.dimension, index.dimension
        ));
    }

    Ok(embedding)
}

fn find_best_faculty_matches(
    index: &FacultyEmbeddingIndex,
    prompt_embedding: &[f32],
    limit: usize,
    allowed_rows: Option<&HashSet<usize>>,
) -> Vec<FacultyMatchResult> {
    if limit == 0 {
        return Vec::new();
    }

    let mut candidates: Vec<FacultyMatchResult> = index
        .entries
        .iter()
        .filter_map(|entry| {
            if let Some(allowed) = allowed_rows {
                if !allowed.contains(&entry.row_index) {
                    return None;
                }
            }

            if entry.embedding.len() != prompt_embedding.len() {
                return None;
            }

            let similarity = cosine_similarity(prompt_embedding, &entry.embedding)?;
            let mut identifiers = entry.identifiers.clone();
            identifiers.retain(|_, value| !value.trim().is_empty());

            Some(FacultyMatchResult {
                row_index: entry.row_index,
                similarity,
                identifiers,
                faculty_text: None,
                student_rank_for_faculty: None,
                student_rank_total: None,
            })
        })
        .collect();

    candidates.sort_by(|a, b| {
        b.similarity
            .partial_cmp(&a.similarity)
            .unwrap_or(Ordering::Equal)
    });
    candidates.truncate(limit);
    candidates
}

fn assign_student_rankings(match_sets: &mut [&mut Vec<FacultyMatchResult>]) {
    if match_sets.is_empty() {
        return;
    }

    let mut occurrences: HashMap<usize, Vec<(usize, usize, f32)>> = HashMap::new();

    for (prompt_index, matches) in match_sets.iter().enumerate() {
        for (match_index, faculty) in matches.iter().enumerate() {
            occurrences.entry(faculty.row_index).or_default().push((
                prompt_index,
                match_index,
                faculty.similarity,
            ));
        }
    }

    if occurrences.is_empty() {
        for matches in match_sets.iter_mut() {
            for faculty in matches.iter_mut() {
                faculty.student_rank_for_faculty = None;
                faculty.student_rank_total = None;
            }
        }
        return;
    }

    let mut rank_map: HashMap<(usize, usize), (usize, usize)> = HashMap::new();

    for (_, mut entries) in occurrences {
        entries.sort_by(|a, b| {
            let primary = b.2.partial_cmp(&a.2).unwrap_or(Ordering::Equal);
            if primary != Ordering::Equal {
                return primary;
            }

            let secondary = a.0.cmp(&b.0);
            if secondary != Ordering::Equal {
                return secondary;
            }

            a.1.cmp(&b.1)
        });

        let total = entries.len();
        for (position, (prompt_index, match_index, _)) in entries.into_iter().enumerate() {
            rank_map.insert((prompt_index, match_index), (position + 1, total));
        }
    }

    for (prompt_index, matches) in match_sets.iter_mut().enumerate() {
        for (match_index, faculty) in matches.iter_mut().enumerate() {
            if let Some((rank, total)) = rank_map.get(&(prompt_index, match_index)) {
                faculty.student_rank_for_faculty = Some(*rank);
                faculty.student_rank_total = Some(*total);
            } else {
                faculty.student_rank_for_faculty = None;
                faculty.student_rank_total = None;
            }
        }
    }
}

fn enrich_matches_with_faculty_text(
    app_handle: &tauri::AppHandle,
    embedding_columns: &[String],
    matches: &mut [FacultyMatchResult],
) -> Result<(), String> {
    if matches.is_empty() {
        return Ok(());
    }

    if embedding_columns.is_empty() {
        return Ok(());
    }

    let dataset_path = dataset_destination(app_handle)?;
    if !dataset_path.exists() {
        return Err("The faculty dataset could not be located.".into());
    }

    let (headers, rows) = read_full_spreadsheet(&dataset_path)?;
    if rows.is_empty() {
        return Err("The faculty dataset does not include any rows.".into());
    }

    let header_map = build_header_index_map(&headers);
    let embedding_indexes = indexes_from_labels(&header_map, embedding_columns)?;
    if embedding_indexes.is_empty() {
        return Err(
            "No embedding columns are available to retrieve faculty text. Re-run the dataset analysis.".into(),
        );
    }

    for faculty_match in matches {
        if let Some(row) = rows.get(faculty_match.row_index) {
            let mut text_parts = Vec::new();
            for &index in &embedding_indexes {
                if let Some(value) = row.get(index) {
                    let trimmed = value.trim();
                    if !trimmed.is_empty() {
                        text_parts.push(trimmed.to_string());
                    }
                }
            }

            if !text_parts.is_empty() {
                faculty_match.faculty_text = Some(text_parts.join("\n\n"));
            }
        }
    }

    Ok(())
}

fn process_directory_documents(
    app_handle: &tauri::AppHandle,
    directory: &Path,
    index: &FacultyEmbeddingIndex,
    limit: usize,
    allowed_rows: Option<&HashSet<usize>>,
) -> Result<DirectoryProcessingOutcome, String> {
    #[derive(Debug)]
    struct DirectoryDocumentContext {
        result_index: usize,
        prompt: String,
    }

    #[derive(Debug)]
    struct DirectoryDocumentResult {
        identifier: String,
        preview: String,
        prompt_label: Option<String>,
        matches: Vec<FacultyMatchResult>,
        status_message: Option<String>,
    }

    let mut warnings = Vec::new();
    let mut document_results: Vec<DirectoryDocumentResult> = Vec::new();
    let mut contexts: Vec<DirectoryDocumentContext> = Vec::new();
    let mut file_paths: Vec<PathBuf> = Vec::new();

    let reader = fs::read_dir(directory).map_err(|err| {
        format!(
            "Unable to read the directory '{}': {err}",
            directory.display()
        )
    })?;

    for entry in reader {
        match entry {
            Ok(entry) => {
                let path = entry.path();
                match entry.file_type() {
                    Ok(file_type) => {
                        if file_type.is_file() {
                            file_paths.push(path);
                        }
                    }
                    Err(err) => warnings.push(format!(
                        "Skipped '{}': unable to determine the file type ({err}).",
                        path.to_string_lossy()
                    )),
                }
            }
            Err(err) => warnings.push(format!(
                "Unable to read an entry in '{}': {err}",
                directory.display()
            )),
        }
    }

    file_paths.sort();

    for path in file_paths {
        let identifier = path
            .file_name()
            .and_then(|name| name.to_str())
            .map(|value| value.to_string())
            .unwrap_or_else(|| path.to_string_lossy().into_owned());

        let mut result = DirectoryDocumentResult {
            identifier: identifier.clone(),
            preview: String::new(),
            prompt_label: None,
            matches: Vec::new(),
            status_message: None,
        };
        let mut prompt_text: Option<String> = None;

        match extract_document_prompt(&path) {
            Ok(DocumentExtractionResult {
                text,
                warnings: extraction_warnings,
            }) => {
                for warning in extraction_warnings {
                    warnings.push(format!("{identifier}: {warning}"));
                }

                if text.trim().is_empty() {
                    let message =
                        format!("Skipped '{identifier}' because it did not contain readable text.");
                    warnings.push(message.clone());
                    result.status_message = Some(message);
                } else {
                    result.preview = build_prompt_preview(&text);
                    if result.preview.is_empty() {
                        result.prompt_label = Some(result.identifier.clone());
                    } else {
                        result.prompt_label =
                            Some(format!("{} — {}", result.identifier, result.preview));
                    }
                    prompt_text = Some(text);
                }
            }
            Err(err) => {
                warnings.push(err.clone());
                result.status_message = Some(err);
            }
        }

        let result_index = document_results.len();
        if let Some(text) = prompt_text {
            contexts.push(DirectoryDocumentContext {
                result_index,
                prompt: text,
            });
        }

        document_results.push(result);
    }

    if document_results.is_empty() {
        warnings.push("The selected directory did not contain any files to process.".into());
    }

    let mut prompt_matches = Vec::new();
    let mut missing_embeddings = 0usize;

    if !contexts.is_empty() {
        let model_name = if index.model.trim().is_empty() {
            DEFAULT_EMBEDDING_MODEL.to_string()
        } else {
            index.model.clone()
        };

        let payload = EmbeddingRequestPayload {
            model: model_name,
            texts: contexts
                .iter()
                .enumerate()
                .map(|(id, context)| EmbeddingRequestRow {
                    id,
                    text: context.prompt.clone(),
                })
                .collect(),
            item_label: Some("document".into()),
            item_label_plural: Some("documents".into()),
        };

        let response = run_embedding_helper(app_handle, &payload)?;
        if response.dimension != index.dimension {
            return Err(format!(
                "The document embedding dimension ({}) does not match the faculty embedding dimension ({}).",
                response.dimension, index.dimension
            ));
        }

        let mut embedding_map: HashMap<usize, Vec<f32>> = HashMap::new();
        for row in response.rows {
            embedding_map.insert(row.id, row.embedding);
        }

        for (context_index, context) in contexts.iter().enumerate() {
            let identifier = document_results[context.result_index].identifier.clone();

            match embedding_map.remove(&context_index) {
                Some(embedding) => {
                    let matches = find_best_faculty_matches(index, &embedding, limit, allowed_rows);

                    if matches.is_empty() {
                        document_results[context.result_index].status_message =
                            Some("No faculty matches were returned.".into());
                    } else {
                        document_results[context.result_index].status_message = None;
                    }

                    document_results[context.result_index].matches = matches;
                }
                None => {
                    missing_embeddings += 1;
                    let message = "The embedding helper did not return a result for this document."
                        .to_string();
                    document_results[context.result_index].status_message = Some(message.clone());
                    warnings.push(format!(
                        "The embedding helper did not return an embedding for '{}'.",
                        identifier
                    ));
                }
            }
        }
    } else if !document_results.is_empty() {
        warnings
            .push("None of the files in the directory contained readable text to embed.".into());
    }

    {
        let mut match_refs: Vec<&mut Vec<FacultyMatchResult>> = document_results
            .iter_mut()
            .map(|result| &mut result.matches)
            .collect();
        assign_student_rankings(&mut match_refs);
    }

    for result in &document_results {
        if let Some(label) = &result.prompt_label {
            prompt_matches.push(PromptMatchResult {
                prompt: label.clone(),
                faculty_matches: result.matches.clone(),
            });
        }
    }

    let processed_documents = if contexts.is_empty() {
        0
    } else {
        contexts.len().saturating_sub(missing_embeddings)
    };
    let skipped_documents = document_results.len().saturating_sub(processed_documents);
    let matched_documents = document_results
        .iter()
        .filter(|result| !result.matches.is_empty())
        .count();

    let student_headers = vec!["Document".to_string()];
    let faculty_headers = index.identifier_columns.clone();
    let headers = build_matches_headers(&student_headers, &faculty_headers);

    let mut preview_rows: Vec<Vec<String>> = Vec::new();
    let mut match_entries: Vec<MatchEntry> = Vec::new();
    let mut student_summary_rows: Vec<Vec<String>> = Vec::new();

    for result in &document_results {
        student_summary_rows.push(vec![result.identifier.clone()]);

        if result.matches.is_empty() {
            let message = result
                .status_message
                .clone()
                .unwrap_or_else(|| "No faculty matches were returned.".into());
            let mut preview_row = Vec::new();
            preview_row.push(String::new());
            preview_row.push(String::new());
            preview_row.push(result.identifier.clone());
            preview_row.extend(vec![String::new(); faculty_headers.len()]);
            preview_row.push(message);
            preview_row.push(String::new());
            preview_row.push(String::new());
            if preview_rows.len() < 20 {
                preview_rows.push(preview_row);
            }
            continue;
        }

        for (rank, faculty) in result.matches.iter().enumerate() {
            let faculty_values: Vec<String> = faculty_headers
                .iter()
                .map(|label| faculty.identifiers.get(label).cloned().unwrap_or_default())
                .collect();
            let similarity = faculty.similarity;
            let student_rank = faculty
                .student_rank_for_faculty
                .map(|value| (value, faculty.student_rank_total));
            let student_rank_text = student_rank
                .map(|(position, total)| match total {
                    Some(limit) => format!("{position} of {limit}"),
                    None => position.to_string(),
                })
                .unwrap_or_default();

            let mut preview_row = Vec::new();
            preview_row.push(String::new());
            preview_row.push(String::new());
            preview_row.push(result.identifier.clone());
            preview_row.extend(faculty_values.clone());
            preview_row.push(format_similarity_percent(similarity));
            preview_row.push(student_rank_text);
            preview_row.push((rank + 1).to_string());
            if preview_rows.len() < 20 {
                preview_rows.push(preview_row);
            }

            match_entries.push(MatchEntry {
                student_values: vec![result.identifier.clone()],
                faculty_values,
                similarity: Some(similarity),
                student_rank,
                faculty_rank: Some(rank + 1),
            });
        }
    }

    let preview = SpreadsheetPreview {
        headers: headers.clone(),
        rows: preview_rows,
        suggested_prompt_columns: Vec::new(),
        suggested_identifier_columns: Vec::new(),
    };

    let workbook_bytes = build_matches_workbook(
        &student_headers,
        &student_summary_rows,
        &faculty_headers,
        &match_entries,
    )?;
    let encoded_workbook = Base64Engine.encode(workbook_bytes);

    let results = DirectoryMatchResults {
        processed_documents,
        matched_documents,
        skipped_documents,
        total_rows: match_entries.len(),
        preview,
        spreadsheet: GeneratedSpreadsheet {
            filename: default_directory_workbook_name(),
            mime_type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet".into(),
            content: encoded_workbook,
            encoding: Some("base64".into()),
        },
    };

    Ok(DirectoryProcessingOutcome {
        warnings,
        prompt_matches,
        results,
    })
}

fn process_prompt_spreadsheet(
    app_handle: &tauri::AppHandle,
    spreadsheet_path: &Path,
    index: &FacultyEmbeddingIndex,
    prompt_columns: &[String],
    identifier_columns: &[String],
    limit: usize,
    allowed_rows: Option<&HashSet<usize>>,
) -> Result<SpreadsheetProcessingOutcome, String> {
    #[derive(Debug)]
    struct SpreadsheetRowContext {
        result_index: usize,
        prompt: String,
    }

    #[derive(Debug)]
    struct SpreadsheetRowResult {
        warning_label: String,
        identifier_values: Vec<String>,
        identifier_label: String,
        prompt_preview: String,
        prompt_label: Option<String>,
        matches: Vec<FacultyMatchResult>,
        status_message: Option<String>,
    }

    let (headers, rows) = read_full_spreadsheet(spreadsheet_path)?;
    let header_map = build_header_index_map(&headers);
    let prompt_indexes = indexes_from_spreadsheet_labels(&header_map, prompt_columns)?;
    let identifier_indexes = indexes_from_spreadsheet_labels(&header_map, identifier_columns)?;
    let include_row_number_column = identifier_indexes.is_empty();

    let mut warnings = Vec::new();
    let mut contexts: Vec<SpreadsheetRowContext> = Vec::new();
    let mut row_results: Vec<SpreadsheetRowResult> = Vec::new();

    if rows.is_empty() {
        warnings.push("The spreadsheet did not include any data rows to process.".into());
    }

    for (row_index, row) in rows.iter().enumerate() {
        let row_number = row_index + 2;
        let mut identifier_values = Vec::new();
        let mut label_segments = Vec::new();

        if include_row_number_column {
            identifier_values.push(row_number.to_string());
            label_segments.push(format!("Row {}", row_number));
        }

        for &index in &identifier_indexes {
            let value = row.get(index).cloned().unwrap_or_default();
            if !value.trim().is_empty() {
                label_segments.push(value.trim().to_string());
            }
            identifier_values.push(value);
        }

        let identifier_label = if label_segments.is_empty() {
            format!("Row {}", row_number)
        } else {
            label_segments.join(" – ")
        };

        let warning_label = if label_segments.is_empty()
            || (include_row_number_column && label_segments.len() == 1)
        {
            format!("row {}", row_number)
        } else {
            format!("row {} ({})", row_number, identifier_label.as_str())
        };

        let mut prompt_parts = Vec::new();
        for &index in &prompt_indexes {
            if let Some(value) = row.get(index) {
                let trimmed = value.trim();
                if !trimmed.is_empty() {
                    prompt_parts.push(trimmed.to_string());
                }
            }
        }

        let mut result = SpreadsheetRowResult {
            warning_label,
            identifier_values,
            identifier_label,
            prompt_preview: String::new(),
            prompt_label: None,
            matches: Vec::new(),
            status_message: None,
        };

        if prompt_parts.is_empty() {
            warnings.push(format!(
                "Skipped {} because the selected prompt columns were empty.",
                result.warning_label
            ));
            result.status_message =
                Some("No prompt content was provided in the selected columns.".into());
        } else {
            let prompt_text = prompt_parts.join("\n\n");
            result.prompt_preview = build_prompt_preview(&prompt_text);
            if result.prompt_preview.is_empty() {
                result.prompt_label = Some(result.identifier_label.clone());
            } else {
                result.prompt_label = Some(format!(
                    "{} — {}",
                    result.identifier_label, result.prompt_preview
                ));
            }
            let result_index = row_results.len();
            contexts.push(SpreadsheetRowContext {
                result_index,
                prompt: prompt_text,
            });
        }

        row_results.push(result);
    }

    let mut prompt_matches = Vec::new();
    let mut missing_embeddings = 0usize;

    if !contexts.is_empty() {
        let model_name = if index.model.trim().is_empty() {
            DEFAULT_EMBEDDING_MODEL.to_string()
        } else {
            index.model.clone()
        };

        let payload = EmbeddingRequestPayload {
            model: model_name,
            texts: contexts
                .iter()
                .enumerate()
                .map(|(id, context)| EmbeddingRequestRow {
                    id,
                    text: context.prompt.clone(),
                })
                .collect(),
            item_label: Some("spreadsheet row".into()),
            item_label_plural: Some("spreadsheet rows".into()),
        };

        let response = run_embedding_helper(app_handle, &payload)?;
        if response.dimension != index.dimension {
            return Err(format!(
                "The spreadsheet embedding dimension ({}) does not match the faculty embedding dimension ({}).",
                response.dimension, index.dimension
            ));
        }

        let mut embedding_map: HashMap<usize, Vec<f32>> = HashMap::new();
        for row in response.rows {
            embedding_map.insert(row.id, row.embedding);
        }

        for (context_index, context) in contexts.iter().enumerate() {
            let result = &mut row_results[context.result_index];

            match embedding_map.remove(&context_index) {
                Some(embedding) => {
                    let matches = find_best_faculty_matches(index, &embedding, limit, allowed_rows);

                    if matches.is_empty() {
                        result.status_message = Some("No faculty matches were returned.".into());
                    } else {
                        result.status_message = None;
                    }

                    result.matches = matches;
                }
                None => {
                    missing_embeddings += 1;
                    let message =
                        "The embedding helper did not return a result for this row.".to_string();
                    warnings.push(format!(
                        "The embedding helper did not return an embedding for {}.",
                        result.warning_label
                    ));
                    result.status_message = Some(message.clone());
                }
            }
        }
    } else if !row_results.is_empty() {
        warnings.push("None of the rows in the spreadsheet contained prompt text to embed.".into());
    }

    {
        let mut match_refs: Vec<&mut Vec<FacultyMatchResult>> = row_results
            .iter_mut()
            .map(|result| &mut result.matches)
            .collect();
        assign_student_rankings(&mut match_refs);
    }

    for result in &row_results {
        if let Some(label) = &result.prompt_label {
            prompt_matches.push(PromptMatchResult {
                prompt: label.clone(),
                faculty_matches: result.matches.clone(),
            });
        }
    }

    let processed_rows = if contexts.is_empty() {
        0
    } else {
        contexts.len().saturating_sub(missing_embeddings)
    };
    let skipped_rows = row_results.len().saturating_sub(processed_rows);
    let matched_rows = row_results
        .iter()
        .filter(|result| !result.matches.is_empty())
        .count();

    let student_headers: Vec<String> = if include_row_number_column {
        vec!["Row Number".into()]
    } else {
        identifier_columns
            .iter()
            .map(|label| label.trim().to_string())
            .collect()
    };
    let faculty_headers: Vec<String> = index.identifier_columns.clone();

    let headers = build_matches_headers(&student_headers, &faculty_headers);

    let mut match_entries: Vec<MatchEntry> = Vec::new();
    let mut preview_rows: Vec<Vec<String>> = Vec::new();

    for result in &row_results {
        if result.matches.is_empty() {
            let message = result
                .status_message
                .clone()
                .unwrap_or_else(|| "No faculty matches were returned.".into());
            let mut preview_row = Vec::new();
            preview_row.push(String::new());
            preview_row.push(String::new());
            preview_row.extend(result.identifier_values.clone());
            preview_row.extend(vec![String::new(); faculty_headers.len()]);
            preview_row.push(message.clone());
            preview_row.push(String::new());
            preview_row.push(String::new());
            if preview_rows.len() < 20 {
                preview_rows.push(preview_row);
            }
            continue;
        }

        for (rank, faculty) in result.matches.iter().enumerate() {
            let faculty_values: Vec<String> = faculty_headers
                .iter()
                .map(|label| faculty.identifiers.get(label).cloned().unwrap_or_default())
                .collect();
            let similarity = faculty.similarity;
            let student_rank = faculty
                .student_rank_for_faculty
                .map(|value| (value, faculty.student_rank_total));
            let student_rank_text = student_rank
                .map(|(position, total)| match total {
                    Some(limit) => format!("{position} of {limit}"),
                    None => position.to_string(),
                })
                .unwrap_or_default();

            let mut preview_row = Vec::new();
            preview_row.push(String::new());
            preview_row.push(String::new());
            preview_row.extend(result.identifier_values.clone());
            preview_row.extend(faculty_values.clone());
            preview_row.push(format_similarity_percent(similarity));
            preview_row.push(student_rank_text.clone());
            preview_row.push((rank + 1).to_string());
            if preview_rows.len() < 20 {
                preview_rows.push(preview_row);
            }

            match_entries.push(MatchEntry {
                student_values: result.identifier_values.clone(),
                faculty_values,
                similarity: Some(similarity),
                student_rank,
                faculty_rank: Some(rank + 1),
            });
        }
    }

    let preview = SpreadsheetPreview {
        headers: headers.clone(),
        rows: preview_rows,
        suggested_prompt_columns: Vec::new(),
        suggested_identifier_columns: Vec::new(),
    };

    let student_summary_rows: Vec<Vec<String>> = row_results
        .iter()
        .map(|result| result.identifier_values.clone())
        .collect();
    let workbook_bytes = build_matches_workbook(
        &student_headers,
        &student_summary_rows,
        &faculty_headers,
        &match_entries,
    )?;
    let encoded_workbook = Base64Engine.encode(workbook_bytes);

    let results = SpreadsheetMatchResults {
        processed_rows,
        matched_rows,
        skipped_rows,
        total_rows: match_entries.len(),
        preview,
        spreadsheet: GeneratedSpreadsheet {
            filename: default_matches_workbook_name(),
            mime_type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet".into(),
            content: encoded_workbook,
            encoding: Some("base64".into()),
        },
    };

    Ok(SpreadsheetProcessingOutcome {
        warnings,
        prompt_matches,
        results,
    })
}

fn build_matches_headers(student_headers: &[String], faculty_headers: &[String]) -> Vec<String> {
    let mut headers: Vec<String> = Vec::new();
    headers.push("First reviewer".into());
    headers.push("Reviewer".into());
    headers.extend(student_headers.iter().cloned());
    headers.extend(faculty_headers.iter().cloned());
    headers.push("Similarity %".into());
    headers.push("Student rank".into());
    headers.push("Faculty rank".into());
    headers
}

fn build_matches_workbook(
    student_headers: &[String],
    student_summary_rows: &[Vec<String>],
    faculty_headers: &[String],
    match_entries: &[MatchEntry],
) -> Result<Vec<u8>, String> {
    let mut workbook = Workbook::new();
    let matches_sheet_name = "Matches";
    let matches_sheet = workbook.add_worksheet();
    matches_sheet
        .set_name(matches_sheet_name)
        .map_err(|err| format!("Unable to configure the matches worksheet: {err}"))?;

    let header_format = Format::new().set_bold();
    let percent_format = Format::new().set_num_format("0.0%");

    let headers = build_matches_headers(student_headers, faculty_headers);
    for (col_index, header) in headers.iter().enumerate() {
        matches_sheet
            .write_string_with_format(0, col_index as u16, header, &header_format)
            .map_err(|err| format!("Unable to write the matches header row: {err}"))?;
    }

    let student_offset = 2u32;
    let faculty_offset = student_offset + student_headers.len() as u32;
    let similarity_col = faculty_offset + faculty_headers.len() as u32;
    let student_rank_col = similarity_col + 1;
    let faculty_rank_col = student_rank_col + 1;

    for (row_index, entry) in match_entries.iter().enumerate() {
        let row = (row_index + 1) as u32;
        matches_sheet
            .write_string(row, 0, "")
            .map_err(|err| format!("Unable to write the first reviewer column: {err}"))?;
        matches_sheet
            .write_string(row, 1, "")
            .map_err(|err| format!("Unable to write the reviewer column: {err}"))?;

        for (offset, value) in entry.student_values.iter().enumerate() {
            matches_sheet
                .write_string(row, (student_offset + offset as u32) as u16, value)
                .map_err(|err| format!("Unable to write a student identifier value: {err}"))?;
        }

        for (offset, value) in entry.faculty_values.iter().enumerate() {
            matches_sheet
                .write_string(row, (faculty_offset + offset as u32) as u16, value)
                .map_err(|err| format!("Unable to write a faculty identifier value: {err}"))?;
        }

        if let Some(value) = entry.similarity {
            matches_sheet
                .write_number_with_format(
                    row,
                    similarity_col as u16,
                    f64::from(value),
                    &percent_format,
                )
                .map_err(|err| format!("Unable to write the similarity percentage: {err}"))?;
        } else {
            matches_sheet
                .write_string(row, similarity_col as u16, "")
                .map_err(|err| format!("Unable to write the similarity placeholder: {err}"))?;
        }

        if let Some((position, total)) = entry.student_rank {
            let text = match total {
                Some(limit) => format!("{position} of {limit}"),
                None => position.to_string(),
            };
            matches_sheet
                .write_string(row, student_rank_col as u16, text)
                .map_err(|err| format!("Unable to write the student rank: {err}"))?;
        } else {
            matches_sheet
                .write_string(row, student_rank_col as u16, "")
                .map_err(|err| format!("Unable to write the student rank placeholder: {err}"))?;
        }

        if let Some(rank) = entry.faculty_rank {
            matches_sheet
                .write_number(row, faculty_rank_col as u16, rank as f64)
                .map_err(|err| format!("Unable to write the faculty rank: {err}"))?;
        } else {
            matches_sheet
                .write_string(row, faculty_rank_col as u16, "")
                .map_err(|err| format!("Unable to write the faculty rank placeholder: {err}"))?;
        }
    }

    let match_row_count = match_entries.len() as u32;
    let mut student_summary_headers = student_headers.to_vec();
    student_summary_headers.push("Total first reviewers".into());
    student_summary_headers.push("Total reviewers".into());

    let student_summary_sheet = workbook.add_worksheet();
    student_summary_sheet
        .set_name("Student Summary")
        .map_err(|err| format!("Unable to configure the student summary worksheet: {err}"))?;
    for (col_index, header) in student_summary_headers.iter().enumerate() {
        student_summary_sheet
            .write_string_with_format(0, col_index as u16, header, &header_format)
            .map_err(|err| format!("Unable to write the student summary header row: {err}"))?;
    }

    let first_reviewer_range = if match_row_count > 0 {
        Some(excel_range_reference(
            matches_sheet_name,
            1,
            0,
            match_row_count,
            0,
        ))
    } else {
        None
    };
    let reviewer_range = if match_row_count > 0 {
        Some(excel_range_reference(
            matches_sheet_name,
            1,
            1,
            match_row_count,
            1,
        ))
    } else {
        None
    };

    for (row_index, identifiers) in student_summary_rows.iter().enumerate() {
        let row = (row_index + 1) as u32;
        for (col_offset, value) in identifiers.iter().enumerate() {
            student_summary_sheet
                .write_string(row, col_offset as u16, value)
                .map_err(|err| {
                    format!("Unable to write a student summary identifier value: {err}")
                })?;
        }

        let first_col = student_headers.len();
        let total_col = first_col + 1;

        if match_row_count == 0 {
            student_summary_sheet
                .write_number(row, first_col as u16, 0.0)
                .map_err(|err| {
                    format!("Unable to write the student first reviewer placeholder: {err}")
                })?;
            student_summary_sheet
                .write_number(row, total_col as u16, 0.0)
                .map_err(|err| {
                    format!("Unable to write the student reviewer count placeholder: {err}")
                })?;
            continue;
        }

        let mut first_factors = Vec::new();
        first_factors.push(format!(
            "--({first}=1)",
            first = first_reviewer_range.as_ref().unwrap()
        ));
        for (col_offset, _) in student_headers.iter().enumerate() {
            let student_range = excel_range_reference(
                matches_sheet_name,
                1,
                student_offset + col_offset as u32,
                match_row_count,
                student_offset + col_offset as u32,
            );
            let summary_cell = excel_cell_reference(row, col_offset as u32, true, false);
            first_factors.push(format!("--({student_range}={summary_cell})"));
        }
        let first_formula = build_sumproduct_formula(&first_factors);
        student_summary_sheet
            .write_formula(row, first_col as u16, first_formula.as_str())
            .map_err(|err| format!("Unable to write the student first reviewer formula: {err}"))?;

        let mut total_factors = Vec::new();
        total_factors.push(format!(
            "--((( {first}=1)+({reviewer}=1))>0)",
            first = first_reviewer_range.as_ref().unwrap(),
            reviewer = reviewer_range.as_ref().unwrap()
        ));
        for (col_offset, _) in student_headers.iter().enumerate() {
            let student_range = excel_range_reference(
                matches_sheet_name,
                1,
                student_offset + col_offset as u32,
                match_row_count,
                student_offset + col_offset as u32,
            );
            let summary_cell = excel_cell_reference(row, col_offset as u32, true, false);
            total_factors.push(format!("--({student_range}={summary_cell})"));
        }
        let total_formula = build_sumproduct_formula(&total_factors);
        student_summary_sheet
            .write_formula(row, total_col as u16, total_formula.as_str())
            .map_err(|err| format!("Unable to write the student reviewer count formula: {err}"))?;
    }

    let mut faculty_summary_headers = faculty_headers.to_vec();
    faculty_summary_headers.push("First reviewer count".into());
    faculty_summary_headers.push("Total reviewer count".into());

    let faculty_summary_sheet = workbook.add_worksheet();
    faculty_summary_sheet
        .set_name("Faculty Summary")
        .map_err(|err| format!("Unable to configure the faculty summary worksheet: {err}"))?;
    for (col_index, header) in faculty_summary_headers.iter().enumerate() {
        faculty_summary_sheet
            .write_string_with_format(0, col_index as u16, header, &header_format)
            .map_err(|err| format!("Unable to write the faculty summary header row: {err}"))?;
    }

    let mut seen_faculty = HashSet::new();
    let mut faculty_summary_rows: Vec<Vec<String>> = Vec::new();
    for entry in match_entries {
        if entry.faculty_rank.is_none() {
            continue;
        }
        let key = entry.faculty_values.join("\u{1f}");
        if seen_faculty.insert(key) {
            faculty_summary_rows.push(entry.faculty_values.clone());
        }
    }

    for (row_index, identifiers) in faculty_summary_rows.iter().enumerate() {
        let row = (row_index + 1) as u32;
        for (col_offset, value) in identifiers.iter().enumerate() {
            faculty_summary_sheet
                .write_string(row, col_offset as u16, value)
                .map_err(|err| {
                    format!("Unable to write a faculty summary identifier value: {err}")
                })?;
        }

        let first_col = faculty_headers.len();
        let total_col = first_col + 1;

        if match_row_count == 0 {
            faculty_summary_sheet
                .write_number(row, first_col as u16, 0.0)
                .map_err(|err| {
                    format!("Unable to write the faculty first reviewer placeholder: {err}")
                })?;
            faculty_summary_sheet
                .write_number(row, total_col as u16, 0.0)
                .map_err(|err| {
                    format!("Unable to write the faculty reviewer placeholder: {err}")
                })?;
            continue;
        }

        let mut first_factors = Vec::new();
        first_factors.push(format!(
            "--({first}=1)",
            first = first_reviewer_range.as_ref().unwrap()
        ));
        for (col_offset, _) in faculty_headers.iter().enumerate() {
            let faculty_range = excel_range_reference(
                matches_sheet_name,
                1,
                faculty_offset + col_offset as u32,
                match_row_count,
                faculty_offset + col_offset as u32,
            );
            let summary_cell = excel_cell_reference(row, col_offset as u32, true, false);
            first_factors.push(format!("--({faculty_range}={summary_cell})"));
        }
        let first_formula = build_sumproduct_formula(&first_factors);
        faculty_summary_sheet
            .write_formula(row, first_col as u16, first_formula.as_str())
            .map_err(|err| format!("Unable to write the faculty first reviewer formula: {err}"))?;

        let mut total_factors = Vec::new();
        total_factors.push(format!(
            "--((( {first}=1)+({reviewer}=1))>0)",
            first = first_reviewer_range.as_ref().unwrap(),
            reviewer = reviewer_range.as_ref().unwrap()
        ));
        for (col_offset, _) in faculty_headers.iter().enumerate() {
            let faculty_range = excel_range_reference(
                matches_sheet_name,
                1,
                faculty_offset + col_offset as u32,
                match_row_count,
                faculty_offset + col_offset as u32,
            );
            let summary_cell = excel_cell_reference(row, col_offset as u32, true, false);
            total_factors.push(format!("--({faculty_range}={summary_cell})"));
        }
        let total_formula = build_sumproduct_formula(&total_factors);
        faculty_summary_sheet
            .write_formula(row, total_col as u16, total_formula.as_str())
            .map_err(|err| format!("Unable to write the faculty reviewer formula: {err}"))?;
    }

    workbook
        .save_to_buffer()
        .map_err(|err| format!("Unable to finalize the match workbook: {err}"))
}

fn build_sumproduct_formula(factors: &[String]) -> String {
    if factors.is_empty() {
        "=0".into()
    } else {
        format!("=SUMPRODUCT({})", factors.join(", "))
    }
}

fn excel_column_name(mut index: u32) -> String {
    let mut name = String::new();
    loop {
        let remainder = index % 26;
        name.push((b'A' + remainder as u8) as char);
        if index < 26 {
            break;
        }
        index = index / 26 - 1;
    }
    name.chars().rev().collect()
}

fn excel_cell_reference(row: u32, col: u32, absolute_col: bool, absolute_row: bool) -> String {
    let mut reference = String::new();
    if absolute_col {
        reference.push('$');
    }
    reference.push_str(&excel_column_name(col));
    if absolute_row {
        reference.push('$');
    }
    reference.push_str(&(row + 1).to_string());
    reference
}

fn excel_range_reference(
    sheet: &str,
    start_row: u32,
    start_col: u32,
    end_row: u32,
    end_col: u32,
) -> String {
    let escaped_sheet = sheet.replace('\'', "''");
    let start = excel_cell_reference(start_row, start_col, true, true);
    let end = excel_cell_reference(end_row, end_col, true, true);
    format!("'{}'!{}:{}", escaped_sheet, start, end)
}

fn default_matches_workbook_name() -> String {
    let timestamp = Utc::now().format("%Y%m%d-%H%M%S");
    format!("DBBS_matches_{timestamp}.xlsx")
}

fn default_directory_workbook_name() -> String {
    let timestamp = Utc::now().format("%Y%m%d-%H%M%S");
    format!("DBBS_directory_matches_{timestamp}.xlsx")
}

fn format_similarity_percent(value: f32) -> String {
    if value.is_finite() {
        format!("{:.1}%", value * 100.0)
    } else {
        "n/a".into()
    }
}

fn cosine_similarity(a: &[f32], b: &[f32]) -> Option<f32> {
    if a.len() != b.len() || a.is_empty() {
        return None;
    }

    let mut dot = 0.0f64;
    let mut norm_a = 0.0f64;
    let mut norm_b = 0.0f64;

    for (&x, &y) in a.iter().zip(b.iter()) {
        let xf = f64::from(x);
        let yf = f64::from(y);
        dot += xf * yf;
        norm_a += xf * xf;
        norm_b += yf * yf;
    }

    if norm_a == 0.0 || norm_b == 0.0 {
        return None;
    }

    Some((dot / (norm_a.sqrt() * norm_b.sqrt())) as f32)
}

fn extract_document_prompt(path: &Path) -> Result<DocumentExtractionResult, String> {
    let data = fs::read(path)
        .map_err(|err| format!("Unable to read document '{}': {err}", path.display()))?;
    let extension = path
        .extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_ascii_lowercase());

    let mut warnings = Vec::new();

    let raw_text = match extension.as_deref() {
        Some("txt") => decode_text_bytes(&data, &mut warnings),
        Some("pdf") => extract_pdf_text(&data)?,
        Some("docx") => extract_docx_text(&data)?,
        Some("doc") => extract_doc_text(&data, &mut warnings)?,
        _ => detect_and_extract_unknown_document(&data, &mut warnings)?,
    };

    let normalized = normalize_document_text(&raw_text);

    Ok(DocumentExtractionResult {
        text: normalized,
        warnings,
    })
}

fn detect_and_extract_unknown_document(
    data: &[u8],
    warnings: &mut Vec<String>,
) -> Result<String, String> {
    if looks_like_pdf(data) {
        return extract_pdf_text(data);
    }
    if looks_like_docx(data) {
        return extract_docx_text(data);
    }
    if looks_like_rtf(data) {
        return extract_rtf_text(data);
    }
    if std::str::from_utf8(data).is_ok() {
        return Ok(decode_text_bytes(data, warnings));
    }

    Err("The selected document format is not supported. Provide a PDF, Word document, or plain text file.".into())
}

fn extract_doc_text(data: &[u8], warnings: &mut Vec<String>) -> Result<String, String> {
    if looks_like_docx(data) {
        return extract_docx_text(data);
    }
    if looks_like_pdf(data) {
        return extract_pdf_text(data);
    }
    if looks_like_rtf(data) {
        return extract_rtf_text(data);
    }

    warnings.push(
        "The .doc file was treated as plain text. Save the document as .docx if formatting is important.".into(),
    );
    Ok(decode_text_bytes(data, warnings))
}

fn extract_pdf_text(data: &[u8]) -> Result<String, String> {
    extract_text_from_mem(data)
        .map(|text| text.trim().to_string())
        .map_err(|err| format!("Unable to extract text from the PDF document: {err}"))
}

fn extract_docx_text(data: &[u8]) -> Result<String, String> {
    let package =
        read_docx(data).map_err(|err| format!("Unable to read the DOCX document: {err}"))?;
    let mut segments = Vec::new();

    for child in &package.document.children {
        collect_docx_child_text(child, &mut segments);
    }

    Ok(segments.join("\n"))
}

fn extract_rtf_text(data: &[u8]) -> Result<String, String> {
    let content = String::from_utf8_lossy(data);
    let document = RtfDocument::try_from(content.as_ref())
        .map_err(|err| format!("Unable to parse the RTF document: {err}"))?;
    Ok(document.get_text())
}

fn decode_text_bytes(data: &[u8], warnings: &mut Vec<String>) -> String {
    match std::str::from_utf8(data) {
        Ok(text) => text.to_string(),
        Err(_) => {
            warnings.push(
                "The document contained invalid UTF-8 characters. Some characters were replaced during decoding.".into(),
            );
            String::from_utf8_lossy(data).into_owned()
        }
    }
}

fn normalize_document_text(text: &str) -> String {
    let mut normalized = text.replace('\u{0000}', "");
    normalized = normalized.trim_start_matches('\u{FEFF}').to_string();
    normalized = normalized.replace("\r\n", "\n");
    normalized = normalized.replace('\r', "\n");

    let lines: Vec<&str> = normalized.lines().map(|line| line.trim_end()).collect();
    lines.join("\n").trim().to_string()
}

fn looks_like_pdf(data: &[u8]) -> bool {
    data.starts_with(b"%PDF-")
}

fn looks_like_docx(data: &[u8]) -> bool {
    data.len() > 4 && data.starts_with(b"PK")
}

fn looks_like_rtf(data: &[u8]) -> bool {
    let sample = String::from_utf8_lossy(data);
    sample
        .trim_start_matches(|c: char| c.is_ascii_whitespace())
        .starts_with("{\\rtf")
}

fn collect_docx_child_text(child: &DocumentChild, segments: &mut Vec<String>) {
    match child {
        DocumentChild::Paragraph(paragraph) => {
            if let Some(text) = collect_docx_paragraph_text(paragraph.as_ref(), segments) {
                segments.push(text);
            }
        }
        DocumentChild::Table(table) => collect_docx_table_text(table.as_ref(), segments),
        DocumentChild::StructuredDataTag(tag) => {
            collect_docx_structured_data_tag(tag.as_ref(), segments)
        }
        _ => {}
    }
}

fn collect_docx_paragraph_text(
    paragraph: &Paragraph,
    segments: &mut Vec<String>,
) -> Option<String> {
    let mut buffer = String::new();
    for child in &paragraph.children {
        append_paragraph_child_text(child, &mut buffer, segments);
    }

    let trimmed = buffer.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

fn collect_docx_table_text(table: &Table, segments: &mut Vec<String>) {
    for row in &table.rows {
        let row = match row {
            TableChild::TableRow(row) => row,
        };

        for cell in &row.cells {
            let cell = match cell {
                TableRowChild::TableCell(cell) => cell,
            };

            for content in &cell.children {
                match content {
                    TableCellContent::Paragraph(paragraph) => {
                        if let Some(text) = collect_docx_paragraph_text(paragraph, segments) {
                            segments.push(text);
                        }
                    }
                    TableCellContent::Table(inner) => {
                        collect_docx_table_text(inner, segments);
                    }
                    TableCellContent::StructuredDataTag(tag) => {
                        collect_docx_structured_data_tag(tag.as_ref(), segments);
                    }
                    TableCellContent::TableOfContents(_) => {}
                }
            }
        }
    }
}

fn collect_docx_structured_data_tag(tag: &StructuredDataTag, segments: &mut Vec<String>) {
    let mut buffer = String::new();
    append_structured_data_tag_text(tag, &mut buffer, segments);
    let trimmed = buffer.trim();
    if !trimmed.is_empty() {
        segments.push(trimmed.to_string());
    }
}

fn append_paragraph_child_text(
    child: &ParagraphChild,
    buffer: &mut String,
    segments: &mut Vec<String>,
) {
    match child {
        ParagraphChild::Run(run) => append_run_text(run.as_ref(), buffer),
        ParagraphChild::Insert(insert) => append_insert_text(insert, buffer),
        ParagraphChild::Hyperlink(hyperlink) => {
            for inner in &hyperlink.children {
                append_paragraph_child_text(inner, buffer, segments);
            }
        }
        ParagraphChild::StructuredDataTag(tag) => {
            append_structured_data_tag_text(tag.as_ref(), buffer, segments);
        }
        ParagraphChild::BookmarkStart(_) | ParagraphChild::BookmarkEnd(_) => {}
        ParagraphChild::CommentStart(_) | ParagraphChild::CommentEnd(_) => {}
        ParagraphChild::Delete(_) => {}
        ParagraphChild::PageNum(_) | ParagraphChild::NumPages(_) => {}
    }
}

fn append_insert_text(insert: &Insert, buffer: &mut String) {
    for child in &insert.children {
        match child {
            InsertChild::Run(run) => append_run_text(run.as_ref(), buffer),
            InsertChild::Delete(_) => {}
            InsertChild::CommentStart(_) | InsertChild::CommentEnd(_) => {}
        }
    }
}

fn append_structured_data_tag_text(
    tag: &StructuredDataTag,
    buffer: &mut String,
    segments: &mut Vec<String>,
) {
    for child in &tag.children {
        match child {
            StructuredDataTagChild::Run(run) => append_run_text(run.as_ref(), buffer),
            StructuredDataTagChild::Paragraph(paragraph) => {
                if let Some(text) = collect_docx_paragraph_text(paragraph.as_ref(), segments) {
                    if !buffer.is_empty() && !buffer.ends_with('\n') && !buffer.ends_with(' ') {
                        buffer.push(' ');
                    }
                    buffer.push_str(&text);
                }
            }
            StructuredDataTagChild::Table(table) => {
                collect_docx_table_text(table.as_ref(), segments)
            }
            StructuredDataTagChild::StructuredDataTag(inner) => {
                append_structured_data_tag_text(inner.as_ref(), buffer, segments);
            }
            StructuredDataTagChild::BookmarkStart(_) | StructuredDataTagChild::BookmarkEnd(_) => {}
            StructuredDataTagChild::CommentStart(_) | StructuredDataTagChild::CommentEnd(_) => {}
        }
    }
}

fn append_run_text(run: &Run, buffer: &mut String) {
    for child in &run.children {
        match child {
            RunChild::Text(text) => buffer.push_str(&text.text),
            RunChild::Break(_) => buffer.push('\n'),
            RunChild::Tab(_) | RunChild::PTab(_) => buffer.push('\t'),
            RunChild::Sym(sym) => {
                if let Ok(value) = u32::from_str_radix(&sym.char, 16) {
                    if let Some(ch) = char::from_u32(value) {
                        buffer.push(ch);
                    }
                }
            }
            RunChild::InstrTextString(value) => buffer.push_str(value),
            RunChild::DeleteText(_) => {}
            RunChild::FieldChar(_) => {}
            RunChild::Drawing(_) => {}
            RunChild::Shape(_) => {}
            RunChild::CommentStart(_) | RunChild::CommentEnd(_) => {}
            RunChild::FootnoteReference(_) => {}
            RunChild::Shading(_) => {}
            RunChild::InstrText(_) => {}
            RunChild::DeleteInstrText(_) => {}
        }
    }
}

struct RowEmbeddingContext {
    row_index: usize,
    text: String,
    identifiers: HashMap<String, String>,
}

fn default_progress_phase() -> String {
    "embedding".to_string()
}

#[derive(Debug, Serialize, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
struct EmbeddingProgressUpdate {
    #[serde(default = "default_progress_phase")]
    phase: String,
    #[serde(default)]
    #[serde(skip_serializing_if = "Option::is_none")]
    message: Option<String>,
    #[serde(default)]
    processed_rows: usize,
    #[serde(default)]
    total_rows: usize,
    #[serde(default)]
    #[serde(skip_serializing_if = "Option::is_none")]
    elapsed_seconds: Option<f64>,
    #[serde(default)]
    #[serde(skip_serializing_if = "Option::is_none")]
    estimated_remaining_seconds: Option<f64>,
}

fn emit_faculty_embedding_progress(
    app_handle: &tauri::AppHandle,
    progress: EmbeddingProgressUpdate,
) {
    let _ = app_handle.emit(FACULTY_EMBEDDING_PROGRESS_EVENT, progress);
}

fn emit_embedding_error(app_handle: &tauri::AppHandle, total_rows: usize, message: &str) {
    emit_faculty_embedding_progress(
        app_handle,
        EmbeddingProgressUpdate {
            phase: "error".into(),
            message: Some(message.to_string()),
            processed_rows: 0,
            total_rows,
            elapsed_seconds: None,
            estimated_remaining_seconds: None,
        },
    );
}

#[tauri::command]
async fn update_faculty_embeddings(app_handle: tauri::AppHandle) -> Result<String, String> {
    let result =
        tauri::async_runtime::spawn_blocking(move || perform_faculty_embedding_refresh(app_handle))
            .await
            .map_err(|err| format!("Embedding refresh task failed: {err}"))?;

    Ok(result?)
}

fn perform_faculty_embedding_refresh(app_handle: tauri::AppHandle) -> Result<String, String> {
    let started_at = Instant::now();
    emit_faculty_embedding_progress(
        &app_handle,
        EmbeddingProgressUpdate {
            phase: "starting".into(),
            message: Some("Preparing to refresh faculty embeddings…".into()),
            processed_rows: 0,
            total_rows: 0,
            elapsed_seconds: Some(0.0),
            estimated_remaining_seconds: None,
        },
    );

    let status = build_faculty_dataset_status(&app_handle)?;

    if !status.is_valid {
        let message = status.message.unwrap_or_else(|| {
            "Provide a valid faculty dataset before refreshing embeddings.".into()
        });
        return Err(message);
    }

    let analysis = status.analysis.clone().ok_or_else(|| {
        "Run the faculty dataset analysis before generating embeddings.".to_string()
    })?;

    let dataset_path = dataset_destination(&app_handle)?;
    if !dataset_path.exists() {
        return Err("No faculty dataset is available. Restore or configure the dataset before generating embeddings.".into());
    }

    let (headers, rows) = read_full_spreadsheet(&dataset_path)?;

    if rows.is_empty() {
        return Err("The faculty dataset does not include any data rows to embed.".into());
    }

    emit_faculty_embedding_progress(
        &app_handle,
        EmbeddingProgressUpdate {
            phase: "preparing".into(),
            message: Some(format!(
                "Scanning {row_count} faculty row{plural} for embedding content…",
                row_count = rows.len(),
                plural = if rows.len() == 1 { "" } else { "s" }
            )),
            processed_rows: 0,
            total_rows: rows.len(),
            elapsed_seconds: Some(started_at.elapsed().as_secs_f64()),
            estimated_remaining_seconds: None,
        },
    );

    let header_map = build_header_index_map(&headers);
    let embedding_indexes = indexes_from_labels(&header_map, &analysis.embedding_columns)?;
    let identifier_indexes = indexes_from_labels(&header_map, &analysis.identifier_columns)?;

    if embedding_indexes.is_empty() {
        return Err("No embedding columns were identified for the faculty dataset.".into());
    }

    let mut contexts = Vec::new();
    let mut skipped_due_to_text = 0usize;

    for (row_index, row) in rows.iter().enumerate() {
        let mut text_parts = Vec::new();
        for &index in &embedding_indexes {
            if let Some(value) = row.get(index) {
                let trimmed = value.trim();
                if trimmed.is_empty() {
                    continue;
                }
                text_parts.push(trimmed.to_string());
            }
        }

        if text_parts.is_empty() {
            skipped_due_to_text += 1;
            continue;
        }

        let text = text_parts.join("\n\n");

        let mut identifiers = HashMap::new();
        for &index in &identifier_indexes {
            if let Some(value) = row.get(index) {
                let trimmed = value.trim();
                if trimmed.is_empty() {
                    continue;
                }
                let label = header_label(&headers, index);
                identifiers
                    .entry(label)
                    .or_insert_with(|| trimmed.to_string());
            }
        }

        contexts.push(RowEmbeddingContext {
            row_index,
            text,
            identifiers,
        });
    }

    if contexts.is_empty() {
        return Err("None of the faculty rows include embedding content. Add research interest details before refreshing embeddings.".into());
    }

    let total_contexts = contexts.len();
    emit_faculty_embedding_progress(
        &app_handle,
        EmbeddingProgressUpdate {
            phase: "preparing".into(),
            message: Some(format!(
                "Prepared {total} faculty row{plural} for embedding.",
                total = total_contexts,
                plural = if total_contexts == 1 { "" } else { "s" }
            )),
            processed_rows: 0,
            total_rows: total_contexts,
            elapsed_seconds: Some(started_at.elapsed().as_secs_f64()),
            estimated_remaining_seconds: None,
        },
    );

    let request_payload = EmbeddingRequestPayload {
        model: DEFAULT_EMBEDDING_MODEL.to_string(),
        texts: contexts
            .iter()
            .map(|context| EmbeddingRequestRow {
                id: context.row_index,
                text: context.text.clone(),
            })
            .collect(),
        item_label: Some("faculty row".into()),
        item_label_plural: Some("faculty rows".into()),
    };

    emit_faculty_embedding_progress(
        &app_handle,
        EmbeddingProgressUpdate {
            phase: "embedding".into(),
            message: Some(format!(
                "Starting embeddings for {total} faculty row{plural}…",
                total = total_contexts,
                plural = if total_contexts == 1 { "" } else { "s" }
            )),
            processed_rows: 0,
            total_rows: total_contexts,
            elapsed_seconds: Some(started_at.elapsed().as_secs_f64()),
            estimated_remaining_seconds: None,
        },
    );

    let response = run_embedding_helper(&app_handle, &request_payload)?;

    if response.dimension == 0 || response.rows.is_empty() {
        return Err("The embedding helper returned an empty result. Verify the Python environment can load the PubMedBERT model.".into());
    }

    let EmbeddingResponsePayload {
        model: response_model,
        dimension: response_dimension,
        rows: response_rows,
    } = response;

    let mut embedding_map: HashMap<usize, Vec<f32>> = HashMap::new();
    let helper_row_count = response_rows.len();
    for row in response_rows {
        embedding_map.insert(row.id, row.embedding);
    }

    emit_faculty_embedding_progress(
        &app_handle,
        EmbeddingProgressUpdate {
            phase: "processing-results".into(),
            message: Some("Aligning embeddings with faculty rows…".into()),
            processed_rows: helper_row_count,
            total_rows: total_contexts,
            elapsed_seconds: Some(started_at.elapsed().as_secs_f64()),
            estimated_remaining_seconds: None,
        },
    );

    let mut entries = Vec::new();
    let mut missing_embeddings = 0usize;

    for context in contexts {
        match embedding_map.remove(&context.row_index) {
            Some(embedding) => {
                entries.push(FacultyEmbeddingEntry {
                    row_index: context.row_index,
                    identifiers: context.identifiers,
                    embedding,
                });
            }
            None => {
                missing_embeddings += 1;
            }
        }
    }

    if entries.is_empty() {
        return Err("No embeddings were generated for the faculty dataset. Confirm the Python helper executed successfully.".into());
    }

    entries.sort_by_key(|entry| entry.row_index);

    let total_rows = rows.len();
    let embedded_rows = entries.len();
    let skipped_rows = total_rows.saturating_sub(embedded_rows);

    let index = FacultyEmbeddingIndex {
        model: response_model,
        generated_at: Some(Utc::now().to_rfc3339()),
        dimension: response_dimension,
        total_rows: Some(total_rows),
        embedded_rows: Some(embedded_rows),
        skipped_rows: Some(skipped_rows),
        embedding_columns: analysis.embedding_columns.clone(),
        identifier_columns: analysis.identifier_columns.clone(),
        entries,
    };

    let embeddings_path = dataset_directory(&app_handle)?.join(FACULTY_EMBEDDINGS_NAME);
    ensure_dataset_directory(&embeddings_path)?;

    emit_faculty_embedding_progress(
        &app_handle,
        EmbeddingProgressUpdate {
            phase: "saving".into(),
            message: Some("Saving faculty embedding index…".into()),
            processed_rows: embedded_rows,
            total_rows: total_contexts,
            elapsed_seconds: Some(started_at.elapsed().as_secs_f64()),
            estimated_remaining_seconds: None,
        },
    );

    let json = serde_json::to_string_pretty(&index)
        .map_err(|err| format!("Unable to serialize faculty embeddings: {err}"))?;
    fs::write(&embeddings_path, json)
        .map_err(|err| format!("Unable to write faculty embeddings: {err}"))?;

    let mut message = format!(
        "Generated embeddings for {embedded_rows} faculty row{plural} using {model}.",
        embedded_rows = embedded_rows,
        plural = if embedded_rows == 1 { "" } else { "s" },
        model = index.model
    );

    if skipped_due_to_text + missing_embeddings > 0 {
        message.push_str(&format!(
            " Skipped {count} row{plural} without usable embedding content.",
            count = skipped_due_to_text + missing_embeddings,
            plural = if skipped_due_to_text + missing_embeddings == 1 {
                ""
            } else {
                "s"
            }
        ));
    }

    message.push_str(&format!(
        " Saved the embedding index to {}.",
        embeddings_path.to_string_lossy()
    ));

    emit_faculty_embedding_progress(
        &app_handle,
        EmbeddingProgressUpdate {
            phase: "complete".into(),
            message: Some(message.clone()),
            processed_rows: embedded_rows,
            total_rows: total_contexts,
            elapsed_seconds: Some(started_at.elapsed().as_secs_f64()),
            estimated_remaining_seconds: None,
        },
    );

    Ok(message)
}

fn build_header_index_map(headers: &[String]) -> HashMap<String, usize> {
    let mut map = HashMap::new();
    for (index, header) in headers.iter().enumerate() {
        let label = header.trim();
        let normalized = if label.is_empty() {
            format!("Column {}", index + 1)
        } else {
            label.to_string()
        };
        map.entry(normalized.to_lowercase()).or_insert(index);
    }
    map
}

fn indexes_from_labels(
    header_map: &HashMap<String, usize>,
    labels: &[String],
) -> Result<Vec<usize>, String> {
    let mut indexes = Vec::new();

    for label in labels {
        let key = label.trim().to_lowercase();
        if let Some(&index) = header_map.get(&key) {
            indexes.push(index);
        } else {
            return Err(format!(
                "The column '{label}' is not available in the faculty dataset. Re-run the dataset analysis before refreshing embeddings."
            ));
        }
    }

    indexes.sort_unstable();
    indexes.dedup();
    Ok(indexes)
}

fn indexes_from_spreadsheet_labels(
    header_map: &HashMap<String, usize>,
    labels: &[String],
) -> Result<Vec<usize>, String> {
    let mut indexes = Vec::new();

    for label in labels {
        let key = label.trim().to_lowercase();
        if let Some(&index) = header_map.get(&key) {
            indexes.push(index);
        } else {
            return Err(format!(
                "The column '{label}' is not available in the spreadsheet. Reload the preview and try again."
            ));
        }
    }

    indexes.sort_unstable();
    indexes.dedup();
    Ok(indexes)
}

impl EmbeddingHelperProcess {
    fn spawn(app_handle: &tauri::AppHandle) -> Result<Self, String> {
        let mut child = spawn_python_helper(app_handle)?;
        let stdin = child
            .stdin
            .take()
            .ok_or_else(|| "Unable to access the embedding helper stdin.".to_string())?;
        let stdout = child
            .stdout
            .take()
            .ok_or_else(|| "Unable to access the embedding helper stdout.".to_string())?;
        let stderr = child
            .stderr
            .take()
            .ok_or_else(|| "Unable to access the embedding helper stderr.".to_string())?;

        let progress_total = Arc::new(Mutex::new(None));
        let stderr_buffer = Arc::new(Mutex::new(Vec::new()));

        let (sender, receiver) = mpsc::channel();

        let stdout_sender = sender.clone();
        let stdout_handle = std::thread::spawn(move || {
            let mut reader = BufReader::new(stdout);
            loop {
                let mut line = String::new();
                match reader.read_line(&mut line) {
                    Ok(0) => {
                        let _ = stdout_sender.send(EmbeddingHelperMessage::Terminated(None));
                        break;
                    }
                    Ok(_) => {
                        let trimmed = line.trim_end();
                        if trimmed.is_empty() {
                            continue;
                        }

                        if let Ok(response) =
                            serde_json::from_str::<EmbeddingResponsePayload>(trimmed)
                        {
                            if stdout_sender
                                .send(EmbeddingHelperMessage::Response(response))
                                .is_err()
                            {
                                break;
                            }
                            continue;
                        }

                        match serde_json::from_str::<EmbeddingHelperEnvelope>(trimmed) {
                            Ok(EmbeddingHelperEnvelope::Result { payload }) => {
                                if stdout_sender
                                    .send(EmbeddingHelperMessage::Response(payload))
                                    .is_err()
                                {
                                    break;
                                }
                            }
                            Ok(EmbeddingHelperEnvelope::Error { message }) => {
                                let _ = stdout_sender.send(EmbeddingHelperMessage::Error(message));
                            }
                            Err(err) => {
                                let message = format!(
                                    "Unable to parse embedding helper output: {err}. Raw: {trimmed}"
                                );
                                let _ = stdout_sender.send(EmbeddingHelperMessage::Error(message));
                            }
                        }
                    }
                    Err(err) => {
                        let _ = stdout_sender.send(EmbeddingHelperMessage::Error(format!(
                            "Unable to read embedding helper stdout: {err}"
                        )));
                        break;
                    }
                }
            }
        });

        let progress_total_for_thread = Arc::clone(&progress_total);
        let stderr_buffer_for_thread = Arc::clone(&stderr_buffer);
        let app_handle_for_progress = app_handle.clone();
        let stderr_sender = sender.clone();
        let stderr_handle = std::thread::spawn(move || {
            let mut reader = BufReader::new(stderr);
            loop {
                let mut line = String::new();
                match reader.read_line(&mut line) {
                    Ok(0) => break,
                    Ok(_) => {
                        if let Some(json_str) = line.strip_prefix("PROGRESS ") {
                            match serde_json::from_str::<EmbeddingProgressUpdate>(json_str.trim()) {
                                Ok(mut update) => {
                                    if update.total_rows == 0 {
                                        if let Ok(total) = progress_total_for_thread.lock() {
                                            if let Some(rows) = *total {
                                                update.total_rows = rows;
                                            }
                                        }
                                    }
                                    emit_faculty_embedding_progress(
                                        &app_handle_for_progress,
                                        update,
                                    );
                                }
                                Err(_) => {
                                    if let Ok(mut buffer) = stderr_buffer_for_thread.lock() {
                                        buffer.extend_from_slice(line.as_bytes());
                                    }
                                }
                            }
                        } else if let Ok(mut buffer) = stderr_buffer_for_thread.lock() {
                            buffer.extend_from_slice(line.as_bytes());
                        }
                    }
                    Err(err) => {
                        let _ = stderr_sender.send(EmbeddingHelperMessage::Error(format!(
                            "Unable to read embedding helper stderr: {err}"
                        )));
                        break;
                    }
                }
            }
        });

        Ok(Self {
            child,
            stdin: BufWriter::new(stdin),
            receiver,
            progress_total,
            stderr_buffer,
            stdout_handle: Some(stdout_handle),
            stderr_handle: Some(stderr_handle),
        })
    }

    fn is_running(&mut self) -> bool {
        match self.child.try_wait() {
            Ok(None) => true,
            Ok(Some(_)) => false,
            Err(_) => false,
        }
    }

    fn send_embedding_request(
        &mut self,
        payload: &EmbeddingRequestPayload,
    ) -> Result<EmbeddingResponsePayload, String> {
        if let Ok(mut total) = self.progress_total.lock() {
            *total = Some(payload.texts.len());
        }
        if let Ok(mut buffer) = self.stderr_buffer.lock() {
            buffer.clear();
        }

        let command = EmbeddingHelperCommand {
            command_type: "embed",
            payload,
        };

        let mut data = serde_json::to_vec(&command)
            .map_err(|err| format!("Unable to serialize the embedding request: {err}"))?;
        data.push(b'\n');

        if let Err(err) = self.stdin.write_all(&data) {
            return Err(self.augment_error(format!(
                "Unable to send data to the embedding helper: {err}"
            )));
        }
        if let Err(err) = self.stdin.flush() {
            return Err(
                self.augment_error(format!("Unable to flush embedding helper input: {err}"))
            );
        }

        match self.receiver.recv() {
            Ok(EmbeddingHelperMessage::Response(response)) => {
                if let Ok(mut total) = self.progress_total.lock() {
                    *total = None;
                }
                Ok(response)
            }
            Ok(EmbeddingHelperMessage::Error(message)) => {
                if let Ok(mut total) = self.progress_total.lock() {
                    *total = None;
                }
                Err(self.augment_error(message))
            }
            Ok(EmbeddingHelperMessage::Terminated(message)) => {
                if let Ok(mut total) = self.progress_total.lock() {
                    *total = None;
                }
                Err(self.describe_termination(message))
            }
            Err(_) => {
                if let Ok(mut total) = self.progress_total.lock() {
                    *total = None;
                }
                Err(self.augment_error("Lost communication with the embedding helper.".into()))
            }
        }
    }

    fn send_preload_request(&mut self, model: &str) -> Result<(), String> {
        if let Ok(mut total) = self.progress_total.lock() {
            *total = Some(0);
        }
        if let Ok(mut buffer) = self.stderr_buffer.lock() {
            buffer.clear();
        }

        let mut data = serde_json::to_vec(&serde_json::json!({
            "type": "preload",
            "model": model,
        }))
        .map_err(|err| format!("Unable to serialize the preload request: {err}"))?;
        data.push(b'\n');

        if let Err(err) = self.stdin.write_all(&data) {
            return Err(self.augment_error(format!(
                "Unable to send data to the embedding helper: {err}"
            )));
        }
        if let Err(err) = self.stdin.flush() {
            return Err(
                self.augment_error(format!("Unable to flush embedding helper input: {err}"))
            );
        }

        let result = match self.receiver.recv() {
            Ok(EmbeddingHelperMessage::Response(_)) => Ok(()),
            Ok(EmbeddingHelperMessage::Error(message)) => Err(self.augment_error(message)),
            Ok(EmbeddingHelperMessage::Terminated(message)) => {
                Err(self.describe_termination(message))
            }
            Err(_) => {
                Err(self.augment_error("Lost communication with the embedding helper.".into()))
            }
        };

        if let Ok(mut total) = self.progress_total.lock() {
            *total = None;
        }

        result
    }

    fn augment_error(&mut self, base: String) -> String {
        let mut message = base;

        if let Ok(buffer) = self.stderr_buffer.lock() {
            if !buffer.is_empty() {
                let stderr_text = String::from_utf8_lossy(&buffer);
                let trimmed = stderr_text.trim();
                if !trimmed.is_empty() {
                    message = format!("{message}\n\nEmbedding helper stderr:\n{trimmed}");
                }
            }
        }

        if let Ok(Some(status)) = self.child.try_wait() {
            message = format!(
                "{message}\n\nEmbedding helper exit status: {status}",
                status = status
            );
        }

        message
    }

    fn describe_termination(&mut self, reason: Option<String>) -> String {
        let base = reason.unwrap_or_else(|| "The embedding helper exited unexpectedly.".into());
        self.augment_error(base)
    }

    fn send_shutdown_message(&mut self) {
        if !self.is_running() {
            return;
        }

        let shutdown_command = serde_json::json!({ "type": "shutdown" });
        let mut data = match serde_json::to_vec(&shutdown_command) {
            Ok(bytes) => bytes,
            Err(_) => return,
        };
        data.push(b'\n');

        let _ = self.stdin.write_all(&data);
        let _ = self.stdin.flush();
    }

    fn shutdown(&mut self) {
        self.send_shutdown_message();
        let _ = self.child.wait();

        if let Some(handle) = self.stdout_handle.take() {
            let _ = handle.join();
        }
        if let Some(handle) = self.stderr_handle.take() {
            let _ = handle.join();
        }
    }
}

impl Drop for EmbeddingHelperProcess {
    fn drop(&mut self) {
        self.shutdown();
    }
}

fn run_embedding_helper(
    app_handle: &tauri::AppHandle,
    payload: &EmbeddingRequestPayload,
) -> Result<EmbeddingResponsePayload, String> {
    let total_rows = payload.texts.len();
    let helper_state: tauri::State<EmbeddingHelperHandle> = app_handle.state();

    let result = (|| {
        let mut guard = helper_state
            .process
            .lock()
            .map_err(|_| "Unable to lock embedding helper state.".to_string())?;

        if let Some(process) = guard.as_mut() {
            if !process.is_running() {
                process.shutdown();
                *guard = None;
            }
        }

        if guard.is_none() {
            let process = EmbeddingHelperProcess::spawn(app_handle)?;
            *guard = Some(process);
        }

        let response = guard.as_mut().unwrap().send_embedding_request(payload);

        if response.is_err() {
            if let Some(mut process) = guard.take() {
                process.shutdown();
            }
        }

        response
    })();

    if let Err(err) = &result {
        emit_embedding_error(app_handle, total_rows, err);
    }

    result
}

fn ensure_embedding_helper_spawned(app_handle: &tauri::AppHandle) -> Result<(), String> {
    let helper_state: tauri::State<EmbeddingHelperHandle> = app_handle.state();
    let mut guard = helper_state
        .process
        .lock()
        .map_err(|_| "Unable to lock embedding helper state.".to_string())?;

    if let Some(process) = guard.as_mut() {
        if process.is_running() {
            return Ok(());
        }
        process.shutdown();
        *guard = None;
    }

    let process = EmbeddingHelperProcess::spawn(app_handle)?;
    *guard = Some(process);
    Ok(())
}

fn warm_up_embedding_helper(app_handle: &tauri::AppHandle) -> Result<(), String> {
    ensure_embedding_helper_spawned(app_handle)?;
    let helper_state: tauri::State<EmbeddingHelperHandle> = app_handle.state();
    let mut guard = helper_state
        .process
        .lock()
        .map_err(|_| "Unable to lock embedding helper state.".to_string())?;

    if let Some(process) = guard.as_mut() {
        process.send_preload_request(DEFAULT_EMBEDDING_MODEL)?;
    }

    Ok(())
}

struct BundledPythonRuntime {
    executable: PathBuf,
    root: PathBuf,
}

fn expected_bundled_runtime_root(app_handle: &tauri::AppHandle) -> Option<PathBuf> {
    app_handle.path().resource_dir().ok().map(|resource_dir| {
        resource_dir.join("python").join(format!(
            "{}-{}",
            std::env::consts::OS,
            std::env::consts::ARCH
        ))
    })
}

fn spawn_python_helper(app_handle: &tauri::AppHandle) -> Result<Child, String> {
    let script = PYTHON_EMBEDDING_HELPER;
    let mut attempt_messages: Vec<String> = Vec::new();

    let bundled_runtime = locate_bundled_python_runtime(app_handle)?;

    if let Some(runtime) = bundled_runtime {
        match Command::new(&runtime.executable)
            .arg("-c")
            .arg(script)
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::piped())
            .env("HF_HUB_DISABLE_PROGRESS_BARS", "1")
            .env("TOKENIZERS_PARALLELISM", "false")
            .env("PYTHONUTF8", "1")
            .env("PYTHONUNBUFFERED", "1")
            .env("VIRTUAL_ENV", runtime.root.as_os_str())
            .spawn()
        {
            Ok(child) => return Ok(child),
            Err(err) => {
                attempt_messages.push(format!(
                    "Bundled runtime at {}: {err}.",
                    runtime.executable.display()
                ));
            }
        }
    } else if let Some(expected_root) = expected_bundled_runtime_root(app_handle) {
        attempt_messages.push(format!(
            "Bundled runtime not found at {}.",
            expected_root.display()
        ));
    } else {
        attempt_messages.push(
            "Bundled runtime path could not be determined from the application resources directory.".to_string(),
        );
    }

    let candidates = ["python3", "python"];

    for candidate in candidates {
        match Command::new(candidate)
            .arg("-c")
            .arg(script)
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::piped())
            .env("HF_HUB_DISABLE_PROGRESS_BARS", "1")
            .env("TOKENIZERS_PARALLELISM", "false")
            .env("PYTHONUTF8", "1")
            .env("PYTHONUNBUFFERED", "1")
            .spawn()
        {
            Ok(child) => return Ok(child),
            Err(err) => {
                let message = if err.kind() == std::io::ErrorKind::NotFound {
                    format!("System interpreter '{candidate}' was not found on the PATH.")
                } else {
                    format!("System interpreter '{candidate}': {err}.")
                };
                attempt_messages.push(message);
                if err.kind() == std::io::ErrorKind::NotFound {
                    continue;
                }
            }
        }
    }

    let remediation =
        "Reinstall the application to restore the bundled runtime or install Python with the 'torch' and 'transformers' packages.";

    let attempt_summary = if attempt_messages.is_empty() {
        "No interpreter launch attempts were recorded.".to_string()
    } else {
        attempt_messages
            .iter()
            .map(|msg| format!("- {msg}"))
            .collect::<Vec<_>>()
            .join("\n")
    };

    Err(format!(
        "Unable to launch a Python 3 runtime.\n{}\n{}",
        attempt_summary, remediation
    ))
}

fn locate_bundled_python_runtime(
    app_handle: &tauri::AppHandle,
) -> Result<Option<BundledPythonRuntime>, String> {
    let resource_dir = match app_handle.path().resource_dir() {
        Ok(path) => path,
        Err(_) => return Ok(None),
    };

    let runtime_root = resource_dir.join("python").join(format!(
        "{}-{}",
        std::env::consts::OS,
        std::env::consts::ARCH
    ));

    if !runtime_root.exists() {
        return Ok(None);
    }

    let scripts_dir = if cfg!(target_os = "windows") {
        runtime_root.join("Scripts")
    } else {
        runtime_root.join("bin")
    };

    if !scripts_dir.exists() {
        return Err(format!(
            "The bundled Python runtime is missing the interpreter directory at {}.",
            scripts_dir.display()
        ));
    }

    let candidate_names: &[&str] = if cfg!(target_os = "windows") {
        &["python.exe", "python"]
    } else {
        &["python3", "python"]
    };

    for name in candidate_names {
        let candidate = scripts_dir.join(name);
        if candidate.exists() {
            return Ok(Some(BundledPythonRuntime {
                executable: candidate,
                root: runtime_root,
            }));
        }
    }

    Err(format!(
        "The bundled Python runtime at {} does not contain a Python interpreter.",
        runtime_root.display()
    ))
}

const PYTHON_EMBEDDING_HELPER: &str = include_str!("../../python/embedding_helper.py");

#[tauri::command]
fn get_faculty_dataset_status(
    app_handle: tauri::AppHandle,
) -> Result<FacultyDatasetStatus, String> {
    build_faculty_dataset_status(&app_handle)
}

#[tauri::command]
fn preview_faculty_roster(
    app_handle: tauri::AppHandle,
    path: String,
) -> Result<FacultyRosterPreviewResponse, String> {
    let trimmed = path.trim();
    if trimmed.is_empty() {
        return Err("Provide a TSV, TXT, or Excel file to analyze for the faculty roster.".into());
    }

    let source = resolve_existing_path(Some(trimmed.to_string()), false, "Faculty roster file")?;
    let (mut headers, mut rows) = read_spreadsheet_with_limit(&source, Some(10))?;
    align_row_lengths(&mut headers, &mut rows);

    let metadata = load_faculty_dataset_metadata(&app_handle)?.ok_or_else(|| {
        "The faculty dataset metadata is unavailable. Refresh the dataset analysis before selecting a roster.".to_string()
    })?;

    let mut warnings = Vec::new();
    let mut suggestions: HashMap<String, Option<usize>> = HashMap::new();

    if metadata.analysis.identifier_columns.is_empty() {
        warnings.push("No identifier columns are defined in the active faculty dataset.".into());
    }

    let header_map = build_header_index_map(&headers);

    for identifier in &metadata.analysis.identifier_columns {
        let normalized_identifier = identifier.trim().to_lowercase();
        let mut column_index = header_map.get(&normalized_identifier).copied();

        if column_index.is_none() {
            let normalized_target = normalize_identifier_label(identifier);
            if !normalized_target.is_empty() {
                for (candidate_index, header) in headers.iter().enumerate() {
                    if normalize_identifier_label(header) == normalized_target {
                        column_index = Some(candidate_index);
                        break;
                    }
                }
            }
        }

        if let Some(index) = column_index {
            suggestions.insert(identifier.clone(), Some(index));
        } else {
            warnings.push(format!(
                "No roster column matched the faculty identifier '{identifier}'.",
            ));
            suggestions.insert(identifier.clone(), None);
        }
    }

    let preview = SpreadsheetPreview {
        headers,
        rows,
        suggested_prompt_columns: Vec::new(),
        suggested_identifier_columns: Vec::new(),
    };

    Ok(FacultyRosterPreviewResponse {
        preview,
        suggested_identifier_matches: suggestions,
        warnings,
    })
}

#[tauri::command]
fn preview_faculty_dataset_replacement(
    path: String,
) -> Result<FacultyDatasetPreviewResponse, String> {
    let trimmed = path.trim();
    if trimmed.is_empty() {
        return Err("Select a TSV, TXT, or Excel file to import for the faculty dataset.".into());
    }

    let source = resolve_existing_path(Some(trimmed.to_string()), false, "Faculty dataset file")?;

    let extension = source
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or("")
        .to_lowercase();

    if !FACULTY_DATASET_EXTENSIONS.contains(&extension.as_str()) {
        return Err(
            "Select a tab-delimited .tsv or .txt file, or an Excel workbook (.xlsx or .xls) to replace the faculty dataset.".into(),
        );
    }

    let preview = build_dataset_preview(&source)?;
    let program_columns = suggest_program_columns(&preview.headers, &preview.rows);

    Ok(FacultyDatasetPreviewResponse {
        suggested_embedding_columns: preview.suggested_prompt_columns.clone(),
        suggested_identifier_columns: preview.suggested_identifier_columns.clone(),
        suggested_program_columns: program_columns,
        preview,
    })
}

#[tauri::command]
fn replace_faculty_dataset(
    app_handle: tauri::AppHandle,
    path: String,
    configuration: Option<FacultyDatasetColumnConfiguration>,
) -> Result<FacultyDatasetStatus, String> {
    let trimmed = path.trim();
    if trimmed.is_empty() {
        return Err("Select a TSV, TXT, or Excel file to import for the faculty dataset.".into());
    }

    let source = resolve_existing_path(Some(trimmed.to_string()), false, "Faculty dataset file")?;

    let extension = source
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or("")
        .to_lowercase();

    if !FACULTY_DATASET_EXTENSIONS.contains(&extension.as_str()) {
        return Err(
            "Select a tab-delimited .tsv or .txt file, or an Excel workbook (.xlsx or .xls) to replace the faculty dataset.".into(),
        );
    }

    let destination = dataset_destination_for_extension(&app_handle, &extension)?;
    ensure_dataset_directory(&destination)?;
    if let Some(directory) = destination.parent() {
        remove_other_dataset_variants(directory, &extension)?;
    }
    fs::copy(&source, &destination)
        .map_err(|err| format!("Unable to replace the faculty dataset: {err}"))?;

    write_faculty_dataset_source_path(&app_handle, &source)?;

    let mut status =
        build_faculty_dataset_status_with_overrides(&app_handle, configuration.as_ref())?;
    if status.message.is_none() {
        status.message = Some("Faculty dataset replaced successfully.".into());
        status.message_variant = Some("success".into());
    } else if status.message_variant.is_none() {
        status.message_variant = Some(if status.is_valid {
            "success".into()
        } else {
            "error".into()
        });
    }

    Ok(status)
}

#[tauri::command]
fn restore_default_faculty_dataset(
    app_handle: tauri::AppHandle,
) -> Result<FacultyDatasetStatus, String> {
    let destination =
        dataset_destination_for_extension(&app_handle, FACULTY_DATASET_DEFAULT_EXTENSION)?;
    ensure_default_faculty_dataset(&app_handle, &destination)?;

    let mut status = build_faculty_dataset_status(&app_handle)?;
    if status.message.is_none() {
        status.message = Some("Faculty dataset reset to the packaged default.".into());
        status.message_variant = Some("success".into());
    } else if status.message_variant.is_none() {
        status.message_variant = Some(if status.is_valid {
            "success".into()
        } else {
            "error".into()
        });
    }

    Ok(status)
}

fn ensure_default_faculty_dataset(
    app_handle: &tauri::AppHandle,
    destination: &Path,
) -> Result<(), String> {
    ensure_dataset_directory(destination)?;
    if let Some(directory) = destination.parent() {
        remove_other_dataset_variants(directory, FACULTY_DATASET_DEFAULT_EXTENSION)?;
    }
    fs::write(destination, DEFAULT_FACULTY_DATASET)
        .map_err(|err| format!("Unable to restore the default faculty dataset: {err}"))?;

    let embeddings_path = dataset_directory(app_handle)?.join(FACULTY_EMBEDDINGS_NAME);
    ensure_dataset_directory(&embeddings_path)?;
    fs::write(&embeddings_path, DEFAULT_FACULTY_EMBEDDINGS)
        .map_err(|err| format!("Unable to restore the default faculty embeddings: {err}"))?;

    let _ = clear_faculty_dataset_source_path(app_handle);
    Ok(())
}

#[tauri::command]
fn save_generated_spreadsheet(
    path: String,
    content: String,
    encoding: Option<String>,
) -> Result<(), String> {
    let trimmed = path.trim();
    if trimmed.is_empty() {
        return Err("Select a location to save the generated spreadsheet.".into());
    }

    let destination = PathBuf::from(trimmed);
    if let Some(parent) = destination.parent() {
        if !parent.exists() {
            return Err("The selected directory does not exist.".into());
        }
    }

    let data = match encoding
        .as_deref()
        .map(|value| value.trim().to_ascii_lowercase())
        .as_deref()
    {
        Some("base64") => Base64Engine
            .decode(content.as_bytes())
            .map_err(|err| format!("Unable to decode the generated spreadsheet: {err}"))?,
        _ => content.into_bytes(),
    };

    fs::write(&destination, data)
        .map_err(|err| format!("Unable to save the generated spreadsheet: {err}"))?;

    Ok(())
}

#[tauri::command]
fn analyze_spreadsheet(path: String) -> Result<SpreadsheetPreview, String> {
    if path.trim().is_empty() {
        return Err("Provide a spreadsheet path to analyze.".into());
    }

    let spreadsheet = resolve_existing_path(Some(path), false, "Spreadsheet file")?;
    let (mut headers, mut rows) = read_spreadsheet(&spreadsheet)?;
    align_row_lengths(&mut headers, &mut rows);
    let (prompt_columns, identifier_columns) = suggest_spreadsheet_columns(&headers, &rows);

    Ok(SpreadsheetPreview {
        headers,
        rows,
        suggested_prompt_columns: prompt_columns,
        suggested_identifier_columns: identifier_columns,
    })
}

fn normalize_programs(programs: Vec<String>) -> Vec<String> {
    let mut seen = HashSet::new();
    let mut cleaned = Vec::new();

    for program in programs {
        let trimmed = program.trim();
        if trimmed.is_empty() {
            continue;
        }

        let key = trimmed.to_lowercase();
        if seen.insert(key) {
            cleaned.push(trimmed.to_string());
        }
    }

    cleaned
}

fn normalize_columns(columns: Vec<String>) -> Vec<String> {
    let mut seen = HashSet::new();
    let mut cleaned = Vec::new();

    for column in columns {
        let trimmed = column.trim();
        if trimmed.is_empty() {
            continue;
        }

        let key = trimmed.to_lowercase();
        if seen.insert(key) {
            cleaned.push(trimmed.to_string());
        }
    }

    cleaned
}

fn normalize_identifier_value(value: &str) -> String {
    value
        .split_whitespace()
        .filter(|segment| !segment.is_empty())
        .collect::<Vec<_>>()
        .join(" ")
        .to_lowercase()
}

fn normalize_identifier_label(value: &str) -> String {
    value
        .chars()
        .filter(|ch| ch.is_alphanumeric())
        .collect::<String>()
        .to_lowercase()
}

fn read_spreadsheet(path: &Path) -> Result<(Vec<String>, Vec<Vec<String>>), String> {
    read_spreadsheet_with_limit(path, Some(10))
}

fn read_full_spreadsheet(path: &Path) -> Result<(Vec<String>, Vec<Vec<String>>), String> {
    read_spreadsheet_with_limit(path, None)
}

fn read_spreadsheet_with_limit(
    path: &Path,
    max_rows: Option<usize>,
) -> Result<(Vec<String>, Vec<Vec<String>>), String> {
    let extension = path
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or("")
        .to_lowercase();

    if matches!(extension.as_str(), "xlsx" | "xlsm" | "xls" | "xlsb") {
        read_excel_spreadsheet_with_limit(path, max_rows)
    } else {
        read_delimited_spreadsheet_with_limit(path, max_rows)
    }
}

fn read_delimited_spreadsheet_with_limit(
    path: &Path,
    max_rows: Option<usize>,
) -> Result<(Vec<String>, Vec<Vec<String>>), String> {
    let delimiter = detect_delimiter(path)?;
    let mut reader = csv::ReaderBuilder::new()
        .delimiter(delimiter)
        .has_headers(true)
        .flexible(true)
        .from_path(path)
        .map_err(|err| format!("Unable to open the spreadsheet: {err}"))?;

    let mut headers: Vec<String> = reader
        .headers()
        .map_err(|err| format!("Unable to read spreadsheet headers: {err}"))?
        .iter()
        .map(|value| value.trim().to_string())
        .collect();

    let mut rows = Vec::new();
    for record in reader.records() {
        let record = record.map_err(|err| format!("Unable to read spreadsheet rows: {err}"))?;
        let values: Vec<String> = record
            .iter()
            .map(|value| value.trim().to_string())
            .collect();
        if values.iter().all(|value| value.is_empty()) {
            continue;
        }
        rows.push(values);
        if let Some(limit) = max_rows {
            if rows.len() >= limit {
                break;
            }
        }
    }

    align_row_lengths(&mut headers, &mut rows);
    Ok((headers, rows))
}

fn read_excel_spreadsheet_with_limit(
    path: &Path,
    max_rows: Option<usize>,
) -> Result<(Vec<String>, Vec<Vec<String>>), String> {
    let mut workbook =
        open_workbook_auto(path).map_err(|err| format!("Unable to open the spreadsheet: {err}"))?;

    let sheet_name = workbook
        .sheet_names()
        .get(0)
        .cloned()
        .ok_or_else(|| "The workbook does not contain any worksheets.".to_string())?;

    let range = workbook
        .worksheet_range(&sheet_name)
        .ok_or_else(|| format!("Unable to read the worksheet named '{sheet_name}'."))?
        .map_err(|err| format!("Unable to read the worksheet data: {err}"))?;

    let mut rows_iter = range.rows();
    let header_row = rows_iter
        .next()
        .ok_or_else(|| "The worksheet is empty.".to_string())?;

    let mut headers: Vec<String> = header_row.iter().map(cell_to_string).collect();
    let mut rows = Vec::new();

    for row in rows_iter {
        let values: Vec<String> = row.iter().map(cell_to_string).collect();
        if values.iter().all(|value| value.is_empty()) {
            continue;
        }
        rows.push(values);
        if let Some(limit) = max_rows {
            if rows.len() >= limit {
                break;
            }
        }
    }

    align_row_lengths(&mut headers, &mut rows);
    Ok((headers, rows))
}

fn cell_to_string(cell: &DataType) -> String {
    match cell {
        DataType::Empty => String::new(),
        _ => cell.to_string().trim().to_string(),
    }
}

fn align_row_lengths(headers: &mut Vec<String>, rows: &mut Vec<Vec<String>>) {
    let mut column_count = headers.len();
    for row in rows.iter() {
        if row.len() > column_count {
            column_count = row.len();
        }
    }

    if headers.len() < column_count {
        headers.resize(column_count, String::new());
    }

    for row in rows.iter_mut() {
        if row.len() < column_count {
            row.resize(column_count, String::new());
        } else if row.len() > column_count {
            row.truncate(column_count);
        }
    }
}

fn detect_delimiter(path: &Path) -> Result<u8, String> {
    let file = File::open(path).map_err(|err| format!("Unable to open the spreadsheet: {err}"))?;
    let mut reader = BufReader::new(file);
    let mut buffer = String::new();

    for _ in 0..5 {
        buffer.clear();
        let bytes_read = reader
            .read_line(&mut buffer)
            .map_err(|err| format!("Unable to inspect the spreadsheet: {err}"))?;
        if bytes_read == 0 {
            break;
        }
        if buffer.trim().is_empty() {
            continue;
        }

        let counts = [
            (b'\t', buffer.matches('\t').count()),
            (b',', buffer.matches(',').count()),
            (b';', buffer.matches(';').count()),
        ];

        if let Some((delimiter, count)) = counts.iter().max_by_key(|(_, count)| *count) {
            if *count > 0 {
                return Ok(*delimiter);
            }
        }
    }

    Ok(b'\t')
}

fn build_faculty_dataset_status(
    app_handle: &tauri::AppHandle,
) -> Result<FacultyDatasetStatus, String> {
    build_faculty_dataset_status_with_overrides(app_handle, None)
}

fn build_faculty_dataset_status_with_overrides(
    app_handle: &tauri::AppHandle,
    overrides: Option<&FacultyDatasetColumnConfiguration>,
) -> Result<FacultyDatasetStatus, String> {
    let dataset_path = dataset_destination(app_handle)?;
    let mut status = FacultyDatasetStatus {
        path: Some(dataset_path.to_string_lossy().into_owned()),
        canonical_path: None,
        source_path: None,
        last_modified: None,
        row_count: None,
        column_count: None,
        is_valid: false,
        is_default: false,
        message: None,
        message_variant: None,
        preview: None,
        analysis: None,
    };

    if !dataset_path.exists() {
        if let Err(init_err) = ensure_default_faculty_dataset(app_handle, &dataset_path) {
            let _ = clear_faculty_dataset_metadata(app_handle);
            let _ = clear_faculty_dataset_source_path(app_handle);
            status.message = Some(format!(
                "Unable to restore the packaged faculty dataset: {init_err}"
            ));
            status.message_variant = Some("error".into());
            return Ok(status);
        }
    }

    if !dataset_path.exists() {
        let _ = clear_faculty_dataset_metadata(app_handle);
        let _ = clear_faculty_dataset_source_path(app_handle);
        status.message = Some(
            "No faculty dataset has been configured. Restore the default file to continue.".into(),
        );
        status.message_variant = Some("info".into());
        return Ok(status);
    }

    status.canonical_path = dataset_path
        .canonicalize()
        .ok()
        .map(|path| path.to_string_lossy().into_owned());

    match read_faculty_dataset_source_path(app_handle) {
        Ok(source_path) => {
            status.source_path = source_path;
        }
        Err(err) => {
            if status.message.is_none() {
                status.message = Some(err);
                status.message_variant = Some("error".into());
            }
        }
    }

    let metadata = fs::metadata(&dataset_path)
        .map_err(|err| format!("Unable to inspect the faculty dataset: {err}"))?;
    status.last_modified = metadata.modified().ok().map(format_system_time);

    let bytes = fs::read(&dataset_path)
        .map_err(|err| format!("Unable to read the faculty dataset: {err}"))?;
    status.is_default = bytes == DEFAULT_FACULTY_DATASET;

    let extension = dataset_path
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or("")
        .to_lowercase();

    let dimensions = match extension.as_str() {
        "xlsx" | "xls" => compute_excel_dimensions(&dataset_path),
        _ => compute_tsv_dimensions(&bytes),
    };

    match dimensions {
        Ok((rows, columns)) => {
            status.row_count = Some(rows);
            status.column_count = Some(columns);
            status.is_valid = rows > 0 && columns > 0;
            if !status.is_valid {
                status.message = Some("The faculty dataset does not contain any data rows.".into());
                status.message_variant = Some("error".into());
            }
        }
        Err(err) => {
            status.message = Some(err);
            status.message_variant = Some("error".into());
        }
    }

    match build_dataset_preview(&dataset_path) {
        Ok(preview) => {
            status.preview = Some(preview);
        }
        Err(err) => {
            if status.message.is_none() {
                status.message = Some(err);
                status.message_variant = Some("error".into());
            }
        }
    }

    if status.is_valid {
        match analyze_faculty_dataset(app_handle, &dataset_path, overrides) {
            Ok(analysis) => {
                status.analysis = Some(analysis);
            }
            Err(err) => {
                let _ = clear_faculty_dataset_metadata(app_handle);
                status.analysis = None;
                status.is_valid = false;
                if status.message.is_none() {
                    status.message = Some(err);
                    status.message_variant = Some("error".into());
                }
            }
        }
    } else if status.analysis.is_some() {
        status.analysis = None;
    }

    if !status.is_valid {
        if let Err(err) = clear_faculty_dataset_metadata(app_handle) {
            if status.message.is_none() {
                status.message = Some(err);
                status.message_variant = Some("error".into());
            }
        }
    }

    Ok(status)
}

fn analyze_faculty_dataset(
    app_handle: &tauri::AppHandle,
    dataset_path: &Path,
    overrides: Option<&FacultyDatasetColumnConfiguration>,
) -> Result<FacultyDatasetAnalysis, String> {
    let (mut headers, mut rows) = read_full_spreadsheet(dataset_path)?;
    if headers.is_empty() {
        return Err("The faculty dataset does not include any columns.".into());
    }

    align_row_lengths(&mut headers, &mut rows);

    let column_count = headers.len();

    let (embedding_indexes, identifier_indexes) = if let Some(config) = overrides {
        (
            normalize_column_selection(&config.embedding_columns, column_count),
            normalize_column_selection(&config.identifier_columns, column_count),
        )
    } else {
        suggest_spreadsheet_columns(&headers, &rows)
    };

    let program_indexes = if let Some(config) = overrides {
        normalize_column_selection(&config.program_columns, column_count)
    } else {
        suggest_program_columns(&headers, &rows)
    };

    if embedding_indexes.is_empty() {
        return Err(
            "Select at least one column containing faculty research interests or other embedding content.".into(),
        );
    }

    if identifier_indexes.is_empty() {
        return Err(
            "Select at least one column that uniquely identifies each faculty member.".into(),
        );
    }

    let analysis = FacultyDatasetAnalysis {
        embedding_columns: indexes_to_headers(&headers, &embedding_indexes),
        identifier_columns: indexes_to_headers(&headers, &identifier_indexes),
        program_columns: indexes_to_headers(&headers, &program_indexes),
        available_programs: collect_program_values(&rows, &program_indexes),
    };

    let memberships =
        build_faculty_program_memberships(&headers, &rows, &identifier_indexes, &program_indexes);

    write_faculty_dataset_metadata(app_handle, &analysis, &memberships)?;

    Ok(analysis)
}

fn suggest_program_columns(headers: &[String], rows: &[Vec<String>]) -> Vec<usize> {
    const PROGRAM_KEYWORDS: &[&str] = &["program", "track", "pathway", "division", "department"];

    let mut program_columns: Vec<usize> = headers
        .iter()
        .enumerate()
        .filter_map(|(index, header)| {
            let lower = header.to_lowercase();
            if lower.is_empty() {
                return None;
            }
            if PROGRAM_KEYWORDS
                .iter()
                .any(|keyword| lower.contains(keyword))
            {
                Some(index)
            } else {
                None
            }
        })
        .collect();

    sort_and_dedup(&mut program_columns);

    if !program_columns.is_empty() {
        return program_columns;
    }

    let mut candidates: Vec<(usize, usize, usize)> = Vec::new();

    for (index, header) in headers.iter().enumerate() {
        if header.trim().is_empty() {
            continue;
        }

        let mut unique = BTreeSet::new();
        let mut non_empty = 0usize;

        for row in rows {
            if let Some(value) = row.get(index) {
                let trimmed = value.trim();
                if trimmed.is_empty() {
                    continue;
                }
                non_empty += 1;
                unique.insert(trimmed.to_lowercase());
            }
        }

        if non_empty == 0 {
            continue;
        }

        let unique_count = unique.len();
        if unique_count > 0 && unique_count <= 25 {
            candidates.push((index, unique_count, non_empty));
        }
    }

    candidates.sort_by(|a, b| a.1.cmp(&b.1).then(b.2.cmp(&a.2)));

    for (index, _, _) in candidates.into_iter().take(4) {
        program_columns.push(index);
    }

    sort_and_dedup(&mut program_columns);
    program_columns
}

fn indexes_to_headers(headers: &[String], indexes: &[usize]) -> Vec<String> {
    let mut values = Vec::new();
    let mut seen = HashSet::new();

    for &index in indexes {
        if index >= headers.len() {
            continue;
        }

        let label = header_label(headers, index);
        let key = label.to_lowercase();
        if seen.insert(key) {
            values.push(label);
        }
    }

    values
}

fn normalize_column_selection(indexes: &[usize], column_count: usize) -> Vec<usize> {
    let mut normalized = Vec::new();
    let mut seen = HashSet::new();

    for &index in indexes {
        if index >= column_count {
            continue;
        }

        if seen.insert(index) {
            normalized.push(index);
        }
    }

    normalized
}

fn header_label(headers: &[String], index: usize) -> String {
    headers
        .get(index)
        .map(|value| value.trim())
        .filter(|value| !value.is_empty())
        .map(|value| value.to_string())
        .unwrap_or_else(|| format!("Column {}", index + 1))
}

fn collect_program_values(rows: &[Vec<String>], program_indexes: &[usize]) -> Vec<String> {
    let mut seen = HashSet::new();
    let mut values = Vec::new();

    for row in rows {
        for &index in program_indexes {
            if let Some(value) = row.get(index) {
                let trimmed = value.trim();
                if trimmed.is_empty() {
                    continue;
                }
                let key = trimmed.to_lowercase();
                if seen.insert(key) {
                    values.push(trimmed.to_string());
                }
            }
        }
    }

    values.sort_by(|a, b| a.to_lowercase().cmp(&b.to_lowercase()));
    values
}

fn build_faculty_program_memberships(
    headers: &[String],
    rows: &[Vec<String>],
    identifier_indexes: &[usize],
    program_indexes: &[usize],
) -> Vec<FacultyProgramMembership> {
    let mut memberships = Vec::new();

    for (row_index, row) in rows.iter().enumerate() {
        let mut identifiers = HashMap::new();
        for &index in identifier_indexes {
            if let Some(value) = row.get(index) {
                let trimmed = value.trim();
                if trimmed.is_empty() {
                    continue;
                }
                let label = header_label(headers, index);
                identifiers
                    .entry(label)
                    .or_insert_with(|| trimmed.to_string());
            }
        }

        let mut program_set = BTreeSet::new();
        for &index in program_indexes {
            if let Some(value) = row.get(index) {
                let trimmed = value.trim();
                if trimmed.is_empty() {
                    continue;
                }
                program_set.insert(trimmed.to_string());
            }
        }

        if identifiers.is_empty() && program_set.is_empty() {
            continue;
        }

        memberships.push(FacultyProgramMembership {
            row_index,
            identifiers,
            programs: program_set.into_iter().collect(),
        });
    }

    memberships
}

fn metadata_path(app_handle: &tauri::AppHandle) -> Result<PathBuf, String> {
    let directory = dataset_directory(app_handle)?;
    Ok(directory.join(FACULTY_DATASET_METADATA_NAME))
}

fn write_faculty_dataset_metadata(
    app_handle: &tauri::AppHandle,
    analysis: &FacultyDatasetAnalysis,
    memberships: &[FacultyProgramMembership],
) -> Result<(), String> {
    let path = metadata_path(app_handle)?;
    ensure_dataset_directory(&path)?;
    let payload = FacultyDatasetMetadata {
        analysis: analysis.clone(),
        memberships: memberships.to_vec(),
    };
    let json = serde_json::to_string_pretty(&payload)
        .map_err(|err| format!("Unable to serialize faculty dataset metadata: {err}"))?;
    fs::write(&path, json)
        .map_err(|err| format!("Unable to persist faculty dataset metadata: {err}"))?;
    Ok(())
}

fn load_faculty_dataset_metadata(
    app_handle: &tauri::AppHandle,
) -> Result<Option<FacultyDatasetMetadata>, String> {
    let path = metadata_path(app_handle)?;
    if !path.exists() {
        return Ok(None);
    }

    let data = fs::read(&path)
        .map_err(|err| format!("Unable to read the faculty dataset metadata: {err}"))?;
    if data.is_empty() {
        return Ok(None);
    }

    let metadata = serde_json::from_slice(&data)
        .map_err(|err| format!("Unable to parse the faculty dataset metadata: {err}"))?;
    Ok(Some(metadata))
}

fn filter_faculty_rows_by_program(
    memberships: &[FacultyProgramMembership],
    programs: &[String],
) -> HashSet<usize> {
    if programs.is_empty() {
        return HashSet::new();
    }

    let normalized_filters: HashSet<String> = programs
        .iter()
        .map(|program| program.to_lowercase())
        .collect();

    let mut allowed_rows = HashSet::new();
    for membership in memberships {
        for program in &membership.programs {
            let normalized_program = program.to_lowercase();
            if normalized_filters.contains(&normalized_program) {
                allowed_rows.insert(membership.row_index);
                break;
            }
        }
    }

    allowed_rows
}

fn clear_faculty_dataset_metadata(app_handle: &tauri::AppHandle) -> Result<(), String> {
    let path = metadata_path(app_handle)?;
    if path.exists() {
        fs::remove_file(&path)
            .map_err(|err| format!("Unable to clear faculty dataset metadata: {err}"))?;
    }
    Ok(())
}

fn dataset_source_record_path(app_handle: &tauri::AppHandle) -> Result<PathBuf, String> {
    let directory = dataset_directory(app_handle)?;
    Ok(directory.join(FACULTY_DATASET_SOURCE_NAME))
}

fn write_faculty_dataset_source_path(
    app_handle: &tauri::AppHandle,
    source: &Path,
) -> Result<(), String> {
    let record_path = dataset_source_record_path(app_handle)?;
    ensure_dataset_directory(&record_path)?;
    let canonical = source
        .canonicalize()
        .unwrap_or_else(|_| source.to_path_buf())
        .to_string_lossy()
        .into_owned();
    fs::write(&record_path, canonical)
        .map_err(|err| format!("Unable to record the faculty dataset source path: {err}"))?;
    Ok(())
}

fn read_faculty_dataset_source_path(
    app_handle: &tauri::AppHandle,
) -> Result<Option<String>, String> {
    let record_path = dataset_source_record_path(app_handle)?;
    if !record_path.exists() {
        return Ok(None);
    }

    let contents = fs::read_to_string(&record_path)
        .map_err(|err| format!("Unable to read the faculty dataset source path: {err}"))?;
    let trimmed = contents.trim();
    if trimmed.is_empty() {
        Ok(None)
    } else {
        Ok(Some(trimmed.to_string()))
    }
}

fn clear_faculty_dataset_source_path(app_handle: &tauri::AppHandle) -> Result<(), String> {
    let record_path = dataset_source_record_path(app_handle)?;
    if record_path.exists() {
        fs::remove_file(&record_path)
            .map_err(|err| format!("Unable to clear the faculty dataset source path: {err}"))?;
    }
    Ok(())
}

fn dataset_destination(app_handle: &tauri::AppHandle) -> Result<PathBuf, String> {
    let directory = dataset_directory(app_handle)?;
    for extension in FACULTY_DATASET_EXTENSIONS {
        let candidate = dataset_path_with_extension(&directory, extension);
        if candidate.exists() {
            return Ok(candidate);
        }
    }

    Ok(dataset_path_with_extension(
        &directory,
        FACULTY_DATASET_DEFAULT_EXTENSION,
    ))
}

fn dataset_destination_for_extension(
    app_handle: &tauri::AppHandle,
    extension: &str,
) -> Result<PathBuf, String> {
    let directory = dataset_directory(app_handle)?;
    Ok(dataset_path_with_extension(&directory, extension))
}

fn dataset_directory(app_handle: &tauri::AppHandle) -> Result<PathBuf, String> {
    app_handle
        .path()
        .app_data_dir()
        .map_err(|err| format!("Unable to resolve the application data directory: {err}"))
}

fn dataset_path_with_extension(directory: &Path, extension: &str) -> PathBuf {
    directory.join(format!("{}.{}", FACULTY_DATASET_BASENAME, extension))
}

fn remove_other_dataset_variants(directory: &Path, keep_extension: &str) -> Result<(), String> {
    for extension in FACULTY_DATASET_EXTENSIONS {
        if extension.eq_ignore_ascii_case(keep_extension) {
            continue;
        }

        let candidate = dataset_path_with_extension(directory, extension);
        if candidate.exists() {
            fs::remove_file(&candidate)
                .map_err(|err| format!("Unable to remove a previous faculty dataset: {err}"))?;
        }
    }

    Ok(())
}

fn ensure_dataset_directory(path: &Path) -> Result<(), String> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)
            .map_err(|err| format!("Unable to prepare the faculty dataset directory: {err}"))?;
    }
    Ok(())
}

fn compute_tsv_dimensions(data: &[u8]) -> Result<(usize, usize), String> {
    let mut reader = csv::ReaderBuilder::new()
        .delimiter(b'\t')
        .has_headers(true)
        .flexible(true)
        .from_reader(Cursor::new(data));

    let headers = reader
        .headers()
        .map_err(|err| format!("Unable to read the faculty dataset headers: {err}"))?
        .clone();

    let mut row_count = 0usize;
    for record in reader.records() {
        let record =
            record.map_err(|err| format!("Unable to read the faculty dataset rows: {err}"))?;
        if record.iter().any(|value| !value.trim().is_empty()) {
            row_count += 1;
        }
    }

    Ok((row_count, headers.len()))
}

fn compute_excel_dimensions(path: &Path) -> Result<(usize, usize), String> {
    let mut workbook =
        open_workbook_auto(path).map_err(|err| format!("Unable to open the dataset: {err}"))?;

    let sheet_name = workbook
        .sheet_names()
        .get(0)
        .cloned()
        .ok_or_else(|| "The workbook does not contain any worksheets.".to_string())?;

    let range = workbook
        .worksheet_range(&sheet_name)
        .ok_or_else(|| format!("Unable to read the worksheet named '{sheet_name}'."))?
        .map_err(|err| format!("Unable to read the worksheet data: {err}"))?;

    let mut rows_iter = range.rows();
    let header_row = rows_iter
        .next()
        .ok_or_else(|| "The worksheet is empty.".to_string())?;

    let column_count = header_row.len();
    let mut row_count = 0usize;

    for row in rows_iter {
        if row.iter().any(|cell| !cell_to_string(cell).is_empty()) {
            row_count += 1;
        }
    }

    Ok((row_count, column_count))
}

fn build_dataset_preview(path: &Path) -> Result<SpreadsheetPreview, String> {
    let (mut headers, mut rows) = read_spreadsheet(path)?;
    align_row_lengths(&mut headers, &mut rows);
    let (prompt_columns, identifier_columns) = suggest_spreadsheet_columns(&headers, &rows);

    Ok(SpreadsheetPreview {
        headers,
        rows,
        suggested_prompt_columns: prompt_columns,
        suggested_identifier_columns: identifier_columns,
    })
}

fn format_system_time(time: SystemTime) -> String {
    let datetime: DateTime<Utc> = time.into();
    datetime.to_rfc3339()
}

fn suggest_spreadsheet_columns(
    headers: &[String],
    rows: &[Vec<String>],
) -> (Vec<usize>, Vec<usize>) {
    const PROMPT_KEYWORDS: &[&str] = &[
        "prompt",
        "interest",
        "research",
        "description",
        "summary",
        "essay",
        "statement",
        "focus",
        "topic",
        "goal",
    ];
    const IDENTIFIER_KEYWORDS: &[&str] = &[
        "id",
        "identifier",
        "name",
        "first",
        "last",
        "student",
        "email",
        "netid",
        "number",
        "uid",
    ];

    let mut prompt_columns = Vec::new();
    let mut identifier_columns = Vec::new();

    for (index, header) in headers.iter().enumerate() {
        let header_lower = header.to_lowercase();
        if header_lower.is_empty() {
            continue;
        }

        if PROMPT_KEYWORDS
            .iter()
            .any(|keyword| header_lower.contains(keyword))
        {
            prompt_columns.push(index);
        }

        if IDENTIFIER_KEYWORDS
            .iter()
            .any(|keyword| header_lower.contains(keyword))
        {
            identifier_columns.push(index);
        }
    }

    sort_and_dedup(&mut prompt_columns);
    sort_and_dedup(&mut identifier_columns);

    let stats = compute_column_stats(headers, rows);

    if prompt_columns.is_empty() {
        let mut candidates: Vec<&ColumnStats> = stats
            .iter()
            .filter(|stat| stat.non_empty > 0 && stat.numeric_ratio < 0.6)
            .collect();

        candidates.sort_by(|a, b| {
            b.average_length
                .partial_cmp(&a.average_length)
                .unwrap_or(Ordering::Equal)
                .then(b.max_length.cmp(&a.max_length))
        });

        let mut fallback = Vec::new();
        for stat in &candidates {
            if stat.average_length >= 18.0 || stat.max_length >= 60 {
                fallback.push(stat.index);
            }
        }

        if fallback.is_empty() {
            if let Some(best) = candidates.first() {
                fallback.push(best.index);
            }
        }

        prompt_columns = fallback;
    }

    if identifier_columns.is_empty() {
        let mut candidates: Vec<&ColumnStats> =
            stats.iter().filter(|stat| stat.non_empty > 0).collect();

        candidates.sort_by(|a, b| {
            a.average_length
                .partial_cmp(&b.average_length)
                .unwrap_or(Ordering::Equal)
                .then(b.non_empty.cmp(&a.non_empty))
        });

        let mut fallback = Vec::new();
        for stat in &candidates {
            if stat.average_length <= 36.0 || stat.numeric_ratio >= 0.5 {
                fallback.push(stat.index);
            }
            if fallback.len() >= 3 {
                break;
            }
        }

        if fallback.is_empty() {
            if let Some(best) = candidates.first() {
                fallback.push(best.index);
            }
        }

        identifier_columns = fallback;
    }

    if prompt_columns.is_empty() && !headers.is_empty() {
        prompt_columns.push(headers.len() - 1);
    }

    if identifier_columns.is_empty() && !headers.is_empty() {
        identifier_columns.push(0);
    }

    sort_and_dedup(&mut prompt_columns);
    sort_and_dedup(&mut identifier_columns);

    (prompt_columns, identifier_columns)
}

fn sort_and_dedup(values: &mut Vec<usize>) {
    values.sort_unstable();
    values.dedup();
}

struct ColumnStats {
    index: usize,
    non_empty: usize,
    average_length: f64,
    max_length: usize,
    numeric_ratio: f64,
}

fn compute_column_stats(headers: &[String], rows: &[Vec<String>]) -> Vec<ColumnStats> {
    let column_count = headers.len();
    let mut stats = Vec::new();

    for index in 0..column_count {
        let mut non_empty = 0usize;
        let mut total_length = 0usize;
        let mut max_length = 0usize;
        let mut numeric_like = 0usize;

        for row in rows {
            if let Some(value) = row.get(index) {
                let trimmed = value.trim();
                if trimmed.is_empty() {
                    continue;
                }
                non_empty += 1;
                let length = trimmed.chars().count();
                total_length += length;
                if length > max_length {
                    max_length = length;
                }
                if is_numeric_like(trimmed) {
                    numeric_like += 1;
                }
            }
        }

        let average_length = if non_empty > 0 {
            total_length as f64 / non_empty as f64
        } else {
            0.0
        };

        let numeric_ratio = if non_empty > 0 {
            numeric_like as f64 / non_empty as f64
        } else {
            0.0
        };

        stats.push(ColumnStats {
            index,
            non_empty,
            average_length,
            max_length,
            numeric_ratio,
        });
    }

    stats
}

fn is_numeric_like(value: &str) -> bool {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        return false;
    }

    let has_alpha = trimmed.chars().any(|c| c.is_ascii_alphabetic());
    !has_alpha
}

fn resolve_existing_path(
    raw_path: Option<String>,
    expects_directory: bool,
    label: &str,
) -> Result<PathBuf, String> {
    let provided = raw_path.as_deref().map(str::trim).unwrap_or_default();

    if provided.is_empty() {
        return Err(format!("{label} path is required."));
    }

    let path = expand_home(provided);
    let metadata =
        fs::metadata(&path).map_err(|_| format!("{label} was not found: {}", path.display()))?;

    if expects_directory && !metadata.is_dir() {
        return Err(format!(
            "{label} is expected to be a directory: {}",
            path.display()
        ));
    }

    if !expects_directory && !metadata.is_file() {
        return Err(format!(
            "{label} is expected to be a file: {}",
            path.display()
        ));
    }

    Ok(path)
}

fn expand_home(path: &str) -> PathBuf {
    if path == "~" {
        if let Some(home) = home_dir() {
            return home;
        }
        return PathBuf::from(path);
    }

    if let Some(stripped) = path.strip_prefix("~/") {
        if let Some(home) = home_dir() {
            return home.join(stripped);
        }
    }

    PathBuf::from(path)
}

fn home_dir() -> Option<PathBuf> {
    std::env::var("HOME")
        .map(PathBuf::from)
        .or_else(|_| std::env::var("USERPROFILE").map(PathBuf::from))
        .ok()
}

fn validate_extension(path: &Path, allowed: &[&str], label: &str) -> Option<String> {
    match path.extension().and_then(|ext| ext.to_str()) {
        Some(ext) if allowed.iter().any(|value| ext.eq_ignore_ascii_case(value)) => None,
        Some(ext) => Some(format!(
            "The selected {label} uses '.{ext}', which is outside the expected extensions: {}.",
            allowed.join(", ")
        )),
        None => Some(format!(
            "The selected {label} does not include an extension. Confirm it is supported."
        )),
    }
}

fn build_prompt_preview(text: &str) -> String {
    let characters: Vec<char> = text.chars().collect();
    let max = 280usize;
    if characters.len() <= max {
        return text.to_string();
    }

    let mut preview: String = characters[..max].iter().collect();
    preview.push('…');
    preview
}

fn build_summary(
    task_type: &TaskType,
    faculty_scope: &FacultyScope,
    faculty_per_student: u32,
    program_count: usize,
    has_custom_roster: bool,
) -> String {
    let input_summary = match task_type {
        TaskType::Prompt => "a single prompt".to_string(),
        TaskType::Document => "one document".to_string(),
        TaskType::Spreadsheet => "a spreadsheet of prompts".to_string(),
        TaskType::Directory => "a directory of documents".to_string(),
    };

    let scope_summary = match faculty_scope {
        FacultyScope::All => "the complete faculty roster".to_string(),
        FacultyScope::Program => format!(
            "faculty filtered to {program_count} program{}",
            if program_count == 1 { "" } else { "s" }
        ),
        FacultyScope::Custom => {
            if has_custom_roster {
                "the provided faculty roster spreadsheet".to_string()
            } else {
                "a custom faculty roster".to_string()
            }
        }
    };

    let summary = format!(
        "Ready to match {input_summary} against {scope_summary}. Each student will receive up to {faculty_per_student} faculty recommendation{plural}.",
        plural = if faculty_per_student == 1 { "" } else { "s" }
    );

    summary
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .manage(EmbeddingHelperHandle::default())
        .setup(|app| {
            let app_handle = app.handle().clone();
            tauri::async_runtime::spawn_blocking(move || {
                if let Err(err) = warm_up_embedding_helper(&app_handle) {
                    eprintln!("⚠️ Unable to warm up the embedding helper: {err}");
                }
            });
            Ok(())
        })
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            submit_matching_request,
            update_faculty_embeddings,
            analyze_spreadsheet,
            get_faculty_dataset_status,
            preview_faculty_roster,
            preview_faculty_dataset_replacement,
            replace_faculty_dataset,
            restore_default_faculty_dataset,
            save_generated_spreadsheet
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
