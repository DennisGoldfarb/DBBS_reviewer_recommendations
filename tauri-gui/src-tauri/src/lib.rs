use calamine::{open_workbook_auto, DataType, Reader};
use chrono::{DateTime, Utc};
use docx_rs::{
    read_docx, DocumentChild, Insert, InsertChild, Paragraph, ParagraphChild, Run, RunChild,
    StructuredDataTag, StructuredDataTagChild, Table, TableCellContent, TableChild, TableRowChild,
};
use pdf_extract::extract_text_from_mem;
use rtf_parser::RtfDocument;
use serde::{Deserialize, Serialize};
use std::char;
use std::cmp::Ordering;
use std::collections::{BTreeSet, HashMap, HashSet};
use std::convert::TryFrom;
use std::fs;
use std::fs::File;
use std::io::{BufRead, BufReader, Cursor, Read, Write};
use std::path::{Path, PathBuf};
use std::process::{Child, Command, Stdio};
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
    student_recs_per_faculty: u32,
    #[serde(default)]
    spreadsheet_prompt_columns: Vec<String>,
    #[serde(default)]
    spreadsheet_identifier_columns: Vec<String>,
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
    recommendations_per_faculty: u32,
    prompt_preview: Option<String>,
    spreadsheet_prompt_columns: Vec<String>,
    spreadsheet_identifier_columns: Vec<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct SubmissionResponse {
    summary: String,
    warnings: Vec<String>,
    details: SubmissionDetails,
    prompt_matches: Vec<PromptMatchResult>,
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
        student_recs_per_faculty,
        spreadsheet_prompt_columns,
        spreadsheet_identifier_columns,
    } = payload;

    if faculty_recs_per_student == 0 {
        return Err("Specify at least one faculty recommendation per student.".into());
    }

    let mut warnings = Vec::new();
    let mut validated_paths = Vec::new();
    let mut prompt_preview = None;
    let mut selected_prompt_columns = Vec::new();
    let mut selected_identifier_columns = Vec::new();
    let mut prepared_prompt_text: Option<String> = None;

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

            selected_prompt_columns = normalize_columns(spreadsheet_prompt_columns);
            selected_identifier_columns = normalize_columns(spreadsheet_identifier_columns);

            if selected_identifier_columns.is_empty() {
                return Err("Select at least one column to identify each student.".into());
            }

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
        recommendations_per_faculty: student_recs_per_faculty,
        prompt_preview,
        spreadsheet_prompt_columns: selected_prompt_columns.clone(),
        spreadsheet_identifier_columns: selected_identifier_columns.clone(),
    };

    let summary = build_summary(
        &task_type,
        &faculty_scope,
        faculty_recs_per_student,
        student_recs_per_faculty,
        details.program_filters.len(),
        faculty_roster_path.is_some(),
    );

    let mut prompt_matches = Vec::new();
    if let Some(prompt_text) = prepared_prompt_text {
        let embedding_index = load_faculty_embedding_index(&app_handle)?;
        if embedding_index.entries.is_empty() {
            return Err(
                "No faculty embeddings are available. Generate embeddings before matching.".into(),
            );
        }

        let limit = faculty_recs_per_student.max(1) as usize;
        let prompt_embedding = embed_prompt(&app_handle, &embedding_index, &prompt_text)?;
        let matches = find_best_faculty_matches(
            &embedding_index,
            &prompt_embedding,
            limit,
            allowed_faculty_rows.as_ref(),
        );

        prompt_matches.push(PromptMatchResult {
            prompt: match task_type {
                TaskType::Document => build_prompt_preview(&prompt_text),
                _ => prompt_text.clone(),
            },
            faculty_matches: matches,
        });
    }

    Ok(SubmissionResponse {
        summary,
        warnings,
        details,
        prompt_matches,
    })
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct EmbeddingRequestPayload {
    model: String,
    texts: Vec<EmbeddingRequestRow>,
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
        let TableChild::TableRow(row) = row;
        let row = row.as_ref();

        for cell in &row.cells {
            let TableRowChild::TableCell(cell) = cell;
            let cell = cell.as_ref();

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

fn run_embedding_helper(
    app_handle: &tauri::AppHandle,
    payload: &EmbeddingRequestPayload,
) -> Result<EmbeddingResponsePayload, String> {
    let total_rows = payload.texts.len();
    let mut child = spawn_python_helper()?;
    let input = serde_json::to_vec(payload)
        .map_err(|err| format!("Unable to serialize the embedding request: {err}"))?;

    let mut stdin = child
        .stdin
        .take()
        .ok_or_else(|| "Unable to access the embedding helper stdin.".to_string())?;
    stdin
        .write_all(&input)
        .map_err(|err| format!("Unable to send data to the embedding helper: {err}"))?;
    stdin
        .flush()
        .map_err(|err| format!("Unable to flush embedding helper input: {err}"))?;
    drop(stdin);

    let stdout = child
        .stdout
        .take()
        .ok_or_else(|| "Unable to access the embedding helper stdout.".to_string())?;
    let stderr = child
        .stderr
        .take()
        .ok_or_else(|| "Unable to access the embedding helper stderr.".to_string())?;

    let stdout_handle = std::thread::spawn(move || -> Result<Vec<u8>, String> {
        let mut buffer = Vec::new();
        let mut reader = BufReader::new(stdout);
        reader
            .read_to_end(&mut buffer)
            .map_err(|err| format!("Unable to read embedding helper stdout: {err}"))?;
        Ok(buffer)
    });

    let app_handle_for_progress = app_handle.clone();
    let total_rows_for_progress = total_rows;
    let stderr_handle = std::thread::spawn(move || -> Result<Vec<u8>, String> {
        let mut reader = BufReader::new(stderr);
        let mut buffer = Vec::new();

        loop {
            let mut line = String::new();
            match reader.read_line(&mut line) {
                Ok(0) => break,
                Ok(_) => {
                    if let Some(json_str) = line.strip_prefix("PROGRESS ") {
                        match serde_json::from_str::<EmbeddingProgressUpdate>(json_str.trim()) {
                            Ok(mut update) => {
                                if update.total_rows == 0 {
                                    update.total_rows = total_rows_for_progress;
                                }
                                emit_faculty_embedding_progress(&app_handle_for_progress, update);
                            }
                            Err(_) => {
                                buffer.extend_from_slice(line.as_bytes());
                            }
                        }
                    } else {
                        buffer.extend_from_slice(line.as_bytes());
                    }
                }
                Err(err) => {
                    return Err(format!("Unable to read embedding helper stderr: {err}"));
                }
            }
        }

        Ok(buffer)
    });

    let status = child
        .wait()
        .map_err(|err| format!("Unable to wait for the embedding helper: {err}"))?;

    let stdout_bytes = match stdout_handle.join() {
        Ok(result) => result?,
        Err(_) => {
            return Err("Unable to join the embedding helper stdout reader.".into());
        }
    };

    let stderr_bytes = match stderr_handle.join() {
        Ok(result) => result?,
        Err(_) => {
            return Err("Unable to join the embedding helper stderr reader.".into());
        }
    };

    if !status.success() {
        let message = String::from_utf8_lossy(&stderr_bytes);
        let trimmed = message.trim();
        let error_message = if trimmed.is_empty() {
            "The embedding helper exited with an error.".to_string()
        } else {
            trimmed.to_string()
        };
        emit_embedding_error(app_handle, total_rows, &error_message);
        return Err(error_message);
    }

    if stdout_bytes.is_empty() {
        let stderr_message = String::from_utf8_lossy(&stderr_bytes);
        let trimmed = stderr_message.trim();
        let error_message = if trimmed.is_empty() {
            "The embedding helper did not produce any output.".to_string()
        } else {
            trimmed.to_string()
        };
        emit_embedding_error(app_handle, total_rows, &error_message);
        return Err(error_message);
    }

    match serde_json::from_slice(&stdout_bytes) {
        Ok(response) => Ok(response),
        Err(err) => {
            let stderr_message = String::from_utf8_lossy(&stderr_bytes);
            let trimmed = stderr_message.trim();
            let error_message = if trimmed.is_empty() {
                format!("Unable to parse embedding helper output: {err}")
            } else {
                format!(
                    "Unable to parse embedding helper output: {err}. Details: {}",
                    trimmed
                )
            };
            emit_embedding_error(app_handle, total_rows, &error_message);
            Err(error_message)
        }
    }
}

fn spawn_python_helper() -> Result<Child, String> {
    let script = PYTHON_EMBEDDING_HELPER;
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
            .spawn()
        {
            Ok(child) => return Ok(child),
            Err(err) => {
                if err.kind() == std::io::ErrorKind::NotFound {
                    continue;
                }
                return Err(format!("Unable to launch {candidate}: {err}"));
            }
        }
    }

    Err(
        "Python 3 is required to generate embeddings. Install Python along with the 'torch' and 'transformers' packages.".into(),
    )
}

const PYTHON_EMBEDDING_HELPER: &str = r#"
import json
import os
import sys
import time

os.environ.setdefault("TOKENIZERS_PARALLELISM", "false")

try:
    from transformers import AutoModel, AutoTokenizer
    from transformers.utils import logging as hf_logging
except ImportError as exc:
    sys.stderr.write("Install the 'transformers' package (with torch) to generate embeddings.\n")
    sys.stderr.write(str(exc) + "\n")
    sys.exit(1)

try:
    import torch
except ImportError as exc:
    sys.stderr.write("Install the 'torch' package to generate embeddings.\n")
    sys.stderr.write(str(exc) + "\n")
    sys.exit(1)

hf_logging.set_verbosity_error()


def emit_progress(payload: dict) -> None:
    sys.stderr.write("PROGRESS " + json.dumps(payload) + "\n")
    sys.stderr.flush()


def main():
    try:
        payload = json.load(sys.stdin)
    except Exception as exc:  # noqa: BLE001
        sys.stderr.write(f"Unable to parse embedding request: {exc}\n")
        sys.exit(1)

    model_name = payload.get("model") or "NeuML/pubmedbert-base-embeddings"
    texts = payload.get("texts") or []
    total = len(texts)

    if not texts:
        json.dump({"model": model_name, "dimension": 0, "rows": []}, sys.stdout)
        return

    emit_progress(
        {
            "phase": "loading-model",
            "message": "Loading PubMedBERT model…",
            "processedRows": 0,
            "totalRows": total,
        }
    )

    tokenizer = AutoTokenizer.from_pretrained(model_name)
    model = AutoModel.from_pretrained(model_name)
    model.eval()

    emit_progress(
        {
            "phase": "embedding",
            "message": f"Starting embeddings for {total} faculty rows…",
            "processedRows": 0,
            "totalRows": total,
            "elapsedSeconds": 0.0,
        }
    )

    start_time = time.time()

    rows = []
    for item in texts:
        text = (item.get("text") or "").strip()
        if not text:
            continue

        inputs = tokenizer(
            text,
            return_tensors="pt",
            truncation=True,
            max_length=512,
            padding=True,
        )

        with torch.no_grad():
            outputs = model(**inputs)

        last_hidden = outputs.last_hidden_state
        attention_mask = inputs["attention_mask"]
        mask = attention_mask.unsqueeze(-1).expand(last_hidden.size()).float()
        masked = last_hidden * mask
        summed = masked.sum(dim=1)
        counts = mask.sum(dim=1).clamp(min=1e-9)
        embedding = (summed / counts).squeeze(0)

        rows.append({"id": item.get("id"), "embedding": embedding.tolist()})

        processed = len(rows)
        elapsed = time.time() - start_time
        remaining = None
        if processed < total and processed > 0 and elapsed > 0:
            remaining = (total - processed) * (elapsed / processed)

        emit_progress(
            {
                "phase": "embedding",
                "message": f"Embedded {processed} of {total} faculty rows",
                "processedRows": processed,
                "totalRows": total,
                "elapsedSeconds": elapsed,
                "estimatedRemainingSeconds": remaining,
            }
        )

    result = {
        "model": model_name,
        "dimension": len(rows[0]["embedding"]) if rows else 0,
        "rows": rows,
    }

    emit_progress(
        {
            "phase": "finalizing",
            "message": "Finalizing embedding response…",
            "processedRows": len(rows),
            "totalRows": total,
            "elapsedSeconds": time.time() - start_time,
        }
    )

    json.dump(result, sys.stdout)


if __name__ == "__main__":
    main()
"#;

#[tauri::command]
fn get_faculty_dataset_status(
    app_handle: tauri::AppHandle,
) -> Result<FacultyDatasetStatus, String> {
    build_faculty_dataset_status(&app_handle)
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
    ensure_dataset_directory(&destination)?;
    if let Some(directory) = destination.parent() {
        remove_other_dataset_variants(directory, FACULTY_DATASET_DEFAULT_EXTENSION)?;
    }
    fs::write(&destination, DEFAULT_FACULTY_DATASET)
        .map_err(|err| format!("Unable to restore the default faculty dataset: {err}"))?;

    let _ = clear_faculty_dataset_source_path(&app_handle);

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
    students_per_faculty: u32,
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

    let mut summary = format!(
        "Ready to match {input_summary} against {scope_summary}. Each student will receive up to {faculty_per_student} faculty recommendation{plural}.",
        plural = if faculty_per_student == 1 { "" } else { "s" }
    );

    if students_per_faculty > 0 {
        summary.push_str(&format!(
            " Each faculty member will receive up to {students_per_faculty} student recommendation{plural}.",
            plural = if students_per_faculty == 1 { "" } else { "s" }
        ));
    }

    summary
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            submit_matching_request,
            update_faculty_embeddings,
            analyze_spreadsheet,
            get_faculty_dataset_status,
            preview_faculty_dataset_replacement,
            replace_faculty_dataset,
            restore_default_faculty_dataset
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
