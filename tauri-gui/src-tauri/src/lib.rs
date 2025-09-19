use calamine::{open_workbook_auto, DataType, Reader};
use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};
use std::cmp::Ordering;
use std::collections::HashSet;
use std::fs;
use std::fs::File;
use std::io::{BufRead, BufReader, Cursor};
use std::path::{Path, PathBuf};
use std::time::SystemTime;
use tauri::Manager;

const FACULTY_DATASET_BASENAME: &str = "faculty_dataset";
const FACULTY_DATASET_DEFAULT_EXTENSION: &str = "tsv";
const FACULTY_DATASET_EXTENSIONS: &[&str] = &["tsv", "txt", "xlsx", "xls"];
const DEFAULT_FACULTY_DATASET: &[u8] = include_bytes!("../assets/default_faculty_dataset.tsv");

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
    update_embeddings: bool,
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
    update_embeddings: bool,
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
struct FacultyDatasetStatus {
    path: Option<String>,
    canonical_path: Option<String>,
    last_modified: Option<String>,
    row_count: Option<usize>,
    column_count: Option<usize>,
    is_valid: bool,
    is_default: bool,
    message: Option<String>,
    message_variant: Option<String>,
    preview: Option<SpreadsheetPreview>,
}

#[tauri::command]
fn submit_matching_request(payload: SubmissionPayload) -> Result<SubmissionResponse, String> {
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
        update_embeddings,
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

    match task_type {
        TaskType::Prompt => {
            let text = prompt_text.as_deref().map(str::trim).unwrap_or_default();
            if text.is_empty() {
                return Err("Provide a prompt describing the student's interests.".into());
            }
            prompt_preview = Some(build_prompt_preview(text));
        }
        TaskType::Document => {
            let document = resolve_existing_path(document_path, false, "Document file")?;
            if let Some(message) =
                validate_extension(&document, &["txt", "pdf", "doc", "docx"], "document")
            {
                warnings.push(message);
            }
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
        update_embeddings,
        prompt_preview,
        spreadsheet_prompt_columns: selected_prompt_columns.clone(),
        spreadsheet_identifier_columns: selected_identifier_columns.clone(),
    };

    let summary = build_summary(
        &task_type,
        &faculty_scope,
        faculty_recs_per_student,
        student_recs_per_faculty,
        update_embeddings,
        details.program_filters.len(),
        faculty_roster_path.is_some(),
    );

    Ok(SubmissionResponse {
        summary,
        warnings,
        details,
    })
}

#[tauri::command]
fn update_faculty_embeddings() -> Result<String, String> {
    Ok("Faculty embeddings update request received. The backend stub only confirms availability in this build.".into())
}

#[tauri::command]
fn get_faculty_dataset_status(
    app_handle: tauri::AppHandle,
) -> Result<FacultyDatasetStatus, String> {
    build_faculty_dataset_status(&app_handle)
}

#[tauri::command]
fn replace_faculty_dataset(
    app_handle: tauri::AppHandle,
    path: String,
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

    let mut status = build_faculty_dataset_status(&app_handle)?;
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
    let extension = path
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or("")
        .to_lowercase();

    if matches!(extension.as_str(), "xlsx" | "xlsm" | "xls" | "xlsb") {
        read_excel_spreadsheet(path)
    } else {
        read_delimited_spreadsheet(path)
    }
}

fn read_delimited_spreadsheet(path: &Path) -> Result<(Vec<String>, Vec<Vec<String>>), String> {
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
        if rows.len() >= 10 {
            break;
        }
    }

    align_row_lengths(&mut headers, &mut rows);
    Ok((headers, rows))
}

fn read_excel_spreadsheet(path: &Path) -> Result<(Vec<String>, Vec<Vec<String>>), String> {
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
        if rows.len() >= 10 {
            break;
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
    let dataset_path = dataset_destination(app_handle)?;
    let mut status = FacultyDatasetStatus {
        path: Some(dataset_path.to_string_lossy().into_owned()),
        canonical_path: None,
        last_modified: None,
        row_count: None,
        column_count: None,
        is_valid: false,
        is_default: false,
        message: None,
        message_variant: None,
        preview: None,
    };

    if !dataset_path.exists() {
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

    Ok(status)
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
    update_embeddings: bool,
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

    if update_embeddings {
        summary.push_str(" Faculty embeddings will be refreshed before matching.");
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
            replace_faculty_dataset,
            restore_default_faculty_dataset
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
