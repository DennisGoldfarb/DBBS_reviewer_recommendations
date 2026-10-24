use crate::config::{
    DEFAULT_FACULTY_DATASET, DEFAULT_FACULTY_EMBEDDINGS, FACULTY_DATASET_BASENAME,
    FACULTY_DATASET_DEFAULT_EXTENSION, FACULTY_DATASET_EXTENSIONS, FACULTY_DATASET_METADATA_NAME,
    FACULTY_DATASET_SOURCE_NAME, FACULTY_EMBEDDINGS_NAME,
};
use crate::models::{
    FacultyDatasetAnalysis, FacultyDatasetColumnConfiguration, FacultyDatasetMetadata,
    FacultyDatasetStatus, FacultyProgramMembership, SpreadsheetPreview,
};
use crate::spreadsheets::{
    align_row_lengths, cell_to_string, read_full_spreadsheet, read_spreadsheet, sort_and_dedup,
    suggest_spreadsheet_columns,
};
use calamine::{open_workbook_auto, Reader};
use chrono::{DateTime, Utc};
use std::collections::BTreeSet;
use std::collections::HashMap;
use std::collections::HashSet;
use std::fs;
use std::io::Cursor;
use std::path::Path;
use std::path::PathBuf;
use std::time::SystemTime;
use tauri::Manager;

pub(crate) fn ensure_default_faculty_dataset(
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

pub(crate) fn build_faculty_dataset_status(
    app_handle: &tauri::AppHandle,
) -> Result<FacultyDatasetStatus, String> {
    build_faculty_dataset_status_with_overrides(app_handle, None)
}

pub(crate) fn build_faculty_dataset_status_with_overrides(
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
    let (analysis, memberships) = inspect_faculty_dataset(dataset_path, overrides)?;
    write_faculty_dataset_metadata(app_handle, &analysis, &memberships)?;
    Ok(analysis)
}

/// Analyze a candidate without changing the active dataset or its metadata.
pub(crate) fn inspect_faculty_dataset(
    dataset_path: &Path,
    overrides: Option<&FacultyDatasetColumnConfiguration>,
) -> Result<(FacultyDatasetAnalysis, Vec<FacultyProgramMembership>), String> {
    let (mut headers, mut rows) = read_full_spreadsheet(dataset_path)?;
    if headers.is_empty() {
        return Err("The faculty dataset does not include any columns.".into());
    }
    if rows.is_empty() {
        return Err("The faculty dataset does not include any faculty rows.".into());
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

    Ok((analysis, memberships))
}

/// Stage new bytes in the same directory, so a failed write leaves the active file intact.
pub(crate) fn persist_dataset_file(source: &Path, destination: &Path) -> Result<(), String> {
    use std::io::Write;
    use std::sync::atomic::{AtomicUsize, Ordering};
    static NEXT_REPLACEMENT: AtomicUsize = AtomicUsize::new(0);

    let bytes = fs::read(source)
        .map_err(|err| format!("Unable to read the replacement faculty dataset: {err}"))?;
    ensure_dataset_directory(destination)?;
    let staging = destination.with_file_name(format!(
        ".faculty-dataset-{}-{}.pending",
        std::process::id(),
        NEXT_REPLACEMENT.fetch_add(1, Ordering::Relaxed),
    ));
    let mut file = fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&staging)
        .map_err(|err| format!("Unable to prepare the replacement faculty dataset: {err}"))?;
    let result = (|| {
        file.write_all(&bytes)?;
        file.sync_all()?;
        drop(file);
        fs::rename(&staging, destination)
    })();
    if result.is_err() {
        let _ = fs::remove_file(staging);
    }
    result.map_err(|err| format!("Unable to replace the faculty dataset: {err}"))
}

pub(crate) fn suggest_program_columns(headers: &[String], rows: &[Vec<String>]) -> Vec<usize> {
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

pub(crate) fn header_label(headers: &[String], index: usize) -> String {
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

    values.sort_by_key(|a| a.to_lowercase());
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

pub(crate) fn load_faculty_dataset_metadata(
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

pub(crate) fn filter_faculty_rows_by_program(
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

pub(crate) fn write_faculty_dataset_source_path(
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

pub(crate) fn dataset_destination(app_handle: &tauri::AppHandle) -> Result<PathBuf, String> {
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

pub(crate) fn dataset_destination_for_extension(
    app_handle: &tauri::AppHandle,
    extension: &str,
) -> Result<PathBuf, String> {
    let directory = dataset_directory(app_handle)?;
    Ok(dataset_path_with_extension(&directory, extension))
}

pub(crate) fn dataset_directory(app_handle: &tauri::AppHandle) -> Result<PathBuf, String> {
    app_handle
        .path()
        .app_data_dir()
        .map_err(|err| format!("Unable to resolve the application data directory: {err}"))
}

fn dataset_path_with_extension(directory: &Path, extension: &str) -> PathBuf {
    directory.join(format!("{}.{}", FACULTY_DATASET_BASENAME, extension))
}

pub(crate) fn remove_other_dataset_variants(
    directory: &Path,
    keep_extension: &str,
) -> Result<(), String> {
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

pub(crate) fn ensure_dataset_directory(path: &Path) -> Result<(), String> {
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
        .first()
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

pub(crate) fn build_dataset_preview(path: &Path) -> Result<SpreadsheetPreview, String> {
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

#[cfg(test)]
#[path = "tests/dataset.rs"]
mod tests;
