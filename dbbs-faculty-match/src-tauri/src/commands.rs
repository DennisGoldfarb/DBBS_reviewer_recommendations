use crate::config::{FACULTY_DATASET_DEFAULT_EXTENSION, FACULTY_DATASET_EXTENSIONS};
use crate::dataset::{
    build_dataset_preview, build_faculty_dataset_status,
    build_faculty_dataset_status_with_overrides, dataset_destination_for_extension,
    dataset_directory, ensure_default_faculty_dataset, inspect_faculty_dataset,
    load_faculty_dataset_metadata, persist_dataset_file, remove_other_dataset_variants,
    suggest_program_columns, write_faculty_dataset_source_path,
};
use crate::embeddings::{
    perform_faculty_embedding_refresh, warmup_embedding_helper, EMBEDDING_HELPER_READY,
};
use crate::matching::perform_matching_request;
use crate::models::{
    FacultyDatasetColumnConfiguration, FacultyDatasetPreviewResponse, FacultyDatasetStatus,
    FacultyRosterPreviewResponse, SpreadsheetPreview, SubmissionPayload, SubmissionResponse,
};
use crate::paths::resolve_existing_path;
use crate::spreadsheets::{
    align_row_lengths, build_header_index_map, normalize_identifier_label, read_spreadsheet,
    read_spreadsheet_with_limit, suggest_spreadsheet_columns,
};
use base64::{engine::general_purpose::STANDARD as Base64Engine, Engine as _};
use std::collections::HashMap;
use std::fs;
use std::path::PathBuf;

#[derive(serde::Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) enum UploadKind {
    Document,
    Spreadsheet,
    Directory,
    Dataset,
}

/// Validate the real filesystem item; extensions alone cannot distinguish a folder from a file.
#[tauri::command]
pub(crate) fn validate_upload_path(path: String, kind: UploadKind) -> Result<String, String> {
    let is_directory = matches!(kind, UploadKind::Directory);
    let path = resolve_existing_path(Some(path), is_directory, "Dropped item")?;
    let allowed: &[&str] = match kind {
        UploadKind::Directory => &[],
        UploadKind::Document => &["txt", "pdf", "doc", "docx", "rtf"],
        UploadKind::Spreadsheet => &["csv", "tsv", "txt", "xlsx", "xls"],
        UploadKind::Dataset => FACULTY_DATASET_EXTENSIONS,
    };
    if !is_directory {
        if let Some(warning) = crate::paths::validate_extension(&path, allowed, "file") {
            return Err(warning);
        }
    }
    Ok(path.to_string_lossy().into_owned())
}
use std::sync::atomic::Ordering as AtomicOrdering;

#[tauri::command]
pub(crate) async fn submit_matching_request(
    app_handle: tauri::AppHandle,
    payload: SubmissionPayload,
) -> Result<SubmissionResponse, String> {
    tauri::async_runtime::spawn_blocking(move || perform_matching_request(app_handle, payload))
        .await
        .map_err(|err| format!("Matching task failed: {err}"))?
}

#[tauri::command]
pub(crate) async fn update_faculty_embeddings(
    app_handle: tauri::AppHandle,
) -> Result<String, String> {
    let result =
        tauri::async_runtime::spawn_blocking(move || perform_faculty_embedding_refresh(app_handle))
            .await
            .map_err(|err| format!("Embedding refresh task failed: {err}"))?;

    result
}

#[tauri::command]
pub(crate) async fn ensure_embedding_helper_ready(
    app_handle: tauri::AppHandle,
) -> Result<bool, String> {
    warmup_embedding_helper(app_handle).await?;
    Ok(EMBEDDING_HELPER_READY.load(AtomicOrdering::SeqCst))
}

#[tauri::command]
pub(crate) fn get_faculty_dataset_status(
    app_handle: tauri::AppHandle,
) -> Result<FacultyDatasetStatus, String> {
    build_faculty_dataset_status(&app_handle)
}

#[tauri::command]
pub(crate) fn preview_faculty_roster(
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
pub(crate) fn preview_faculty_dataset_replacement(
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
pub(crate) fn replace_faculty_dataset(
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

    inspect_faculty_dataset(&source, configuration.as_ref())?;
    let destination = dataset_destination_for_extension(&app_handle, &extension)?;
    let embeddings = dataset_directory(&app_handle)?.join(crate::config::FACULTY_EMBEDDINGS_NAME);
    if embeddings.exists() {
        fs::remove_file(embeddings)
            .map_err(|err| format!("Unable to clear the previous faculty embeddings: {err}"))?;
    }
    persist_dataset_file(&source, &destination)?;
    if let Some(directory) = destination.parent() {
        remove_other_dataset_variants(directory, &extension)?;
    }

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
pub(crate) fn restore_default_faculty_dataset(
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

#[tauri::command]
pub(crate) fn save_generated_spreadsheet(
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
pub(crate) fn analyze_spreadsheet(path: String) -> Result<SpreadsheetPreview, String> {
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

#[cfg(test)]
#[path = "tests/commands.rs"]
mod tests;
