use crate::config::{DEFAULT_EMBEDDING_MODEL, FACULTY_EMBEDDINGS_NAME};
use crate::dataset::{
    dataset_destination, dataset_directory, filter_faculty_rows_by_program, header_label,
    load_faculty_dataset_metadata,
};
use crate::directory_matching::process_directory_documents;
use crate::documents::extract_document_prompt;
use crate::embeddings::run_embedding_helper;
use crate::exclusions::FacultyExclusions;
use crate::models::{
    EmbeddingRequestPayload, EmbeddingRequestRow, FacultyEmbeddingIndex, FacultyMatchResult,
    FacultyScope, PathConfirmation, PromptMatchResult, SubmissionDetails, SubmissionPayload,
    SubmissionResponse, TaskType,
};
use crate::paths::{build_prompt_preview, resolve_existing_path, validate_extension};
use crate::scoring::{assign_student_rankings, find_best_faculty_matches};
use crate::spreadsheet_matching::{process_prompt_spreadsheet, StudentSpreadsheetColumns};
use crate::spreadsheets::{
    align_row_lengths, build_header_index_map, indexes_from_labels, normalize_columns,
    normalize_identifier_label, normalize_identifier_value, normalize_programs,
    read_full_spreadsheet,
};
use std::collections::HashMap;
use std::collections::HashSet;
use std::fs;
use std::path::PathBuf;

pub(crate) fn perform_matching_request(
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
        faculty_exclusions,
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
                validate_extension(&document, &["txt", "pdf", "doc", "docx", "rtf"], "document")
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
            if let Some(message) = validate_extension(
                &spreadsheet,
                &["csv", "tsv", "txt", "xlsx", "xls"],
                "spreadsheet",
            ) {
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
        if let Some(message) = validate_extension(
            &roster,
            &["csv", "tsv", "txt", "xlsx", "xls"],
            "faculty list",
        ) {
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

    let exclusions = faculty_exclusions
        .as_ref()
        .map(|configuration| {
            let index = faculty_embedding_index
                .as_ref()
                .ok_or_else(|| "The faculty embedding index was not loaded.".to_string())?;
            FacultyExclusions::load(configuration, index, &task_type)
        })
        .transpose()?;

    if let Some(prompt_text) = prepared_prompt_text {
        let limit = faculty_recs_per_student.max(1) as usize;
        let embedding_index = faculty_embedding_index
            .as_ref()
            .ok_or_else(|| "The faculty embedding index was not loaded.".to_string())?;
        let prompt_embeddings = embed_prompt_chunks(&app_handle, embedding_index, &prompt_text)?;
        let eligible_rows = exclusions
            .as_ref()
            .map(|rules| rules.eligible_rows(&rules.student_id, allowed_faculty_rows.as_ref()));
        let mut matches = find_best_faculty_matches(
            embedding_index,
            &prompt_embeddings,
            limit,
            eligible_rows.as_ref().or(allowed_faculty_rows.as_ref()),
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
            exclusions.as_ref(),
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
            StudentSpreadsheetColumns {
                prompts: &selected_prompt_columns,
                identifiers: &selected_identifier_columns,
            },
            limit,
            allowed_faculty_rows.as_ref(),
            exclusions.as_ref(),
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

fn load_faculty_embedding_index(
    app_handle: &tauri::AppHandle,
) -> Result<FacultyEmbeddingIndex, String> {
    let embeddings_path = dataset_directory(app_handle)?.join(FACULTY_EMBEDDINGS_NAME);
    let data = if embeddings_path.exists() {
        fs::read(&embeddings_path)
            .map_err(|err| format!("Unable to read faculty embeddings: {err}"))?
    } else {
        return Err("Faculty embeddings are unavailable. Update faculty embeddings before matching against this dataset.".into());
    };

    parse_faculty_embedding_index(&data)
}

fn parse_faculty_embedding_index(data: &[u8]) -> Result<FacultyEmbeddingIndex, String> {
    let mut index: FacultyEmbeddingIndex = serde_json::from_slice(data)
        .map_err(|err| format!("Unable to parse faculty embeddings: {err}"))?;
    if !index.model.trim().is_empty() && index.model != DEFAULT_EMBEDDING_MODEL {
        return Err("This faculty index uses an experimental embedding model. Select Update faculty embeddings to rebuild it with PubMedBERT.".into());
    }
    index.model = DEFAULT_EMBEDDING_MODEL.to_string();
    Ok(index)
}

fn embed_prompt_chunks(
    app_handle: &tauri::AppHandle,
    index: &FacultyEmbeddingIndex,
    prompt: &str,
) -> Result<Vec<Vec<f32>>, String> {
    let payload = EmbeddingRequestPayload {
        model: DEFAULT_EMBEDDING_MODEL.to_string(),
        texts: vec![EmbeddingRequestRow {
            id: 0,
            text: prompt.to_string(),
        }],
        item_label: Some("text query".into()),
        item_label_plural: Some("text queries".into()),
        faculty_preparation: None,
    };

    let response = run_embedding_helper(app_handle, &payload)?;
    if response.rows.is_empty() {
        return Err("The embedding helper did not return an embedding for the prompt.".into());
    }

    let embeddings = response.rows.into_iter().next().unwrap().into_embeddings();
    if let Some(embedding) = embeddings
        .iter()
        .find(|chunk| chunk.len() != index.dimension)
    {
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

    Ok(embeddings)
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

#[cfg(test)]
#[path = "tests/matching.rs"]
mod tests;
