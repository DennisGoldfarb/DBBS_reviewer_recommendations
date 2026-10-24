use crate::config::DEFAULT_EMBEDDING_MODEL;
use crate::embeddings::run_embedding_helper;
use crate::exclusions::FacultyExclusions;
use crate::exports::{
    build_matches_headers, build_matches_workbook, default_matches_workbook_name,
    format_similarity_percent,
};
use crate::models::{
    EmbeddingRequestPayload, EmbeddingRequestRow, FacultyEmbeddingIndex, FacultyMatchResult,
    GeneratedSpreadsheet, MatchEntry, PromptMatchResult, SpreadsheetMatchResults,
    SpreadsheetPreview, SpreadsheetProcessingOutcome,
};
use crate::paths::build_prompt_preview;
use crate::scoring::{assign_student_rankings, find_best_faculty_matches};
use crate::spreadsheets::{
    build_header_index_map, indexes_from_spreadsheet_labels, read_full_spreadsheet,
};
use base64::{engine::general_purpose::STANDARD as Base64Engine, Engine as _};
use std::collections::HashMap;
use std::collections::HashSet;
use std::path::Path;

pub(crate) struct StudentSpreadsheetColumns<'a> {
    pub(crate) prompts: &'a [String],
    pub(crate) identifiers: &'a [String],
}

pub(crate) fn process_prompt_spreadsheet(
    app_handle: &tauri::AppHandle,
    spreadsheet_path: &Path,
    index: &FacultyEmbeddingIndex,
    columns: StudentSpreadsheetColumns<'_>,
    limit: usize,
    allowed_rows: Option<&HashSet<usize>>,
    exclusions: Option<&FacultyExclusions>,
) -> Result<SpreadsheetProcessingOutcome, String> {
    let StudentSpreadsheetColumns {
        prompts: prompt_columns,
        identifiers: identifier_columns,
    } = columns;

    #[derive(Debug)]
    struct SpreadsheetRowContext {
        result_index: usize,
        prompt: String,
        eligible_rows: Option<HashSet<usize>>,
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
    let exclusion_id_index = exclusions
        .map(|rules| {
            indexes_from_spreadsheet_labels(
                &header_map,
                std::slice::from_ref(&rules.student_id_column),
            )
            .map(|indexes| indexes[0])
        })
        .transpose()?;
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
                eligible_rows: exclusions.zip(exclusion_id_index).map(|(rules, column)| {
                    rules.eligible_rows(
                        row.get(column).map(String::as_str).unwrap_or_default(),
                        allowed_rows,
                    )
                }),
                result_index,
                prompt: prompt_text,
            });
        }

        row_results.push(result);
    }

    let mut prompt_matches = Vec::new();
    let mut missing_embeddings = 0usize;

    if !contexts.is_empty() {
        let payload = EmbeddingRequestPayload {
            model: DEFAULT_EMBEDDING_MODEL.to_string(),
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
            faculty_preparation: None,
        };

        let response = run_embedding_helper(app_handle, &payload)?;
        if response.dimension != index.dimension {
            return Err(format!(
                "The spreadsheet embedding dimension ({}) does not match the faculty embedding dimension ({}).",
                response.dimension, index.dimension
            ));
        }

        let mut embedding_map: HashMap<usize, Vec<Vec<f32>>> = HashMap::new();
        for row in response.rows {
            embedding_map.insert(row.id, row.into_embeddings());
        }

        for (context_index, context) in contexts.iter().enumerate() {
            let result = &mut row_results[context.result_index];

            match embedding_map.remove(&context_index) {
                Some(embeddings) => {
                    let matches = find_best_faculty_matches(
                        index,
                        &embeddings,
                        limit,
                        context.eligible_rows.as_ref().or(allowed_rows),
                    );

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
