use crate::config::DEFAULT_EMBEDDING_MODEL;
use crate::documents::extract_document_prompt;
use crate::embeddings::run_embedding_helper;
use crate::exclusions::FacultyExclusions;
use crate::exports::{
    build_matches_headers, build_matches_workbook, default_directory_workbook_name,
    format_similarity_percent,
};
use crate::models::{
    DirectoryMatchResults, DirectoryProcessingOutcome, DocumentExtractionResult,
    EmbeddingRequestPayload, EmbeddingRequestRow, FacultyEmbeddingIndex, FacultyMatchResult,
    GeneratedSpreadsheet, MatchEntry, PromptMatchResult, SpreadsheetPreview,
};
use crate::paths::build_prompt_preview;
use crate::scoring::{assign_student_rankings, find_best_faculty_matches};
use base64::{engine::general_purpose::STANDARD as Base64Engine, Engine as _};
use std::collections::HashMap;
use std::collections::HashSet;
use std::fs;
use std::path::Path;
use std::path::PathBuf;

pub(crate) fn process_directory_documents(
    app_handle: &tauri::AppHandle,
    directory: &Path,
    index: &FacultyEmbeddingIndex,
    limit: usize,
    allowed_rows: Option<&HashSet<usize>>,
    exclusions: Option<&FacultyExclusions>,
) -> Result<DirectoryProcessingOutcome, String> {
    #[derive(Debug)]
    struct DirectoryDocumentContext {
        result_index: usize,
        prompt: String,
        eligible_rows: Option<HashSet<usize>>,
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
                eligible_rows: exclusions.map(|rules| {
                    rules.eligible_rows(
                        path.file_stem()
                            .and_then(|value| value.to_str())
                            .unwrap_or_default(),
                        allowed_rows,
                    )
                }),
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
            item_label: Some("document".into()),
            item_label_plural: Some("documents".into()),
            faculty_preparation: None,
        };

        let response = run_embedding_helper(app_handle, &payload)?;
        if response.dimension != index.dimension {
            return Err(format!(
                "The document embedding dimension ({}) does not match the faculty embedding dimension ({}).",
                response.dimension, index.dimension
            ));
        }

        let mut embedding_map: HashMap<usize, Vec<Vec<f32>>> = HashMap::new();
        for row in response.rows {
            embedding_map.insert(row.id, row.into_embeddings());
        }

        for (context_index, context) in contexts.iter().enumerate() {
            let identifier = document_results[context.result_index].identifier.clone();

            match embedding_map.remove(&context_index) {
                Some(embeddings) => {
                    let matches = find_best_faculty_matches(
                        index,
                        &embeddings,
                        limit,
                        context.eligible_rows.as_ref().or(allowed_rows),
                    );

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
