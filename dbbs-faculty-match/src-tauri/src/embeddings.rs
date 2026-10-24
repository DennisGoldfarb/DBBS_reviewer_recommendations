use crate::config::{
    DEFAULT_EMBEDDING_MODEL, FACULTY_EMBEDDINGS_NAME, FACULTY_EMBEDDING_PROGRESS_EVENT,
    FACULTY_EMBEDDING_STANDBY_MESSAGE,
};
use crate::dataset::{
    build_faculty_dataset_status, dataset_destination, dataset_directory, ensure_dataset_directory,
    header_label,
};
use crate::models::{
    EmbeddingRequestPayload, EmbeddingRequestRow, EmbeddingResponsePayload, FacultyEmbeddingEntry,
    FacultyEmbeddingIndex, FacultyPreparationOptions,
};
use crate::spreadsheets::{build_header_index_map, indexes_from_labels, read_full_spreadsheet};
use chrono::Utc;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;
use std::sync::atomic::{AtomicBool, Ordering as AtomicOrdering};
use std::sync::{Mutex, OnceLock};
use std::time::Instant;
use tauri::Emitter;
use tauri_plugin_shell::{
    process::{CommandChild, CommandEvent, TerminatedPayload},
    ShellExt,
};
use tokio::sync::mpsc::Receiver;

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

pub(crate) fn perform_faculty_embedding_refresh(
    app_handle: tauri::AppHandle,
) -> Result<String, String> {
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
        faculty_preparation: Some(FacultyPreparationOptions {}),
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
        return Err("The embedding helper returned an empty result. Verify the Python environment can load the embedding model.".into());
    }

    let EmbeddingResponsePayload {
        model: response_model,
        dimension: response_dimension,
        rows: response_rows,
    } = response;

    let mut embedding_map = HashMap::new();
    let helper_row_count = response_rows.len();
    for row in response_rows {
        embedding_map.insert(row.id, row);
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
    let mut prepared_descriptions = 0usize;

    for context in contexts {
        match embedding_map.remove(&context.row_index) {
            Some(row) => {
                if row.preparation.is_some() {
                    prepared_descriptions += 1;
                }
                entries.push(FacultyEmbeddingEntry {
                    row_index: context.row_index,
                    identifiers: context.identifiers,
                    embedding: row.embedding,
                    preparation: row.preparation,
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

    if prepared_descriptions > 0 {
        message.push_str(&format!(
            " Prepared {prepared_descriptions} faculty description{plural} locally for topic matching within the 512-token limit. Original descriptions are preserved in the uploaded dataset.",
            plural = if prepared_descriptions == 1 { "" } else { "s" },
        ));
    }

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

static EMBEDDING_HELPER: OnceLock<Mutex<EmbeddingHelperManager>> = OnceLock::new();
pub(crate) static EMBEDDING_HELPER_READY: AtomicBool = AtomicBool::new(false);
static EMBEDDING_HELPER_WARMUP_LOCK: OnceLock<Mutex<()>> = OnceLock::new();

struct EmbeddingHelperManager {
    child: Option<CommandChild>,
    receiver: Option<Receiver<CommandEvent>>,
    stdout_buffer: Vec<u8>,
}

impl EmbeddingHelperManager {
    fn new() -> Self {
        Self {
            child: None,
            receiver: None,
            stdout_buffer: Vec::new(),
        }
    }

    fn ensure_process(&mut self, app_handle: &tauri::AppHandle) -> Result<(), String> {
        if self.child.is_none() {
            let (receiver, child) = spawn_embedding_helper(app_handle)?;
            self.child = Some(child);
            self.receiver = Some(receiver);
            self.stdout_buffer.clear();
        }

        Ok(())
    }

    fn write_with_newline(&mut self, bytes: &[u8]) -> std::io::Result<()> {
        if let Some(child) = self.child.as_mut() {
            let mut payload = Vec::with_capacity(bytes.len() + 1);
            payload.extend_from_slice(bytes);
            payload.push(b'\n');

            child
                .write(&payload)
                .map_err(|err| std::io::Error::other(err.to_string()))
        } else {
            Err(std::io::Error::new(
                std::io::ErrorKind::BrokenPipe,
                "Embedding helper is not running",
            ))
        }
    }

    fn receive_event(&mut self) -> Result<Option<CommandEvent>, String> {
        let receiver = match self.receiver.as_mut() {
            Some(receiver) => receiver,
            None => return Ok(None),
        };

        Ok(tauri::async_runtime::block_on(receiver.recv()))
    }

    fn reset(&mut self) {
        self.child = None;
        self.receiver = None;
        self.stdout_buffer.clear();
    }

    fn shutdown(&mut self) {
        if let Some(child) = self.child.as_mut() {
            let _ = child.write(b"{\"command\":\"shutdown\"}\n");
        }

        if let Some(receiver) = self.receiver.as_mut() {
            while let Some(event) = tauri::async_runtime::block_on(receiver.recv()) {
                if matches!(event, CommandEvent::Terminated(_)) {
                    break;
                }
            }
        }

        self.reset();
    }
}

fn embedding_helper_manager() -> &'static Mutex<EmbeddingHelperManager> {
    EMBEDDING_HELPER.get_or_init(|| Mutex::new(EmbeddingHelperManager::new()))
}

fn embedding_helper_warmup_lock() -> &'static Mutex<()> {
    EMBEDDING_HELPER_WARMUP_LOCK.get_or_init(|| Mutex::new(()))
}

fn emit_embedding_helper_standby(app_handle: &tauri::AppHandle) {
    emit_faculty_embedding_progress(
        app_handle,
        EmbeddingProgressUpdate {
            phase: "standby".into(),
            message: Some(FACULTY_EMBEDDING_STANDBY_MESSAGE.into()),
            processed_rows: 0,
            total_rows: 0,
            elapsed_seconds: None,
            estimated_remaining_seconds: None,
        },
    );
}

fn perform_embedding_helper_warmup(app_handle: tauri::AppHandle) -> Result<(), String> {
    if EMBEDDING_HELPER_READY.load(AtomicOrdering::SeqCst) {
        emit_embedding_helper_standby(&app_handle);
        return Ok(());
    }

    let lock = embedding_helper_warmup_lock()
        .lock()
        .map_err(|err| format!("Unable to warm up the embedding helper: {err}"))?;

    if EMBEDDING_HELPER_READY.load(AtomicOrdering::SeqCst) {
        drop(lock);
        emit_embedding_helper_standby(&app_handle);
        return Ok(());
    }

    let payload = EmbeddingRequestPayload {
        model: DEFAULT_EMBEDDING_MODEL.to_string(),
        texts: vec![EmbeddingRequestRow {
            id: 0,
            text: "Embedding helper warm-up text".into(),
        }],
        item_label: Some("warm-up text".into()),
        item_label_plural: Some("warm-up texts".into()),
        faculty_preparation: None,
    };

    let result = run_embedding_helper(&app_handle, &payload);
    drop(lock);

    result.map(|_| {
        EMBEDDING_HELPER_READY.store(true, AtomicOrdering::SeqCst);
        emit_embedding_helper_standby(&app_handle);
    })
}

pub(crate) async fn warmup_embedding_helper(app_handle: tauri::AppHandle) -> Result<(), String> {
    if EMBEDDING_HELPER_READY.load(AtomicOrdering::SeqCst) {
        emit_embedding_helper_standby(&app_handle);
        return Ok(());
    }

    tauri::async_runtime::spawn_blocking(move || perform_embedding_helper_warmup(app_handle))
        .await
        .map_err(|err| format!("Unable to warm up the embedding helper: {err}"))?
}

pub(crate) fn shutdown_embedding_helper() {
    if let Some(lock) = EMBEDDING_HELPER.get() {
        if let Ok(mut manager) = lock.lock() {
            manager.shutdown();
        }
    }
    EMBEDDING_HELPER_READY.store(false, AtomicOrdering::SeqCst);
}

pub(crate) fn run_embedding_helper(
    app_handle: &tauri::AppHandle,
    payload: &EmbeddingRequestPayload,
) -> Result<EmbeddingResponsePayload, String> {
    let total_rows = payload.texts.len();
    let input = serde_json::to_vec(payload)
        .map_err(|err| format!("Unable to serialize the embedding request: {err}"))?;

    let manager_lock = embedding_helper_manager();
    let mut manager = manager_lock
        .lock()
        .map_err(|err| format!("Unable to lock the embedding helper: {err}"))?;

    manager.ensure_process(app_handle)?;

    if let Err(err) = manager.write_with_newline(&input) {
        let message = handle_sidecar_write_error(
            format!("Unable to send data to the embedding helper: {err}"),
            &mut manager,
            app_handle,
            total_rows,
        );
        emit_embedding_error(app_handle, total_rows, &message);
        return Err(message);
    }

    let output = match collect_sidecar_output(&mut manager, app_handle, total_rows, true) {
        Ok(output) => output,
        Err(err) => {
            emit_embedding_error(app_handle, total_rows, &err);
            return Err(err);
        }
    };

    if !sidecar_exited_successfully(&output) {
        let error_message = trimmed_message_or_default(
            &output.stderr,
            "The embedding helper exited with an error.",
        );
        emit_embedding_error(app_handle, total_rows, &error_message);
        return Err(error_message);
    }

    if output.stdout.is_empty() {
        let error_message = trimmed_message_or_default(
            &output.stderr,
            if output.termination.is_some() {
                "The embedding helper exited before returning a response."
            } else {
                "The embedding helper did not produce any output."
            },
        );
        emit_embedding_error(app_handle, total_rows, &error_message);
        return Err(error_message);
    }

    match parse_embedding_response(&output.stdout) {
        Ok(response) => Ok(response),
        Err(error_message) => {
            emit_embedding_error(app_handle, total_rows, &error_message);
            Err(error_message)
        }
    }
}

fn parse_embedding_response(output: &[u8]) -> Result<EmbeddingResponsePayload, String> {
    let value: serde_json::Value = serde_json::from_slice(output)
        .map_err(|err| format!("Invalid embedding helper response: {err}"))?;
    if let Some(error) = value.get("error").and_then(serde_json::Value::as_str) {
        return Err(error.to_string());
    }
    serde_json::from_value(value).map_err(|err| format!("Invalid embedding helper response: {err}"))
}

fn handle_sidecar_write_error(
    base_message: String,
    manager: &mut EmbeddingHelperManager,
    app_handle: &tauri::AppHandle,
    total_rows: usize,
) -> String {
    let result = match collect_sidecar_output(manager, app_handle, total_rows, false) {
        Ok(output) => summarize_sidecar_failure(&output, base_message),
        Err(err) => format!("{base_message}\n\nAdditionally, {err}"),
    };

    manager.reset();
    result
}

struct SidecarOutput {
    stdout: Vec<u8>,
    stderr: Vec<u8>,
    termination: Option<TerminatedPayload>,
}

fn sidecar_exited_successfully(output: &SidecarOutput) -> bool {
    output.termination.is_none()
}

fn trim_trailing_newline(buffer: &mut Vec<u8>) {
    while matches!(buffer.last(), Some(b'\n') | Some(b'\r')) {
        buffer.pop();
    }
}

fn collect_sidecar_output(
    manager: &mut EmbeddingHelperManager,
    app_handle: &tauri::AppHandle,
    total_rows: usize,
    expect_response: bool,
) -> Result<SidecarOutput, String> {
    let mut stdout = Vec::new();
    let mut stderr = Vec::new();

    if !expect_response && !manager.stdout_buffer.is_empty() {
        stdout.append(&mut manager.stdout_buffer);
    }

    loop {
        if expect_response {
            if let Some(position) = manager.stdout_buffer.iter().position(|&byte| byte == b'\n') {
                stdout.extend(manager.stdout_buffer.drain(..=position));
                trim_trailing_newline(&mut stdout);
                return Ok(SidecarOutput {
                    stdout,
                    stderr,
                    termination: None,
                });
            }
        }

        match manager.receive_event()? {
            Some(CommandEvent::Stdout(bytes)) => {
                if expect_response {
                    manager.stdout_buffer.extend_from_slice(&bytes);
                } else {
                    stdout.extend_from_slice(&bytes);
                }
            }
            Some(CommandEvent::Stderr(bytes)) => {
                if !emit_progress_from_line(app_handle, total_rows, &bytes) {
                    stderr.extend_from_slice(&bytes);
                }
            }
            Some(CommandEvent::Terminated(payload)) => {
                if expect_response {
                    stdout.append(&mut manager.stdout_buffer);
                    trim_trailing_newline(&mut stdout);
                }
                manager.reset();
                return Ok(SidecarOutput {
                    stdout,
                    stderr,
                    termination: Some(payload),
                });
            }
            Some(CommandEvent::Error(err)) => {
                manager.reset();
                return Err(format!("Unable to read embedding helper output: {err}"));
            }
            Some(_) => {}
            None => {
                manager.reset();
                if expect_response {
                    return Err(
                        "The embedding helper output stream closed unexpectedly.".to_string()
                    );
                } else {
                    trim_trailing_newline(&mut stdout);
                    return Ok(SidecarOutput {
                        stdout,
                        stderr,
                        termination: None,
                    });
                }
            }
        }
    }
}

fn summarize_sidecar_failure(output: &SidecarOutput, base_message: String) -> String {
    let mut message = base_message;

    let stderr_text = String::from_utf8_lossy(&output.stderr);
    let stderr_trimmed = stderr_text.trim_end();
    if !stderr_trimmed.is_empty() {
        message = format!("{message}\n\nEmbedding helper stderr:\n{stderr_trimmed}");
    }

    let stdout_text = String::from_utf8_lossy(&output.stdout);
    let stdout_trimmed = stdout_text.trim_end();
    if !stdout_trimmed.is_empty() {
        message = format!("{message}\n\nEmbedding helper stdout:\n{stdout_trimmed}");
    }

    if let Some(termination) = &output.termination {
        let status_summary = match (termination.code, termination.signal) {
            (Some(code), Some(signal)) => format!("code {code}, signal {signal}"),
            (Some(code), None) => format!("code {code}"),
            (None, Some(signal)) => format!("signal {signal}"),
            (None, None) => "no exit code recorded".to_string(),
        };
        message = format!("{message}\n\nEmbedding helper exit status: {status_summary}");
    }

    message
}

fn emit_progress_from_line(app_handle: &tauri::AppHandle, total_rows: usize, bytes: &[u8]) -> bool {
    let text = String::from_utf8_lossy(bytes);
    let trimmed = text.trim_end_matches(['\r', '\n']);
    if let Some(json_str) = trimmed.strip_prefix("PROGRESS ") {
        match serde_json::from_str::<EmbeddingProgressUpdate>(json_str.trim()) {
            Ok(mut update) => {
                if update.total_rows == 0 {
                    update.total_rows = total_rows;
                }
                emit_faculty_embedding_progress(app_handle, update);
                true
            }
            Err(_) => false,
        }
    } else {
        false
    }
}

fn trimmed_message_or_default(buffer: &[u8], default_message: &str) -> String {
    let text = String::from_utf8_lossy(buffer);
    let trimmed = text.trim();
    if trimmed.is_empty() {
        default_message.to_string()
    } else {
        trimmed.to_string()
    }
}

fn spawn_embedding_helper(
    app_handle: &tauri::AppHandle,
) -> Result<(Receiver<CommandEvent>, CommandChild), String> {
    app_handle
        .shell()
        .sidecar("embedding-helper")
        .map_err(|err| format!("Unable to configure the embedding helper: {err}"))?
        .env("HF_HUB_DISABLE_PROGRESS_BARS", "1")
        .env("TOKENIZERS_PARALLELISM", "false")
        .env("PYTHONUTF8", "1")
        .env("PYTHONUNBUFFERED", "1")
        .spawn()
        .map_err(|err| format!("Unable to launch the embedding helper: {err}"))
}

#[cfg(test)]
#[path = "tests/embeddings.rs"]
mod tests;
