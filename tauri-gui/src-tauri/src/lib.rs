use serde::{Deserialize, Serialize};
use std::collections::HashSet;
use std::fs;
use std::path::{Path, PathBuf};

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

#[derive(Debug, Deserialize, Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct FacultyIdentifier {
    first_name: String,
    last_name: String,
    #[serde(default)]
    identifier: Option<String>,
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
    custom_faculty: Vec<FacultyIdentifier>,
    faculty_recs_per_student: u32,
    student_recs_per_faculty: u32,
    update_embeddings: bool,
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
    custom_faculty: Vec<FacultyIdentifier>,
    recommendations_per_student: u32,
    recommendations_per_faculty: u32,
    update_embeddings: bool,
    prompt_preview: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct SubmissionResponse {
    summary: String,
    warnings: Vec<String>,
    details: SubmissionDetails,
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
        custom_faculty,
        faculty_recs_per_student,
        student_recs_per_faculty,
        update_embeddings,
    } = payload;

    if faculty_recs_per_student == 0 {
        return Err("Specify at least one faculty recommendation per student.".into());
    }

    let mut warnings = Vec::new();
    let mut validated_paths = Vec::new();
    let mut prompt_preview = None;

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
            if let Some(message) = validate_extension(
                &spreadsheet,
                &["csv", "tsv", "xlsx", "xls", "ods"],
                "spreadsheet",
            ) {
                warnings.push(message);
            }
            validated_paths.push(PathConfirmation::new("Spreadsheet", &spreadsheet));
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
    let sanitized_faculty = sanitize_faculty_list(custom_faculty);

    match faculty_scope {
        FacultyScope::Program if normalized_programs.is_empty() => {
            return Err("Provide at least one program to limit the faculty list.".into());
        }
        FacultyScope::Custom if sanitized_faculty.is_empty() => {
            return Err("Add at least one faculty member to the custom list.".into());
        }
        _ => {}
    }

    let details = SubmissionDetails {
        task_type: task_type.clone(),
        faculty_scope: faculty_scope.clone(),
        validated_paths,
        program_filters: match faculty_scope {
            FacultyScope::Program => normalized_programs.clone(),
            _ => Vec::new(),
        },
        custom_faculty: match faculty_scope {
            FacultyScope::Custom => sanitized_faculty.clone(),
            _ => Vec::new(),
        },
        recommendations_per_student: faculty_recs_per_student,
        recommendations_per_faculty: student_recs_per_faculty,
        update_embeddings,
        prompt_preview,
    };

    let summary = build_summary(
        &task_type,
        &faculty_scope,
        faculty_recs_per_student,
        student_recs_per_faculty,
        update_embeddings,
        details.program_filters.len(),
        details.custom_faculty.len(),
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

fn sanitize_faculty_list(entries: Vec<FacultyIdentifier>) -> Vec<FacultyIdentifier> {
    let mut seen = HashSet::new();
    let mut cleaned = Vec::new();

    for entry in entries {
        let first = entry.first_name.trim();
        let last = entry.last_name.trim();
        if first.is_empty() || last.is_empty() {
            continue;
        }

        let identifier = entry
            .identifier
            .as_deref()
            .map(str::trim)
            .filter(|value| !value.is_empty())
            .map(String::from);

        let key = format!(
            "{}::{}::{}",
            first.to_lowercase(),
            last.to_lowercase(),
            identifier.as_deref().unwrap_or_default().to_lowercase()
        );

        if seen.insert(key) {
            cleaned.push(FacultyIdentifier {
                first_name: first.to_string(),
                last_name: last.to_string(),
                identifier,
            });
        }
    }

    cleaned
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
    custom_count: usize,
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
        FacultyScope::Custom => format!(
            "{custom_count} custom faculty member{}",
            if custom_count == 1 { "" } else { "s" }
        ),
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
            update_faculty_embeddings
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
