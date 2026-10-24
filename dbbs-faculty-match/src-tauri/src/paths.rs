use std::fs;
use std::path::Path;
use std::path::PathBuf;

pub(crate) fn resolve_existing_path(
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

pub(crate) fn validate_extension(path: &Path, allowed: &[&str], label: &str) -> Option<String> {
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

pub(crate) fn build_prompt_preview(text: &str) -> String {
    let characters: Vec<char> = text.chars().collect();
    let max = 280usize;
    if characters.len() <= max {
        return text.to_string();
    }

    let mut preview: String = characters[..max].iter().collect();
    preview.push('…');
    preview
}
