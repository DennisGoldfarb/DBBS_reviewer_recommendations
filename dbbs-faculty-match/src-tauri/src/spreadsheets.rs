use calamine::{open_workbook_auto, DataType, Reader};
use std::cmp::Ordering;
use std::collections::HashMap;
use std::collections::HashSet;
use std::fs::File;
use std::io::{BufRead, BufReader};
use std::path::Path;

pub(crate) fn build_header_index_map(headers: &[String]) -> HashMap<String, usize> {
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

pub(crate) fn indexes_from_labels(
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

pub(crate) fn indexes_from_spreadsheet_labels(
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
                "The column '{label}' is not available in the spreadsheet. Reload the preview and try again."
            ));
        }
    }

    indexes.sort_unstable();
    indexes.dedup();
    Ok(indexes)
}

pub(crate) fn normalize_programs(programs: Vec<String>) -> Vec<String> {
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

pub(crate) fn normalize_columns(columns: Vec<String>) -> Vec<String> {
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

pub(crate) fn normalize_identifier_value(value: &str) -> String {
    value
        .split_whitespace()
        .filter(|segment| !segment.is_empty())
        .collect::<Vec<_>>()
        .join(" ")
        .to_lowercase()
}

pub(crate) fn normalize_identifier_label(value: &str) -> String {
    value
        .chars()
        .filter(|ch| ch.is_alphanumeric())
        .collect::<String>()
        .to_lowercase()
}

pub(crate) fn read_spreadsheet(path: &Path) -> Result<(Vec<String>, Vec<Vec<String>>), String> {
    read_spreadsheet_with_limit(path, Some(10))
}

pub(crate) fn read_full_spreadsheet(
    path: &Path,
) -> Result<(Vec<String>, Vec<Vec<String>>), String> {
    read_spreadsheet_with_limit(path, None)
}

pub(crate) fn read_spreadsheet_with_limit(
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

pub(crate) fn cell_to_string(cell: &DataType) -> String {
    match cell {
        DataType::Empty => String::new(),
        _ => cell.to_string().trim().to_string(),
    }
}

pub(crate) fn align_row_lengths(headers: &mut Vec<String>, rows: &mut [Vec<String>]) {
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

pub(crate) fn suggest_spreadsheet_columns(
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

pub(crate) fn sort_and_dedup(values: &mut Vec<usize>) {
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

#[cfg(test)]
#[path = "tests/spreadsheets.rs"]
mod tests;
