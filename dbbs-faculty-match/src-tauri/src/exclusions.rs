use crate::models::{FacultyEmbeddingIndex, FacultyExclusionConfig, TaskType};
use crate::paths::resolve_existing_path;
use crate::spreadsheets::{
    build_header_index_map, indexes_from_spreadsheet_labels, read_full_spreadsheet,
};
use std::collections::{HashMap, HashSet};

/// Both "Last, First" and "First Last" become lowercase "first last".
/// Keep name parts intact: nicknames and missing middle names are not inferred.
pub(crate) fn normalize_person_name(value: &str) -> String {
    let ordered = match value.split_once(',') {
        Some((last, first)) => format!("{first} {last}"),
        None => value.to_string(),
    };
    ordered
        .split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
        .to_lowercase()
}

pub(crate) struct FacultyExclusions {
    pub(crate) student_id_column: String,
    pub(crate) student_id: String,
    all_faculty_rows: HashSet<usize>,
    forbidden_rows: HashMap<String, HashSet<usize>>,
}

impl FacultyExclusions {
    pub(crate) fn load(
        configuration: &FacultyExclusionConfig,
        index: &FacultyEmbeddingIndex,
        task: &TaskType,
    ) -> Result<Self, String> {
        if matches!(task, TaskType::Spreadsheet)
            && configuration.student_id_column.trim().is_empty()
        {
            return Err("Select the student ID column for faculty exclusions.".into());
        }
        if matches!(task, TaskType::Prompt | TaskType::Document)
            && configuration.student_id.trim().is_empty()
        {
            return Err("Enter the student ID for faculty exclusions.".into());
        }
        if configuration.excluded_name_columns.is_empty() {
            return Err("Select at least one faculty name column to exclude.".into());
        }
        if !index
            .identifier_columns
            .contains(&configuration.faculty_name_column)
        {
            return Err("Select a faculty name column from the active dataset identifiers.".into());
        }
        let path = resolve_existing_path(
            Some(configuration.path.clone()),
            false,
            "Exclusion spreadsheet",
        )?;
        let (headers, rows) = read_full_spreadsheet(&path)?;
        let columns = build_header_index_map(&headers);
        let id_column = indexes_from_spreadsheet_labels(
            &columns,
            std::slice::from_ref(&configuration.exclusion_student_id_column),
        )?[0];
        let writer_columns =
            indexes_from_spreadsheet_labels(&columns, &configuration.excluded_name_columns)?;
        let mut faculty_by_name: HashMap<String, HashSet<usize>> = HashMap::new();
        for entry in &index.entries {
            if let Some(name) = entry.identifiers.get(&configuration.faculty_name_column) {
                let name = normalize_person_name(name);
                if !name.is_empty() {
                    // Store every faculty row with this name, including duplicates.
                    faculty_by_name
                        .entry(name)
                        .or_default()
                        .insert(entry.row_index);
                }
            }
        }
        let mut forbidden_rows: HashMap<String, HashSet<usize>> = HashMap::new();
        for (row_number, row) in rows.iter().enumerate() {
            let id = row
                .get(id_column)
                .map(|value| value.trim())
                .unwrap_or_default();
            let names: Vec<String> = writer_columns
                .iter()
                .filter_map(|column| row.get(*column))
                .map(|name| normalize_person_name(name))
                .filter(|name| !name.is_empty())
                .collect();
            if id.is_empty() {
                if !names.is_empty() {
                    return Err(format!(
                        "Exclusion spreadsheet row {} has names but no student ID.",
                        row_number + 2
                    ));
                }
                continue;
            }
            let forbidden = forbidden_rows.entry(id.to_string()).or_default();
            for name in names {
                if let Some(faculty_rows) = faculty_by_name.get(&name) {
                    forbidden.extend(faculty_rows);
                }
            }
        }
        if forbidden_rows.is_empty() {
            return Err("The exclusion spreadsheet contains no student IDs.".into());
        }
        Ok(Self {
            student_id_column: configuration.student_id_column.clone(),
            student_id: configuration.student_id.trim().into(),
            all_faculty_rows: index.entries.iter().map(|entry| entry.row_index).collect(),
            forbidden_rows,
        })
    }

    pub(crate) fn eligible_rows(
        &self,
        student_id: &str,
        allowed: Option<&HashSet<usize>>,
    ) -> HashSet<usize> {
        let mut eligible = allowed.unwrap_or(&self.all_faculty_rows).clone();
        if let Some(forbidden) = self.forbidden_rows.get(student_id.trim()) {
            eligible.retain(|row| !forbidden.contains(row));
        }
        eligible
    }
}

#[cfg(test)]
#[path = "tests/exclusions.rs"]
mod tests;
