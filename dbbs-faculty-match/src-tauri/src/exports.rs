use crate::models::MatchEntry;
use chrono::Utc;
use rust_xlsxwriter::{Format, Workbook};
use std::collections::HashSet;

pub(crate) fn build_matches_headers(
    student_headers: &[String],
    faculty_headers: &[String],
) -> Vec<String> {
    let mut headers: Vec<String> = Vec::new();
    headers.push("First reviewer".into());
    headers.push("Reviewer".into());
    headers.extend(student_headers.iter().cloned());
    headers.extend(faculty_headers.iter().cloned());
    headers.push("Similarity %".into());
    headers.push("Student rank".into());
    headers.push("Faculty rank".into());
    headers
}

pub(crate) fn build_matches_workbook(
    student_headers: &[String],
    student_summary_rows: &[Vec<String>],
    faculty_headers: &[String],
    match_entries: &[MatchEntry],
) -> Result<Vec<u8>, String> {
    let mut workbook = Workbook::new();
    let matches_sheet_name = "Matches";
    let matches_sheet = workbook.add_worksheet();
    matches_sheet
        .set_name(matches_sheet_name)
        .map_err(|err| format!("Unable to configure the matches worksheet: {err}"))?;

    let header_format = Format::new().set_bold();
    let percent_format = Format::new().set_num_format("0.0%");

    let headers = build_matches_headers(student_headers, faculty_headers);
    for (col_index, header) in headers.iter().enumerate() {
        matches_sheet
            .write_string_with_format(0, col_index as u16, header, &header_format)
            .map_err(|err| format!("Unable to write the matches header row: {err}"))?;
    }

    let student_offset = 2u32;
    let faculty_offset = student_offset + student_headers.len() as u32;
    let similarity_col = faculty_offset + faculty_headers.len() as u32;
    let student_rank_col = similarity_col + 1;
    let faculty_rank_col = student_rank_col + 1;

    for (row_index, entry) in match_entries.iter().enumerate() {
        let row = (row_index + 1) as u32;
        matches_sheet
            .write_string(row, 0, "")
            .map_err(|err| format!("Unable to write the first reviewer column: {err}"))?;
        matches_sheet
            .write_string(row, 1, "")
            .map_err(|err| format!("Unable to write the reviewer column: {err}"))?;

        for (offset, value) in entry.student_values.iter().enumerate() {
            matches_sheet
                .write_string(row, (student_offset + offset as u32) as u16, value)
                .map_err(|err| format!("Unable to write a student identifier value: {err}"))?;
        }

        for (offset, value) in entry.faculty_values.iter().enumerate() {
            matches_sheet
                .write_string(row, (faculty_offset + offset as u32) as u16, value)
                .map_err(|err| format!("Unable to write a faculty identifier value: {err}"))?;
        }

        if let Some(value) = entry.similarity {
            matches_sheet
                .write_number_with_format(
                    row,
                    similarity_col as u16,
                    f64::from(value),
                    &percent_format,
                )
                .map_err(|err| format!("Unable to write the similarity percentage: {err}"))?;
        } else {
            matches_sheet
                .write_string(row, similarity_col as u16, "")
                .map_err(|err| format!("Unable to write the similarity placeholder: {err}"))?;
        }

        if let Some((position, total)) = entry.student_rank {
            let text = match total {
                Some(limit) => format!("{position} of {limit}"),
                None => position.to_string(),
            };
            matches_sheet
                .write_string(row, student_rank_col as u16, text)
                .map_err(|err| format!("Unable to write the student rank: {err}"))?;
        } else {
            matches_sheet
                .write_string(row, student_rank_col as u16, "")
                .map_err(|err| format!("Unable to write the student rank placeholder: {err}"))?;
        }

        if let Some(rank) = entry.faculty_rank {
            matches_sheet
                .write_number(row, faculty_rank_col as u16, rank as f64)
                .map_err(|err| format!("Unable to write the faculty rank: {err}"))?;
        } else {
            matches_sheet
                .write_string(row, faculty_rank_col as u16, "")
                .map_err(|err| format!("Unable to write the faculty rank placeholder: {err}"))?;
        }
    }

    let match_row_count = match_entries.len() as u32;
    let mut student_summary_headers = student_headers.to_vec();
    student_summary_headers.push("Total first reviewers".into());
    student_summary_headers.push("Total reviewers".into());

    let student_summary_sheet = workbook.add_worksheet();
    student_summary_sheet
        .set_name("Student Summary")
        .map_err(|err| format!("Unable to configure the student summary worksheet: {err}"))?;
    for (col_index, header) in student_summary_headers.iter().enumerate() {
        student_summary_sheet
            .write_string_with_format(0, col_index as u16, header, &header_format)
            .map_err(|err| format!("Unable to write the student summary header row: {err}"))?;
    }

    let first_reviewer_range = if match_row_count > 0 {
        Some(excel_range_reference(
            matches_sheet_name,
            1,
            0,
            match_row_count,
            0,
        ))
    } else {
        None
    };
    let reviewer_range = if match_row_count > 0 {
        Some(excel_range_reference(
            matches_sheet_name,
            1,
            1,
            match_row_count,
            1,
        ))
    } else {
        None
    };

    for (row_index, identifiers) in student_summary_rows.iter().enumerate() {
        let row = (row_index + 1) as u32;
        for (col_offset, value) in identifiers.iter().enumerate() {
            student_summary_sheet
                .write_string(row, col_offset as u16, value)
                .map_err(|err| {
                    format!("Unable to write a student summary identifier value: {err}")
                })?;
        }

        let first_col = student_headers.len();
        let total_col = first_col + 1;

        if match_row_count == 0 {
            student_summary_sheet
                .write_number(row, first_col as u16, 0.0)
                .map_err(|err| {
                    format!("Unable to write the student first reviewer placeholder: {err}")
                })?;
            student_summary_sheet
                .write_number(row, total_col as u16, 0.0)
                .map_err(|err| {
                    format!("Unable to write the student reviewer count placeholder: {err}")
                })?;
            continue;
        }

        let mut first_factors = Vec::new();
        first_factors.push(format!(
            "--({first}=1)",
            first = first_reviewer_range.as_ref().unwrap()
        ));
        for (col_offset, _) in student_headers.iter().enumerate() {
            let student_range = excel_range_reference(
                matches_sheet_name,
                1,
                student_offset + col_offset as u32,
                match_row_count,
                student_offset + col_offset as u32,
            );
            let summary_cell = excel_cell_reference(row, col_offset as u32, true, false);
            first_factors.push(format!("--({student_range}={summary_cell})"));
        }
        let first_formula = build_sumproduct_formula(&first_factors);
        student_summary_sheet
            .write_formula(row, first_col as u16, first_formula.as_str())
            .map_err(|err| format!("Unable to write the student first reviewer formula: {err}"))?;

        let mut total_factors = Vec::new();
        total_factors.push(format!(
            "--((( {first}=1)+({reviewer}=1))>0)",
            first = first_reviewer_range.as_ref().unwrap(),
            reviewer = reviewer_range.as_ref().unwrap()
        ));
        for (col_offset, _) in student_headers.iter().enumerate() {
            let student_range = excel_range_reference(
                matches_sheet_name,
                1,
                student_offset + col_offset as u32,
                match_row_count,
                student_offset + col_offset as u32,
            );
            let summary_cell = excel_cell_reference(row, col_offset as u32, true, false);
            total_factors.push(format!("--({student_range}={summary_cell})"));
        }
        let total_formula = build_sumproduct_formula(&total_factors);
        student_summary_sheet
            .write_formula(row, total_col as u16, total_formula.as_str())
            .map_err(|err| format!("Unable to write the student reviewer count formula: {err}"))?;
    }

    let mut faculty_summary_headers = faculty_headers.to_vec();
    faculty_summary_headers.push("First reviewer count".into());
    faculty_summary_headers.push("Total reviewer count".into());

    let faculty_summary_sheet = workbook.add_worksheet();
    faculty_summary_sheet
        .set_name("Faculty Summary")
        .map_err(|err| format!("Unable to configure the faculty summary worksheet: {err}"))?;
    for (col_index, header) in faculty_summary_headers.iter().enumerate() {
        faculty_summary_sheet
            .write_string_with_format(0, col_index as u16, header, &header_format)
            .map_err(|err| format!("Unable to write the faculty summary header row: {err}"))?;
    }

    let mut seen_faculty = HashSet::new();
    let mut faculty_summary_rows: Vec<Vec<String>> = Vec::new();
    for entry in match_entries {
        if entry.faculty_rank.is_none() {
            continue;
        }
        let key = entry.faculty_values.join("\u{1f}");
        if seen_faculty.insert(key) {
            faculty_summary_rows.push(entry.faculty_values.clone());
        }
    }

    for (row_index, identifiers) in faculty_summary_rows.iter().enumerate() {
        let row = (row_index + 1) as u32;
        for (col_offset, value) in identifiers.iter().enumerate() {
            faculty_summary_sheet
                .write_string(row, col_offset as u16, value)
                .map_err(|err| {
                    format!("Unable to write a faculty summary identifier value: {err}")
                })?;
        }

        let first_col = faculty_headers.len();
        let total_col = first_col + 1;

        if match_row_count == 0 {
            faculty_summary_sheet
                .write_number(row, first_col as u16, 0.0)
                .map_err(|err| {
                    format!("Unable to write the faculty first reviewer placeholder: {err}")
                })?;
            faculty_summary_sheet
                .write_number(row, total_col as u16, 0.0)
                .map_err(|err| {
                    format!("Unable to write the faculty reviewer placeholder: {err}")
                })?;
            continue;
        }

        let mut first_factors = Vec::new();
        first_factors.push(format!(
            "--({first}=1)",
            first = first_reviewer_range.as_ref().unwrap()
        ));
        for (col_offset, _) in faculty_headers.iter().enumerate() {
            let faculty_range = excel_range_reference(
                matches_sheet_name,
                1,
                faculty_offset + col_offset as u32,
                match_row_count,
                faculty_offset + col_offset as u32,
            );
            let summary_cell = excel_cell_reference(row, col_offset as u32, true, false);
            first_factors.push(format!("--({faculty_range}={summary_cell})"));
        }
        let first_formula = build_sumproduct_formula(&first_factors);
        faculty_summary_sheet
            .write_formula(row, first_col as u16, first_formula.as_str())
            .map_err(|err| format!("Unable to write the faculty first reviewer formula: {err}"))?;

        let mut total_factors = Vec::new();
        total_factors.push(format!(
            "--((( {first}=1)+({reviewer}=1))>0)",
            first = first_reviewer_range.as_ref().unwrap(),
            reviewer = reviewer_range.as_ref().unwrap()
        ));
        for (col_offset, _) in faculty_headers.iter().enumerate() {
            let faculty_range = excel_range_reference(
                matches_sheet_name,
                1,
                faculty_offset + col_offset as u32,
                match_row_count,
                faculty_offset + col_offset as u32,
            );
            let summary_cell = excel_cell_reference(row, col_offset as u32, true, false);
            total_factors.push(format!("--({faculty_range}={summary_cell})"));
        }
        let total_formula = build_sumproduct_formula(&total_factors);
        faculty_summary_sheet
            .write_formula(row, total_col as u16, total_formula.as_str())
            .map_err(|err| format!("Unable to write the faculty reviewer formula: {err}"))?;
    }

    workbook
        .save_to_buffer()
        .map_err(|err| format!("Unable to finalize the match workbook: {err}"))
}

fn build_sumproduct_formula(factors: &[String]) -> String {
    if factors.is_empty() {
        "=0".into()
    } else {
        format!("=SUMPRODUCT({})", factors.join(", "))
    }
}

fn excel_column_name(mut index: u32) -> String {
    let mut name = String::new();
    loop {
        let remainder = index % 26;
        name.push((b'A' + remainder as u8) as char);
        if index < 26 {
            break;
        }
        index = index / 26 - 1;
    }
    name.chars().rev().collect()
}

fn excel_cell_reference(row: u32, col: u32, absolute_col: bool, absolute_row: bool) -> String {
    let mut reference = String::new();
    if absolute_col {
        reference.push('$');
    }
    reference.push_str(&excel_column_name(col));
    if absolute_row {
        reference.push('$');
    }
    reference.push_str(&(row + 1).to_string());
    reference
}

fn excel_range_reference(
    sheet: &str,
    start_row: u32,
    start_col: u32,
    end_row: u32,
    end_col: u32,
) -> String {
    let escaped_sheet = sheet.replace('\'', "''");
    let start = excel_cell_reference(start_row, start_col, true, true);
    let end = excel_cell_reference(end_row, end_col, true, true);
    format!("'{}'!{}:{}", escaped_sheet, start, end)
}

pub(crate) fn default_matches_workbook_name() -> String {
    let timestamp = Utc::now().format("%Y%m%d-%H%M%S");
    format!("DBBS_matches_{timestamp}.xlsx")
}

pub(crate) fn default_directory_workbook_name() -> String {
    let timestamp = Utc::now().format("%Y%m%d-%H%M%S");
    format!("DBBS_directory_matches_{timestamp}.xlsx")
}

pub(crate) fn format_similarity_percent(value: f32) -> String {
    if value.is_finite() {
        format!("{:.1}%", value * 100.0)
    } else {
        "n/a".into()
    }
}

#[cfg(test)]
#[path = "tests/exports.rs"]
mod tests;
