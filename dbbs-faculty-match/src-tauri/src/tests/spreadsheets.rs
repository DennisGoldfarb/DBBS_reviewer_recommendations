use super::*;
use crate::test_support::TestDirectory;

#[test]
fn reads_csv_with_quoted_commas_and_tab_delimited_text() {
    let directory = TestDirectory::new();
    for (filename, content) in [
        (
            "students.csv",
            "Name,Research interests\nRené,\"Protein structure, signaling\"\n",
        ),
        (
            "students.tsv",
            "Name\tResearch interests\nRené\tProtein structure, signaling\n",
        ),
        (
            "students.txt",
            "Name\tResearch interests\nRené\tProtein structure, signaling\n",
        ),
    ] {
        let path = directory.write(filename, content);
        let (headers, rows) = read_full_spreadsheet(&path).unwrap();
        assert_eq!(headers, vec!["Name", "Research interests"]);
        assert_eq!(rows, vec![vec!["René", "Protein structure, signaling"]]);
        assert_eq!(
            suggest_spreadsheet_columns(&headers, &rows),
            (vec![1], vec![0])
        );
    }
}

#[test]
fn preview_limits_do_not_truncate_full_matching_inputs() {
    let directory = TestDirectory::new();
    let mut content = "Name\tResearch interests\n".to_string();
    for index in 0..15 {
        content.push_str(&format!("Student {index}\tProtein structure\n"));
    }
    let path = directory.write("students.tsv", content);
    assert_eq!(read_spreadsheet(&path).unwrap().1.len(), 10);
    assert_eq!(read_full_spreadsheet(&path).unwrap().1.len(), 15);
}

#[test]
fn aligns_uneven_rows_and_supplies_missing_headers() {
    let mut headers = vec!["Name".into()];
    let mut rows = vec![vec!["A".into(), "Research".into()], vec!["B".into()]];
    align_row_lengths(&mut headers, &mut rows);
    assert_eq!(headers, vec!["Name", ""]);
    assert_eq!(rows[1], vec!["B", ""]);
}

#[test]
fn normalizes_identifier_whitespace_and_label_punctuation() {
    assert_eq!(
        normalize_identifier_value("  René   Faculty \n"),
        "rené faculty"
    );
    assert_eq!(normalize_identifier_label("Faculty E-mail"), "facultyemail");
}
