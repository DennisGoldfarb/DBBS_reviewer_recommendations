use super::*;
use crate::test_support::TestDirectory;
use calamine::{open_workbook_auto, Reader};

#[test]
fn generated_workbook_can_be_reopened_with_matching_and_summary_sheets() {
    let directory = TestDirectory::new();
    let bytes = build_matches_workbook(
        &["Student".into()],
        &[vec!["Example student".into()]],
        &["Faculty".into()],
        &[MatchEntry {
            student_values: vec!["Example student".into()],
            faculty_values: vec!["Example faculty".into()],
            similarity: Some(0.9),
            student_rank: Some((1, Some(1))),
            faculty_rank: Some(1),
        }],
    )
    .unwrap();
    assert!(bytes.starts_with(b"PK"));
    let path = directory.write("matches.xlsx", bytes);
    let mut workbook = open_workbook_auto(path).unwrap();
    assert_eq!(
        workbook.sheet_names(),
        &["Matches", "Student Summary", "Faculty Summary"]
    );
    let matches = workbook.worksheet_range("Matches").unwrap().unwrap();
    assert!(matches.rows().any(|row| row
        .iter()
        .any(|cell| cell.get_string() == Some("Example faculty"))));
    assert!(workbook
        .worksheet_formula("Faculty Summary")
        .unwrap()
        .unwrap()
        .rows()
        .any(|row| row.iter().any(|value| value.contains("SUMPRODUCT"))));
}

#[test]
fn excel_references_cover_column_boundaries_and_quote_sheet_names() {
    for (index, expected) in [(0, "A"), (25, "Z"), (26, "AA"), (701, "ZZ"), (702, "AAA")] {
        assert_eq!(excel_column_name(index), expected);
    }
    assert_eq!(
        excel_range_reference("Student's matches", 1, 0, 5, 26),
        "'Student''s matches'!$A$2:$AA$6"
    );
}
