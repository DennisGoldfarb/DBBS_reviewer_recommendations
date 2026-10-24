use super::*;
use crate::models::FacultyEmbeddingEntry;
use crate::scoring::find_best_faculty_matches;
use crate::test_support::TestDirectory;

fn faculty_index() -> FacultyEmbeddingIndex {
    serde_json::from_value(serde_json::json!({
        "model": "test", "dimension": 2, "embeddingColumns": [],
        "identifierColumns": ["Name"], "entries": []
    }))
    .map(|mut index: FacultyEmbeddingIndex| {
        index.entries = ["Dietmann, Sabine", " SABINE  DIETMANN ", "Other, Faculty"]
            .iter()
            .enumerate()
            .map(|(row_index, name)| FacultyEmbeddingEntry {
                row_index,
                embedding: vec![1.0, row_index as f32 * 0.1],
                identifiers: HashMap::from([("Name".into(), name.to_string())]),
                preparation: None,
            })
            .collect();
        index
    })
    .unwrap()
}

fn configuration(directory: &TestDirectory) -> FacultyExclusionConfig {
    let path = directory.write("writers.tsv", "Person Reference ID\tRef 1 Name\tRef 2 Name\n00123\tSabine  Dietmann\tExternal Person\n00123\tDietmann, Sabine\t\n00456\tFaculty Other\t\n");
    FacultyExclusionConfig {
        path: path.to_string_lossy().into(),
        exclusion_student_id_column: "Person Reference ID".into(),
        excluded_name_columns: vec!["Ref 1 Name".into(), "Ref 2 Name".into()],
        faculty_name_column: "Name".into(),
        student_id_column: "ID".into(),
        student_id: "00123".into(),
    }
}

#[test]
fn normalizes_first_last_order_case_and_whitespace_without_fuzzy_matching() {
    assert_eq!(
        normalize_person_name("  Dietmann,  Sabine "),
        "sabine dietmann"
    );
    assert_eq!(
        normalize_person_name(" SABINE\tDietmann "),
        "sabine dietmann"
    );
    assert_eq!(
        normalize_person_name("Smith-Jones, Anne Marie"),
        "anne marie smith-jones"
    );
    assert_ne!(
        normalize_person_name("Smith, A."),
        normalize_person_name("Anne Smith")
    );
    assert_eq!(normalize_person_name("  "), "");
    assert_ne!(
        normalize_person_name("Fuhai Li"),
        normalize_person_name("Li Fuhai")
    );
}

#[test]
fn excludes_all_duplicate_faculty_before_selecting_top_results_for_each_student() {
    let directory = TestDirectory::new();
    let index = faculty_index();
    let rules = FacultyExclusions::load(&configuration(&directory), &index, &TaskType::Spreadsheet)
        .unwrap();
    let eligible = rules.eligible_rows(" 00123 ", None);
    assert_eq!(eligible, HashSet::from([2]));
    let matches = find_best_faculty_matches(&index, &[vec![1.0, 0.0]], 1, Some(&eligible));
    assert_eq!(matches.len(), 1);
    assert_eq!(matches[0].row_index, 2);
    assert_eq!(rules.eligible_rows("00456", None), HashSet::from([0, 1]));
    // IDs remain strings, including leading zeros. Unlisted students retain all candidates.
    assert_eq!(rules.eligible_rows("123", None), HashSet::from([0, 1, 2]));
    let allowed = HashSet::from([0, 1]);
    assert!(rules.eligible_rows("00123", Some(&allowed)).is_empty());
}

#[test]
fn validates_column_mappings_and_required_student_identity() {
    let directory = TestDirectory::new();
    let index = faculty_index();
    let mut config = configuration(&directory);
    config.student_id_column.clear();
    assert!(FacultyExclusions::load(&config, &index, &TaskType::Spreadsheet).is_err());
    config.student_id.clear();
    assert!(FacultyExclusions::load(&config, &index, &TaskType::Prompt).is_err());
    config.faculty_name_column = "Unknown".into();
    assert!(FacultyExclusions::load(&config, &index, &TaskType::Directory).is_err());
    config.faculty_name_column = "Name".into();
    config.excluded_name_columns = vec!["Missing".into()];
    assert!(FacultyExclusions::load(&config, &index, &TaskType::Directory).is_err());
}

#[test]
fn rejects_letter_writer_rows_without_applicant_ids() {
    let directory = TestDirectory::new();
    let mut config = configuration(&directory);
    directory.write(
        "writers.tsv",
        "Person Reference ID\tRef 1 Name\tRef 2 Name\n\tSabine Dietmann\t\n",
    );
    assert!(FacultyExclusions::load(&config, &faculty_index(), &TaskType::Spreadsheet).is_err());
    config.path = directory
        .write("empty.tsv", "Person Reference ID\tRef 1 Name\tRef 2 Name\n")
        .to_string_lossy()
        .into();
    assert!(FacultyExclusions::load(&config, &faculty_index(), &TaskType::Spreadsheet).is_err());
}

#[test]
fn reads_xlsx_writer_pairs_and_preserves_string_applicant_ids() {
    let directory = TestDirectory::new();
    let mut config = configuration(&directory);
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let sheet = workbook.add_worksheet();
    for (row, values) in [
        ["Ref 1 Name", "Ref 2 Name", "Person Reference ID"],
        ["Sabine  Dietmann", "External Writer", "00123"],
        ["Faculty Other", "", "00456"],
    ]
    .iter()
    .enumerate()
    {
        for (column, value) in values.iter().enumerate() {
            sheet
                .write_string(row as u32, column as u16, *value)
                .unwrap();
        }
    }
    let path = directory.0.join("writers.xlsx");
    workbook.save(&path).unwrap();
    config.path = path.to_string_lossy().into();
    let rules = FacultyExclusions::load(&config, &faculty_index(), &TaskType::Spreadsheet).unwrap();
    assert_eq!(rules.eligible_rows("00123", None), HashSet::from([2]));
    assert_eq!(rules.eligible_rows("00456", None), HashSet::from([0, 1]));
}
