use super::*;
use crate::test_support::TestDirectory;

#[test]
fn candidate_analysis_validates_column_mapping_and_program_memberships() {
    let directory = TestDirectory::new();
    let source = directory.write(
        "faculty.tsv",
        "Name\tResearch interests\tProgram\nExample faculty\tProtein structure\tBiochemistry\n",
    );
    let configuration = FacultyDatasetColumnConfiguration {
        embedding_columns: vec![1],
        identifier_columns: vec![0],
        program_columns: vec![2],
    };
    let (analysis, memberships) = inspect_faculty_dataset(&source, Some(&configuration)).unwrap();
    assert_eq!(analysis.embedding_columns, vec!["Research interests"]);
    assert_eq!(analysis.available_programs, vec!["Biochemistry"]);
    assert_eq!(
        filter_faculty_rows_by_program(&memberships, &["biochemistry".into()]),
        HashSet::from([0])
    );
    assert!(filter_faculty_rows_by_program(&memberships, &["Other program".into()]).is_empty());
    let invalid = FacultyDatasetColumnConfiguration {
        embedding_columns: vec![99],
        ..configuration
    };
    assert!(inspect_faculty_dataset(&source, Some(&invalid)).is_err());
}

#[test]
fn rejects_empty_faculty_datasets_before_replacement() {
    let directory = TestDirectory::new();
    let source = directory.write("faculty.tsv", "Name\tResearch interests\n");
    assert!(inspect_faculty_dataset(&source, None).is_err());
}

#[test]
fn replacement_handles_the_active_file_as_its_own_source_and_cleans_staging() {
    let directory = TestDirectory::new();
    let active = directory.write("faculty.tsv", "Original content");
    persist_dataset_file(&active, &active).unwrap();
    assert_eq!(fs::read_to_string(&active).unwrap(), "Original content");
    let replacement = directory.write("replacement.tsv", "New content");
    persist_dataset_file(&replacement, &active).unwrap();
    assert_eq!(fs::read_to_string(&active).unwrap(), "New content");
    assert_eq!(fs::read_dir(&directory.0).unwrap().count(), 2);
}

#[test]
fn a_failed_replacement_leaves_the_original_file_intact() {
    let directory = TestDirectory::new();
    let active = directory.write("faculty.tsv", "Original content");
    assert!(persist_dataset_file(&directory.0.join("missing.tsv"), &active).is_err());
    assert_eq!(fs::read_to_string(&active).unwrap(), "Original content");
}
