use super::*;
use crate::test_support::TestDirectory;

#[test]
fn validates_file_kind_extension_and_unicode_paths() {
    let directory = TestDirectory::new();
    let file = directory.write("René student.PDF", b"%PDF");
    let value = file.to_string_lossy().into_owned();
    assert_eq!(
        validate_upload_path(value.clone(), UploadKind::Document).unwrap(),
        value
    );
    assert!(validate_upload_path(value, UploadKind::Directory).is_err());
    assert!(validate_upload_path(
        directory.0.to_string_lossy().into_owned(),
        UploadKind::Document
    )
    .is_err());
    assert!(validate_upload_path(
        directory
            .write("bad.exe", b"text")
            .to_string_lossy()
            .into_owned(),
        UploadKind::Document
    )
    .is_err());
    assert!(validate_upload_path(
        directory
            .0
            .join("missing.pdf")
            .to_string_lossy()
            .into_owned(),
        UploadKind::Document
    )
    .is_err());
    assert!(validate_upload_path(" ".into(), UploadKind::Document).is_err());
}

#[test]
fn accepts_a_folder_with_spaces_and_an_extension() {
    let directory = TestDirectory::new();
    let folder = directory.0.join("Student submissions.v2");
    fs::create_dir(&folder).unwrap();
    let value = folder.to_string_lossy().into_owned();
    assert_eq!(
        validate_upload_path(value.clone(), UploadKind::Directory).unwrap(),
        value
    );
}

#[test]
fn saves_base64_workbooks_and_utf8_results_without_corruption() {
    let directory = TestDirectory::new();
    let bytes = b"PK\x03\x04\0\xffworkbook";
    let binary = directory.0.join("matches.xlsx");
    save_generated_spreadsheet(
        binary.to_string_lossy().into_owned(),
        Base64Engine.encode(bytes),
        Some("base64".into()),
    )
    .unwrap();
    assert_eq!(fs::read(binary).unwrap(), bytes);
    let text = directory.0.join("matches.tsv");
    save_generated_spreadsheet(
        text.to_string_lossy().into_owned(),
        "Name\tScore\nRené\t90".into(),
        None,
    )
    .unwrap();
    assert_eq!(fs::read_to_string(text).unwrap(), "Name\tScore\nRené\t90");
    assert!(save_generated_spreadsheet(
        directory.0.join("bad.xlsx").to_string_lossy().into_owned(),
        "!invalid!".into(),
        Some("base64".into())
    )
    .is_err());
}
