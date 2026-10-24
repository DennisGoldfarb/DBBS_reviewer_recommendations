use super::*;
use crate::test_support::TestDirectory;

#[test]
fn extracts_unicode_plain_text_and_normalizes_newlines() {
    let directory = TestDirectory::new();
    let path = directory.write(
        "research.txt",
        "  René studies proteins.\r\n\r\nStructure and signaling.  ",
    );
    let extracted = extract_document_prompt(&path).unwrap();
    assert!(extracted.text.contains("René studies proteins."));
    assert!(!extracted.text.contains('\r'));
    assert!(extracted.warnings.is_empty());
}

#[test]
fn extracts_docx_paragraphs() {
    let directory = TestDirectory::new();
    let path = directory.0.join("research.docx");
    docx_rs::Docx::new()
        .add_paragraph(Paragraph::new().add_run(Run::new().add_text("Protein structure research")))
        .build()
        .pack(fs::File::create(&path).unwrap())
        .unwrap();
    assert!(extract_document_prompt(&path)
        .unwrap()
        .text
        .contains("Protein structure research"));
}

#[test]
fn reports_invalid_pdf_instead_of_treating_binary_as_text() {
    let directory = TestDirectory::new();
    let path = directory.write("invalid.pdf", b"this is not a PDF");
    assert!(extract_document_prompt(&path).is_err());
}
