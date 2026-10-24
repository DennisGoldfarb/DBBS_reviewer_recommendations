use super::*;

fn index_payload(model: &str) -> Vec<u8> {
    serde_json::to_vec(&serde_json::json!({
        "model": model,
        "dimension": 768,
        "entries": []
    }))
    .unwrap()
}

#[test]
fn accepts_pubmedbert_and_legacy_empty_model_identifiers() {
    for model in [DEFAULT_EMBEDDING_MODEL, ""] {
        let index = parse_faculty_embedding_index(&index_payload(model)).unwrap();
        assert_eq!(index.model, DEFAULT_EMBEDDING_MODEL);
    }
}

#[test]
fn experimental_index_requires_refresh_before_any_matching() {
    let error = parse_faculty_embedding_index(&index_payload("other/model")).unwrap_err();
    assert!(error.contains("Update faculty embeddings"));
    assert!(error.contains("PubMedBERT"));
}

#[test]
fn malformed_index_keeps_an_actionable_parse_error() {
    let error = parse_faculty_embedding_index(b"not-json").unwrap_err();
    assert!(error.starts_with("Unable to parse faculty embeddings:"));
}
