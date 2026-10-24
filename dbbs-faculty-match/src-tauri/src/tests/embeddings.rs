use super::*;

#[test]
fn helper_reports_model_or_preparation_errors_directly() {
    let result = parse_embedding_response(
        br#"{"error":"Unable to shorten faculty row 3 locally: model unavailable"}"#,
    );
    assert_eq!(
        result.err().unwrap(),
        "Unable to shorten faculty row 3 locally: model unavailable"
    );
}

#[test]
fn parses_legacy_embeddings_without_preparation_metadata() {
    let response = parse_embedding_response(
        br#"{"model":"test","dimension":2,"rows":[{"id":1,"embedding":[0.1,0.2]}]}"#,
    )
    .unwrap();
    assert!(response.rows[0].preparation.is_none());
    assert!(response.rows[0].additional_embeddings.is_empty());
}

#[test]
fn retains_the_text_and_model_used_to_prepare_faculty_embeddings() {
    let response = parse_embedding_response(br#"{"model":"test","dimension":2,"rows":[{"id":1,"embedding":[0.1,0.2],"preparation":{"text":"Research summary","originalTokens":600,"tokens":200,"model":"local-summary-model","revision":"pinned-revision"}}]}"#).unwrap();
    let prepared = response.rows[0].preparation.as_ref().unwrap();
    assert_eq!(prepared.text, "Research summary");
    assert_eq!(prepared.original_tokens, 600);
    assert_eq!(prepared.tokens, 200);
    assert!(!prepared.truncated);
}

#[test]
fn packaged_faculty_index_matches_the_default_model_and_has_complete_vectors() {
    let index: FacultyEmbeddingIndex =
        serde_json::from_slice(crate::config::DEFAULT_FACULTY_EMBEDDINGS).unwrap();
    assert_eq!(index.entries.len(), 612);
    assert_eq!(index.model, crate::config::DEFAULT_EMBEDDING_MODEL);
    assert_eq!(index.dimension, 768);
    assert!(index
        .entries
        .iter()
        .all(|entry| entry.embedding.len() == index.dimension
            && entry.embedding.iter().all(|value| value.is_finite())));
}

#[test]
fn retains_cleanup_and_truncation_metadata_without_a_summary_cache() {
    let response = parse_embedding_response(br#"{"model":"test","dimension":2,"rows":[{"id":1,"embedding":[0.1,0.2],"preparation":{"text":"Original research wording","originalTokens":700,"tokens":600,"truncated":true,"model":"deterministic-text-cleanup","revision":"4"}}]}"#).unwrap();
    let prepared = response.rows[0].preparation.as_ref().unwrap();
    assert!(prepared.truncated);
    assert_eq!(prepared.tokens, 600);
    assert_eq!(prepared.revision, "4");
    assert_eq!(
        serde_json::to_value(crate::models::FacultyPreparationOptions {}).unwrap(),
        serde_json::json!({})
    );
}

#[test]
fn groups_all_chunk_embeddings_under_the_original_student_id() {
    let response = parse_embedding_response(br#"{"model":"test","dimension":2,"rows":[{"id":7,"embedding":[1.0,0.0],"additionalEmbeddings":[[0.0,1.0],[0.5,0.5]]},{"id":42,"embedding":[0.2,0.8]}]}"#).unwrap();
    let mut rows = response.rows.into_iter();
    let long = rows.next().unwrap();
    assert_eq!(long.id, 7);
    assert_eq!(
        long.into_embeddings(),
        vec![vec![1.0, 0.0], vec![0.0, 1.0], vec![0.5, 0.5]]
    );
    let short = rows.next().unwrap();
    assert_eq!(short.id, 42);
    assert_eq!(short.into_embeddings(), vec![vec![0.2, 0.8]]);
}
