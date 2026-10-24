use super::*;
use crate::models::FacultyEmbeddingEntry;

fn index() -> FacultyEmbeddingIndex {
    FacultyEmbeddingIndex {
        model: "test".into(),
        generated_at: None,
        dimension: 2,
        total_rows: None,
        embedded_rows: None,
        skipped_rows: None,
        embedding_columns: vec![],
        identifier_columns: vec![],
        entries: vec![
            (0, vec![1.0, 0.0]),
            (1, vec![0.0, 1.0]),
            (2, vec![0.8, 0.2]),
        ]
        .into_iter()
        .map(|(row_index, embedding)| FacultyEmbeddingEntry {
            row_index,
            embedding,
            identifiers: HashMap::new(),
            preparation: None,
        })
        .collect(),
    }
}

#[test]
fn ranks_by_similarity_and_respects_faculty_filters_and_limit() {
    let index = index();
    let matches = find_best_faculty_matches(&index, &[vec![1.0, 0.0]], 2, None);
    assert_eq!(
        matches
            .iter()
            .map(|item| item.row_index)
            .collect::<Vec<_>>(),
        vec![0, 2]
    );
    let allowed = HashSet::from([1, 2]);
    let filtered = find_best_faculty_matches(&index, &[vec![1.0, 0.0]], 1, Some(&allowed));
    assert_eq!(filtered[0].row_index, 2);
    assert!(find_best_faculty_matches(&index, &[vec![1.0, 0.0]], 0, None).is_empty());
}

#[test]
fn invalid_and_zero_vectors_cannot_produce_matches() {
    for vector in [
        vec![],
        vec![0.0, 0.0],
        vec![f32::NAN, 1.0],
        vec![f32::INFINITY, 1.0],
        vec![1.0],
    ] {
        assert_eq!(cosine_similarity(&[1.0, 0.0], &vector), None);
    }
    assert_eq!(cosine_similarity(&[1.0, 0.0], &[1.0, 0.0]), Some(1.0));
    assert_eq!(cosine_similarity(&[1.0, 0.0], &[-1.0, 0.0]), Some(-1.0));
}

#[test]
fn assigns_student_rankings_for_each_faculty_across_prompts() {
    let index = index();
    let mut first = find_best_faculty_matches(&index, &[vec![1.0, 0.0]], 3, None);
    let mut second = find_best_faculty_matches(&index, &[vec![0.0, 1.0]], 3, None);
    assign_student_rankings(&mut [&mut first, &mut second]);
    assert_eq!(
        first
            .iter()
            .find(|item| item.row_index == 0)
            .unwrap()
            .student_rank_for_faculty,
        Some(1)
    );
    assert_eq!(
        second
            .iter()
            .find(|item| item.row_index == 0)
            .unwrap()
            .student_rank_for_faculty,
        Some(2)
    );
    assert!(first.iter().all(|item| item.student_rank_total == Some(2)));
}

#[test]
fn uses_each_facultys_best_chunk_and_returns_each_faculty_only_once() {
    let index = index();
    let chunks = vec![vec![1.0, 0.0], vec![0.0, 1.0], vec![1.0, 0.0]];
    let matches = find_best_faculty_matches(&index, &chunks, 10, None);
    assert_eq!(matches.len(), 3);
    assert_eq!(matches[0].row_index, 0);
    assert_eq!(matches[1].row_index, 1);
    assert_eq!(matches[0].similarity, 1.0);
    assert_eq!(matches[1].similarity, 1.0);
    let allowed = HashSet::from([1, 2]);
    let filtered = find_best_faculty_matches(&index, &chunks, 1, Some(&allowed));
    assert_eq!(filtered.len(), 1);
    assert_eq!(filtered[0].row_index, 1);
}

#[test]
fn ignores_invalid_chunks_but_retains_other_valid_matches() {
    let index = index();
    let mut chunks = vec![vec![], vec![0.0, 0.0], vec![f32::NAN, 1.0], vec![1.0]];
    assert!(find_best_faculty_matches(&index, &chunks, 10, None).is_empty());
    chunks.push(vec![0.0, 1.0]);
    let matches = find_best_faculty_matches(&index, &chunks, 1, None);
    assert_eq!(matches[0].row_index, 1);
    assert!(find_best_faculty_matches(&index, &[], 10, None).is_empty());
}

#[test]
fn negative_chunk_similarities_are_not_replaced_by_zero_or_averaged() {
    let index = index();
    let chunks = vec![vec![-1.0, 0.0], vec![0.0, -1.0]];
    let matches = find_best_faculty_matches(&index, &chunks, 3, None);
    let score = matches
        .iter()
        .find(|entry| entry.row_index == 2)
        .unwrap()
        .similarity;
    assert!(
        (score - cosine_similarity(&chunks[1], &index.entries[2].embedding).unwrap()).abs() < 1e-6
    );
    assert!(score < 0.0);
}

#[test]
fn reciprocal_rankings_count_students_instead_of_their_chunks() {
    let index = index();
    let mut long = find_best_faculty_matches(&index, &[vec![1.0, 0.0], vec![0.0, 1.0]], 3, None);
    let mut short = find_best_faculty_matches(&index, &[vec![0.0, 1.0]], 3, None);
    assign_student_rankings(&mut [&mut long, &mut short]);
    assert!(long
        .iter()
        .chain(short.iter())
        .all(|entry| entry.student_rank_total == Some(2)));
}
