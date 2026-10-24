use crate::models::{FacultyEmbeddingIndex, FacultyMatchResult};
use std::cmp::Ordering;
use std::collections::HashMap;
use std::collections::HashSet;

pub(crate) fn find_best_faculty_matches(
    index: &FacultyEmbeddingIndex,
    prompt_embeddings: &[Vec<f32>],
    limit: usize,
    allowed_rows: Option<&HashSet<usize>>,
) -> Vec<FacultyMatchResult> {
    if limit == 0 {
        return Vec::new();
    }

    let mut candidates: Vec<FacultyMatchResult> = index
        .entries
        .iter()
        .filter_map(|entry| {
            if let Some(allowed) = allowed_rows {
                if !allowed.contains(&entry.row_index) {
                    return None;
                }
            }

            let similarity = prompt_embeddings
                .iter()
                .filter_map(|chunk| cosine_similarity(chunk, &entry.embedding))
                .reduce(f32::max)?;
            let mut identifiers = entry.identifiers.clone();
            identifiers.retain(|_, value| !value.trim().is_empty());

            Some(FacultyMatchResult {
                row_index: entry.row_index,
                similarity,
                identifiers,
                faculty_text: None,
                student_rank_for_faculty: None,
                student_rank_total: None,
            })
        })
        .collect();

    candidates.sort_by(|a, b| {
        b.similarity
            .partial_cmp(&a.similarity)
            .unwrap_or(Ordering::Equal)
    });
    candidates.truncate(limit);
    candidates
}

pub(crate) fn assign_student_rankings(match_sets: &mut [&mut Vec<FacultyMatchResult>]) {
    if match_sets.is_empty() {
        return;
    }

    let mut occurrences: HashMap<usize, Vec<(usize, usize, f32)>> = HashMap::new();

    for (prompt_index, matches) in match_sets.iter().enumerate() {
        for (match_index, faculty) in matches.iter().enumerate() {
            occurrences.entry(faculty.row_index).or_default().push((
                prompt_index,
                match_index,
                faculty.similarity,
            ));
        }
    }

    if occurrences.is_empty() {
        for matches in match_sets.iter_mut() {
            for faculty in matches.iter_mut() {
                faculty.student_rank_for_faculty = None;
                faculty.student_rank_total = None;
            }
        }
        return;
    }

    let mut rank_map: HashMap<(usize, usize), (usize, usize)> = HashMap::new();

    for (_, mut entries) in occurrences {
        entries.sort_by(|a, b| {
            let primary = b.2.partial_cmp(&a.2).unwrap_or(Ordering::Equal);
            if primary != Ordering::Equal {
                return primary;
            }

            let secondary = a.0.cmp(&b.0);
            if secondary != Ordering::Equal {
                return secondary;
            }

            a.1.cmp(&b.1)
        });

        let total = entries.len();
        for (position, (prompt_index, match_index, _)) in entries.into_iter().enumerate() {
            rank_map.insert((prompt_index, match_index), (position + 1, total));
        }
    }

    for (prompt_index, matches) in match_sets.iter_mut().enumerate() {
        for (match_index, faculty) in matches.iter_mut().enumerate() {
            if let Some((rank, total)) = rank_map.get(&(prompt_index, match_index)) {
                faculty.student_rank_for_faculty = Some(*rank);
                faculty.student_rank_total = Some(*total);
            } else {
                faculty.student_rank_for_faculty = None;
                faculty.student_rank_total = None;
            }
        }
    }
}

fn cosine_similarity(a: &[f32], b: &[f32]) -> Option<f32> {
    if a.len() != b.len() || a.is_empty() {
        return None;
    }

    let mut dot = 0.0f64;
    let mut norm_a = 0.0f64;
    let mut norm_b = 0.0f64;

    for (&x, &y) in a.iter().zip(b.iter()) {
        let xf = f64::from(x);
        let yf = f64::from(y);
        dot += xf * yf;
        norm_a += xf * xf;
        norm_b += yf * yf;
    }

    if norm_a == 0.0 || norm_b == 0.0 {
        return None;
    }

    let similarity = dot / (norm_a.sqrt() * norm_b.sqrt());
    similarity
        .is_finite()
        .then_some(similarity.clamp(-1.0, 1.0) as f32)
}

#[cfg(test)]
#[path = "tests/scoring.rs"]
mod tests;
