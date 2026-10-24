"""Keep all student text through overlapping token windows for local embeddings."""

from embedding_config import MAX_INPUT_TOKENS
CHUNK_OVERLAP = 50


def token_windows(token_ids, window_size, overlap=CHUNK_OVERLAP):
    """Yield bounded windows, stopping once the final source token is included."""
    if window_size <= 0 or not 0 <= overlap < window_size:
        raise ValueError("The chunk overlap must be smaller than the positive window size.")
    start = 0
    while start < len(token_ids):
        end = min(start + window_size, len(token_ids))
        yield token_ids[start:end]
        if end == len(token_ids):
            break
        start = end - overlap


def student_embedding_inputs(tokenizer, text):
    """Keep the short-text path; window long inputs without decoding their IDs."""
    ids = tokenizer.encode(text, add_special_tokens=False, truncation=False)
    budget = MAX_INPUT_TOKENS - tokenizer.num_special_tokens_to_add(pair=False)
    if len(ids) <= budget:
        yield tokenizer(
            text, return_tensors="pt", truncation=True,
            max_length=MAX_INPUT_TOKENS, padding=True,
        )
        return
    for window in token_windows(ids, budget):
        yield tokenizer.prepare_for_model(
            window, add_special_tokens=True, return_tensors="pt",
            return_attention_mask=True, truncation=False, padding=False,
            prepend_batch_axis=True,
        )
