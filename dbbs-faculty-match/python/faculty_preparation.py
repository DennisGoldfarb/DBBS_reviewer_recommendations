"""Prune faculty descriptions locally with the small LLMLingua-2 classifier."""

from bisect import bisect_right
import gc
import math
import os
import re

from embedding_config import MAX_INPUT_TOKENS

PRUNING_MODEL = "microsoft/llmlingua-2-bert-base-multilingual-cased-meetingbank"
PRUNING_REVISION = "5f0c82792b7ea14c6484e015b6a072009496b7f2"
PREPARATION_VERSION = 5
PRUNING_WINDOW_TOKENS = 512
PRUNING_WINDOW_OVERLAP = 64


def count_tokens(tokenizer, text):
    return len(tokenizer.encode(text, add_special_tokens=False, truncation=False))


def remove_lab_website(text):
    return re.sub(r"(?i)\bLab[ \t]+Website", "", text).strip()


class LocalWordPruner:
    """Lazy CPU-only BERT classifier; score the entire source in overlapping windows."""

    def __init__(self):
        self.tokenizer = None
        self.model = None
        self.previous_threads = None

    def _load(self):
        if self.model is not None:
            return
        import torch
        from transformers import AutoModelForTokenClassification, AutoTokenizer

        self.previous_threads = torch.get_num_threads()
        torch.set_num_threads(min(2, os.cpu_count() or 1))
        self.tokenizer = AutoTokenizer.from_pretrained(
            PRUNING_MODEL, revision=PRUNING_REVISION, use_fast=True
        )
        self.model = AutoModelForTokenClassification.from_pretrained(
            PRUNING_MODEL, revision=PRUNING_REVISION,
            torch_dtype=torch.float32, use_safetensors=True,
        ).to("cpu").eval()

    def score_words(self, text, spans):
        import torch

        self._load()
        ends = [span.end() for span in spans]
        totals = [0.0] * len(spans)
        counts = [0] * len(spans)
        inputs = self.tokenizer(
            text, return_offsets_mapping=True, return_overflowing_tokens=True,
            max_length=PRUNING_WINDOW_TOKENS, truncation=True,
            stride=PRUNING_WINDOW_OVERLAP,
        )
        with torch.inference_mode():
            for index, offsets in enumerate(inputs["offset_mapping"]):
                batch = {
                    key: torch.tensor([values[index]], dtype=torch.long)
                    for key, values in inputs.items()
                    if key in ("input_ids", "attention_mask", "token_type_ids")
                }
                # LLMLingua-2's class 1 is the probability of preserving a token.
                probabilities = self.model(**batch).logits.softmax(dim=-1)[0, :, 1].tolist()
                for (start, end), probability in zip(offsets, probabilities):
                    if end <= start:
                        continue
                    word = bisect_right(ends, start)
                    if word < len(spans):
                        totals[word] += probability
                        counts[word] += 1
        if not all(counts):
            raise ValueError("The pruning tokenizer did not score every source word.")
        return [total / count for total, count in zip(totals, counts)]

    def close(self):
        self.model = None
        self.tokenizer = None
        gc.collect()
        if self.previous_threads is not None:
            import torch
            torch.set_num_threads(self.previous_threads)
            self.previous_threads = None


def prune_to_budget(text, tokenizer, pruner, budget):
    """Delete low-scoring whole words, retaining source order and the actual token budget.

    Words are selected entirely by LLMLingua-2 scores, without manual term
    protection or a stopword list. Keep whole words in their source order.
    """
    spans = list(re.finditer(r"\S+", text))
    scores = pruner.score_words(text, spans)
    if len(scores) != len(spans) or any(not math.isfinite(score) for score in scores):
        raise ValueError("The pruning model returned invalid word scores.")
    order = sorted(range(len(spans)), key=lambda index: (scores[index], -index))
    kept = [True] * len(spans)
    edited, tokens = text, count_tokens(tokenizer, text)
    for index in order:
        if tokens <= budget:
            break
        kept[index] = False
        edited = " ".join(span.group() for i, span in enumerate(spans) if kept[i])
        tokens = count_tokens(tokenizer, edited)
    if not edited or not 0 < tokens <= budget:
        raise ValueError("Unable to retain faculty topic text within the embedding limit.")
    return edited, tokens


def prepare_faculty_texts(texts, tokenizer, emit_progress, pruner_factory=LocalWordPruner):
    """Remove website footers, then prune only descriptions exceeding the input window."""
    budget = MAX_INPUT_TOKENS - tokenizer.num_special_tokens_to_add(pair=False)
    prepared = {}
    overlong = []
    for row in texts:
        original = row["text"]
        cleaned = remove_lab_website(original)
        original_tokens = count_tokens(tokenizer, original)
        tokens = count_tokens(tokenizer, cleaned)
        if tokens > budget:
            overlong.append((row, cleaned, original_tokens))
        elif cleaned != original:
            prepared[row["id"]] = {
                "text": cleaned, "originalTokens": original_tokens, "tokens": tokens,
                "truncated": False, "model": "deterministic-text-cleanup",
                "revision": str(PREPARATION_VERSION),
            }
    if not overlong:
        return prepared

    emit_progress({
        "phase": "loading-pruning-model",
        "message": "Loading the small local LLMLingua-2 model; first use may download model files…",
        "processedRows": 0, "totalRows": len(overlong),
    })
    pruner = pruner_factory()
    try:
        for position, (row, cleaned, original_tokens) in enumerate(overlong):
            emit_progress({
                "phase": "pruning",
                "message": f"Pruning faculty topic text: {position + 1} of {len(overlong)}…",
                "processedRows": position, "totalRows": len(overlong),
            })
            try:
                text, tokens = prune_to_budget(cleaned, tokenizer, pruner, budget)
            except Exception as exc:
                raise RuntimeError(
                    f"Unable to prune faculty row {row['id'] + 1} locally: {exc}"
                ) from exc
            prepared[row["id"]] = {
                "text": text, "originalTokens": original_tokens, "tokens": tokens,
                "truncated": False, "model": PRUNING_MODEL, "revision": PRUNING_REVISION,
            }
        emit_progress({
            "phase": "pruning", "message": "Finished pruning long faculty descriptions.",
            "processedRows": len(overlong), "totalRows": len(overlong),
        })
    finally:
        pruner.close()
    return prepared
