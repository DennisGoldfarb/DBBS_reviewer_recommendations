import json
import os
import re
import sys
import time
from collections.abc import Mapping, Sequence

os.environ.setdefault("TOKENIZERS_PARALLELISM", "false")
os.environ.setdefault("USE_TF", "0")
os.environ.setdefault("USE_FLAX", "0")

from faculty_preparation import prepare_faculty_texts
from student_chunking import student_embedding_inputs
from embedding_config import MODEL_NAME, MODEL_REVISION, MAX_INPUT_TOKENS

_TOKENIZER = None
_MODEL = None


def load_runtime():
    """Load optional ML dependencies only when an embedding is requested."""
    try:
        import torch
        from transformers import AutoModel, AutoTokenizer
        from transformers.utils import logging as hf_logging
    except ImportError as exc:
        raise RuntimeError(
            "Install the 'transformers' and 'torch' packages to generate embeddings."
        ) from exc
    hf_logging.set_verbosity_error()
    return torch, AutoTokenizer, AutoModel


def sanitize_utf8_text(text: str) -> str:
    """Ensure a text string can be encoded as UTF-8, dropping bad words."""

    if not text:
        return ""

    cleaned_parts = []
    last_index = 0

    for match in re.finditer(r"\S+", text):
        prefix = text[last_index : match.start()]
        if prefix:
            cleaned_parts.append(prefix)

        word = match.group(0)
        try:
            word.encode("utf-8")
        except UnicodeEncodeError:
            pass
        else:
            cleaned_parts.append(word)

        last_index = match.end()

    suffix = text[last_index:]
    if suffix:
        cleaned_parts.append(suffix)

    cleaned = "".join(cleaned_parts)

    try:
        cleaned.encode("utf-8")
    except UnicodeEncodeError:
        cleaned = cleaned.encode("utf-8", "ignore").decode("utf-8", "ignore")

    return cleaned.strip()


def normalize_text_value(value) -> str:
    """Coerce arbitrary JSON payload values into a clean text string."""

    if value is None:
        return ""

    if isinstance(value, str):
        return sanitize_utf8_text(value.strip())

    if isinstance(value, (bytes, bytearray)):
        try:
            text = value.decode("utf-8", errors="surrogateescape")
        except Exception:  # noqa: BLE001
            text = value.decode("utf-8", "ignore")
        return sanitize_utf8_text(text.strip())

    if isinstance(value, Mapping):
        parts = []
        for key in ("text", "value", "content"):
            if key in value:
                part = normalize_text_value(value[key])
                if part:
                    parts.append(part)
        if not parts:
            for item in value.values():
                part = normalize_text_value(item)
                if part:
                    parts.append(part)
        return sanitize_utf8_text("\n\n".join(parts))

    if isinstance(value, Sequence) and not isinstance(value, (str, bytes, bytearray)):
        parts = []
        for item in value:
            part = normalize_text_value(item)
            if part:
                parts.append(part)
        return sanitize_utf8_text("\n\n".join(parts))

    return sanitize_utf8_text(str(value).strip())


def emit_progress(payload: dict) -> None:
    sys.stderr.write("PROGRESS " + json.dumps(payload) + "\n")
    sys.stderr.flush()


def load_model():
    global _TOKENIZER, _MODEL
    if _TOKENIZER is None or _MODEL is None:
        _, AutoTokenizer, AutoModel = load_runtime()
        tokenizer = AutoTokenizer.from_pretrained(
            MODEL_NAME, revision=MODEL_REVISION, trust_remote_code=False,
        )
        model = AutoModel.from_pretrained(
            MODEL_NAME, revision=MODEL_REVISION,
            use_safetensors=True, trust_remote_code=False,
        )
        model.eval()
        _TOKENIZER, _MODEL = tokenizer, model
    return _TOKENIZER, _MODEL


def normalize_embedding_row(index, item):
    row_id = item.get("id") if isinstance(item, Mapping) else None
    text = item.get("text") if isinstance(item, Mapping) else item
    try:
        identifier = row_id if isinstance(row_id, int) else int(str(row_id))
    except (TypeError, ValueError):
        identifier = index
    return {"id": identifier, "text": normalize_text_value(text)}


def embed_tokenized_inputs(inputs, model, torch):
    """Mean pooling shared by short inputs, student chunks, and faculty text."""
    with torch.no_grad():
        outputs = model(**inputs)
    last_hidden = outputs.last_hidden_state
    attention_mask = inputs["attention_mask"]
    mask = attention_mask.unsqueeze(-1).expand(last_hidden.size()).float()
    masked = last_hidden * mask
    summed = masked.sum(dim=1)
    counts = mask.sum(dim=1).clamp(min=1e-9)
    return (summed / counts).squeeze(0).tolist()


def process_embedding_request(payload: dict) -> dict:
    model_name = payload.get("model") or MODEL_NAME
    if model_name != MODEL_NAME:
        raise ValueError("This app uses PubMedBERT. Update faculty embeddings to replace an experimental model index before matching.")
    texts = payload.get("texts") or []
    total = len(texts)

    raw_label = payload.get("itemLabel")
    raw_plural = payload.get("itemLabelPlural")

    if isinstance(raw_label, str):
        raw_label = raw_label.strip()
    else:
        raw_label = ""

    if isinstance(raw_plural, str):
        raw_plural = raw_plural.strip()
    else:
        raw_plural = ""

    singular_label = raw_label or "text entry"
    if raw_plural:
        plural_label = raw_plural
    elif singular_label.endswith("s"):
        plural_label = singular_label
    else:
        plural_label = singular_label + "s"

    label_for_total = singular_label if total == 1 else plural_label

    if not texts:
        return {"model": model_name, "dimension": 0, "rows": []}

    emit_progress(
        {
            "phase": "loading-model",
            "message": f"Loading {model_name.rsplit('/', 1)[-1]} model…",
            "processedRows": 0,
            "totalRows": total,
        }
    )

    tokenizer, model = load_model()
    torch, _, _ = load_runtime()

    normalized_texts = [normalize_embedding_row(index, item) for index, item in enumerate(texts)]
    faculty_options = payload.get("facultyPreparation")
    prepared = {}
    if faculty_options is not None:
        if not isinstance(faculty_options, Mapping):
            raise ValueError("Faculty preparation options must be a JSON object.")
        emit_progress({
            "phase": "checking-lengths",
            "message": "Checking faculty descriptions against the embedding model's token limit…",
            "processedRows": 0,
            "totalRows": total,
        })
        prepared = prepare_faculty_texts(normalized_texts, tokenizer, emit_progress)

    emit_progress(
        {
            "phase": "embedding",
            "message": f"Starting embeddings for {total} {label_for_total}…",
            "processedRows": 0,
            "totalRows": total,
            "elapsedSeconds": 0.0,
        }
    )

    start_time = time.time()

    rows = []
    skipped_rows = []
    for item in normalized_texts:
        identifier = item["id"]
        preparation = prepared.get(identifier)
        text = preparation["text"] if preparation else item["text"]

        if not text:
            skipped_rows.append({
                "id": identifier,
                "reason": "empty-text",
            })
            sys.stderr.write(
                f"WARNING Skipping row {identifier}: empty text value\n"
            )
            sys.stderr.flush()
            emit_progress(
                {
                    "phase": "embedding",
                    "message": f"Skipped row {identifier}: empty text value",
                    "processedRows": len(rows),
                    "totalRows": total,
                    "elapsedSeconds": time.time() - start_time,
                    "skippedRows": len(skipped_rows),
                }
            )
            continue

        try:
            if faculty_options is None:
                input_chunks = list(student_embedding_inputs(tokenizer, text))
            else:
                input_chunks = [tokenizer(
                    text, return_tensors="pt", truncation=True,
                    max_length=MAX_INPUT_TOKENS, padding=True,
                )]
        except Exception as exc:  # noqa: BLE001
            if faculty_options is not None:
                raise RuntimeError(f"Unable to tokenize faculty row {identifier + 1}: {exc}") from exc
            reason = str(exc).splitlines()[0]
            skipped_rows.append(
                {
                    "id": identifier,
                    "reason": f"tokenizer-error: {reason}",
                }
            )
            sys.stderr.write(
                f"WARNING Skipping row {identifier}: unable to tokenize ({reason})\n"
            )
            sys.stderr.flush()
            emit_progress(
                {
                    "phase": "embedding",
                    "message": f"Skipped row {identifier}: unable to tokenize",
                    "processedRows": len(rows),
                    "totalRows": total,
                    "elapsedSeconds": time.time() - start_time,
                    "skippedRows": len(skipped_rows),
                }
            )
            continue

        embeddings = []
        for chunk_index, inputs in enumerate(input_chunks):
            embeddings.append(embed_tokenized_inputs(inputs, model, torch))
            if len(input_chunks) > 1:
                emit_progress({
                    "phase": "embedding",
                    "message": f"Embedded section {chunk_index + 1} of {len(input_chunks)} for {singular_label} {identifier + 1}…",
                    "processedRows": len(rows), "totalRows": total,
                    "elapsedSeconds": time.time() - start_time,
                })

        row = {"id": identifier, "embedding": embeddings[0]}
        if len(embeddings) > 1:
            row["additionalEmbeddings"] = embeddings[1:]
        if preparation:
            row["preparation"] = preparation
        rows.append(row)

        processed = len(rows)
        elapsed = time.time() - start_time
        remaining = None
        if processed < total and processed > 0 and elapsed > 0:
            remaining = (total - processed) * (elapsed / processed)

        emit_progress(
            {
                "phase": "embedding",
                "message": (
                    f"Embedded {processed} of {total} {label_for_total}"
                    + (
                        f" ({len(skipped_rows)} skipped)"
                        if skipped_rows
                        else ""
                    )
                ),
                "processedRows": processed,
                "totalRows": total,
                "elapsedSeconds": elapsed,
                "estimatedRemainingSeconds": remaining,
                "skippedRows": len(skipped_rows),
            }
        )

    result = {
        "model": model_name,
        "dimension": len(rows[0]["embedding"]) if rows else 0,
        "rows": rows,
    }

    if skipped_rows:
        result["skippedRows"] = skipped_rows

    emit_progress(
        {
            "phase": "finalizing",
            "message": "Finalizing embedding response…",
            "processedRows": len(rows),
            "totalRows": total,
            "elapsedSeconds": time.time() - start_time,
            "skippedRows": len(skipped_rows),
        }
    )

    return result


def handle_command(payload: dict) -> dict:
    if not isinstance(payload, Mapping):
        return {"error": "The embedding request must be a JSON object."}
    command = payload.get("command")

    if command == "shutdown":
        return {"status": "shutdown"}

    if command not in (None, "embed"):
        return {"error": f"Unknown command: {command}"}

    return process_embedding_request(payload)


def main() -> None:
    for raw_line in sys.stdin:
        line = raw_line.strip()
        if not line:
            continue

        try:
            payload = json.loads(line)
        except Exception as exc:  # noqa: BLE001
            message = f"Unable to parse embedding request: {exc}"
            sys.stderr.write(message + "\n")
            sys.stderr.flush()
            sys.stdout.write(json.dumps({"error": message}) + "\n")
            sys.stdout.flush()
            continue

        try:
            response = handle_command(payload)
        except Exception as exc:  # noqa: BLE001
            message = str(exc)
            sys.stderr.write(f"Embedding helper error: {message}\n")
            sys.stderr.flush()
            response = {"error": message}

        sys.stdout.write(json.dumps(response) + "\n")
        sys.stdout.flush()

        if response.get("status") == "shutdown":
            break


if __name__ == "__main__":
    main()
