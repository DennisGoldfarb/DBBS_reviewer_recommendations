import json
import os
import re
import sys
import time
from collections.abc import Mapping, Sequence

os.environ.setdefault("TOKENIZERS_PARALLELISM", "false")

try:
    from transformers import AutoModel, AutoTokenizer
    from transformers.utils import logging as hf_logging
except ImportError as exc:
    sys.stderr.write("Install the 'transformers' package (with torch) to generate embeddings.\n")
    sys.stderr.write(str(exc) + "\n")
    sys.exit(1)

try:
    import torch
except ImportError as exc:
    sys.stderr.write("Install the 'torch' package to generate embeddings.\n")
    sys.stderr.write(str(exc) + "\n")
    sys.exit(1)

hf_logging.set_verbosity_error()


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


def main() -> None:
    try:
        payload = json.load(sys.stdin)
    except Exception as exc:  # noqa: BLE001
        sys.stderr.write(f"Unable to parse embedding request: {exc}\n")
        sys.exit(1)

    model_name = payload.get("model") or "NeuML/pubmedbert-base-embeddings"
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
        json.dump({"model": model_name, "dimension": 0, "rows": []}, sys.stdout)
        return

    emit_progress(
        {
            "phase": "loading-model",
            "message": "Loading PubMedBERT model…",
            "processedRows": 0,
            "totalRows": total,
        }
    )

    tokenizer = AutoTokenizer.from_pretrained(model_name)
    model = AutoModel.from_pretrained(model_name)
    model.eval()

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
    for index, item in enumerate(texts):
        if isinstance(item, Mapping):
            row_id = item.get("id")
            raw_text = item.get("text")
        else:
            row_id = None
            raw_text = item

        text = normalize_text_value(raw_text)
        if isinstance(row_id, int):
            identifier = row_id
        else:
            try:
                identifier = int(str(row_id))
            except (TypeError, ValueError):
                identifier = index

        if not isinstance(text, str):
            text = str(text)

        text = text.strip()

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
            inputs = tokenizer(
                text,
                return_tensors="pt",
                truncation=True,
                max_length=512,
                padding=True,
            )
        except Exception as exc:  # noqa: BLE001
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

        with torch.no_grad():
            outputs = model(**inputs)

        last_hidden = outputs.last_hidden_state
        attention_mask = inputs["attention_mask"]
        mask = attention_mask.unsqueeze(-1).expand(last_hidden.size()).float()
        masked = last_hidden * mask
        summed = masked.sum(dim=1)
        counts = mask.sum(dim=1).clamp(min=1e-9)
        embedding = (summed / counts).squeeze(0)

        rows.append({"id": identifier, "embedding": embedding.tolist()})

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

    json.dump(result, sys.stdout)


if __name__ == "__main__":
    main()
