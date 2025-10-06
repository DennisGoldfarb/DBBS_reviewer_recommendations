import json
import os
import sys
import time
from typing import Dict, Tuple

os.environ.setdefault("TOKENIZERS_PARALLELISM", "false")

try:
    from transformers import AutoModel, AutoTokenizer
    from transformers.utils import logging as hf_logging
except ImportError as exc:  # pragma: no cover - dependency is required at runtime
    sys.stderr.write("Install the 'transformers' package (with torch) to generate embeddings.\n")
    sys.stderr.write(str(exc) + "\n")
    sys.exit(1)

try:
    import torch
except ImportError as exc:  # pragma: no cover - dependency is required at runtime
    sys.stderr.write("Install the 'torch' package to generate embeddings.\n")
    sys.stderr.write(str(exc) + "\n")
    sys.exit(1)

hf_logging.set_verbosity_error()

DEFAULT_MODEL = "NeuML/pubmedbert-base-embeddings"
_MODEL_CACHE: Dict[str, Tuple[AutoTokenizer, AutoModel]] = {}


def emit_progress(payload: dict) -> None:
    sys.stderr.write("PROGRESS " + json.dumps(payload) + "\n")
    sys.stderr.flush()


def normalize_label(value) -> str:
    if isinstance(value, str):
        return value.strip()
    return ""


def ensure_model(model_name: str) -> Tuple[AutoTokenizer, AutoModel]:
    cached = _MODEL_CACHE.get(model_name)
    if cached is not None:
        return cached

    emit_progress(
        {
            "phase": "loading-model",
            "message": f"Loading {model_name}…",
            "processedRows": 0,
            "totalRows": 0,
        }
    )

    tokenizer = AutoTokenizer.from_pretrained(model_name)
    model = AutoModel.from_pretrained(model_name)
    model.eval()

    _MODEL_CACHE[model_name] = (tokenizer, model)
    return tokenizer, model


def compute_embeddings(payload: dict) -> dict:
    model_name = (payload.get("model") or DEFAULT_MODEL).strip() or DEFAULT_MODEL
    texts = payload.get("texts") or []
    total = len(texts)

    raw_label = normalize_label(payload.get("itemLabel"))
    raw_plural = normalize_label(payload.get("itemLabelPlural"))

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

    tokenizer, model = ensure_model(model_name)

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

    for item in texts:
        text = normalize_label(item.get("text"))
        if not text:
            continue

        inputs = tokenizer(
            text,
            return_tensors="pt",
            truncation=True,
            max_length=512,
            padding=True,
        )

        with torch.no_grad():
            outputs = model(**inputs)

        last_hidden = outputs.last_hidden_state
        attention_mask = inputs["attention_mask"]
        mask = attention_mask.unsqueeze(-1).expand(last_hidden.size()).float()
        masked = last_hidden * mask
        summed = masked.sum(dim=1)
        counts = mask.sum(dim=1).clamp(min=1e-9)
        embedding = (summed / counts).squeeze(0)

        rows.append({"id": item.get("id"), "embedding": embedding.tolist()})

        processed = len(rows)
        elapsed = time.time() - start_time
        remaining = None
        if processed < total and processed > 0 and elapsed > 0:
            remaining = (total - processed) * (elapsed / processed)

        emit_progress(
            {
                "phase": "embedding",
                "message": f"Embedded {processed} of {total} {label_for_total}",
                "processedRows": processed,
                "totalRows": total,
                "elapsedSeconds": elapsed,
                "estimatedRemainingSeconds": remaining,
            }
        )

    emit_progress(
        {
            "phase": "finalizing",
            "message": "Finalizing embedding response…",
            "processedRows": len(rows),
            "totalRows": total,
            "elapsedSeconds": time.time() - start_time,
        }
    )

    return {
        "model": model_name,
        "dimension": len(rows[0]["embedding"]) if rows else 0,
        "rows": rows,
    }


def main() -> None:
    while True:
        line = sys.stdin.readline()
        if not line:
            break

        stripped = line.strip()
        if not stripped:
            continue

        try:
            message = json.loads(stripped)
        except Exception as exc:  # noqa: BLE001
            sys.stderr.write(f"Unable to parse embedding request: {exc}\n")
            sys.stderr.flush()
            continue

        command_type = message.get("type")
        if isinstance(command_type, str):
            normalized_type = command_type.strip().lower()
        else:
            normalized_type = "embed"

        if normalized_type == "shutdown":
            break
        if normalized_type not in {"embed", "preload"}:
            sys.stderr.write(f"Unknown command type: {command_type}\n")
            sys.stderr.flush()
            continue

        try:
            if normalized_type == "preload":
                model_name = (message.get("model") or DEFAULT_MODEL).strip() or DEFAULT_MODEL
                ensure_model(model_name)
                response = {"model": model_name, "dimension": 0, "rows": []}
            else:
                response = compute_embeddings(message)
            sys.stdout.write(json.dumps({"type": "result", "payload": response}) + "\n")
            sys.stdout.flush()
        except Exception as exc:  # noqa: BLE001
            import traceback

            traceback.print_exc(file=sys.stderr)
            sys.stderr.flush()
            sys.stdout.write(
                json.dumps({"type": "error", "message": str(exc)}) + "\n"
            )
            sys.stdout.flush()


if __name__ == "__main__":
    main()
