"""Rebuild packaged PubMedBERT vectors directly from the bundled faculty TSV.

Run with the Python helper's requirements installed. Model files are downloaded
on first use, or use HF_HUB_OFFLINE=1 when the pinned models are cached.
"""

import csv
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import sys

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP / "python"))
import embedding_helper as helper
from embedding_config import MODEL_NAME


def main():
    source = APP / "src-tauri/assets/default_faculty_dataset.tsv"
    source_hash = hashlib.sha256(source.read_bytes()).digest()
    with source.open(encoding="utf-8-sig", newline="") as stream:
        faculty = list(csv.DictReader(stream, delimiter="\t"))
    if not faculty:
        raise ValueError("The bundled faculty dataset is empty.")
    response = helper.process_embedding_request({
        "model": MODEL_NAME, "facultyPreparation": {},
        "itemLabel": "faculty row", "itemLabelPlural": "faculty rows",
        "texts": [{"id": index, "text": row["Interests"]} for index, row in enumerate(faculty)],
    })
    if response["dimension"] != 768 or len(response["rows"]) != len(faculty):
        raise ValueError("Every bundled faculty row must have a complete PubMedBERT embedding.")
    if [row["id"] for row in response["rows"]] != list(range(len(faculty))):
        raise ValueError("The helper returned mismatched faculty row identities.")
    index = {
        "model": MODEL_NAME, "generatedAt": datetime.now(timezone.utc).isoformat(),
        "dimension": response["dimension"], "totalRows": len(faculty),
        "embeddedRows": len(response["rows"]), "skippedRows": 0,
        "embeddingColumns": ["Interests"], "identifierColumns": ["Name"],
        "entries": [{"rowIndex": row["id"], "identifiers": {"Name": faculty[row["id"]]["Name"].strip()},
                     "embedding": row["embedding"],
                     **({"preparation": row["preparation"]} if "preparation" in row else {})}
                    for row in response["rows"]],
    }
    if hashlib.sha256(source.read_bytes()).digest() != source_hash:
        raise ValueError("The bundled faculty source changed during embedding; rerun the rebuild.")
    destination = source.with_name("default_faculty_embeddings.json")
    temporary = destination.with_suffix(".tmp")
    temporary.write_text(json.dumps(index, ensure_ascii=False), encoding="utf-8")
    temporary.replace(destination)
    print(f"Rebuilt {len(faculty)} PubMedBERT faculty embeddings.")


if __name__ == "__main__":
    main()
