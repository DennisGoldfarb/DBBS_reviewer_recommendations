"""The finalized model stays pinned and experimental requests cannot load models."""

import contextlib
import io
import json
from pathlib import Path
import sys
import unittest
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import embedding_helper as helper
from embedding_config import MODEL_NAME, MODEL_REVISION, MAX_INPUT_TOKENS


class EmbeddingConfigurationTests(unittest.TestCase):
    def test_default_protocol_model_is_pubmedbert(self):
        self.assertEqual(MODEL_NAME, "NeuML/pubmedbert-base-embeddings")
        self.assertEqual(MAX_INPUT_TOKENS, 512)
        self.assertEqual(helper.process_embedding_request({"texts": []})["model"], MODEL_NAME)

    def test_tokenizer_and_model_share_a_pinned_revision_and_are_reused(self):
        tokenizer_factory, model_factory = Mock(), Mock()
        with patch.object(helper, "_TOKENIZER", None), patch.object(helper, "_MODEL", None), \
                patch.object(helper, "load_runtime", return_value=(None, tokenizer_factory, model_factory)):
            first = helper.load_model()
            self.assertEqual(helper.load_model(), first)
        tokenizer_factory.from_pretrained.assert_called_once_with(MODEL_NAME, revision=MODEL_REVISION)
        model_factory.from_pretrained.assert_called_once_with(MODEL_NAME, revision=MODEL_REVISION)
        model_factory.from_pretrained.return_value.eval.assert_called_once()

    def test_unsupported_model_fails_before_loading_any_runtime(self):
        with patch.object(helper, "load_runtime") as runtime, patch.object(helper, "load_model") as model:
            with self.assertRaisesRegex(ValueError, "Update faculty embeddings"):
                helper.process_embedding_request({"model": "other/model", "texts": [{"text": "Cancer"}]})
        runtime.assert_not_called()
        model.assert_not_called()

    def test_worker_recovers_after_an_unsupported_model_request(self):
        requests = '\n'.join([json.dumps({"model": "other/model", "texts": []}),
                              json.dumps({"texts": []}), json.dumps({"command": "shutdown"})])
        output = io.StringIO()
        with patch.object(helper.sys, "stdin", io.StringIO(requests)), \
                contextlib.redirect_stdout(output), contextlib.redirect_stderr(io.StringIO()):
            helper.main()
        rows = [json.loads(line) for line in output.getvalue().splitlines()]
        self.assertIn("Update faculty embeddings", rows[0]["error"])
        self.assertEqual(rows[1]["model"], MODEL_NAME)
        self.assertEqual(rows[2], {"status": "shutdown"})


if __name__ == "__main__":
    unittest.main()
