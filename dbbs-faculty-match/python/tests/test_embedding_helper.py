"""Protocol and text regressions run without downloading a model or installing torch."""
import contextlib
import importlib.util
import io
import json
from pathlib import Path
import sys
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

helper_path = Path(__file__).resolve().parents[1] / "embedding_helper.py"
sys.path.insert(0, str(helper_path.parent))
spec = importlib.util.spec_from_file_location("embedding_helper", helper_path)
helper = importlib.util.module_from_spec(spec)
spec.loader.exec_module(helper)


class TextNormalizationTests(unittest.TestCase):
    def test_unicode_and_whitespace(self):
        self.assertEqual(helper.normalize_text_value("  René studies proteins  "), "René studies proteins")

    def test_nested_json_content(self):
        self.assertEqual(helper.normalize_text_value({"text": "Protein structure", "content": ["Signaling", None]}), "Protein structure\n\nSignaling")

    def test_bytes_and_invalid_surrogates(self):
        self.assertEqual(helper.normalize_text_value(b"Protein research"), "Protein research")
        self.assertEqual(helper.sanitize_utf8_text("Protein bad\udcffword structure"), "Protein  structure")

    def test_empty_and_numeric_values(self):
        self.assertEqual(helper.normalize_text_value(None), "")
        self.assertEqual(helper.normalize_text_value([None, " "]), "")
        self.assertEqual(helper.normalize_text_value(123), "123")


class FacultyEmbeddingTests(unittest.TestCase):
    def test_pruning_precedes_embedding_and_ignores_legacy_summary_cache(self):
        class StopAfterTokenization(Exception):
            pass

        tokenizer = Mock()
        tokenizer.encode.side_effect = lambda text, **kwargs: text.split()
        tokenizer.num_special_tokens_to_add.return_value = 2
        tokenizer.return_value = {"captured": True}
        model = Mock(side_effect=StopAfterTokenization)
        torch = SimpleNamespace(no_grad=contextlib.nullcontext)
        source = "the receptor " + "research " * 600 + "the pathway\nLab Website"
        prepared = {0: {"text": "Renal cancer proteomics", "tokens": 3, "truncated": False}}
        with patch.object(helper, "prepare_faculty_texts", return_value=prepared) as prepare, \
                patch.object(helper, "load_model", return_value=(tokenizer, model)), \
                patch.object(helper, "load_runtime", return_value=(torch, None, None)), \
                patch.object(helper, "emit_progress"), \
                patch("builtins.open", side_effect=AssertionError("Must not read an old summary cache")):
            with self.assertRaises(StopAfterTokenization):
                helper.process_embedding_request({
                    "model": "NeuML/pubmedbert-base-embeddings",
                    "facultyPreparation": {"cachePath": "/old-model-summaries.json"},
                    "texts": [{"id": 0, "text": source}],
                })
        tokenizer.assert_called_once_with(
            "Renal cancer proteomics",
            return_tensors="pt", truncation=True, max_length=512, padding=True,
        )
        self.assertEqual(prepare.call_args.args[0], [{"id": 0, "text": source}])
        model.assert_called_once_with(captured=True)


class StudentEmbeddingTests(unittest.TestCase):
    def tokenizer(self):
        tokenizer = Mock()
        tokenizer.encode.side_effect = lambda text, **kwargs: list(range(600 if text == "Long interests" else 2))
        tokenizer.num_special_tokens_to_add.return_value = 2
        tokenizer.prepare_for_model.side_effect = [{"chunk": 0}, {"chunk": 1}]
        tokenizer.return_value = {"short": True}
        return tokenizer

    def test_batch_keeps_chunk_vectors_with_their_student_and_preserves_short_protocol(self):
        tokenizer = self.tokenizer()
        with patch.object(helper, "load_model", return_value=(tokenizer, object())), \
                patch.object(helper, "load_runtime", return_value=(None, None, None)), \
                patch.object(helper, "embed_tokenized_inputs", side_effect=[[1.0, 0.0], [0.0, 1.0], [0.5, 0.5]]) as embed, \
                patch.object(helper, "prepare_faculty_texts") as faculty, \
                patch.object(helper, "emit_progress"):
            result = helper.process_embedding_request({"model": "NeuML/pubmedbert-base-embeddings", "texts": [
                {"id": 7, "text": "Long interests"},
                {"id": 42, "text": "Cancer"},
            ]})
        self.assertEqual(result["rows"], [
            {"id": 7, "embedding": [1.0, 0.0], "additionalEmbeddings": [[0.0, 1.0]]},
            {"id": 42, "embedding": [0.5, 0.5]},
        ])
        self.assertEqual(result["dimension"], 2)
        self.assertEqual(embed.call_count, 3)
        faculty.assert_not_called()

    def test_failed_later_chunk_skips_the_whole_student_and_continues_the_batch(self):
        tokenizer = self.tokenizer()
        tokenizer.prepare_for_model.side_effect = [{"chunk": 0}, ValueError("bad later chunk")]
        with patch.object(helper, "load_model", return_value=(tokenizer, object())), \
                patch.object(helper, "load_runtime", return_value=(None, None, None)), \
                patch.object(helper, "embed_tokenized_inputs", return_value=[0.5, 0.5]) as embed, \
                patch.object(helper, "emit_progress"), contextlib.redirect_stderr(io.StringIO()):
            result = helper.process_embedding_request({"model": "NeuML/pubmedbert-base-embeddings", "texts": [
                {"id": 7, "text": "Long interests"},
                {"id": 42, "text": "Cancer"},
            ]})
        self.assertEqual(result["rows"], [{"id": 42, "embedding": [0.5, 0.5]}])
        self.assertEqual(result["skippedRows"], [{"id": 7, "reason": "tokenizer-error: bad later chunk"}])
        embed.assert_called_once()


class ProtocolTests(unittest.TestCase):
    def test_empty_embedding_request_needs_no_model(self):
        with patch.object(helper, "load_model") as model:
            result = helper.handle_command({"texts": []})
        self.assertEqual(result["rows"], [])
        self.assertEqual(result["dimension"], 0)
        model.assert_not_called()

    def test_shutdown_and_unknown_commands(self):
        self.assertEqual(helper.handle_command({"command": "shutdown"}), {"status": "shutdown"})
        self.assertIn("Unknown command", helper.handle_command({"command": "invalid"})["error"])
        self.assertIn("JSON object", helper.handle_command([])["error"])

    def test_line_protocol_recovers_from_invalid_json_and_stops_on_shutdown(self):
        output = io.StringIO()
        requests = 'invalid json\n{"texts": []}\n{"command": "shutdown"}\n{"command": "ignored"}\n'
        with patch.object(helper.sys, "stdin", io.StringIO(requests)), contextlib.redirect_stdout(output), contextlib.redirect_stderr(io.StringIO()):
            helper.main()
        responses = [json.loads(line) for line in output.getvalue().splitlines()]
        self.assertEqual(len(responses), 3)
        self.assertIn("error", responses[0])
        self.assertEqual(responses[1]["rows"], [])
        self.assertEqual(responses[2], {"status": "shutdown"})


if __name__ == "__main__":
    unittest.main()
