"""Topic-pruning boundaries and full-source scoring without ML dependencies."""

import contextlib
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import faculty_preparation as preparation


class WordTokenizer:
    def encode(self, text, add_special_tokens=False, truncation=False):
        assert add_special_tokens is False
        assert truncation is False
        return text.split()

    def num_special_tokens_to_add(self, pair=False):
        return 2


def words(n, word="research"):
    return " ".join([word] * n)


class FacultyPreparationTests(unittest.TestCase):
    def setUp(self):
        self.tokenizer = WordTokenizer()
        self.progress = []
        self.pruner = Mock()
        self.pruner.score_words.side_effect = lambda text, spans: [0.5] * len(spans)
        self.factory = Mock(return_value=self.pruner)

    def prepare(self, rows):
        return preparation.prepare_faculty_texts(
            rows, self.tokenizer, self.progress.append, pruner_factory=self.factory,
        )

    def test_inputs_at_or_below_budget_need_no_pruning_model(self):
        rows = [{"id": 0, "text": ""}, {"id": 1, "text": words(508) + " the receptor"}]
        originals = [dict(row) for row in rows]
        self.assertEqual(self.prepare(rows), {})
        self.assertEqual(rows, originals)
        self.factory.assert_not_called()
        self.assertEqual(self.progress, [])

    def test_footer_is_removed_even_from_short_descriptions(self):
        result = self.prepare([{"id": 7, "text": "Renal cancer proteomics\nLab Website"}])[7]
        self.assertEqual(result["text"], "Renal cancer proteomics")
        self.assertEqual(result["tokens"], 3)
        self.assertEqual(result["model"], "deterministic-text-cleanup")
        self.factory.assert_not_called()

    def test_footer_cleanup_can_avoid_loading_the_model(self):
        result = self.prepare([{"id": 0, "text": words(509) + "\nLab Website"}])[0]
        self.assertEqual(result["text"], words(509))
        self.factory.assert_not_called()

    def test_website_labels_are_removed_without_losing_trailing_department_text(self):
        source = "Genomics\nLab WebsitePsychological and Brain Sciences\nLab Website√ä"
        result = self.prepare([{"id": 0, "text": source}])[0]
        self.assertEqual(result["text"], "Genomics\nPsychological and Brain Sciences\n√ä")

    def test_footer_only_source_becomes_empty_without_loading_the_model(self):
        result = self.prepare([{"id": 0, "text": "Lab Website"}])[0]
        self.assertEqual(result["text"], "")
        self.assertEqual(result["tokens"], 0)
        self.factory.assert_not_called()

    def test_pruning_keeps_high_scoring_topic_words_from_the_end(self):
        topics = {"acute", "myeloid", "leukemia", "proteomics", "genomics"}
        self.pruner.score_words.side_effect = lambda text, spans: [
            1.0 if span.group() in topics else 0.0 for span in spans
        ]
        source = words(600, "filler") + " acute myeloid leukemia proteomics genomics\nLab Website"
        rows = [{"id": 7, "text": source}, {"id": 8, "text": "Cancer research"}]
        originals = [dict(row) for row in rows]
        result = self.prepare(rows)
        self.assertEqual(set(result), {7})
        self.assertTrue(result[7]["text"].endswith("acute myeloid leukemia proteomics genomics"))
        self.assertEqual(result[7]["tokens"], 510)
        self.assertFalse(result[7]["truncated"])
        self.assertEqual(result[7]["model"], preparation.PRUNING_MODEL)
        self.assertEqual(result[7]["revision"], preparation.PRUNING_REVISION)
        self.assertEqual(rows, originals)
        self.assertNotIn("Lab Website", self.pruner.score_words.call_args.args[0])
        self.pruner.close.assert_called_once()

    def test_mechanism_polarity_and_gene_ids_are_not_forcibly_retained(self):
        self.pruner.score_words.side_effect = lambda text, spans: [
            0.0 if span.group() in {"not", "may", "will", "PROTEIN1"} else 1.0 for span in spans
        ]
        source = "not may will PROTEIN1 " + words(510)
        result = self.prepare([{"id": 0, "text": source}])[0]
        self.assertEqual(result["text"], words(510))

    def test_articles_are_selected_by_model_scores_not_a_manual_deletion_rule(self):
        self.pruner.score_words.side_effect = lambda text, spans: [
            0.0 if span.group() == "filler" else 1.0 for span in spans
        ]
        result = self.prepare([{"id": 0, "text": "the " + words(509) + " filler"}])[0]
        self.assertTrue(result["text"].startswith("the "))
        self.assertNotIn("filler", result["text"])

    def test_pruning_only_deletes_whole_source_words_and_keeps_their_order(self):
        source = "René renal-cell AKT2-NOX2 " + words(510) + " single-cell proteomics"
        result = self.prepare([{"id": 0, "text": source}])[0]
        original_words = iter(source.split())
        for word in result["text"].split():
            self.assertTrue(any(original == word for original in original_words))
        self.assertEqual(result["tokens"], 510)

    def test_special_tokens_are_reserved_from_the_512_total_budget(self):
        self.tokenizer.num_special_tokens_to_add = lambda pair=False: 4
        result = self.prepare([{"id": 0, "text": words(509)}])[0]
        self.assertEqual(result["tokens"], 508)

    def test_one_model_is_used_for_all_long_rows_and_then_released(self):
        self.prepare([{"id": 0, "text": words(600)}, {"id": 1, "text": words(700)}])
        self.factory.assert_called_once()
        self.assertEqual(self.pruner.score_words.call_count, 2)
        self.pruner.close.assert_called_once()
        self.assertEqual(self.progress[-1]["processedRows"], 2)
        self.assertEqual(self.progress[-1]["totalRows"], 2)

    def test_failed_model_reports_the_row_and_releases_resources(self):
        self.pruner.score_words.side_effect = OSError("model unavailable")
        with self.assertRaisesRegex(RuntimeError, "faculty row 8 locally: model unavailable"):
            self.prepare([{"id": 7, "text": words(600)}])
        self.pruner.close.assert_called_once()

    def test_invalid_model_scores_fail_instead_of_silently_truncating(self):
        for scores in ([0.5], [float("nan")] * 511, [float("inf")] * 511):
            with self.subTest(scores=len(scores)):
                self.pruner.score_words.side_effect = None
                self.pruner.score_words.return_value = scores
                with self.assertRaisesRegex(RuntimeError, "invalid word scores"):
                    self.prepare([{"id": 0, "text": words(511)}])


class FullSourceScoringTests(unittest.TestCase):
    def test_pruning_model_requires_safe_weights_and_disables_remote_code(self):
        tokenizer_factory, model_factory = Mock(), Mock()
        torch = SimpleNamespace(get_num_threads=lambda: 4, set_num_threads=Mock(), float32="float32")
        transformers = SimpleNamespace(
            AutoTokenizer=tokenizer_factory, AutoModelForTokenClassification=model_factory,
        )
        pruner = preparation.LocalWordPruner()
        with patch.dict(sys.modules, {"torch": torch, "transformers": transformers}):
            pruner._load()
            tokenizer_factory.from_pretrained.assert_called_once_with(
                preparation.PRUNING_MODEL, revision=preparation.PRUNING_REVISION,
                use_fast=True, trust_remote_code=False,
            )
            model_factory.from_pretrained.assert_called_once_with(
                preparation.PRUNING_MODEL, revision=preparation.PRUNING_REVISION,
                torch_dtype="float32", use_safetensors=True, trust_remote_code=False,
            )
            pruner.close()

    def test_all_windows_are_scored_and_subwords_and_overlap_are_averaged(self):
        class Logits:
            def __init__(self, probabilities):
                self.probabilities = probabilities

            def softmax(self, dim):
                return self

            def __getitem__(self, key):
                return SimpleNamespace(tolist=lambda: self.probabilities)

        pruner = preparation.LocalWordPruner()
        pruner._load = Mock()
        pruner.tokenizer = Mock(return_value={
            "input_ids": [[0, 1, 2, 0], [0, 2, 3, 4, 0]],
            "attention_mask": [[1] * 4, [1] * 5],
            "offset_mapping": [
                [(0, 0), (0, 5), (6, 12), (0, 0)],
                [(0, 0), (6, 12), (13, 17), (17, 23), (0, 0)],
            ],
        })
        pruner.model = Mock(side_effect=[
            SimpleNamespace(logits=Logits([0, 0.9, 0.3, 0])),
            SimpleNamespace(logits=Logits([0, 0.7, 0.8, 1.0, 0])),
        ])
        torch = SimpleNamespace(inference_mode=contextlib.nullcontext, long="long",
                                tensor=lambda values, **kwargs: values)
        text = "renal cancer proteomics"
        with patch.dict(sys.modules, {"torch": torch}):
            scores = pruner.score_words(text, list(preparation.re.finditer(r"\S+", text)))
        for actual, expected in zip(scores, [0.9, 0.5, 0.9]):
            self.assertAlmostEqual(actual, expected)
        self.assertEqual(pruner.model.call_count, 2)
        self.assertTrue(pruner.tokenizer.call_args.kwargs["return_overflowing_tokens"])
        self.assertEqual(pruner.tokenizer.call_args.kwargs["stride"], 64)
        self.assertEqual(pruner.tokenizer.call_args.kwargs["max_length"], 512)


if __name__ == "__main__":
    unittest.main()
