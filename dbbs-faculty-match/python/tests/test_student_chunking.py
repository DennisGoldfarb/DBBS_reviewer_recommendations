"""Exact token coverage, overlap, and unchanged short inputs without ML downloads."""

from pathlib import Path
import sys
import unittest
from unittest.mock import Mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from student_chunking import student_embedding_inputs, token_windows


class TokenWindowTests(unittest.TestCase):
    def test_boundaries_include_the_tail_without_an_overlap_only_extra_chunk(self):
        for length, expected_lengths in (
            (0, []), (509, [509]), (510, [510]), (511, [510, 51]),
            (970, [510, 510]), (971, [510, 510, 51]), (1000, [510, 510, 80]),
        ):
            with self.subTest(length=length):
                windows = list(token_windows(list(range(length)), 510))
                self.assertEqual([len(window) for window in windows], expected_lengths)

    def test_every_source_token_is_retained_and_overlaps_are_exactly_50_tokens(self):
        source = list(range(12000))
        windows = list(token_windows(source, 510))
        rebuilt = windows[0][:]
        for previous, current in zip(windows, windows[1:]):
            self.assertEqual(previous[-50:], current[:50])
            rebuilt.extend(current[50:])
        self.assertEqual(rebuilt, source)
        self.assertTrue(all(len(window) <= 510 for window in windows))

    def test_invalid_windows_fail_instead_of_looping_forever(self):
        for window, overlap in ((0, 0), (10, -1), (50, 50), (20, 50)):
            with self.subTest(window=window, overlap=overlap):
                with self.assertRaises(ValueError):
                    list(token_windows([1, 2, 3], window, overlap))


class StudentInputsTests(unittest.TestCase):
    def tokenizer(self, ids):
        tokenizer = Mock()
        tokenizer.encode.return_value = ids
        tokenizer.num_special_tokens_to_add.return_value = 2
        tokenizer.return_value = {"original": True}
        tokenizer.prepare_for_model.side_effect = lambda ids, **kwargs: {
            "input_ids": [[101] + ids + [102]] if kwargs.get("prepend_batch_axis") else [101] + ids + [102],
        }
        tokenizer.decode.side_effect = AssertionError("Must not round-trip source tokens through text")
        return tokenizer

    def test_short_inputs_use_the_original_tokenization_path(self):
        for length in (2, 509, 510):
            with self.subTest(length=length):
                tokenizer = self.tokenizer(list(range(length)))
                chunks = list(student_embedding_inputs(tokenizer, "René studies cancer."))
                self.assertEqual(chunks, [{"original": True}])
                tokenizer.assert_called_once_with(
                    "René studies cancer.", return_tensors="pt", truncation=True,
                    max_length=512, padding=True,
                )
                tokenizer.prepare_for_model.assert_not_called()

    def test_long_inputs_use_original_ids_with_special_tokens_for_every_window(self):
        source = list(range(1000))
        tokenizer = self.tokenizer(source)
        chunks = list(student_embedding_inputs(tokenizer, "Long research document"))
        self.assertEqual([len(chunk["input_ids"][0]) for chunk in chunks], [512, 512, 82])
        self.assertEqual(chunks[0]["input_ids"][0], [101] + source[:510] + [102])
        self.assertEqual(chunks[-1]["input_ids"][0], [101] + source[920:] + [102])
        tokenizer.assert_not_called()
        tokenizer.decode.assert_not_called()
        for call in tokenizer.prepare_for_model.call_args_list:
            self.assertTrue(call.kwargs["add_special_tokens"])
            self.assertFalse(call.kwargs["truncation"])
            self.assertTrue(call.kwargs["prepend_batch_axis"])

    def test_special_tokens_are_reserved_from_the_actual_tokenizer_budget(self):
        tokenizer = self.tokenizer(list(range(509)))
        tokenizer.num_special_tokens_to_add.return_value = 4
        chunks = list(student_embedding_inputs(tokenizer, "Long research interests"))
        self.assertEqual(len(chunks), 2)
        self.assertEqual(len(tokenizer.prepare_for_model.call_args_list[0].args[0]), 508)


if __name__ == "__main__":
    unittest.main()
