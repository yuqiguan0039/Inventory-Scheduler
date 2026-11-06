import tempfile
import unittest
from pathlib import Path

from order_history import (
    load_order_answers,
    load_order_history,
    record_order_answer,
    record_order_detail,
)


class OrderHistoryTests(unittest.TestCase):
    def test_each_submission_creates_a_separate_record(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            history_path = Path(temp_dir) / "order_history.jsonl"
            first, _ = record_order_detail(
                "MT15K\nMiami Mint 2",
                use_ai=False,
                history_path=history_path,
            )
            second, _ = record_order_detail(
                "MT15K\nMiami Mint 2",
                use_ai=True,
                history_path=history_path,
            )

            records = load_order_history(history_path)
            self.assertEqual(len(records), 2)
            self.assertNotEqual(first["record_id"], second["record_id"])
            self.assertFalse(records[0]["use_ai"])
            self.assertTrue(records[1]["use_ai"])
            self.assertEqual(records[0]["order_detail"], "MT15K\nMiami Mint 2")

    def test_empty_submission_is_not_recorded(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            history_path = Path(temp_dir) / "order_history.jsonl"
            with self.assertRaisesRegex(ValueError, "empty"):
                record_order_detail("   ", history_path=history_path)
            self.assertFalse(history_path.exists())

    def test_confirmed_answer_is_linked_to_order_record(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            answer_path = Path(temp_dir) / "order_answers.jsonl"
            answer, _ = record_order_answer(
                "order-123",
                [
                    {"SKU": "C-GBP35K-ML", "Quantity": 1},
                    {"SKU": "C-GBP35K-SDL", "Quantity": "2"},
                ],
                answer_path=answer_path,
            )

            records = load_order_answers(answer_path)
            self.assertEqual(len(records), 1)
            self.assertEqual(answer["order_record_id"], "order-123")
            self.assertEqual(
                records[0]["correct_answer"],
                [
                    {"SKU": "C-GBP35K-ML", "Quantity": 1.0},
                    {"SKU": "C-GBP35K-SDL", "Quantity": 2.0},
                ],
            )

    def test_empty_confirmed_answer_is_not_recorded(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            answer_path = Path(temp_dir) / "order_answers.jsonl"
            with self.assertRaisesRegex(ValueError, "empty"):
                record_order_answer(
                    "order-123",
                    [{"SKU": "", "Quantity": 0}],
                    answer_path=answer_path,
                )
            self.assertFalse(answer_path.exists())


if __name__ == "__main__":
    unittest.main()
