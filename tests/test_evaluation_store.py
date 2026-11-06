import unittest

from evaluation_store import comparison_fields, input_fingerprint, normalize_answer_rows


class EvaluationStoreTests(unittest.TestCase):
    def test_fingerprint_normalizes_outer_space_and_newlines(self):
        self.assertEqual(input_fingerprint(" A\r\nB "), input_fingerprint("A\nB"))

    def test_answer_rows_keep_positive_sku_quantities(self):
        rows = normalize_answer_rows(
            [
                {"SKU": " C-A ", "Quantity": "2"},
                {"SKU": "", "Quantity": 1},
                {"SKU": "C-B", "Quantity": 0},
            ]
        )
        self.assertEqual(rows, [{"sku": "C-A", "quantity": 2.0, "variant": None}])

    def test_comparison_identifies_quantity_error(self):
        result = comparison_fields(
            {"machine_sku": "c-a", "machine_quantity": 1},
            {"sku": "C-A", "quantity": 2, "variant": None},
        )
        self.assertTrue(result["sku_correct"])
        self.assertFalse(result["quantity_correct"])
        self.assertEqual(result["error_type"], "quantity")


if __name__ == "__main__":
    unittest.main()
