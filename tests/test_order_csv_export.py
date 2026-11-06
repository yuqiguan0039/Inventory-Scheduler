import io
import unittest

import pandas as pd

from run_case_regressions import load_parser_namespace


class OrderCsvExportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.namespace = load_parser_namespace()

    def test_csv_contains_resolved_values_without_formulas(self):
        guide = pd.DataFrame([
            {
                "Warehouse": "TX_SALE",
                "Order SKU": "C-GBPX25-BR",
                "Quantity": 1.0,
            },
            {
                "Warehouse": "PHX_SALE",
                "Order SKU": "C-GBPX25-MM",
                "Quantity": 0.5,
            },
        ])
        payload = self.namespace["build_order_csv"](guide)
        self.assertTrue(payload.startswith(b"\xef\xbb\xbf"))

        text = payload.decode("utf-8-sig")
        self.assertNotIn("=VLOOKUP", text)
        self.assertNotIn("=TODAY", text)

        result = pd.read_csv(io.StringIO(text), dtype=str)
        self.assertEqual(len(result), 2)
        self.assertEqual(result.loc[0, "SKU"], "C-GBPX25-BR")
        self.assertEqual(result.loc[0, "Product Code"], "GKPULX25KZ")
        self.assertEqual(result.loc[0, "Quantity"], "1.0")
        self.assertEqual(result.loc[1, "Quantity"], "0.5")
        self.assertNotEqual(
            result.loc[0, "SalesOrder Number"],
            result.loc[1, "SalesOrder Number"],
        )

    def test_missing_product_code_blocks_incomplete_csv(self):
        guide = pd.DataFrame([
            {
                "Warehouse": "TX_SALE",
                "Order SKU": "C-NOT-IN-TEMPLATE",
                "Quantity": 1.0,
            }
        ])
        with self.assertRaisesRegex(ValueError, "Product Code is missing"):
            self.namespace["build_order_csv"](guide)


if __name__ == "__main__":
    unittest.main()
