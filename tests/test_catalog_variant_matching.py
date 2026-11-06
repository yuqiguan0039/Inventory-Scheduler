import unittest

from run_case_regressions import load_parser_namespace


AVEN_0806001 = """Pulse liq 30mg
Blue Razz Ice 1
Fcuking Fab 2
Miami Mint 2
Pink Lemonade 1
Sour Strawberry 1
Triple Berries 1
Watermelon Ice 1

pulse liq 50mg

Blue Razz Ice 1
Fcuking Fab 2
Miami Mint 2
Strawberry Banana 1"""


class CatalogVariantMatchingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.namespace = load_parser_namespace()
        cls.catalog_df, catalog_error = cls.namespace["load_product_catalog"](
            str(cls.namespace["DEFAULT_PRODUCT_CATALOG_PATH"])
        )
        if catalog_error:
            raise RuntimeError(catalog_error)

    def test_explicit_mg_strength_filters_catalog_candidates(self):
        choose = self.namespace["choose_catalog_match"]
        row, _ = choose(
            self.catalog_df,
            "pulse liq 50mg",
            "Blue Razz Ice",
            "pulse liq 50mg Blue Razz Ice",
        )
        self.assertEqual(row["SKU"], "C-PLTL50-BRI")

    def test_aven_0806001_keeps_30mg_and_50mg_blocks_separate(self):
        parse = self.namespace["parse_demand_with_local_fallback"]
        rows, _, review_rows = parse(AVEN_0806001, self.catalog_df)
        quantities = {
            row["SKU"]: float(row["Quantity"])
            for row in rows
        }
        expected = {
            "C-PLTL30-BRI": 1.0,
            "C-PLTL30-FF": 2.0,
            "C-PLTL30-MM": 2.0,
            "C-PLTL30-PL": 1.0,
            "C-PLTL30-SS": 1.0,
            "C-PLTL30-TB": 1.0,
            "C-PLTL30-WI": 1.0,
            "C-PLTL50-BRI": 1.0,
            "C-PLTL50-FF": 2.0,
            "C-PLTL50-MM": 2.0,
            "C-PLTL50-SB": 1.0,
        }
        self.assertEqual(quantities, expected)
        self.assertFalse(any(row.get("Need Review") for row in review_rows))


if __name__ == "__main__":
    unittest.main()
