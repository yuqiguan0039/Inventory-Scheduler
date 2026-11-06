import unittest

from run_case_regressions import load_parser_namespace


class BrandAliasMappingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.namespace = load_parser_namespace()
        cls.catalog_df, catalog_error = cls.namespace["load_product_catalog"](
            str(cls.namespace["DEFAULT_PRODUCT_CATALOG_PATH"])
        )
        if catalog_error:
            raise RuntimeError(catalog_error)

    def test_excel_mapping_is_loaded(self):
        aliases = self.namespace["PRODUCT_HEADER_ALIASES"]
        self.assertGreaterEqual(len(aliases), 100)
        self.assertIn("MT15K", aliases)

    def test_alias_normalization_ignores_spaces_and_symbols(self):
        normalize = self.namespace["normalize_brand_alias"]
        self.assertEqual(normalize("MT 15K"), normalize("mt-15k"))

    def test_one_alias_can_map_to_multiple_brands(self):
        brands = self.namespace["PRODUCT_HEADER_ALIASES"]["MT15K"]
        self.assertIn("LOST MARY (TURBO)_MT15000", brands)
        self.assertIn("LOST MARY (THERMAL EDITION)_MT15000", brands)
        self.assertGreaterEqual(len(brands), 4)

    def test_plain_header_prefers_non_variant_brand(self):
        resolve = self.namespace["resolve_product_header_alias"]
        rank = self.namespace["rank_brand_candidates"]
        candidates = rank("NERA KIT", resolve("NERA KIT"))
        self.assertEqual(candidates[0], "LOST MARY NERA FULLVIEW KIT_70000")

    def test_prefixed_heading_uses_longest_suffix_alias(self):
        resolve = self.namespace["resolve_product_header_alias"]
        rank = self.namespace["rank_brand_candidates"]
        candidates = rank(
            "Multiline CO: PulseX 2 50k",
            resolve("Multiline CO: PulseX 2 50k"),
        )
        self.assertEqual(candidates[0], "GEEK BAR PULSE X 2_50000")

    def test_numbered_product_alias_is_checked_before_quantity_fallback(self):
        resolve = self.namespace["resolve_product_header_alias"]
        rank = self.namespace["rank_brand_candidates"]
        candidates = rank("PULSE 2", resolve("PULSE 2"))
        self.assertEqual(candidates[0], "GEEK BAR PULSE 2_25000")

    def test_pulse_2_toasted_pina_colada_uses_tpc_sku(self):
        parse = self.namespace["parse_demand_with_local_fallback"]
        rows, _, review_rows = parse(
            "PULSE 2\nToasted Pina Colada 1",
            self.catalog_df,
        )
        self.assertIn(
            {"SKU": "C-GBP225-TPC", "Quantity": 1.0},
            rows,
        )
        self.assertFalse(any(row.get("Need Review") for row in review_rows))

    def test_pulse_2_typo_suggests_tpc_instead_of_pulse_15k(self):
        parse = self.namespace["parse_demand_with_local_fallback"]
        _, _, review_rows = parse(
            "PULSE 2\nToasd Pina Colada 1",
            self.catalog_df,
        )
        typo_row = next(
            row for row in review_rows if row["Input"] == "Toasd Pina Colada 1"
        )
        self.assertEqual(typo_row["Parsed Brand"], "GEEK BAR PULSE 2_25000")
        self.assertEqual(typo_row["Suggested SKU"], "C-GBP225-TPC")

    def test_chinese_global_default_fills_only_missing_quantities(self):
        parse = self.namespace["parse_demand_with_local_fallback"]
        rows, _, review_rows = parse(
            "PULSEX 2\nBlueberry Jam 2\nBlue Razz Ice\nWatermelon Ice\n没标数字的全部为1",
            self.catalog_df,
        )
        quantities = {row["SKU"]: row["Quantity"] for row in rows}
        self.assertEqual(quantities["C-GBPX250-BJ"], 2.0)
        self.assertEqual(quantities["C-GBPX250-BRI"], 1.0)
        self.assertEqual(quantities["C-GBPX250-WI"], 1.0)
        self.assertFalse(any(row.get("Need Review") for row in review_rows))

    def test_chinese_default_wording_is_supported(self):
        extract = self.namespace["extract_default_quantity_instruction"]
        quantity, lines = extract("没标数字的默认为1")
        self.assertEqual(quantity, 1.0)
        self.assertEqual(lines, ["没标数字的默认为1"])

    def test_english_case_each_is_not_treated_as_a_flavor(self):
        parse = self.namespace["parse_demand_with_local_fallback"]
        rows, _, review_rows = parse(
            "PULSE 2\nMiami Mint\nPeach Lemonade\n1 case each",
            self.catalog_df,
        )
        self.assertEqual(len(rows), 2)
        self.assertTrue(all(row["Quantity"] == 1.0 for row in rows))
        self.assertFalse(any(row["Input"] == "1 case each" for row in review_rows))

    def test_mixed_block_without_global_instruction_stays_safe(self):
        parse = self.namespace["parse_demand_with_local_fallback"]
        _, _, review_rows = parse(
            "PULSEX 2\nBlueberry Jam 2\nBlue Razz Ice",
            self.catalog_df,
        )
        missing = next(row for row in review_rows if row["Input"] == "Blue Razz Ice")
        self.assertTrue(missing["Need Review"])
        self.assertEqual(missing["Quantity"], 0.0)

    def test_yn_product_header_selects_yn_sku_and_preference(self):
        parse = self.namespace["parse_demand_with_local_fallback"]
        enrich = self.namespace["enrich_demand_version_preferences"]
        raw_text = "PULSE YN\nBlue Razz Ice 1"
        rows, _, review_rows = parse(raw_text, self.catalog_df)
        enriched = enrich(rows, review_rows, raw_text)
        self.assertEqual(enriched[0]["SKU"], "C-YNGBP15K-BRI")
        self.assertEqual(enriched[0]["Requested Version"], "YN")
        self.assertEqual(enriched[0]["Version Policy"], "Preferred")

    def test_global_only_yn_mg_excludes_regular_by_policy(self):
        parse = self.namespace["parse_demand_with_local_fallback"]
        enrich = self.namespace["enrich_demand_version_preferences"]
        raw_text = "PULSE\nBlue Razz Ice 1\n本单只看YN/MG库存"
        rows, _, review_rows = parse(raw_text, self.catalog_df)
        enriched = enrich(rows, review_rows, raw_text)
        self.assertEqual(enriched[0]["Requested Version"], "YN/MG")
        self.assertEqual(enriched[0]["Version Policy"], "Only")

    def test_flavor_line_version_overrides_block_default(self):
        parse = self.namespace["parse_demand_with_local_fallback"]
        enrich = self.namespace["enrich_demand_version_preferences"]
        raw_text = "PULSE\nBlue Razz Ice YN 2\nMiami Mint 1"
        rows, _, review_rows = parse(raw_text, self.catalog_df)
        enriched = enrich(rows, review_rows, raw_text)
        by_sku = {row["SKU"]: row for row in enriched}
        self.assertEqual(by_sku["C-GBP15K-BRI"]["Requested Version"], "YN")
        self.assertEqual(by_sku["C-GBP15K-MMI"]["Requested Version"], "Regular")

    def test_zero_mg_wording_is_not_classified_as_mg(self):
        detect = self.namespace["detect_version_request"]
        self.assertEqual(detect("PULSE 0MG"), ("0MG", "Preferred"))
        self.assertEqual(detect("只看 YN 0MG 库存"), ("YN 0MG", "Only"))

    def test_zero_mg_header_selects_zero_nicotine_catalog_sku(self):
        parse = self.namespace["parse_demand_with_local_fallback"]
        enrich = self.namespace["enrich_demand_version_preferences"]
        raw_text = "PULSE 0MG\nBlue Razz Ice 1"
        rows, _, review_rows = parse(raw_text, self.catalog_df)
        enriched = enrich(rows, review_rows, raw_text)
        self.assertEqual(enriched[0]["SKU"], "C-0MGGBP15K-BRI")
        self.assertEqual(enriched[0]["Requested Version"], "0MG")

    def test_mt35_percent_heading_selects_two_percent_family(self):
        parse = self.namespace["parse_demand_with_local_fallback"]
        rows, _, review_rows = parse(
            "Lost Mary MT35000 · 2% — 6箱\nWintermint — 1\nMiami Mint — 5",
            self.catalog_df,
        )
        self.assertEqual(
            {row["SKU"]: row["Quantity"] for row in rows},
            {"C-LMMT352-WM": 1.0, "C-LMMT352-MM": 5.0},
        )
        self.assertFalse(any(row.get("Need Review") for row in review_rows))

    def test_mt35_lemonade_flavors_select_lemonade_edition(self):
        parse = self.namespace["parse_demand_with_local_fallback"]
        rows, _, review_rows = parse(
            """Lost Mary MT35000 · 5% — 5箱
Blue Razz Lemonade — 2
Strawberry Dragon Lemonade — 3""",
            self.catalog_df,
        )
        self.assertEqual(
            {row["SKU"]: row["Quantity"] for row in rows},
            {"C-LMMT35LE-BRL": 2.0, "C-LMMT35LE-SDL": 3.0},
        )
        self.assertFalse(any(row.get("Need Review") for row in review_rows))

    def test_crystal_replacement_pod_heading_never_selects_kit(self):
        parse = self.namespace["parse_demand_with_local_fallback"]
        raw_text = """Off-Stamp Crystal Cube Replacement Pod 35K · 5%
Bangin Strawberry — 1
Maui Punch — 2
Miami Mint — 4
Vanilla — 1"""
        rows, _, review_rows = parse(raw_text, self.catalog_df)
        self.assertEqual(len(rows), 4)
        self.assertTrue(all("CSPOD" in row["SKU"] for row in rows))
        self.assertTrue(
            all("POD (CRYSTAL)" in row["Matched Brand"] for row in review_rows)
        )
        self.assertFalse(any(row.get("Need Review") for row in review_rows))

    def test_b_gum_flavor_does_not_request_box_sku(self):
        choose = self.namespace["choose_catalog_match"]
        row, _ = choose(
            self.catalog_df,
            "LOST MARY_MT35000",
            "B-Gum Cotton Carnival",
            "LOST MARY_MT35000 B-Gum Cotton Carnival",
        )
        self.assertIsNotNone(row)
        self.assertEqual(row["Prefix"], "C")

        parse = self.namespace["parse_demand_with_local_fallback"]
        _, _, review_rows = parse(
            "Lost Mary MT35000 · 5% — 1箱\nB-Gum Cotton Carnival — 1",
            self.catalog_df,
        )
        review = next(row for row in review_rows if row["Input"].startswith("B-Gum"))
        self.assertFalse(str(review.get("Suggested SKU", "")).startswith("B-"))

    def test_all_brand_values_match_catalog_exactly(self):
        unmatched = self.namespace["find_unmatched_alias_brands"](
            self.namespace["PRODUCT_HEADER_ALIASES"],
            self.catalog_df,
        )
        self.assertEqual(unmatched, [])


if __name__ == "__main__":
    unittest.main()
