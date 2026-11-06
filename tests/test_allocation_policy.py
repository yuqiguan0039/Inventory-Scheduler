import unittest

from run_case_regressions import load_parser_namespace


class AllocationPolicyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.namespace = load_parser_namespace()

    def test_subhalf_quantity_is_blocked_by_default(self):
        usable = self.namespace["allocatable_quantity"]
        self.assertEqual(usable(0.4), 0.0)
        self.assertEqual(usable(0.4, allow_subhalf=True), 0.4)

    def test_single_warehouse_fulfillment_beats_first_priority_half_case(self):
        prioritize = self.namespace["prioritize_fulfillment_warehouses"]
        result = prioritize(
            1.0,
            [
                {"Warehouse": "TX", "Available": 0.5, "Priority": 0},
                {"Warehouse": "PHX", "Available": 1.0, "Priority": 1},
            ],
        )
        self.assertEqual(result, ["PHX", "TX"])

    def test_priority_is_preserved_when_first_warehouse_can_fulfill(self):
        prioritize = self.namespace["prioritize_fulfillment_warehouses"]
        result = prioritize(
            1.0,
            [
                {"Warehouse": "TX", "Available": 1.0, "Priority": 0},
                {"Warehouse": "PHX", "Available": 2.0, "Priority": 1},
            ],
        )
        self.assertEqual(result, ["TX", "PHX"])

    def test_split_priority_remains_when_no_warehouse_can_fulfill(self):
        prioritize = self.namespace["prioritize_fulfillment_warehouses"]
        result = prioritize(
            1.0,
            [
                {"Warehouse": "TX", "Available": 0.5, "Priority": 0},
                {"Warehouse": "PHX", "Available": 0.5, "Priority": 1},
            ],
        )
        self.assertEqual(result, ["TX", "PHX"])

    def test_subhalf_shipment_is_only_used_when_enabled(self):
        choose = self.namespace["choose_variant_shipments"]
        options = [{"SKU": "C-TEST", "Priority": 0, "Available": 0.4}]
        self.assertEqual(choose(0.4, options), [])
        self.assertEqual(
            choose(0.4, options, allow_subhalf=True),
            [{"SKU": "C-TEST", "Priority": 0, "Quantity": 0.4}],
        )

    def test_requested_variant_controls_inventory_order(self):
        order = self.namespace["version_priority_order"]
        self.assertEqual(order("Regular", "Preferred", True), [0, 1, 2])
        self.assertEqual(order("YN", "Preferred", True), [1, 0, 2])
        self.assertEqual(order("MG", "Preferred", True), [2, 0, 1])

    def test_only_variant_policy_excludes_regular_inventory(self):
        order = self.namespace["version_priority_order"]
        self.assertEqual(order("YN", "Only", True), [1])
        self.assertEqual(order("MG", "Only", True), [2])
        self.assertEqual(order("YN/MG", "Only", True), [1, 2])

    def test_disabling_fallback_keeps_explicit_requested_version(self):
        order = self.namespace["version_priority_order"]
        self.assertEqual(order("YN", "Preferred", False), [1])
        self.assertEqual(order("MG", "Preferred", False), [2])

    def test_zero_mg_inventory_is_a_separate_nicotine_free_variant(self):
        detail = self.namespace["get_sku_detail"]
        self.assertEqual(
            detail("B-0MGGBP15K-BRI"),
            ("B-GBP15K-BRI", 3, "B-0MGGBP15K-BRI"),
        )
        self.assertEqual(
            self.namespace["version_from_sku"]("B-0MGGBP15K-BRI"),
            "0MG",
        )

    def test_yn_zero_mg_combination_is_preserved_separately(self):
        detail = self.namespace["get_sku_detail"]
        self.assertEqual(
            detail("C-YN0MGGBP15K-BRI"),
            ("C-GBP15K-BRI", 4, "C-YN0MGGBP15K-BRI"),
        )
        self.assertEqual(
            self.namespace["version_from_sku"]("C-YN0MGGBP15K-BRI"),
            "YN 0MG",
        )

    def test_nicotine_versions_never_fall_back_to_zero_mg(self):
        order = self.namespace["version_priority_order"]
        self.assertNotIn(3, order("Regular", "Preferred", True))
        self.assertNotIn(3, order("YN", "Preferred", True))
        self.assertNotIn(3, order("MG", "Preferred", True))
        self.assertEqual(order("0MG", "Preferred", True), [3, 4])
        self.assertEqual(order("YN 0MG", "Preferred", True), [4, 3])

    def test_region_order_precedes_later_region_full_case(self):
        prioritize = self.namespace["prioritize_fulfillment_warehouses"]
        result = prioritize(
            1.0,
            [
                {
                    "Warehouse": "TX_SALE",
                    "Available": 0.5,
                    "Priority": 0,
                    "Group": "TX",
                },
                {
                    "Warehouse": "TX_STORAGE",
                    "Available": 0.5,
                    "Priority": 1,
                    "Group": "TX",
                },
                {
                    "Warehouse": "PHX_SALE",
                    "Available": 1.0,
                    "Priority": 2,
                    "Group": "PHX",
                },
            ],
        )
        self.assertEqual(
            result,
            ["TX_SALE", "TX_STORAGE", "PHX_SALE"],
        )

    def test_storage_can_avoid_split_within_same_region(self):
        prioritize = self.namespace["prioritize_fulfillment_warehouses"]
        result = prioritize(
            1.0,
            [
                {
                    "Warehouse": "TX_SALE",
                    "Available": 0.5,
                    "Priority": 0,
                    "Group": "TX",
                },
                {
                    "Warehouse": "TX_STORAGE",
                    "Available": 1.0,
                    "Priority": 1,
                    "Group": "TX",
                },
            ],
        )
        self.assertEqual(result, ["TX_STORAGE", "TX_SALE"])


if __name__ == "__main__":
    unittest.main()
