import io
import unittest

import pandas as pd

from upload_processing import (
    SOURCE_FILE_COLUMN,
    build_warehouse_sequence,
    common_inventory_columns,
    deduplicate_uploaded_files,
    filter_inventory_warehouses,
    find_cross_file_inventory_conflicts,
    is_eligible_warehouse,
    read_inventory_upload,
    warehouse_region,
)


class FakeUpload:
    def __init__(self, name, content):
        self.name = name
        self._content = content

    def getvalue(self):
        return self._content


class UploadProcessingTests(unittest.TestCase):
    def test_standard_sale_is_default_and_local_is_always_ignored(self):
        self.assertTrue(is_eligible_warehouse("TX_SALE"))
        self.assertFalse(is_eligible_warehouse("TX_STORAGE"))
        self.assertTrue(is_eligible_warehouse("TX_STORAGE", include_storage=True))
        self.assertFalse(is_eligible_warehouse("LAS_LOCAL_SALE"))
        self.assertFalse(
            is_eligible_warehouse(
                "HOUSTON_LOCAL_STORAGE", include_storage=True
            )
        )
        self.assertFalse(is_eligible_warehouse("CENTRAL_HUB"))

    def test_inventory_filter_reports_ignored_warehouses(self):
        inventory = pd.DataFrame(
            {
                "Warehouse": [
                    "TX_SALE",
                    "TX_STORAGE",
                    "LAS_LOCAL_SALE",
                    "PHX_PICK_UP",
                ],
                "SKU": ["A", "B", "C", "D"],
            }
        )
        filtered, ignored = filter_inventory_warehouses(
            inventory,
            "Warehouse",
            include_storage=True,
        )
        self.assertEqual(
            filtered["Warehouse"].tolist(), ["TX_SALE", "TX_STORAGE"]
        )
        self.assertEqual(ignored, ["LAS_LOCAL_SALE", "PHX_PICK_UP"])

    def test_each_region_uses_sale_then_storage(self):
        sequence = build_warehouse_sequence(
            [
                "PHX_STORAGE",
                "TX_STORAGE",
                "PHX_SALE",
                "TX_SALE",
                "LAS_LOCAL_SALE",
            ],
            ["TX", "PHX"],
            include_storage=True,
        )
        self.assertEqual(
            sequence,
            ["TX_SALE", "TX_STORAGE", "PHX_SALE", "PHX_STORAGE"],
        )

    def test_storage_is_absent_from_default_sequence(self):
        sequence = build_warehouse_sequence(
            ["TX_STORAGE", "TX_SALE"],
            ["TX"],
        )
        self.assertEqual(sequence, ["TX_SALE"])
        self.assertEqual(warehouse_region("TX_STORAGE"), "TX")

    def test_identical_uploads_are_skipped(self):
        first = FakeUpload("first.png", b"same")
        duplicate = FakeUpload("copy.png", b"same")
        different = FakeUpload("other.png", b"different")

        unique, duplicate_names = deduplicate_uploaded_files(
            [first, duplicate, different]
        )

        self.assertEqual([item.name for item in unique], ["first.png", "other.png"])
        self.assertEqual(duplicate_names, ["copy.png"])

    def test_common_columns_exclude_source_metadata(self):
        first = pd.DataFrame(
            {"Warehouse": ["TX"], "SKU": ["A"], SOURCE_FILE_COLUMN: ["a.xlsx"]}
        )
        second = pd.DataFrame(
            {"Warehouse": ["ID"], "SKU": ["B"], SOURCE_FILE_COLUMN: ["b.xlsx"]}
        )
        self.assertEqual(common_inventory_columns([first, second]), ["Warehouse", "SKU"])

    def test_sales_inventory_title_rows_and_physical_quantity_are_adapted(self):
        source = pd.DataFrame(
            [
                ["Main Data", None, None, None],
                [None, None, None, None],
                [
                    "Warehouse Name",
                    "Brand",
                    "SKU",
                    "Physical Available For Sale Qty",
                ],
                ["ID_SALE", "TEST", "C-TEST-A", 0.5],
            ]
        )
        payload = io.BytesIO()
        source.to_excel(payload, index=False, header=False)

        result = read_inventory_upload(
            FakeUpload("Main_Data.xlsx", payload.getvalue())
        )

        self.assertEqual(
            result.columns.tolist(),
            [
                "Warehouse Name",
                "Brand",
                "SKU",
                "Available For Sale",
                SOURCE_FILE_COLUMN,
            ],
        )
        self.assertEqual(result.iloc[0]["Warehouse Name"], "ID_SALE")
        self.assertEqual(result.iloc[0]["SKU"], "C-TEST-A")
        self.assertEqual(result.iloc[0]["Available For Sale"], 0.5)

    def test_accounting_and_physical_exports_share_quantity_column(self):
        accounting = pd.DataFrame(
            {
                "Warehouse Name": ["TX_SALE"],
                "SKU": ["C-A"],
                "Accounting Available For Sale": [2],
            }
        )
        physical = pd.DataFrame(
            {
                "Warehouse Name": ["PHX_SALE"],
                "SKU": ["C-B"],
                "Physical Available For Sale Qty": [3],
            }
        )
        accounting_payload = io.BytesIO()
        physical_payload = io.BytesIO()
        accounting.to_excel(accounting_payload, index=False)
        physical.to_excel(physical_payload, index=False)

        frames = [
            read_inventory_upload(
                FakeUpload("accounting.xlsx", accounting_payload.getvalue())
            ),
            read_inventory_upload(
                FakeUpload("physical.xlsx", physical_payload.getvalue())
            ),
        ]

        self.assertIn("Available For Sale", common_inventory_columns(frames))

    def test_cross_file_overlap_is_reported_after_summing_within_file(self):
        inventory = pd.DataFrame(
            {
                "Warehouse": ["TX", "TX", "TX", "ID"],
                "SKU": ["A", "A", "A", "B"],
                "Qty": [2, 3, 7, 4],
                SOURCE_FILE_COLUMN: ["one.xlsx", "one.xlsx", "two.xlsx", "two.xlsx"],
            }
        )

        conflicts = find_cross_file_inventory_conflicts(
            inventory, "Warehouse", "SKU", "Qty"
        )

        self.assertEqual(len(conflicts), 1)
        self.assertEqual(conflicts.iloc[0]["Warehouse Name"], "TX")
        self.assertEqual(conflicts.iloc[0]["SKU"], "A")
        self.assertIn("one.xlsx: 5", conflicts.iloc[0]["Quantity by File"])
        self.assertIn("two.xlsx: 7", conflicts.iloc[0]["Quantity by File"])

    def test_same_key_in_one_file_is_not_a_cross_file_conflict(self):
        inventory = pd.DataFrame(
            {
                "Warehouse": ["TX", "TX"],
                "SKU": ["A", "A"],
                "Qty": [2, 3],
                SOURCE_FILE_COLUMN: ["one.xlsx", "one.xlsx"],
            }
        )
        conflicts = find_cross_file_inventory_conflicts(
            inventory, "Warehouse", "SKU", "Qty"
        )
        self.assertTrue(conflicts.empty)

    def test_conflict_matching_ignores_case_and_outer_spaces(self):
        inventory = pd.DataFrame(
            {
                "Warehouse": ["tx_sale", " TX_SALE "],
                "SKU": ["c-test-a", " C-TEST-A "],
                "Qty": [2, 3],
                SOURCE_FILE_COLUMN: ["one.xlsx", "two.xlsx"],
            }
        )
        conflicts = find_cross_file_inventory_conflicts(
            inventory, "Warehouse", "SKU", "Qty"
        )
        self.assertEqual(len(conflicts), 1)
        self.assertEqual(conflicts.iloc[0]["Warehouse Name"], "TX_SALE")
        self.assertEqual(conflicts.iloc[0]["SKU"], "C-TEST-A")


if __name__ == "__main__":
    unittest.main()
