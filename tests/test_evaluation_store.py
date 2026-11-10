import unittest

from evaluation_store import (
    EvaluationStore,
    comparison_fields,
    input_fingerprint,
    normalize_answer_rows,
)


class FakeResponse:
    def __init__(self, data=None):
        self.data = data or []


class FakeTable:
    def __init__(self, client, name):
        self.client = client
        self.name = name
        self.operation = None
        self.payload = None

    def select(self, *_args, **_kwargs):
        self.operation = "select"
        return self

    def eq(self, *_args, **_kwargs):
        return self

    def order(self, *_args, **_kwargs):
        return self

    def upsert(self, payload, **_kwargs):
        self.operation = "upsert"
        self.payload = payload
        return self

    def update(self, payload, **_kwargs):
        self.operation = "update"
        self.payload = payload
        return self

    def execute(self):
        if self.name == "evaluation_items" and self.operation == "select":
            return FakeResponse(self.client.machine_rows)
        if self.name == "evaluation_items" and self.operation == "upsert":
            self.client.item_upsert = self.payload
        return FakeResponse()


class FakeClient:
    def __init__(self, machine_rows=None):
        self.machine_rows = machine_rows or []
        self.item_upsert = None

    def table(self, name):
        return FakeTable(self, name)


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

    def test_confirm_case_marks_missing_machine_prediction_for_review(self):
        client = FakeClient(
            machine_rows=[
                {
                    "case_id": "case-1",
                    "line_number": 1,
                    "machine_sku": "C-A",
                    "machine_quantity": 1,
                    "needs_review": None,
                }
            ]
        )
        store = EvaluationStore(client)

        confirmed_count = store.confirm_case(
            "case-1",
            [
                {"SKU": "C-A", "Quantity": 1},
                {"SKU": "C-NEW", "Quantity": 1},
            ],
        )

        self.assertEqual(confirmed_count, 2)
        self.assertIs(client.item_upsert[0]["needs_review"], False)
        self.assertIs(client.item_upsert[1]["needs_review"], True)
        self.assertEqual(
            client.item_upsert[1]["error_type"], "missing_prediction"
        )


if __name__ == "__main__":
    unittest.main()
