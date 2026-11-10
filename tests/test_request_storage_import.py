import unittest

from import_request_storage import DEFAULT_SOURCE, load_cases, parse_output_rows


class RequestStorageImportTests(unittest.TestCase):
    def test_output_parser_reads_confirmed_sku_quantities(self):
        rows = parse_output_rows("C-TEST-A\t1\n\nC-TEST-B  0.5")
        self.assertEqual(
            rows,
            [
                {"SKU": "C-TEST-A", "Quantity": 1.0},
                {"SKU": "C-TEST-B", "Quantity": 0.5},
            ],
        )

    def test_workbook_contains_seventeen_cases_and_one_confirmed_case(self):
        cases = load_cases(DEFAULT_SOURCE)
        self.assertEqual(len(cases), 17)
        confirmed = [case for case in cases if case["expected_rows"]]
        self.assertEqual(len(confirmed), 1)
        self.assertEqual(confirmed[0]["case_number"], "Case001")
        self.assertEqual(len(confirmed[0]["expected_rows"]), 8)


if __name__ == "__main__":
    unittest.main()
