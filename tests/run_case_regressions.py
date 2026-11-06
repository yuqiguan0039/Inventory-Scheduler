import ast
import getpass
import io
import json
import os
import re
import sys
from copy import copy
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook


APP_PATH = Path(__file__).resolve().parents[1] / "app0501.py"
# The consolidated Markdown fixture is now the canonical test source.  The
# older parser_regression_examples.md remains as historical documentation.
EXAMPLES_PATH = Path(__file__).with_name("test_data_cases.md")


class FakeStreamlit:
    def cache_data(self, *args, **kwargs):
        def decorator(fn):
            return fn

        return decorator


def load_parser_namespace():
    source = APP_PATH.read_text()
    tree = ast.parse(source)
    keep_assignments = {
        "DEFAULT_PRODUCT_CATALOG_PATH",
        "DEFAULT_BRAND_ALIAS_PATH",
        "DEFAULT_TEMPLATE_PATH",
        "PRODUCT_HEADER_ALIASES",
        "PRODUCT_GROUP_ALIASES",
        "FLAVOR_TEXT_ALIASES",
        "ALL_IN_STOCK_SERIES_PREFIX",
        "SHIPMENT_INCREMENT",
    }
    selected_nodes = []
    for node in tree.body:
        if isinstance(node, ast.Assign):
            names = {
                child.id
                for target in node.targets
                for child in ast.walk(target)
                if isinstance(child, ast.Name)
            }
            if names & keep_assignments:
                selected_nodes.append(node)
        elif isinstance(node, ast.FunctionDef):
            selected_nodes.append(node)

    module = ast.Module(body=selected_nodes, type_ignores=[])
    ast.fix_missing_locations(module)
    namespace = {
        "__file__": str(APP_PATH),
        "st": FakeStreamlit(),
        "pd": pd,
        "re": re,
        "json": json,
        "os": os,
        "sys": sys,
        "io": io,
        "getpass": getpass,
        "Path": Path,
        "SequenceMatcher": SequenceMatcher,
        "copy": copy,
        "datetime": datetime,
        "load_workbook": load_workbook,
    }
    exec(compile(module, str(APP_PATH), "exec"), namespace)
    return namespace


def iter_examples(markdown_text):
    parts = re.split(r"^## ", markdown_text, flags=re.MULTILINE)
    for part in parts[1:]:
        title = part.splitlines()[0].strip()
        blocks = re.findall(r"```text\n(.*?)\n```", part, flags=re.DOTALL)
        if len(blocks) < 2:
            continue
        raw_text = blocks[0]
        expected_lines = [line.strip() for line in blocks[1].splitlines() if line.strip()]
        yield title, raw_text, expected_lines


def parse_expected_records(expected_lines):
    """Extract optional SKU/quantity assertions from a case's expected block."""
    records = {}
    for line in expected_lines:
        if "NEEDS REVIEW" in line.upper():
            continue
        parts = re.split(r"\s+", line.strip(), maxsplit=1)
        if not parts or not re.match(r"^[A-Z0-9]+-[A-Z0-9+&-]+$", parts[0], re.IGNORECASE):
            continue
        quantity = None
        if len(parts) == 2:
            try:
                quantity = float(parts[1].strip())
            except ValueError:
                pass
        records[parts[0].strip()] = quantity
    return records


def main():
    namespace = load_parser_namespace()
    catalog_df, catalog_error = namespace["load_product_catalog"](
        str(namespace["DEFAULT_PRODUCT_CATALOG_PATH"])
    )
    if catalog_error:
        raise RuntimeError(catalog_error)

    failures = []
    for title, raw_text, expected_lines in iter_examples(EXAMPLES_PATH.read_text()):
        if "[image_text]" in title:
            print(f"SKIP: {title} | image fixture requires OCR text")
            continue
        expected_records = parse_expected_records(expected_lines)
        if not expected_records:
            print(f"SKIP: {title} | no structured SKU assertions")
            continue
        rows, parser_name, review_rows = namespace["parse_demand_with_local_fallback"](
            raw_text,
            catalog_df,
        )
        found_quantities = {}
        for row in rows:
            sku = str(row.get("SKU", "")).strip()
            found_quantities[sku] = found_quantities.get(sku, 0.0) + float(
                row.get("Quantity", 0) or 0
            )
        missing = [sku for sku in expected_records if sku not in found_quantities]
        quantity_mismatch = [
            sku for sku, expected_quantity in expected_records.items()
            if expected_quantity is not None
            and sku in found_quantities
            and abs(found_quantities[sku] - expected_quantity) > 1e-9
        ]
        total_quantity = sum(float(row.get("Quantity", 0) or 0) for row in rows)
        review_count = sum(1 for row in review_rows if row.get("Need Review"))
        status = "PASS" if not missing and not quantity_mismatch else "FAIL"
        print(
            f"{status}: {title} | parser={parser_name} | "
            f"rows={len(rows)} | qty={total_quantity:g} | review={review_count}"
        )
        if missing:
            print(f"  Missing expected SKU(s): {', '.join(missing)}")
        if quantity_mismatch:
            print(f"  Quantity mismatch: {', '.join(quantity_mismatch)}")
        if missing or quantity_mismatch:
            failures.append((title, missing, quantity_mismatch))

    if failures:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
