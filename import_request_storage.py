"""Import the request-storage workbook into the evaluation dataset."""

from __future__ import annotations

import argparse
import re
import tomllib
from pathlib import Path

import pandas as pd

from evaluation_store import EvaluationStore


PROJECT_DIR = Path(__file__).resolve().parent
DEFAULT_SOURCE = Path(
    "/Users/guanyuqi/Library/CloudStorage/GoogleDrive-yuqiguan0039@gmail.com/"
    "Other computers/我的笔记本电脑/Documents/1.做单模板/"
    "request storage for test/request storage.xlsx"
)
DATASET_NAME = "request_storage_archive"


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def parse_output_rows(value) -> list[dict]:
    """Parse confirmed C-SKU and positive quantity rows from the output cell."""
    rows = []
    for line in clean_text(value).splitlines():
        match = re.match(
            r"^\s*([A-Z0-9]+-[A-Z0-9+&-]+)\s+([0-9]+(?:\.[0-9]+)?)\s*$",
            line,
            flags=re.IGNORECASE,
        )
        if match:
            rows.append(
                {"SKU": match.group(1).upper(), "Quantity": float(match.group(2))}
            )
    return rows


def load_cases(path: Path) -> list[dict]:
    frame = pd.read_excel(path, sheet_name="Sheet1", dtype=object)
    frame.columns = [str(column).strip() for column in frame.columns]
    cases = []
    for _, row in frame.iterrows():
        case_number = clean_text(row.get("Case#"))
        if not re.fullmatch(r"Case\d+", case_number, flags=re.IGNORECASE):
            continue
        raw_input = clean_text(row.get("request input"))
        document_link = clean_text(row.get("request input doc link"))
        if not raw_input:
            raw_input = f"FILE_REFERENCE: {document_link}" if document_link else (
                f"MISSING_SOURCE_TEXT: {case_number}"
            )
        request_date = row.get("request date")
        if pd.isna(request_date):
            request_date_text = None
        else:
            request_date_text = pd.Timestamp(request_date).date().isoformat()
        cases.append(
            {
                "case_number": case_number,
                "raw_input": raw_input,
                "document_link": document_link or None,
                "request_date": request_date_text,
                "sales_rep": clean_text(row.get("sales rep")) or None,
                "output_so_number": clean_text(row.get("output SO#")) or None,
                "complex_sample": clean_text(row.get("complex sample")) or None,
                "expected_rows": parse_output_rows(row.get("output")),
            }
        )
    return cases


def load_store() -> EvaluationStore:
    with (PROJECT_DIR / ".streamlit" / "secrets.toml").open("rb") as handle:
        settings = tomllib.load(handle)["supabase"]
    return EvaluationStore.from_settings(settings)


def import_cases(path: Path, *, dry_run: bool = False) -> dict:
    cases = load_cases(path)
    summary = {
        "source_cases": len(cases),
        "confirmed_cases": sum(bool(case["expected_rows"]) for case in cases),
        "pending_cases": sum(not case["expected_rows"] for case in cases),
        "confirmed_items": sum(len(case["expected_rows"]) for case in cases),
        "imported": 0,
        "skipped": 0,
    }
    if dry_run:
        return summary

    store = load_store()
    for case in cases:
        metadata = {
            "dataset": DATASET_NAME,
            "external_id": case["case_number"],
            "source_file": path.name,
            "request_date": case["request_date"],
            "sales_rep": case["sales_rep"],
            "output_so_number": case["output_so_number"],
            "complex_sample": case["complex_sample"],
            "request_document_link": case["document_link"],
        }
        existing_id = store.find_imported_case(
            case["raw_input"],
            dataset=DATASET_NAME,
            external_id=case["case_number"],
        )
        if existing_id:
            summary["skipped"] += 1
            continue

        if case["expected_rows"]:
            store.import_confirmed_case(
                title=case["case_number"],
                raw_input=case["raw_input"],
                expected_rows=case["expected_rows"],
                source_file=path.name,
                dataset=DATASET_NAME,
                external_id=case["case_number"],
                extra_metadata=metadata,
            )
        else:
            store.create_case(
                case["raw_input"],
                source_type="excel",
                status="pending",
                operator_name=case["sales_rep"],
                notes="Imported from request storage; confirmed output is pending.",
                input_metadata=metadata,
            )
        summary["imported"] += 1
    summary.update(store.counts())
    return summary


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    summary = import_cases(args.source, dry_run=args.dry_run)
    print(" ".join(f"{key}={value}" for key, value in summary.items()))


if __name__ == "__main__":
    main()
