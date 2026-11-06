"""Import structured historical order-regression cases into Supabase."""

from __future__ import annotations

import re
import tomllib
from pathlib import Path

from evaluation_store import EvaluationStore


PROJECT_DIR = Path(__file__).resolve().parent
SOURCE_FILES = [
    PROJECT_DIR / "tests" / "test_data_cases.md",
    PROJECT_DIR / "parser_regression_examples.md",
]


def iter_markdown_cases(path: Path):
    parts = re.split(r"^## ", path.read_text(encoding="utf-8"), flags=re.MULTILINE)
    for part in parts[1:]:
        title = part.splitlines()[0].strip()
        blocks = re.findall(r"```text\n(.*?)\n```", part, flags=re.DOTALL)
        if len(blocks) < 2:
            continue
        raw_input = blocks[0].strip()
        expected_rows = []
        for line in blocks[1].splitlines():
            if "NEEDS REVIEW" in line.upper():
                continue
            match = re.match(
                r"^([A-Z0-9]+-[A-Z0-9+&-]+)\s+([0-9]+(?:\.[0-9]+)?)$",
                line.strip(),
                flags=re.IGNORECASE,
            )
            if match:
                expected_rows.append(
                    {"SKU": match.group(1), "Quantity": float(match.group(2))}
                )
        if raw_input and expected_rows:
            yield title, raw_input, expected_rows


def load_store() -> EvaluationStore:
    secrets_path = PROJECT_DIR / ".streamlit" / "secrets.toml"
    with secrets_path.open("rb") as handle:
        settings = tomllib.load(handle)["supabase"]
    return EvaluationStore.from_settings(settings)


def main():
    store = load_store()
    seen_inputs = set()
    imported = 0
    skipped = 0
    item_count = 0
    for source_path in SOURCE_FILES:
        for title, raw_input, expected_rows in iter_markdown_cases(source_path):
            normalized_input = raw_input.strip().replace("\r\n", "\n")
            if normalized_input in seen_inputs:
                skipped += 1
                continue
            seen_inputs.add(normalized_input)
            _, created = store.import_confirmed_case(
                title=title,
                raw_input=raw_input,
                expected_rows=expected_rows,
                source_file=source_path.name,
            )
            if created:
                imported += 1
                item_count += len(expected_rows)
                print(f"IMPORTED: {title} ({len(expected_rows)} items)")
            else:
                skipped += 1
                print(f"SKIPPED: {title} (already imported)")
    totals = store.counts()
    print(
        f"complete imported={imported} skipped={skipped} "
        f"new_items={item_count} total_cases={totals['cases']} "
        f"total_items={totals['items']}"
    )


if __name__ == "__main__":
    main()
