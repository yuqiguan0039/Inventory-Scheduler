import json
import os
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4


DEFAULT_HISTORY_PATH = Path(__file__).resolve().parent / "data" / "order_history.jsonl"
DEFAULT_ANSWER_PATH = Path(__file__).resolve().parent / "data" / "order_answers.jsonl"


def _append_json_line(path, record):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    line = (json.dumps(record, ensure_ascii=False) + "\n").encode("utf-8")

    file_descriptor = os.open(
        path,
        os.O_WRONLY | os.O_CREAT | os.O_APPEND,
        0o600,
    )
    try:
        os.write(file_descriptor, line)
    finally:
        os.close(file_descriptor)


def record_order_detail(raw_text, use_ai=False, history_path=None):
    """Append one submitted pasted order to a local JSON Lines history file."""
    order_detail = str(raw_text or "")
    if not order_detail.strip():
        raise ValueError("Order detail is empty and was not recorded.")

    path = Path(history_path or DEFAULT_HISTORY_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)

    record = {
        "record_id": uuid4().hex,
        "submitted_at_utc": datetime.now(timezone.utc).isoformat(),
        "source": "paste_text",
        "use_ai": bool(use_ai),
        "order_detail": order_detail,
    }
    _append_json_line(path, record)

    return record, path


def load_order_history(history_path=None):
    """Read valid order-history records, skipping blank lines."""
    path = Path(history_path or DEFAULT_HISTORY_PATH)
    if not path.exists():
        return []

    records = []
    with path.open("r", encoding="utf-8") as history_file:
        for line_number, line in enumerate(history_file, start=1):
            if not line.strip():
                continue
            try:
                records.append(json.loads(line))
            except json.JSONDecodeError as exc:
                raise ValueError(
                    f"Invalid order-history record on line {line_number}."
                ) from exc
    return records


def record_order_answer(record_id, answer_rows, answer_path=None):
    """Append a human-confirmed SKU/quantity answer linked to an order record."""
    linked_record_id = str(record_id or "").strip()
    if not linked_record_id:
        raise ValueError("Order record ID is required to save a correct answer.")

    normalized_rows = []
    for row in answer_rows or []:
        sku = str(row.get("SKU", "") or "").strip()
        try:
            quantity = float(row.get("Quantity"))
        except (TypeError, ValueError):
            continue
        if sku and quantity > 0:
            normalized_rows.append({"SKU": sku, "Quantity": quantity})

    if not normalized_rows:
        raise ValueError("Correct answer is empty and was not recorded.")

    path = Path(answer_path or DEFAULT_ANSWER_PATH)
    record = {
        "answer_id": uuid4().hex,
        "order_record_id": linked_record_id,
        "confirmed_at_utc": datetime.now(timezone.utc).isoformat(),
        "source": "confirmed_demand_editor",
        "correct_answer": normalized_rows,
    }
    _append_json_line(path, record)
    return record, path


def load_order_answers(answer_path=None):
    """Read saved correct-answer snapshots, skipping blank lines."""
    path = Path(answer_path or DEFAULT_ANSWER_PATH)
    if not path.exists():
        return []

    records = []
    with path.open("r", encoding="utf-8") as answer_file:
        for line_number, line in enumerate(answer_file, start=1):
            if not line.strip():
                continue
            try:
                records.append(json.loads(line))
            except json.JSONDecodeError as exc:
                raise ValueError(
                    f"Invalid order-answer record on line {line_number}."
                ) from exc
    return records
