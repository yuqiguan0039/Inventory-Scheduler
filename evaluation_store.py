"""Supabase persistence for order-recognition evaluation data."""

from __future__ import annotations

import hashlib
from datetime import datetime, timezone
from typing import Any, Mapping, Sequence

from supabase import Client, create_client


def input_fingerprint(raw_input: str) -> str:
    """Return a stable hash without storing another copy of the source text."""
    normalized = str(raw_input or "").strip().replace("\r\n", "\n")
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


def normalize_answer_rows(rows: Sequence[Mapping[str, Any]] | None) -> list[dict]:
    """Keep positive SKU/quantity rows in their user-visible order."""
    normalized = []
    for row in rows or []:
        sku = str(row.get("SKU", "") or "").strip()
        try:
            quantity = float(row.get("Quantity"))
        except (TypeError, ValueError):
            continue
        if not sku or quantity <= 0:
            continue
        normalized.append(
            {
                "sku": sku,
                "quantity": quantity,
                "variant": str(row.get("Requested Version", "") or "").strip()
                or None,
            }
        )
    return normalized


def comparison_fields(machine: Mapping[str, Any], confirmed: Mapping[str, Any]) -> dict:
    """Compare one prediction with one confirmed answer."""
    machine_sku = str(machine.get("machine_sku", "") or "").strip().upper()
    confirmed_sku = str(confirmed.get("sku", "") or "").strip().upper()
    machine_quantity = machine.get("machine_quantity")
    confirmed_quantity = confirmed.get("quantity")
    machine_variant = str(machine.get("machine_variant", "") or "").strip().upper()
    confirmed_variant = str(confirmed.get("variant", "") or "").strip().upper()

    sku_correct = bool(machine_sku and machine_sku == confirmed_sku)
    quantity_correct = (
        machine_quantity is not None
        and confirmed_quantity is not None
        and abs(float(machine_quantity) - float(confirmed_quantity)) < 1e-9
    )
    variant_correct = (
        True
        if not confirmed_variant
        else bool(machine_variant and machine_variant == confirmed_variant)
    )
    completely_correct = sku_correct and quantity_correct and variant_correct
    if completely_correct:
        error_type = None
    elif not sku_correct:
        error_type = "sku"
    elif not quantity_correct:
        error_type = "quantity"
    else:
        error_type = "variant"
    return {
        "sku_correct": sku_correct,
        "quantity_correct": quantity_correct,
        "variant_correct": variant_correct,
        "completely_correct": completely_correct,
        "error_type": error_type,
    }


class EvaluationStore:
    """Small database boundary used by both Streamlit and import scripts."""

    def __init__(self, client: Client):
        self.client = client

    @classmethod
    def from_settings(cls, settings: Mapping[str, Any]) -> "EvaluationStore":
        url = str(settings.get("url", "") or "").strip()
        key = str(settings.get("key", "") or "").strip()
        if not url or not key or "PASTE_YOUR_" in url or "PASTE_YOUR_" in key:
            raise ValueError("Supabase URL and Secret Key are not configured.")
        return cls(create_client(url, key))

    def create_case(
        self,
        raw_input: str,
        *,
        source_type: str = "text",
        ai_enabled: bool = False,
        ai_model: str | None = None,
        parser_version: str | None = None,
        status: str = "pending",
        operator_name: str | None = None,
        notes: str | None = None,
        input_metadata: Mapping[str, Any] | None = None,
    ) -> str:
        payload = {
            "source_type": source_type,
            "raw_input": str(raw_input),
            "input_metadata": dict(input_metadata or {}),
            "parser_version": parser_version,
            "ai_enabled": bool(ai_enabled),
            "ai_model": ai_model,
            "status": status,
            "operator_name": operator_name,
            "notes": notes,
            "input_hash": input_fingerprint(raw_input),
        }
        response = self.client.table("evaluation_cases").insert(payload).execute()
        return str(response.data[0]["id"])

    def find_imported_case(
        self,
        raw_input: str,
        *,
        dataset: str,
        external_id: str | None = None,
    ) -> str | None:
        """Find a previously imported case by source dataset and external ID."""
        fingerprint = input_fingerprint(raw_input)
        response = (
            self.client.table("evaluation_cases")
            .select("id,input_metadata")
            .eq("input_hash", fingerprint)
            .execute()
        )
        for row in response.data or []:
            metadata = row.get("input_metadata") or {}
            if metadata.get("dataset") != dataset:
                continue
            if external_id and metadata.get("external_id") != external_id:
                continue
            return str(row["id"])
        return None

    def find_historical_case(self, raw_input: str) -> str | None:
        return self.find_imported_case(
            raw_input,
            dataset="historical_regression",
        )

    def save_machine_items(
        self,
        case_id: str,
        rows: Sequence[Mapping[str, Any]] | None,
        *,
        parser_name: str | None = None,
    ) -> int:
        normalized = normalize_answer_rows(rows)
        if not normalized:
            return 0
        payload = []
        for index, row in enumerate(normalized, start=1):
            payload.append(
                {
                    "case_id": case_id,
                    "line_number": index,
                    "machine_sku": row["sku"],
                    "machine_quantity": row["quantity"],
                    "machine_variant": row["variant"],
                    "machine_match_method": parser_name,
                    "needs_review": False,
                }
            )
        self.client.table("evaluation_items").upsert(
            payload, on_conflict="case_id,line_number"
        ).execute()
        return len(payload)

    def confirm_case(
        self,
        case_id: str,
        rows: Sequence[Mapping[str, Any]] | None,
    ) -> int:
        confirmed_rows = normalize_answer_rows(rows)
        if not confirmed_rows:
            raise ValueError("Confirmed answer contains no positive SKU quantities.")

        response = (
            self.client.table("evaluation_items")
            .select("*")
            .eq("case_id", case_id)
            .order("line_number")
            .execute()
        )
        machine_rows = list(response.data or [])
        total_rows = max(len(machine_rows), len(confirmed_rows))
        payload = []
        for index in range(total_rows):
            machine = machine_rows[index] if index < len(machine_rows) else {}
            confirmed = confirmed_rows[index] if index < len(confirmed_rows) else None
            base = {
                "case_id": case_id,
                "line_number": index + 1,
                # Supabase requires this column to be non-null. A confirmed row
                # with no matching machine prediction is, by definition, a row
                # that needs review.
                "needs_review": (
                    bool(machine.get("needs_review")) if machine else True
                ),
            }
            if machine:
                for field in (
                    "machine_brand",
                    "machine_sku",
                    "machine_quantity",
                    "machine_variant",
                    "machine_match_method",
                    "machine_confidence",
                ):
                    base[field] = machine.get(field)
            if confirmed is None:
                base.update(
                    {
                        "confirmed_sku": None,
                        "confirmed_quantity": None,
                        "confirmed_variant": None,
                        "sku_correct": False,
                        "quantity_correct": False,
                        "variant_correct": False,
                        "completely_correct": False,
                        "error_type": "extra_prediction",
                    }
                )
            else:
                base.update(
                    {
                        "confirmed_sku": confirmed["sku"],
                        "confirmed_quantity": confirmed["quantity"],
                        "confirmed_variant": confirmed["variant"],
                    }
                )
                if machine:
                    base.update(comparison_fields(machine, confirmed))
                else:
                    base.update(
                        {
                            "sku_correct": False,
                            "quantity_correct": False,
                            "variant_correct": False,
                            "completely_correct": False,
                            "error_type": "missing_prediction",
                        }
                    )
            payload.append(base)

        self.client.table("evaluation_items").upsert(
            payload, on_conflict="case_id,line_number"
        ).execute()
        self.client.table("evaluation_cases").update(
            {
                "status": "confirmed",
                "updated_at": datetime.now(timezone.utc).isoformat(),
            }
        ).eq("id", case_id).execute()
        return len(confirmed_rows)

    def import_confirmed_case(
        self,
        *,
        title: str,
        raw_input: str,
        expected_rows: Sequence[Mapping[str, Any]],
        source_file: str,
        dataset: str = "historical_regression",
        external_id: str | None = None,
        extra_metadata: Mapping[str, Any] | None = None,
    ) -> tuple[str, bool]:
        existing_id = self.find_imported_case(
            raw_input,
            dataset=dataset,
            external_id=external_id,
        )
        if existing_id:
            return existing_id, False
        case_id = self.create_case(
            raw_input,
            source_type="text",
            status="confirmed",
            notes="Imported historical regression case; no original machine output.",
            input_metadata={
                "dataset": dataset,
                "external_id": external_id,
                "title": title,
                "source_file": source_file,
                **dict(extra_metadata or {}),
            },
        )
        payload = []
        for index, row in enumerate(expected_rows, start=1):
            payload.append(
                {
                    "case_id": case_id,
                    "line_number": index,
                    "confirmed_sku": str(row["SKU"]),
                    "confirmed_quantity": float(row["Quantity"]),
                    "sku_correct": None,
                    "quantity_correct": None,
                    "variant_correct": None,
                    "completely_correct": None,
                }
            )
        try:
            self.client.table("evaluation_items").insert(payload).execute()
        except Exception:
            self.client.table("evaluation_cases").delete().eq("id", case_id).execute()
            raise
        return case_id, True

    def counts(self) -> dict[str, int]:
        cases = (
            self.client.table("evaluation_cases")
            .select("id", count="exact")
            .limit(1)
            .execute()
        )
        items = (
            self.client.table("evaluation_items")
            .select("id", count="exact")
            .limit(1)
            .execute()
        )
        return {"cases": int(cases.count or 0), "items": int(items.count or 0)}
