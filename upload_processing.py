import hashlib
import io
from pathlib import Path

import pandas as pd


MAX_BATCH_UPLOADS = 10
SOURCE_FILE_COLUMN = "_Source File"

INVENTORY_COLUMN_ALIASES = {
    "Warehouse Name": ("Warehouse Name", "Warehouse", "Warehouse Code"),
    "SKU": ("SKU", "Item SKU", "Product SKU", "Item Number"),
    "Available For Sale": (
        "Available For Sale",
        "Accounting Available For Sale",
        "Physical Available For Sale Qty",
        "Physical Available For Sale",
        "Available For Sale Qty",
    ),
}


def _normalized_column_name(value):
    """Normalize a possible inventory header for exact alias matching."""
    return " ".join(str(value or "").strip().casefold().split())


def _inventory_header_row(raw_frame):
    """Locate a header row containing warehouse, SKU, and quantity fields."""
    alias_groups = [
        {_normalized_column_name(alias) for alias in aliases}
        for aliases in INVENTORY_COLUMN_ALIASES.values()
    ]
    best_row = None
    best_score = 0
    for row_index, row in raw_frame.iterrows():
        values = {_normalized_column_name(value) for value in row if pd.notna(value)}
        score = sum(bool(values & aliases) for aliases in alias_groups)
        if score > best_score:
            best_row = int(row_index)
            best_score = score
    return best_row if best_score >= 2 else 0


def _standardize_inventory_columns(frame):
    """Map supported inventory-export column names to shared canonical names."""
    normalized_columns = {
        _normalized_column_name(column): column for column in frame.columns
    }
    rename_map = {}
    for canonical, aliases in INVENTORY_COLUMN_ALIASES.items():
        if canonical in frame.columns:
            continue
        for alias in aliases:
            source = normalized_columns.get(_normalized_column_name(alias))
            if source is not None:
                rename_map[source] = canonical
                break
    return frame.rename(columns=rename_map)


def warehouse_region(warehouse_name):
    """Return the region portion of a standard SALE/STORAGE warehouse name."""
    normalized = str(warehouse_name or "").strip().upper()
    for suffix in ("_SALE", "_STORAGE"):
        if normalized.endswith(suffix):
            return normalized[: -len(suffix)]
    return ""


def is_eligible_warehouse(warehouse_name, include_storage=False):
    """Allow standard SALE warehouses and optional STORAGE, excluding LOCAL."""
    normalized = str(warehouse_name or "").strip().upper()
    if "_LOCAL_" in normalized:
        return False
    if normalized.endswith("_SALE"):
        return True
    return bool(include_storage and normalized.endswith("_STORAGE"))


def filter_inventory_warehouses(
    inventory_frame,
    warehouse_column,
    include_storage=False,
):
    """Return eligible inventory rows plus the ignored warehouse names."""
    if warehouse_column not in inventory_frame.columns:
        raise ValueError(f"Warehouse column not found: {warehouse_column}")

    warehouse_names = inventory_frame[warehouse_column].astype(str).str.strip()
    eligible_mask = warehouse_names.apply(
        lambda name: is_eligible_warehouse(name, include_storage)
    )
    ignored = sorted(
        name for name in warehouse_names[~eligible_mask].dropna().unique()
        if str(name).strip()
    )
    return inventory_frame.loc[eligible_mask].copy(), ignored


def build_warehouse_sequence(
    warehouse_names,
    region_priority,
    include_storage=False,
):
    """Order each region as SALE then optional STORAGE before the next region."""
    names = [] if warehouse_names is None else warehouse_names
    priorities = [] if region_priority is None else region_priority
    normalized_names = {
        str(name).strip().upper(): str(name).strip()
        for name in names
        if is_eligible_warehouse(name, include_storage)
    }
    sequence = []
    for region in priorities:
        region_name = str(region or "").strip().upper()
        if not region_name:
            continue
        candidates = [f"{region_name}_SALE"]
        if include_storage:
            candidates.append(f"{region_name}_STORAGE")
        for candidate in candidates:
            actual_name = normalized_names.get(candidate)
            if actual_name and actual_name not in sequence:
                sequence.append(actual_name)
    return sequence


def uploaded_file_fingerprint(uploaded_file):
    """Return a stable content fingerprint for an uploaded file."""
    return hashlib.sha256(uploaded_file.getvalue()).hexdigest()


def deduplicate_uploaded_files(uploaded_files):
    """Keep upload order while removing byte-for-byte duplicate files."""
    unique_files = []
    duplicate_names = []
    seen_fingerprints = set()

    for uploaded_file in uploaded_files or []:
        fingerprint = uploaded_file_fingerprint(uploaded_file)
        if fingerprint in seen_fingerprints:
            duplicate_names.append(str(uploaded_file.name))
            continue
        seen_fingerprints.add(fingerprint)
        unique_files.append(uploaded_file)

    return unique_files, duplicate_names


def read_inventory_upload(uploaded_file):
    """Read an inventory file and adapt supported export/header layouts."""
    suffix = Path(str(uploaded_file.name)).suffix.lower()
    content = uploaded_file.getvalue()
    if suffix == ".csv":
        preview = pd.read_csv(io.BytesIO(content), header=None, nrows=25)
        header_row = _inventory_header_row(preview)
        frame = pd.read_csv(io.BytesIO(content), header=header_row)
    elif suffix == ".xlsx":
        preview = pd.read_excel(
            io.BytesIO(content), header=None, nrows=25, sheet_name=0
        )
        header_row = _inventory_header_row(preview)
        frame = pd.read_excel(
            io.BytesIO(content), header=header_row, sheet_name=0
        )
    else:
        raise ValueError(f"Unsupported inventory file type: {suffix or 'unknown'}")

    frame.columns = [str(column).strip() for column in frame.columns]
    frame = frame.dropna(axis=0, how="all").dropna(axis=1, how="all")
    frame = _standardize_inventory_columns(frame)
    if frame.empty:
        raise ValueError("Inventory file contains no rows.")
    frame[SOURCE_FILE_COLUMN] = str(uploaded_file.name)
    return frame


def common_inventory_columns(frames):
    """Return columns shared by every inventory file, in first-file order."""
    if not frames:
        return []
    shared = set(frames[0].columns)
    for frame in frames[1:]:
        shared &= set(frame.columns)
    shared.discard(SOURCE_FILE_COLUMN)
    return [column for column in frames[0].columns if column in shared]


def find_cross_file_inventory_conflicts(
    inventory_frame,
    warehouse_column,
    sku_column,
    quantity_column,
):
    """Find Warehouse+SKU keys present in more than one source file."""
    required = {
        SOURCE_FILE_COLUMN,
        warehouse_column,
        sku_column,
        quantity_column,
    }
    missing = required - set(inventory_frame.columns)
    if missing:
        raise ValueError(
            "Inventory conflict check is missing column(s): "
            + ", ".join(sorted(missing))
        )

    working = inventory_frame[
        [SOURCE_FILE_COLUMN, warehouse_column, sku_column, quantity_column]
    ].copy()
    working[warehouse_column] = (
        working[warehouse_column].astype(str).str.strip().str.upper()
    )
    working[sku_column] = working[sku_column].astype(str).str.strip().str.upper()
    working[quantity_column] = pd.to_numeric(
        working[quantity_column], errors="coerce"
    ).fillna(0)
    working = working[
        working[warehouse_column].ne("") & working[sku_column].ne("")
    ]

    per_file = (
        working.groupby(
            [warehouse_column, sku_column, SOURCE_FILE_COLUMN],
            as_index=False,
            dropna=False,
        )[quantity_column]
        .sum()
    )
    source_counts = (
        per_file.groupby([warehouse_column, sku_column])[SOURCE_FILE_COLUMN]
        .nunique()
    )
    conflict_keys = source_counts[source_counts > 1].index
    if conflict_keys.empty:
        return pd.DataFrame(
            columns=[
                "Warehouse Name",
                "SKU",
                "Source Files",
                "Quantity by File",
            ]
        )

    conflict_key_set = set(conflict_keys.tolist())
    conflict_rows = per_file[
        per_file.apply(
            lambda row: (row[warehouse_column], row[sku_column])
            in conflict_key_set,
            axis=1,
        )
    ]

    output_rows = []
    for (warehouse, sku), group in conflict_rows.groupby(
        [warehouse_column, sku_column], sort=False
    ):
        details = [
            f"{row[SOURCE_FILE_COLUMN]}: {float(row[quantity_column]):g}"
            for _, row in group.iterrows()
        ]
        output_rows.append(
            {
                "Warehouse Name": warehouse,
                "SKU": sku,
                "Source Files": ", ".join(group[SOURCE_FILE_COLUMN].astype(str)),
                "Quantity by File": " | ".join(details),
            }
        )
    return pd.DataFrame(output_rows)
