"""
Sheet Row Inserter operation.

Reads a master Excel sheet with three key columns:
  • Category   — matches a target sheet by name
  • Particulars — the value to write in the inserted row (or the anchor search term)
  • Value        — another value to write in the inserted row

For every (category, particulars) group in the master the engine:
  1. Finds the target sheet whose name matches the category value
  2. Locates an anchor row inside that sheet (fixed value OR derived from the
     Particulars value itself)
  3. Inserts new rows at the anchor position (before / after / after-last)
  4. Writes master column values into the newly inserted rows

This is distinct from Sheet Updater (which updates existing cells).
Sheet Row Inserter physically inserts new rows into the middle of the sheet,
allowing values to be placed inside an existing report structure rather than
appended to the end.

Multiple insertions in the same sheet are applied bottom-to-top so that
earlier row positions are never affected by later insertions.
"""

from __future__ import annotations

import os
from typing import Optional

import pandas as pd

from core import wb_cache
from operations.sheet_updater import (
    SHEET_MATCH_MODES,
    SHEET_MATCH_KEYS,
    SHEET_MATCH_LABELS,
    _match_sheet,
)

# ---------------------------------------------------------------------------
# Anchor match mode registry
# ---------------------------------------------------------------------------

ANCHOR_MATCH_MODES: dict[str, str] = {
    "exact":       "Exact match (case-insensitive)",
    "contains":    "Contains (case-insensitive)",
    "starts_with": "Starts with",
    "ends_with":   "Ends with",
}
ANCHOR_MATCH_KEYS:   list[str] = list(ANCHOR_MATCH_MODES.keys())
ANCHOR_MATCH_LABELS: list[str] = list(ANCHOR_MATCH_MODES.values())

# ---------------------------------------------------------------------------
# Insert position registry
# ---------------------------------------------------------------------------

INSERT_POSITIONS: dict[str, str] = {
    "before_first": "Before first matching row",
    "after_first":  "After first matching row",
    "after_last":   "After last matching row  *(most common)*",
    "before_last":  "Before last matching row",
}
INSERT_POSITION_KEYS:   list[str] = list(INSERT_POSITIONS.keys())
INSERT_POSITION_LABELS: list[str] = list(INSERT_POSITIONS.values())


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _resolve_col_idx(ws, col_ref: str) -> Optional[int]:
    """
    Resolve a column reference to a 1-based column index.
    Accepts: column letter (A, B, AA), integer string (1, 3),
             or header name found in row 1.
    """
    col_ref = str(col_ref).strip()
    if not col_ref:
        return None
    try:
        from openpyxl.utils import column_index_from_string
        return column_index_from_string(col_ref)
    except Exception:
        pass
    try:
        idx = int(col_ref)
        if idx >= 1:
            return idx
    except ValueError:
        pass
    col_lower = col_ref.lower()
    for cell in ws[1]:
        if cell.value is not None and str(cell.value).strip().lower() == col_lower:
            return cell.column
    return None


def _find_df_col(df: pd.DataFrame, ref: str) -> Optional[str]:
    """Resolve a column reference to an actual DataFrame column name."""
    ref = str(ref).strip()
    ref_lower = ref.lower()
    for c in df.columns:
        if str(c).strip().lower() == ref_lower:
            return c
    try:
        from openpyxl.utils import column_index_from_string
        idx = column_index_from_string(ref) - 1
        if 0 <= idx < len(df.columns):
            return df.columns[idx]
    except Exception:
        pass
    try:
        idx = int(ref) - 1
        if 0 <= idx < len(df.columns):
            return df.columns[idx]
    except ValueError:
        pass
    return None


def _match_anchor(cell_val, anchor_value: str, anchor_match: str) -> bool:
    """Return True if cell_val satisfies the anchor condition."""
    cv = str(cell_val).strip().lower() if cell_val is not None else ""
    av = str(anchor_value).strip().lower()
    if not cv or not av:
        return False
    if anchor_match == "contains":
        return av in cv
    if anchor_match == "starts_with":
        return cv.startswith(av)
    if anchor_match == "ends_with":
        return cv.endswith(av)
    return cv == av   # exact (default)


def _find_anchor_rows(ws, anchor_col_idx: int, anchor_value: str,
                      anchor_match: str, header_row: int) -> list[int]:
    """Return a list of 1-based row numbers in the sheet where the anchor matches."""
    matches: list[int] = []
    for row in ws.iter_rows(min_row=header_row + 1):
        cell = row[anchor_col_idx - 1]
        if cell.value is not None and _match_anchor(cell.value, anchor_value, anchor_match):
            matches.append(cell.row)
    return matches


def _insert_row_idx(anchor_rows: list[int], insert_position: str) -> Optional[int]:
    """
    Return the 1-based row index at which ws.insert_rows() should be called.
    insert_rows(n, count) inserts `count` blank rows BEFORE row n.
    """
    if not anchor_rows:
        return None
    if insert_position == "before_first":
        return anchor_rows[0]
    if insert_position == "after_first":
        return anchor_rows[0] + 1
    if insert_position == "before_last":
        return anchor_rows[-1]
    return anchor_rows[-1] + 1   # after_last (default)


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def run_sheet_row_inserter_step(
    src_file_path: str,
    src_sheet: str,
    src_col_category: str,
    src_col_particulars: str,
    src_col_value: str,
    tgt_file: str,
    sheet_match_mode: str = "cat_in_sheet",
    # Anchor configuration
    anchor_col: str = "",
    anchor_mode: str = "particulars",   # "particulars" | "fixed"
    anchor_value: str = "",             # used when anchor_mode = "fixed"
    anchor_match: str = "exact",
    insert_position: str = "after_last",
    # Target column mapping
    tgt_col_particulars: str = "",
    tgt_col_value: str = "",
    extra_col_map: list = None,         # [{"src_col": "X", "tgt_col": "Y"}, ...]
    tgt_header_row: int = 1,
    # Particulars filter
    include_particulars: list = None,   # if non-empty, only these Particulars are processed
) -> tuple[bool, str]:
    """
    Insert rows from a master sheet into matching sheets of a target workbook.

    Parameters
    ----------
    src_file_path       : path to master workbook
    src_sheet           : sheet in master (empty = first sheet)
    src_col_category    : category column header/letter in master
    src_col_particulars : particulars column header/letter in master
    src_col_value       : value column header/letter in master
    tgt_file            : path to target workbook
    sheet_match_mode    : how to match category → sheet name
    anchor_col          : column in target sheets to search for the anchor
    anchor_mode         : "particulars" — use the Particulars value as the anchor;
                          "fixed" — use anchor_value for all groups
    anchor_value        : the search term when anchor_mode = "fixed"
    anchor_match        : how to compare anchor (exact/contains/starts_with/ends_with)
    insert_position     : where relative to the anchor row to insert
                          (before_first / after_first / after_last / before_last)
    tgt_col_particulars : target column where the Particulars value is written
    tgt_col_value       : target column where the Value is written
    extra_col_map       : additional master→target column mappings
    tgt_header_row      : 1-based header row index in target sheets
    """

    # ── Validation ────────────────────────────────────────────────────────────
    if not src_file_path or not os.path.exists(src_file_path):
        return False, f"Master file not found: {src_file_path}"
    if not tgt_file or not os.path.exists(tgt_file):
        return False, f"Target file not found: {tgt_file}"
    if not src_col_category:
        return False, "Category column (master) is required"
    if not src_col_particulars:
        return False, "Particulars column (master) is required"
    if not src_col_value:
        return False, "Value column (master) is required"
    if not anchor_col:
        return False, "Anchor column (target) is required"
    if anchor_mode == "fixed" and not anchor_value:
        return False, "Anchor value is required when anchor mode is 'Fixed value'"

    # ── Read master ───────────────────────────────────────────────────────────
    try:
        df = wb_cache.read_excel(src_file_path, sheet_name=src_sheet or 0, dtype=str)
    except Exception as exc:
        return False, f"Could not read master sheet: {exc}"

    df.columns = [str(c).strip() for c in df.columns]

    cat_col  = _find_df_col(df, src_col_category)
    part_col = _find_df_col(df, src_col_particulars)
    val_col  = _find_df_col(df, src_col_value)

    missing = []
    if not cat_col:  missing.append(f"category '{src_col_category}'")
    if not part_col: missing.append(f"particulars '{src_col_particulars}'")
    if not val_col:  missing.append(f"value '{src_col_value}'")
    if missing:
        return False, (
            f"Column(s) not found in master: {', '.join(missing)}. "
            f"Available: {list(df.columns)}"
        )

    df = df.dropna(subset=[cat_col, part_col], how="all")
    if df.empty:
        return False, "Master sheet has no data rows."

    # ── Filter to selected Particulars only ───────────────────────────────────
    if include_particulars:
        _inc_lower = {str(p).strip().lower() for p in include_particulars if str(p).strip()}
        df = df[df[part_col].apply(lambda v: str(v).strip().lower() in _inc_lower)]
        if df.empty:
            return False, (
                f"No master rows match the selected Particulars filter: "
                f"{include_particulars}"
            )

    # ── Load target workbook ──────────────────────────────────────────────────
    try:
        tgt_wb = wb_cache.load(tgt_file)
    except Exception as exc:
        return False, f"Could not open target file: {exc}"

    sheet_names = tgt_wb.sheetnames

    # ── Resolve extra column map src column names once ────────────────────────
    extra_maps_resolved: list[tuple] = []   # [(src_df_col, tgt_col_ref), ...]
    for m in (extra_col_map or []):
        sc = _find_df_col(df, m.get("src_col", ""))
        tc = m.get("tgt_col", "").strip()
        if sc and tc:
            extra_maps_resolved.append((sc, tc))

    # ── Build insertion jobs ──────────────────────────────────────────────────
    # Group by category (and particulars when anchor_mode = "particulars")
    if anchor_mode == "particulars":
        group_keys = [cat_col, part_col]
    else:
        group_keys = [cat_col]

    # sheet_jobs: { sheet_name: [(insert_at_row, rows_data_list), ...] }
    sheet_jobs: dict[str, list] = {}
    skipped:    list[str]       = []
    sheets_used: set[str]       = set()

    for group_key, group_df in df.groupby(group_keys, sort=False):
        if isinstance(group_key, str):
            category   = group_key
            particular = ""
        else:
            category   = group_key[0]
            particular = group_key[1] if len(group_key) > 1 else ""

        category   = str(category).strip()
        particular = str(particular).strip()
        if not category:
            continue

        # 1. Match sheet
        matched = _match_sheet(category, sheet_names, sheet_match_mode)
        if not matched:
            skipped.append(f"No sheet matched category '{category}'")
            continue

        ws = tgt_wb[matched]

        # 2. Resolve anchor column
        anchor_col_idx = _resolve_col_idx(ws, anchor_col)
        if not anchor_col_idx:
            skipped.append(f"Anchor col '{anchor_col}' not found in '{matched}'")
            continue

        # 3. Determine anchor search value
        search_value = particular if anchor_mode == "particulars" else anchor_value

        # 4. Find anchor rows
        anchor_rows = _find_anchor_rows(ws, anchor_col_idx, search_value, anchor_match, tgt_header_row)
        if not anchor_rows:
            skipped.append(
                f"Anchor '{search_value}' not found in '{matched}' (col '{anchor_col}')"
            )
            continue

        # 5. Determine insert index
        insert_at = _insert_row_idx(anchor_rows, insert_position)
        if insert_at is None:
            skipped.append(f"Could not determine insert position for '{category}/{particular}' in '{matched}'")
            continue

        # 6. Resolve target write columns (per-sheet, cached by sheet name)
        tgt_part_idx = _resolve_col_idx(ws, tgt_col_particulars) if tgt_col_particulars else None
        tgt_val_idx  = _resolve_col_idx(ws, tgt_col_value)        if tgt_col_value        else None
        extra_idx_map: list[tuple[int, str]] = []   # (col_idx, df_col_name)
        for src_df_col, tgt_col_ref in extra_maps_resolved:
            idx = _resolve_col_idx(ws, tgt_col_ref)
            if idx:
                extra_idx_map.append((idx, src_df_col))

        # 7. Build row data list
        rows_data: list[dict] = []
        for _, master_row in group_df.iterrows():
            row_data: dict[int, object] = {}
            part_val = str(master_row[part_col]).strip() if pd.notna(master_row[part_col]) else ""
            raw_val  = master_row[val_col] if pd.notna(master_row[val_col]) else ""

            if tgt_part_idx and part_val:
                row_data[tgt_part_idx] = part_val
            if tgt_val_idx is not None:
                # Try to preserve numeric type
                try:
                    row_data[tgt_val_idx] = float(raw_val) if str(raw_val).strip() else ""
                    if row_data[tgt_val_idx] == int(row_data[tgt_val_idx]):
                        row_data[tgt_val_idx] = int(row_data[tgt_val_idx])
                except (ValueError, TypeError):
                    row_data[tgt_val_idx] = raw_val
            for ci, sc in extra_idx_map:
                v = master_row.get(sc)
                if pd.notna(v):
                    row_data[ci] = v

            rows_data.append(row_data)

        if matched not in sheet_jobs:
            sheet_jobs[matched] = []
        sheet_jobs[matched].append((insert_at, rows_data))
        sheets_used.add(matched)

    if not sheet_jobs and not skipped:
        return False, "No data to insert."

    # ── Apply insertions — bottom to top within each sheet ────────────────────
    total_inserted = 0

    for sn, jobs in sheet_jobs.items():
        ws = tgt_wb[sn]
        # Sort descending by insert_at so lower positions are unaffected
        for insert_at, rows_data in sorted(jobs, key=lambda x: x[0], reverse=True):
            n = len(rows_data)
            ws.insert_rows(insert_at, n)
            for i, row_data in enumerate(rows_data):
                for col_idx, value in row_data.items():
                    ws.cell(row=insert_at + i, column=col_idx).value = value
            total_inserted += n

    # ── Save ──────────────────────────────────────────────────────────────────
    wb_cache.save(tgt_wb, tgt_file)

    # ── Summary ───────────────────────────────────────────────────────────────
    summary = f"Inserted {total_inserted} row(s) across {len(sheets_used)} sheet(s)"
    if skipped:
        preview = skipped[:5]
        if len(skipped) > 5:
            preview.append(f"…and {len(skipped) - 5} more")
        summary += f". Unmatched: {'; '.join(preview)}"

    return True, summary
