"""
Unpivot (Wide → Long) operation.

Converts an Excel sheet from "wide" format (one row per entity, one column
per category) into "long / database" format (one row per entity-category pair).

Example
-------
Wide (input):
    Company  | Rent | Salary | Travel
    ABC Ltd  | 1000 |  5000  |  700

Long (output):
    Company  | Category | Value
    ABC Ltd  | Rent     | 1000
    ABC Ltd  | Salary   | 5000
    ABC Ltd  | Travel   |  700

Internally uses pandas.melt() for the transformation.

The user specifies:
  • id_cols      — columns to keep as row identifiers (e.g. ["Company"])
  • value_cols   — columns to unpivot (leave empty = all non-id columns)
  • var_name     — name for the new category column  (default "Category")
  • value_name   — name for the new value column     (default "Value")
"""

from __future__ import annotations

import os
from typing import Optional

import pandas as pd
import openpyxl

from core import wb_cache


def run_unpivot_step(
    src_file_path: str,
    src_sheet: str,
    src_header_row: int = 1,
    id_cols: list[str] = None,
    value_cols: list[str] = None,
    var_name: str = "Category",
    value_name: str = "Value",
    tgt_file: str = "",
    tgt_sheet: str = "Unpivot",
    append_mode: bool = False,
    drop_na: bool = True,
) -> tuple[bool, str]:
    """
    Unpivot (melt) a wide-format Excel sheet into long format.

    Parameters
    ----------
    src_file_path : path to source workbook
    src_sheet     : sheet name in source (empty = first sheet)
    src_header_row: 1-based row number of the header in source sheet
    id_cols       : column names to keep as identifiers (stay as rows)
    value_cols    : columns to unpivot; empty list = all non-id columns
    var_name      : name for the new "category" column
    value_name    : name for the new "value" column
    tgt_file      : path to target workbook (can be same as source)
    tgt_sheet     : sheet name to write the result into
    append_mode   : if True, append to existing sheet; else overwrite
    drop_na       : if True, drop rows where value_name is NaN/blank
    """

    # ── Validation ────────────────────────────────────────────────────────────
    if not src_file_path or not os.path.exists(src_file_path):
        return False, f"Source file not found: {src_file_path}"
    if not tgt_file:
        return False, "Target file path is required"
    if not id_cols:
        return False, "At least one ID column (identifier) is required"

    # ── Read source ───────────────────────────────────────────────────────────
    try:
        df = wb_cache.read_excel(
            src_file_path,
            sheet_name=src_sheet or 0,
            header=int(src_header_row) - 1,
        )
    except Exception as exc:
        return False, f"Could not read source sheet: {exc}"

    df.columns = [str(c).strip() for c in df.columns]

    # Validate id_cols exist
    missing_ids = [c for c in id_cols if c not in df.columns]
    if missing_ids:
        return False, (
            f"ID column(s) not found: {missing_ids}. "
            f"Available: {list(df.columns)}"
        )

    # Resolve value_cols — default = all columns that are not id_cols
    if value_cols:
        missing_vals = [c for c in value_cols if c not in df.columns]
        if missing_vals:
            return False, (
                f"Value column(s) not found: {missing_vals}. "
                f"Available: {list(df.columns)}"
            )
        vc = value_cols
    else:
        vc = [c for c in df.columns if c not in id_cols]

    if not vc:
        return False, "No value columns to unpivot — all columns are ID columns"

    # ── Melt ──────────────────────────────────────────────────────────────────
    try:
        melted = df.melt(
            id_vars=id_cols,
            value_vars=vc,
            var_name=var_name or "Category",
            value_name=value_name or "Value",
        )
    except Exception as exc:
        return False, f"Unpivot (melt) failed: {exc}"

    if drop_na:
        melted = melted.dropna(subset=[value_name or "Value"])
        melted = melted[melted[value_name or "Value"] != ""]

    # Preserve column order: id_cols, var_name, value_name
    melted = melted.reset_index(drop=True)

    rows_out = len(melted)

    # ── Write to target ───────────────────────────────────────────────────────
    try:
        if os.path.exists(tgt_file) and append_mode:
            # Append: load existing workbook, add/replace the sheet
            tgt_wb = wb_cache.load(tgt_file)
            if tgt_sheet in tgt_wb.sheetnames:
                del tgt_wb[tgt_sheet]
            ws_new = tgt_wb.create_sheet(tgt_sheet)
            # Write header
            for ci, col in enumerate(melted.columns, 1):
                ws_new.cell(row=1, column=ci, value=col)
            # Write data
            for ri, row_data in enumerate(melted.itertuples(index=False), 2):
                for ci, val in enumerate(row_data, 1):
                    ws_new.cell(row=ri, column=ci, value=val)
            wb_cache.save(tgt_wb, tgt_file)
        else:
            # Overwrite: use ExcelWriter in 'a' (append) mode if file exists,
            # create new file otherwise.
            if os.path.exists(tgt_file):
                with wb_cache.excel_writer(tgt_file, engine="openpyxl", mode="a",
                                           if_sheet_exists="replace") as writer:
                    melted.to_excel(writer, sheet_name=tgt_sheet, index=False)
            else:
                # New file — write directly
                import io
                buf = io.BytesIO()
                melted.to_excel(buf, sheet_name=tgt_sheet, index=False)
                buf.seek(0)
                with open(tgt_file, "wb") as fh:
                    fh.write(buf.read())

    except Exception as exc:
        return False, f"Could not write to target: {exc}"

    return True, (
        f"Unpivoted {len(vc)} column(s) × {len(df)} row(s) → "
        f"{rows_out} long-format row(s) written to '{tgt_sheet}'"
    )
