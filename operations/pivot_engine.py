"""
Pivot Table Engine.

Uses pandas pivot_table to compute pivots and writes the result as structured
data to an Excel sheet.  No win32com — pure pandas + openpyxl.

Supports:
  - Multiple row fields (multi-level index)
  - Optional column pivot fields
  - Multiple value fields with independent aggregation functions
  - Pre-pivot row filtering (reuses filter_utils)
  - Grand Total row / column
  - Configurable output location (file, sheet, start cell)
  - Named template save / load
  - Optional header + subtotal formatting
"""

import os
import json
import traceback

import numpy as np
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string

from core import wb_cache
from operations.filter_utils import apply_filters


# ---------------------------------------------------------------------------
# Aggregation catalogue
# ---------------------------------------------------------------------------

AGGFUNC_KEYS = [
    "sum", "count", "mean", "median", "min", "max",
    "std", "var", "nunique", "first", "last",
]

AGGFUNC_LABELS = {
    "sum":     "Sum",
    "count":   "Count",
    "mean":    "Average (Mean)",
    "median":  "Median",
    "min":     "Min",
    "max":     "Max",
    "std":     "Std Dev",
    "var":     "Variance",
    "nunique": "Count Unique",
    "first":   "First Value",
    "last":    "Last Value",
}

# Map key → callable/string understood by pandas
_AGGFUNC_MAP = {
    "sum":     "sum",
    "count":   "count",
    "mean":    "mean",
    "median":  "median",
    "min":     "min",
    "max":     "max",
    "std":     "std",
    "var":     "var",
    "nunique": pd.Series.nunique,
    "first":   "first",
    "last":    "last",
}

# ---------------------------------------------------------------------------
# Template storage
# ---------------------------------------------------------------------------

TEMPLATES_DIR = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "templates", "pivot"
)


def list_pivot_templates() -> list:
    """Return list of saved pivot template names."""
    try:
        os.makedirs(TEMPLATES_DIR, exist_ok=True)
        return sorted(f[:-5] for f in os.listdir(TEMPLATES_DIR) if f.endswith(".json"))
    except Exception:
        return []


def save_pivot_template(name: str, config: dict) -> tuple:
    """Save pivot config (without file paths) as a named template."""
    try:
        os.makedirs(TEMPLATES_DIR, exist_ok=True)
        safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in name).strip()
        if not safe_name:
            return False, "Invalid template name."
        path = os.path.join(TEMPLATES_DIR, f"{safe_name}.json")
        # Strip file paths — templates describe structure, not file locations
        storable = {k: v for k, v in config.items()
                    if k not in ("src_file", "tgt_file", "src_sheet", "tgt_sheet")}
        with open(path, "w", encoding="utf-8") as f:
            json.dump(storable, f, indent=2)
        return True, f"Template '{safe_name}' saved."
    except Exception as e:
        return False, f"Save failed: {e}"


def load_pivot_template(name: str) -> tuple:
    """Load a named pivot template.  Returns (ok, config_dict | error_str)."""
    try:
        path = os.path.join(TEMPLATES_DIR, f"{name}.json")
        with open(path, "r", encoding="utf-8") as f:
            return True, json.load(f)
    except FileNotFoundError:
        return False, f"Template '{name}' not found."
    except Exception as e:
        return False, f"Load failed: {e}"


def delete_pivot_template(name: str) -> tuple:
    try:
        path = os.path.join(TEMPLATES_DIR, f"{name}.json")
        os.remove(path)
        return True, f"Template '{name}' deleted."
    except Exception as e:
        return False, f"Delete failed: {e}"


# ---------------------------------------------------------------------------
# Column detection helper
# ---------------------------------------------------------------------------

def detect_columns(file_path: str, sheet_name: str, header_row: int = 1) -> dict:
    """
    Read source sheet and classify each column.
    Returns {col_name: {is_numeric, n_unique, sample, dtype}}
    """
    try:
        df = wb_cache.read_excel(file_path, sheet_name=sheet_name, header=header_row - 1)
        # Flatten MultiIndex columns (merged header cells)
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [
                " ".join(str(p) for p in col if str(p) not in ("", "nan", "None")).strip()
                for col in df.columns
            ]
        df.columns = [str(c).strip() for c in df.columns]
        # Deduplicate
        seen: dict = {}
        new_cols = []
        for c in df.columns:
            if c in seen:
                seen[c] += 1
                new_cols.append(f"{c}.{seen[c]}")
            else:
                seen[c] = 0
                new_cols.append(c)
        df.columns = new_cols
        result = {}
        for col in df.columns:
            series = df[col].dropna()
            num_ratio = pd.to_numeric(series, errors="coerce").notna().sum() / max(len(series), 1)
            result[col] = {
                "is_numeric": num_ratio > 0.7,
                "n_unique":   int(df[col].nunique()),
                "n_rows":     len(df),
                "sample":     [str(v) for v in series.head(3).tolist()],
                "dtype":      str(df[col].dtype),
            }
        return result
    except Exception:
        return {}


# ---------------------------------------------------------------------------
# Core pivot executor
# ---------------------------------------------------------------------------

def run_pivot_step(
    src_file_path: str,
    src_sheet: str,
    src_header_row: int = 1,
    tgt_file_path: str = None,
    tgt_sheet: str = "Pivot",
    tgt_start_row: int = 1,
    tgt_start_col: str = "A",
    overwrite_sheet: bool = True,
    row_fields: list = None,        # index columns  → list[str]
    col_fields: list = None,        # column pivot   → list[str]
    value_fields: list = None,      # [{column, aggfunc, label}]
    filter_conditions: list = None,
    filter_combine: str = "AND",
    fill_value=0,
    grand_total_rows: bool = False,
    grand_total_cols: bool = False,
    sort_field: str = "",
    sort_ascending: bool = True,
    apply_formatting: bool = True,
    header_bg_color: str = "1F3864",
    header_font_color: str = "FFFFFF",
    total_bg_color: str = "D9E1F2",
) -> tuple:
    """
    Compute a pivot table and write results to tgt_sheet.
    Returns (success: bool, message: str).
    """
    try:
        # ── Load source ────────────────────────────────────────────────────
        df = wb_cache.read_excel(
            src_file_path, sheet_name=src_sheet, header=src_header_row - 1
        )

        # ── Flatten any MultiIndex columns (merged headers in Excel) ──────
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [
                " ".join(str(p) for p in col if str(p) not in ("", "nan", "None")).strip()
                for col in df.columns
            ]

        # Ensure all column names are plain strings
        df.columns = [str(c).strip() for c in df.columns]

        # Deduplicate column names (suffix .1, .2 … on repeats)
        seen: dict = {}
        new_cols = []
        for c in df.columns:
            if c in seen:
                seen[c] += 1
                new_cols.append(f"{c}.{seen[c]}")
            else:
                seen[c] = 0
                new_cols.append(c)
        df.columns = new_cols

        # Reset any MultiIndex on rows
        if isinstance(df.index, pd.MultiIndex):
            df = df.reset_index(drop=True)

        total_src_rows = len(df)

        # ── Pre-pivot filters ──────────────────────────────────────────────
        if filter_conditions:
            df = apply_filters(df, filter_conditions, combine=filter_combine)

        if df.empty:
            return False, "No rows remain after applying filters."

        # ── Normalise field lists ──────────────────────────────────────────
        row_fields   = [f for f in (row_fields   or []) if f and f in df.columns]
        col_fields   = [f for f in (col_fields   or []) if f and f in df.columns]
        value_fields = [v for v in (value_fields or []) if v.get("column") in df.columns]

        if not row_fields:
            return False, "At least one Row Field is required."
        if not value_fields:
            return False, "At least one Value Field is required."

        # ── Resolve overlap: value col also used as row/col field ─────────
        # pandas can't group by a column AND aggregate it simultaneously.
        # Fix: duplicate any conflicting value column under a temp name,
        # pivot using the temp name, then rename the output back.
        grouper_cols = set(row_fields) | set(col_fields)
        temp_renames: dict = {}   # temp_name → original user-facing column name

        for vf in value_fields:
            orig = vf["column"]
            if orig in grouper_cols:
                temp = f"__val_{orig}__"
                df[temp] = df[orig]
                vf["column"] = temp
                temp_renames[temp] = orig

        # ── Build aggfunc map ──────────────────────────────────────────────
        # Group by column; multiple aggfuncs for the same col → list
        agg_build: dict = {}
        for vf in value_fields:
            col = vf["column"]
            fn  = _AGGFUNC_MAP.get(vf.get("aggfunc", "sum"), "sum")
            if col not in agg_build:
                agg_build[col] = [fn]
            else:
                agg_build[col].append(fn)

        aggfunc_final = {
            col: funcs[0] if len(funcs) == 1 else funcs
            for col, funcs in agg_build.items()
        }
        values_list = list(aggfunc_final.keys())

        # ── Compute pivot ──────────────────────────────────────────────────
        margins = grand_total_rows or grand_total_cols
        try:
            pivot_df = pd.pivot_table(
                df,
                values=values_list,
                index=row_fields,
                columns=col_fields if col_fields else None,
                aggfunc=aggfunc_final,
                fill_value=fill_value,
                margins=margins,
                margins_name="Grand Total",
                observed=True,
            )
        except TypeError:
            # Older pandas may not support observed=
            pivot_df = pd.pivot_table(
                df,
                values=values_list,
                index=row_fields,
                columns=col_fields if col_fields else None,
                aggfunc=aggfunc_final,
                fill_value=fill_value,
                margins=margins,
                margins_name="Grand Total",
            )

        # Restore temp column names in value_fields (so label logic works)
        for vf in value_fields:
            if vf["column"] in temp_renames:
                vf["column"] = temp_renames[vf["column"]]

        # ── Flatten MultiIndex columns ────────────────────────────────────
        if isinstance(pivot_df.columns, pd.MultiIndex):
            pivot_df.columns = [
                " | ".join(str(p) for p in col_tuple if str(p)).strip(" | ")
                for col_tuple in pivot_df.columns
            ]
        else:
            pivot_df.columns = [str(c) for c in pivot_df.columns]

        # Rename temp placeholder columns back to original names in output
        if temp_renames:
            reverse = {f"__val_{orig}__": orig for orig in temp_renames.values()}
            # Also handle when the temp name is embedded in a "col | agg" flattened header
            rename_out = {}
            for c in pivot_df.columns:
                for tmp, orig in reverse.items():
                    if c == tmp or c.startswith(tmp + " |") or c.startswith(tmp + "|"):
                        rename_out[c] = c.replace(tmp, orig)
                        break
            if rename_out:
                pivot_df.rename(columns=rename_out, inplace=True)

        # ── Apply user-defined labels if single-aggfunc, no col_fields ──
        if not col_fields:
            label_map = {}
            for vf in value_fields:
                col   = vf["column"]   # already restored to original name
                label = (vf.get("label") or "").strip()
                agg   = vf.get("aggfunc", "sum")
                # How many aggfuncs were mapped to this column?
                # agg_build used temp names — check both temp and orig
                tmp_key = f"__val_{col}__" if col in grouper_cols else col
                n_agg = len(agg_build.get(tmp_key, agg_build.get(col, [None])))
                if n_agg > 1:
                    flat_name = f"{col} | {agg}"
                else:
                    flat_name = col
                if label and flat_name in pivot_df.columns:
                    label_map[flat_name] = label
            if label_map:
                pivot_df.rename(columns=label_map, inplace=True)

        # ── Flatten MultiIndex index ──────────────────────────────────────
        if isinstance(pivot_df.index, pd.MultiIndex):
            pivot_df.index = [
                " | ".join(str(v) for v in idx_tuple) for idx_tuple in pivot_df.index
            ]
        pivot_df.index = [str(v) for v in pivot_df.index]

        pivot_df = pivot_df.reset_index()
        # Rename generic "index" column back to row field(s) label
        if len(row_fields) == 1 and pivot_df.columns[0] == "index":
            pivot_df.rename(columns={"index": row_fields[0]}, inplace=True)

        # ── Sort ──────────────────────────────────────────────────────────
        if sort_field and sort_field in pivot_df.columns:
            non_total = pivot_df[pivot_df.iloc[:, 0] != "Grand Total"]
            total_row = pivot_df[pivot_df.iloc[:, 0] == "Grand Total"]
            non_total = non_total.sort_values(sort_field, ascending=sort_ascending)
            pivot_df  = pd.concat([non_total, total_row], ignore_index=True)

        # ── Write to workbook ─────────────────────────────────────────────
        tgt_path = tgt_file_path or src_file_path
        wb       = wb_cache.load(tgt_path)

        if overwrite_sheet and tgt_sheet in wb.sheetnames:
            pos = wb.sheetnames.index(tgt_sheet)
            del wb[tgt_sheet]
            ws  = wb.create_sheet(tgt_sheet, pos)
        elif tgt_sheet in wb.sheetnames:
            ws = wb[tgt_sheet]
        else:
            ws = wb.create_sheet(tgt_sheet)

        # Resolve start column
        sc = str(tgt_start_col or "A").strip()
        start_col = int(sc) if sc.isdigit() else column_index_from_string(sc.upper())
        start_row = max(1, int(tgt_start_row or 1))

        n_cols = len(pivot_df.columns)
        n_rows = len(pivot_df)

        # -- Header row --
        for c_off, col_name in enumerate(pivot_df.columns):
            cell = ws.cell(row=start_row, column=start_col + c_off, value=col_name)
            if apply_formatting:
                _style_header(cell, header_bg_color, header_font_color)

        # -- Data rows --
        for r_off, row_data in enumerate(pivot_df.itertuples(index=False), start=1):
            xl_row      = start_row + r_off
            is_total    = margins and r_off == n_rows  # last row is Grand Total

            for c_off, val in enumerate(row_data):
                # Convert numpy scalars → Python natives for openpyxl
                if isinstance(val, (np.integer,)):
                    val = int(val)
                elif isinstance(val, (np.floating,)):
                    val = float(val)
                elif isinstance(val, np.bool_):
                    val = bool(val)

                cell = ws.cell(row=xl_row, column=start_col + c_off, value=val)
                if apply_formatting and is_total:
                    _style_total(cell, total_bg_color)

        # -- Auto column width --
        if apply_formatting:
            for c_off, col_name in enumerate(pivot_df.columns):
                col_letter = get_column_letter(start_col + c_off)
                max_w = max(len(str(col_name)), 8)
                for r_off in range(n_rows):
                    v = pivot_df.iloc[r_off, c_off]
                    max_w = max(max_w, len(str(v)) if v is not None else 0)
                ws.column_dimensions[col_letter].width = min(max_w + 3, 45)

        wb_cache.save(wb, tgt_path)
        wb.close()

        filter_note = (
            f" ({total_src_rows - len(df)} row(s) excluded by filters)"
            if filter_conditions and total_src_rows != len(df)
            else ""
        )
        return True, (
            f"Pivot: {n_rows} row(s) × {n_cols} col(s) written to "
            f"'{tgt_sheet}' at {tgt_start_col}{tgt_start_row} "
            f"(source: {len(df)} row(s){filter_note})."
        )

    except Exception as e:
        return False, f"Pivot error: {e}\n{traceback.format_exc()}"


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------

def _style_header(cell, bg_hex: str, fg_hex: str):
    bg = _to_argb(bg_hex)
    fg = _to_argb(fg_hex)
    cell.font      = Font(bold=True, color=fg, size=11)
    cell.fill      = PatternFill(fill_type="solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    _thin_border(cell)


def _style_total(cell, bg_hex: str):
    bg = _to_argb(bg_hex)
    cell.font = Font(bold=True)
    cell.fill = PatternFill(fill_type="solid", fgColor=bg)


def _thin_border(cell):
    s = Side(style="thin")
    cell.border = Border(bottom=s)


def _to_argb(hex_color: str) -> str:
    h = str(hex_color or "").strip().lstrip("#").upper()
    if len(h) == 6:
        return "FF" + h
    if len(h) == 8:
        return h
    return "FF000000"
