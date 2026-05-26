"""
Split Export Engine.

Slice a source sheet and export to:
  single   → one filtered sheet in one file
  sheets   → one file, N sheets  (split by column values)
  files    → N files, one sheet each
  files_sheets → N files, M sheets each  (two-level split)

Features
--------
- Pre-export row filtering (reuses filter_utils)
- Column include / exclude
- Multi-column sort
- Name templates:  {ColumnName}  {value}  {date}  {seq}
- Row-count limit per sheet (auto-splits oversized groups)
- Optional Summary sheet / sheet index
- Overwrite or append mode
- Full header formatting
"""

import os
import re
import copy
import traceback
from datetime import date as _today

import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from core import wb_cache
from operations.filter_utils import apply_filters


# ---------------------------------------------------------------------------
# Mode catalogue
# ---------------------------------------------------------------------------

SPLIT_MODES = {
    "single":       "Single Export  (filter → one sheet → one file)",
    "sheets":       "One File  →  Multiple Sheets  (split by column)",
    "files":        "Multiple Files  →  One Sheet each  (split by column)",
    "files_sheets": "Multiple Files  →  Multiple Sheets each  (two-level split)",
}

SPLIT_MODE_KEYS   = list(SPLIT_MODES.keys())
SPLIT_MODE_LABELS = list(SPLIT_MODES.values())


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _sanitize_sheet(name: str) -> str:
    """Make a valid Excel sheet name (max 31 chars, no special chars)."""
    s = re.sub(r"[\\/*?:\[\]]", "_", str(name or "Sheet")).strip().strip("'")
    return (s[:31].strip()) or "Sheet"


def _sanitize_filename(name: str) -> str:
    """Remove characters illegal in filenames."""
    s = re.sub(r'[\\/*?:"<>|]', "_", str(name or "output")).strip()
    return s or "output"


def _ensure_xlsx(name: str) -> str:
    lower = name.lower()
    if not (lower.endswith(".xlsx") or lower.endswith(".xls")):
        return name + ".xlsx"
    return name


def _unique_sheet_name(wb: openpyxl.Workbook, base: str) -> str:
    """Return base if not taken, otherwise base_2, base_3, …"""
    existing = {s.lower() for s in wb.sheetnames}
    name = _sanitize_sheet(base)
    if name.lower() not in existing:
        return name
    for i in range(2, 9999):
        candidate = _sanitize_sheet(f"{base[:28]}_{i}")
        if candidate.lower() not in existing:
            return candidate
    return _sanitize_sheet(f"{base[:27]}_{id(base)}")


def _render_name(template: str, group_vals: dict, seq: int = 1) -> str:
    """
    Resolve a name template.
    group_vals = {col_name: value, …}  (from current split group)
    Special tokens: {date}, {seq}/{n}, {value} (alias for first split col value)
    """
    result = str(template or "{value}")
    for k, v in group_vals.items():
        result = result.replace(f"{{{k}}}", _sanitize_filename(str(v)))
    # fallback {value} → first group value
    if "{value}" in result and group_vals:
        result = result.replace("{value}", _sanitize_filename(str(next(iter(group_vals.values())))))
    result = result.replace("{date}", _today.today().strftime("%Y-%m-%d"))
    result = result.replace("{seq}",  str(seq))
    result = result.replace("{n}",    str(seq))
    return result


def _apply_column_selection(df: pd.DataFrame,
                             include: list, exclude: list) -> pd.DataFrame:
    cols = list(df.columns)
    if include:
        cols = [c for c in include if c in df.columns]
    if exclude:
        cols = [c for c in cols if c not in set(exclude)]
    return df[cols] if cols else df


def _apply_sort(df: pd.DataFrame, sort_cols: list,
                sort_ascending) -> pd.DataFrame:
    if not sort_cols:
        return df
    valid = [c for c in sort_cols if c in df.columns]
    if not valid:
        return df
    if isinstance(sort_ascending, list):
        asc = [sort_ascending[i] if i < len(sort_ascending) else True
               for i in range(len(valid))]
    else:
        asc = bool(sort_ascending)
    try:
        return df.sort_values(valid, ascending=asc).reset_index(drop=True)
    except Exception:
        return df


def _write_df_to_sheet(ws, df: pd.DataFrame,
                        include_header: bool = True,
                        apply_formatting: bool = True,
                        header_bg: str = "1F3864",
                        header_fg: str = "FFFFFF",
                        start_row: int = 1,
                        freeze_header: bool = True):
    """Write DataFrame to openpyxl worksheet starting at start_row."""
    cur_row = start_row

    if include_header:
        bg_argb = "FF" + header_bg.lstrip("#").upper().zfill(6)
        fg_argb = "FF" + header_fg.lstrip("#").upper().zfill(6)
        thin    = Side(style="thin")
        for c_idx, col_name in enumerate(df.columns, start=1):
            cell = ws.cell(row=cur_row, column=c_idx, value=str(col_name))
            if apply_formatting:
                cell.font      = Font(bold=True, color=fg_argb, size=11)
                cell.fill      = PatternFill(fill_type="solid", fgColor=bg_argb)
                cell.alignment = Alignment(horizontal="center",
                                           vertical="center", wrap_text=True)
                cell.border    = Border(bottom=thin)
        if apply_formatting and freeze_header:
            ws.freeze_panes = ws.cell(row=cur_row + 1, column=1)
        cur_row += 1

    for row_data in df.itertuples(index=False):
        for c_idx, val in enumerate(row_data, start=1):
            # Coerce numpy scalars → Python natives
            if isinstance(val, (np.integer,)):
                val = int(val)
            elif isinstance(val, (np.floating,)):
                val = None if np.isnan(val) else float(val)
            elif isinstance(val, np.bool_):
                val = bool(val)
            elif val is pd.NaT:
                val = None
            ws.cell(row=cur_row, column=c_idx, value=val)
        cur_row += 1

    # Auto column widths
    if apply_formatting and len(df) > 0:
        for c_idx, col_name in enumerate(df.columns, start=1):
            col_letter = get_column_letter(c_idx)
            sample_rows = min(len(df), 200)
            max_w = max(len(str(col_name)), 8)
            for r in range(sample_rows):
                v = df.iloc[r, c_idx - 1]
                max_w = max(max_w, len(str(v)) if v is not None else 0)
            ws.column_dimensions[col_letter].width = min(max_w + 3, 55)

    return cur_row  # next available row


def _write_summary_sheet(wb: openpyxl.Workbook,
                          summary_data: list,   # [{sheet, group, rows, …}]
                          sheet_name: str = "_Summary",
                          apply_formatting: bool = True,
                          header_bg: str = "1F3864",
                          header_fg: str = "FFFFFF"):
    """Add a summary index sheet to wb."""
    name = _unique_sheet_name(wb, sheet_name)
    ws = wb.create_sheet(name, 0)   # insert at front

    if not summary_data:
        return

    headers = list(summary_data[0].keys())
    _write_df_to_sheet(ws, pd.DataFrame(summary_data, columns=headers),
                       include_header=True,
                       apply_formatting=apply_formatting,
                       header_bg=header_bg, header_fg=header_fg)


def _open_or_create_wb(path: str, overwrite: bool = False) -> openpyxl.Workbook:
    """
    Always load an existing workbook when the file is present.
    'overwrite' is handled at the sheet level by callers — it does NOT
    mean wipe the whole file, only that individual sheets may be replaced.
    """
    if os.path.exists(path):
        try:
            return openpyxl.load_workbook(path)
        except Exception:
            pass
    return openpyxl.Workbook()


# ---------------------------------------------------------------------------
# Core
# ---------------------------------------------------------------------------

def run_split_export_step(
    src_file_path: str,
    src_sheet: str,
    src_header_row: int   = 1,
    # Mode
    split_mode: str       = "single",   # single | sheets | files | files_sheets
    split_col: str        = "",         # primary split column
    split_col2: str       = "",         # secondary (sheet-level) for files_sheets
    # Output paths
    output_folder: str    = "",         # used for files / files_sheets modes
    tgt_file: str         = "",         # used for single / sheets modes
    tgt_sheet: str        = "Export",   # used for single mode only
    # Name templates
    file_name_template: str  = "{value}.xlsx",
    sheet_name_template: str = "{value}",
    sheet_name_template2: str= "{value}",   # sheet-level for files_sheets
    # Data selection
    filter_conditions: list  = None,
    filter_combine: str      = "AND",
    include_columns: list    = None,
    exclude_columns: list    = None,
    sort_cols: list          = None,
    sort_ascending           = True,
    # Options
    include_header: bool     = True,
    overwrite: bool          = True,
    max_rows_per_sheet: int  = 0,       # 0 = unlimited; >0 splits group into chunks
    add_summary_sheet: bool  = False,
    summary_sheet_name: str  = "_Summary",
    # Formatting
    apply_formatting: bool   = True,
    header_bg_color: str     = "1F3864",
    header_font_color: str   = "FFFFFF",
) -> tuple:
    """
    Main entry point.  Returns (success: bool, message: str).
    """
    try:
        # ── Load source ────────────────────────────────────────────────────
        df = wb_cache.read_excel(
            src_file_path, sheet_name=src_sheet, header=src_header_row - 1
        )

        # Flatten MultiIndex columns
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [
                " ".join(str(p) for p in c if str(p) not in ("", "nan", "None")).strip()
                for c in df.columns
            ]
        df.columns = [str(c).strip() for c in df.columns]

        # Reset MultiIndex rows
        if isinstance(df.index, pd.MultiIndex):
            df = df.reset_index(drop=True)

        total_src = len(df)

        # ── Pre-export filters ─────────────────────────────────────────────
        if filter_conditions:
            df = apply_filters(df, filter_conditions, combine=filter_combine)

        if df.empty:
            return False, "No rows remain after applying filters."

        # ── Column selection ───────────────────────────────────────────────
        df = _apply_column_selection(df, include_columns or [], exclude_columns or [])

        # ── Sort ───────────────────────────────────────────────────────────
        df = _apply_sort(df, sort_cols or [], sort_ascending)

        filtered_rows = len(df)

        # ── Dispatch to mode handler ───────────────────────────────────────
        if split_mode == "single":
            return _mode_single(
                df, tgt_file, tgt_sheet, overwrite,
                include_header, apply_formatting,
                header_bg_color, header_font_color,
                total_src, filtered_rows,
            )
        elif split_mode == "sheets":
            return _mode_sheets(
                df, split_col, sheet_name_template,
                tgt_file, overwrite, max_rows_per_sheet,
                include_header, apply_formatting,
                header_bg_color, header_font_color,
                add_summary_sheet, summary_sheet_name,
                total_src, filtered_rows,
            )
        elif split_mode == "files":
            return _mode_files(
                df, split_col, file_name_template, sheet_name_template,
                output_folder, src_file_path, overwrite, max_rows_per_sheet,
                include_header, apply_formatting,
                header_bg_color, header_font_color,
                add_summary_sheet, summary_sheet_name,
                total_src, filtered_rows,
            )
        elif split_mode == "files_sheets":
            return _mode_files_sheets(
                df, split_col, split_col2,
                file_name_template, sheet_name_template, sheet_name_template2,
                output_folder, src_file_path, overwrite, max_rows_per_sheet,
                include_header, apply_formatting,
                header_bg_color, header_font_color,
                add_summary_sheet, summary_sheet_name,
                total_src, filtered_rows,
            )
        else:
            return False, f"Unknown split mode: {split_mode}"

    except Exception as e:
        return False, f"Split Export error: {e}\n{traceback.format_exc()}"


# ---------------------------------------------------------------------------
# Mode handlers
# ---------------------------------------------------------------------------

def _mode_single(df, tgt_file, tgt_sheet, overwrite,
                  include_header, apply_formatting, hbg, hfg,
                  total_src, filtered_rows):
    """Filter → one sheet → one file."""
    if not tgt_file:
        return False, "Target file is required for Single Export mode."
    wb = _open_or_create_wb(tgt_file)
    sheet_name = _sanitize_sheet(tgt_sheet or "Export")

    if overwrite and sheet_name in wb.sheetnames:
        pos = wb.sheetnames.index(sheet_name)
        del wb[sheet_name]
        ws = wb.create_sheet(sheet_name, pos)
    elif sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(sheet_name)

    # Remove default empty sheet if present (brand-new workbook only)
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        del wb["Sheet"]

    _write_df_to_sheet(ws, df, include_header=include_header,
                        apply_formatting=apply_formatting,
                        header_bg=hbg, header_fg=hfg)
    wb.save(tgt_file)
    wb.close()
    skipped = total_src - filtered_rows
    return True, (
        f"Single Export: {filtered_rows} row(s) → '{sheet_name}' in "
        f"{os.path.basename(tgt_file)}"
        + (f"  ({skipped} row(s) excluded by filters)" if skipped else "")
    )


def _mode_sheets(df, split_col, sheet_tmpl, tgt_file, overwrite,
                  max_rows, include_header, apply_formatting, hbg, hfg,
                  add_summary, summary_name, total_src, filtered_rows):
    """Split → one file, N sheets.  Existing sheets NOT being replaced are preserved."""
    if not split_col or split_col not in df.columns:
        return False, f"Split column '{split_col}' not found in data."
    if not tgt_file:
        return False, "Target file is required for Sheets mode."

    # Always load existing workbook — never wipe the whole file.
    wb = _open_or_create_wb(tgt_file)
    # Remove the default blank "Sheet" only when it's the sole sheet in a brand-new workbook.
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) == 1:
        del wb["Sheet"]

    groups     = df.groupby(split_col, sort=False, dropna=False)
    summary    = []
    sheet_seq  = 0
    total_written = 0

    for grp_val, grp_df in groups:
        sheet_seq += 1
        group_vals = {split_col: grp_val, "value": grp_val}

        chunks = _chunk_df(grp_df, max_rows)
        for chunk_idx, chunk in enumerate(chunks, start=1):
            raw_name = _render_name(sheet_tmpl, group_vals, seq=sheet_seq)
            if len(chunks) > 1:
                raw_name = f"{raw_name}_{chunk_idx}"
            sheet_name = _sanitize_sheet(raw_name)

            if sheet_name in wb.sheetnames:
                if overwrite:
                    # Replace this sheet only — all others are untouched.
                    pos = wb.sheetnames.index(sheet_name)
                    del wb[sheet_name]
                    ws = wb.create_sheet(sheet_name, pos)
                else:
                    # Don't overwrite — give it a unique name instead.
                    sheet_name = _unique_sheet_name(wb, sheet_name)
                    ws = wb.create_sheet(sheet_name)
            else:
                ws = wb.create_sheet(sheet_name)

            _write_df_to_sheet(ws, chunk, include_header=include_header,
                                apply_formatting=apply_formatting,
                                header_bg=hbg, header_fg=hfg)
            total_written += len(chunk)
            summary.append({
                "Sheet Name":  sheet_name,
                "Group Value": str(grp_val),
                "Rows":        len(chunk),
            })

    if add_summary and summary:
        _write_summary_sheet(wb, summary, summary_name,
                              apply_formatting=apply_formatting,
                              header_bg=hbg, header_fg=hfg)

    wb.save(tgt_file)
    wb.close()

    n_sheets = len(summary)
    return True, (
        f"Sheets Export: {n_sheets} sheet(s), {total_written} row(s) → "
        f"{os.path.basename(tgt_file)}"
        + (f"\n  Groups: {', '.join(str(s['Group Value']) for s in summary[:8])}"
           + ("…" if n_sheets > 8 else ""))
    )


def _mode_files(df, split_col, file_tmpl, sheet_tmpl,
                 output_folder, src_path, overwrite, max_rows,
                 include_header, apply_formatting, hbg, hfg,
                 add_summary, summary_name, total_src, filtered_rows):
    """Split → N files, one sheet each."""
    if not split_col or split_col not in df.columns:
        return False, f"Split column '{split_col}' not found in data."

    out_dir = output_folder.strip() or os.path.dirname(src_path)
    os.makedirs(out_dir, exist_ok=True)

    groups       = df.groupby(split_col, sort=False, dropna=False)
    files_written = []
    total_written  = 0
    file_seq = 0
    errors   = []

    for grp_val, grp_df in groups:
        file_seq += 1
        group_vals = {split_col: grp_val, "value": grp_val}
        raw_fname  = _render_name(file_tmpl, group_vals, seq=file_seq)
        fname      = _ensure_xlsx(_sanitize_filename(raw_fname))
        fpath      = os.path.join(out_dir, fname)

        try:
            chunks     = _chunk_df(grp_df, max_rows)
            wb         = _open_or_create_wb(fpath, overwrite)
            if "Sheet" in wb.sheetnames and overwrite:
                del wb["Sheet"]

            for chunk_idx, chunk in enumerate(chunks, start=1):
                raw_sname  = _render_name(sheet_tmpl, group_vals, seq=file_seq)
                sname      = _unique_sheet_name(wb, raw_sname)
                ws         = wb.create_sheet(sname)
                _write_df_to_sheet(ws, chunk, include_header=include_header,
                                    apply_formatting=apply_formatting,
                                    header_bg=hbg, header_fg=hfg)
                total_written += len(chunk)

            if add_summary:
                summary_data = [{
                    "Sheet": _sanitize_sheet(_render_name(sheet_tmpl, group_vals, file_seq)),
                    "Group": str(grp_val),
                    "Rows":  len(grp_df),
                }]
                _write_summary_sheet(wb, summary_data, summary_name,
                                      apply_formatting=apply_formatting,
                                      header_bg=hbg, header_fg=hfg)

            wb.save(fpath)
            wb.close()
            files_written.append(fname)

        except Exception as e:
            errors.append(f"{fname}: {e}")

    n_files = len(files_written)
    msg = (
        f"Files Export: {n_files} file(s), {total_written} row(s) → {out_dir}"
        + (f"\n  Files: {', '.join(files_written[:8])}"
           + ("…" if n_files > 8 else ""))
    )
    if errors:
        msg += f"\n  Errors ({len(errors)}): {'; '.join(errors[:3])}"
    return len(files_written) > 0, msg


def _mode_files_sheets(df, split_col, split_col2,
                        file_tmpl, sheet_tmpl, sheet_tmpl2,
                        output_folder, src_path, overwrite, max_rows,
                        include_header, apply_formatting, hbg, hfg,
                        add_summary, summary_name, total_src, filtered_rows):
    """Two-level split → N files, M sheets each."""
    if not split_col or split_col not in df.columns:
        return False, f"File-level split column '{split_col}' not found."
    if not split_col2 or split_col2 not in df.columns:
        return False, f"Sheet-level split column '{split_col2}' not found."

    out_dir = output_folder.strip() or os.path.dirname(src_path)
    os.makedirs(out_dir, exist_ok=True)

    file_groups   = df.groupby(split_col, sort=False, dropna=False)
    files_written  = []
    total_written  = 0
    file_seq = 0
    errors   = []

    for file_val, file_df in file_groups:
        file_seq += 1
        file_gvals = {split_col: file_val, "value": file_val}
        raw_fname  = _render_name(file_tmpl, file_gvals, seq=file_seq)
        fname      = _ensure_xlsx(_sanitize_filename(raw_fname))
        fpath      = os.path.join(out_dir, fname)

        try:
            wb = _open_or_create_wb(fpath, overwrite)
            if "Sheet" in wb.sheetnames and overwrite:
                del wb["Sheet"]

            sheet_groups = file_df.groupby(split_col2, sort=False, dropna=False)
            summary_data = []
            sheet_seq    = 0

            for sheet_val, sheet_df in sheet_groups:
                sheet_seq += 1
                sheet_gvals = {
                    split_col:  file_val,
                    split_col2: sheet_val,
                    "value":    sheet_val,
                }
                chunks = _chunk_df(sheet_df, max_rows)
                for chunk_idx, chunk in enumerate(chunks, start=1):
                    raw_sname = _render_name(sheet_tmpl2, sheet_gvals, seq=sheet_seq)
                    if len(chunks) > 1:
                        raw_sname = f"{raw_sname}_{chunk_idx}"
                    sname = _unique_sheet_name(wb, raw_sname)
                    ws    = wb.create_sheet(sname)
                    _write_df_to_sheet(ws, chunk, include_header=include_header,
                                        apply_formatting=apply_formatting,
                                        header_bg=hbg, header_fg=hfg)
                    total_written += len(chunk)
                    summary_data.append({
                        "Sheet": sname,
                        "Group": f"{file_val} / {sheet_val}",
                        "Rows":  len(chunk),
                    })

            if add_summary and summary_data:
                _write_summary_sheet(wb, summary_data, summary_name,
                                      apply_formatting=apply_formatting,
                                      header_bg=hbg, header_fg=hfg)

            wb.save(fpath)
            wb.close()
            files_written.append(fname)

        except Exception as e:
            errors.append(f"{fname}: {e}")

    n_files = len(files_written)
    msg = (
        f"Files+Sheets Export: {n_files} file(s), {total_written} row(s) → {out_dir}"
        + (f"\n  Files: {', '.join(files_written[:8])}"
           + ("…" if n_files > 8 else ""))
    )
    if errors:
        msg += f"\n  Errors ({len(errors)}): {'; '.join(errors[:3])}"
    return len(files_written) > 0, msg


# ---------------------------------------------------------------------------
# Util
# ---------------------------------------------------------------------------

def _chunk_df(df: pd.DataFrame, max_rows: int):
    """Split df into chunks of max_rows.  Returns [df] when max_rows == 0."""
    if not max_rows or max_rows <= 0 or len(df) <= max_rows:
        return [df]
    return [df.iloc[i: i + max_rows] for i in range(0, len(df), max_rows)]
