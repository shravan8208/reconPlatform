"""
Folder Summary operation.

Three source modes:
  folder — multiple Excel files × one fixed sheet name  (original behaviour)
  sheets — one Excel file × multiple sheets
  both   — multiple Excel files × multiple sheets per file

In every mode the output is one row per (file, sheet) task written into a
consolidated target workbook.

Designed as a platform-generic operation — every config knob is exposed.
"""

import os
import re

import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

from core import wb_cache


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _col_idx(col_spec) -> int:
    """
    Convert a column spec to a 1-based integer index.
    Accepts:
      - A letter / multi-letter string: "A", "AB"
      - A 1-based integer string: "1", "3"
      - An integer: 3
    Returns 1 on any failure.
    """
    if col_spec is None:
        return 1
    if isinstance(col_spec, int):
        return max(1, col_spec)
    s = str(col_spec).strip()
    if not s:
        return 1
    if s.isdigit():
        return max(1, int(s))
    try:
        return column_index_from_string(s)
    except Exception:
        return 1


def _match(cell_val, search_text: str, match_type: str) -> bool:
    """
    Return True if cell_val matches search_text according to match_type.

    match_type values (case-insensitive):
      exact         – exact string equality (case-sensitive)
      iexact        – exact string equality (case-insensitive)
      contains      – substring match (case-sensitive)
      icontains     – substring match (case-insensitive)
      starts_with   – prefix match (case-insensitive)
      ends_with     – suffix match (case-insensitive)
      regex         – full Python re.search match
    """
    if cell_val is None:
        return False
    cell_str = str(cell_val).strip()
    mt = (match_type or "icontains").lower().replace("-", "_").replace(" ", "_")

    if mt == "exact":
        return cell_str == search_text
    if mt == "iexact":
        return cell_str.lower() == search_text.lower()
    if mt == "contains":
        return search_text in cell_str
    if mt == "icontains":
        return search_text.lower() in cell_str.lower()
    if mt in ("starts_with", "startswith"):
        return cell_str.lower().startswith(search_text.lower())
    if mt in ("ends_with", "endswith"):
        return cell_str.lower().endswith(search_text.lower())
    if mt == "regex":
        try:
            return bool(re.search(search_text, cell_str))
        except re.error:
            return False
    # Fallback: icontains
    return search_text.lower() in cell_str.lower()


def _extract_variable(
    df: pd.DataFrame,
    search_col_idx: int,   # 1-based; maps to df column by position
    value_col_idx: int,    # 1-based; maps to df column by position
    search_text: str,
    match_type: str,
    multi_match: str,       # first | last | sum | avg | max | min
    row_offset: int = 0,    # +N / -N rows relative to matched row
    header_row: int = 1,    # used for offset bounds only
):
    """
    Search df[search_col_idx-1] for rows matching search_text, then read df[value_col_idx-1].
    Handles row_offset for cases where the value is N rows below/above the label row.

    Returns (value, found: bool, note: str)
    """
    try:
        # Map 1-based column index to df positional column
        cols = list(df.columns)
        s_col_pos = search_col_idx - 1
        v_col_pos = value_col_idx - 1

        if s_col_pos < 0 or s_col_pos >= len(cols):
            return None, False, f"Search column index {search_col_idx} out of range (df has {len(cols)} cols)"
        if v_col_pos < 0 or v_col_pos >= len(cols):
            return None, False, f"Value column index {value_col_idx} out of range (df has {len(cols)} cols)"

        s_col = cols[s_col_pos]
        v_col = cols[v_col_pos]

        # Find matching rows
        matching_idxs = [
            i for i, val in enumerate(df[s_col])
            if _match(val, search_text, match_type)
        ]

        if not matching_idxs:
            return None, False, f"No match for '{search_text}' ({match_type})"

        # Apply row offset — shift to the value row
        if row_offset != 0:
            shifted = []
            for idx in matching_idxs:
                target_idx = idx + row_offset
                if 0 <= target_idx < len(df):
                    shifted.append(target_idx)
            if not shifted:
                return None, False, f"Row offset {row_offset:+d} put all matches out of bounds"
            matching_idxs = shifted

        # Extract values at the resolved rows
        raw_values = []
        for idx in matching_idxs:
            v = df[v_col].iloc[idx]
            # Coerce to numeric: handle plain floats, ints, and formatted strings
            # (e.g. Indian commas "1,23,456.78", parenthesised negatives "(500)")
            if v is None or (isinstance(v, float) and pd.isna(v)):
                v = None
            elif not isinstance(v, (int, float)):
                s = str(v).strip()
                if s == "" or s == "-":
                    v = None
                else:
                    # Strip thousands-separators and parenthetical negatives
                    s = s.replace(",", "")
                    negative = s.startswith("(") and s.endswith(")")
                    if negative:
                        s = s[1:-1]
                    try:
                        v = float(s)
                        if negative:
                            v = -v
                    except (ValueError, TypeError):
                        v = None   # un-parseable string → treat as missing
            if v is not None:
                raw_values.append(v)

        if not raw_values:
            return None, False, f"Matched rows have no value in value column"

        mm = (multi_match or "first").lower()
        if mm == "first":
            result = raw_values[0]
        elif mm == "last":
            result = raw_values[-1]
        elif mm == "sum":
            try:
                result = sum(float(v) for v in raw_values)
            except Exception:
                result = raw_values[0]
        elif mm == "avg":
            try:
                nums = [float(v) for v in raw_values]
                result = sum(nums) / len(nums)
            except Exception:
                result = raw_values[0]
        elif mm == "max":
            try:
                result = max(float(v) for v in raw_values)
            except Exception:
                result = raw_values[0]
        elif mm == "min":
            try:
                result = min(float(v) for v in raw_values)
            except Exception:
                result = raw_values[0]
        else:
            result = raw_values[0]

        return result, True, f"Found ({len(matching_idxs)} match(es), applied '{mm}')"

    except Exception as e:
        return None, False, f"Error extracting variable: {e}"


def _safe_eval(expression: str, variables: dict):
    """
    Evaluate a BODMAS expression with the given variable name→value mapping.
    Only numeric operations are allowed; string builtins are blocked.

    Returns (result, error_str_or_None)
    """
    if not expression or not expression.strip():
        return None, "Empty expression"
    try:
        safe_globals = {
            "__builtins__": {},
            "abs": abs, "round": round, "min": min, "max": max,
            "sum": sum, "int": int, "float": float,
        }
        # Replace None values with 0 so a single missing variable
        # doesn't kill the whole formula — arithmetic on None raises TypeError.
        clean_vars = {
            k: (float(v) if v is not None else 0.0)
            for k, v in variables.items()
        }
        result = eval(expression, safe_globals, clean_vars)
        return result, None
    except ZeroDivisionError:
        return None, "Division by zero"
    except NameError as e:
        return None, f"Unknown variable: {e}"
    except Exception as e:
        return None, str(e)


# ---------------------------------------------------------------------------
# Sheet-list helpers (for "sheets" and "both" modes)
# ---------------------------------------------------------------------------

def _get_sheet_names(file_path: str) -> list:
    """Return sheet names for any supported Excel format."""
    lower = str(file_path).lower()
    try:
        if lower.endswith(".xlsb"):
            try:
                import pyxlsb
            except ImportError:
                raise ImportError(
                    "pyxlsb is required to read .xlsb files. "
                    "Run:  pip install pyxlsb"
                )
            with pyxlsb.open_workbook(file_path) as wb:
                return list(wb.sheets)
        else:
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            names = list(wb.sheetnames)
            wb.close()
            return names
    except Exception:
        return []


def _filter_sheets(
    sheets: list,
    sheet_selection: str,   # all | contains | starts_with | ends_with | regex | list
    sheet_pattern: str = "",
    sheet_list: list = None,
    sheet_exclude: str = "",
) -> list:
    """
    Filter a list of sheet names according to the selection rules.
    sheet_exclude is always applied last; it's a comma-separated list of
    substrings — any sheet whose name contains any of them is removed.
    """
    sel = (sheet_selection or "all").lower().replace("-", "_").replace(" ", "_")
    pat = (sheet_pattern or "").strip()

    if sel == "all":
        result = list(sheets)
    elif sel == "list":
        wanted = [s.strip() for s in (sheet_list or []) if str(s).strip()]
        result = [s for s in sheets if s in wanted]
    elif sel == "contains":
        result = [s for s in sheets if pat.lower() in s.lower()]
    elif sel in ("starts_with", "startswith"):
        result = [s for s in sheets if s.lower().startswith(pat.lower())]
    elif sel in ("ends_with", "endswith"):
        result = [s for s in sheets if s.lower().endswith(pat.lower())]
    elif sel == "regex":
        try:
            result = [s for s in sheets if re.search(pat, s)]
        except re.error:
            result = list(sheets)
    else:
        result = list(sheets)

    # Apply exclusion list
    if sheet_exclude:
        excl = [e.strip().lower() for e in sheet_exclude.split(",") if e.strip()]
        result = [s for s in result if not any(e in s.lower() for e in excl)]

    return result


def _build_tasks(
    source_mode: str,
    src_file_paths: list,
    src_sheet: str,
    sheet_selection: str,
    sheet_pattern: str,
    sheet_list: list,
    sheet_exclude: str,
) -> list:
    """
    Build a list of (file_path, sheet_name, identifier_label) tuples that the
    main loop will process — one output row per tuple.

    source_mode:
      folder — many files × one fixed sheet
      sheets — one file × many sheets
      both   — many files × many sheets per file
    """
    tasks = []
    mode = (source_mode or "folder").lower()

    if mode == "folder":
        sheet = src_sheet or ""
        for fp in (src_file_paths or []):
            tasks.append((fp, sheet, os.path.basename(fp)))

    elif mode == "sheets":
        fp = ((src_file_paths or [""])[0]) or ""
        if fp:
            all_sheets = _get_sheet_names(fp)
            selected = _filter_sheets(all_sheets, sheet_selection, sheet_pattern,
                                      sheet_list, sheet_exclude)
            for sh in selected:
                tasks.append((fp, sh, sh))

    elif mode == "both":
        for fp in (src_file_paths or []):
            all_sheets = _get_sheet_names(fp)
            selected = _filter_sheets(all_sheets, sheet_selection, sheet_pattern,
                                      sheet_list, sheet_exclude)
            basename = os.path.basename(fp)
            for sh in selected:
                label = f"{basename} — {sh}"
                tasks.append((fp, sh, label))

    else:
        # Unknown mode — fall back to folder behaviour
        sheet = src_sheet or ""
        for fp in (src_file_paths or []):
            tasks.append((fp, sheet, os.path.basename(fp)))

    return tasks


# ---------------------------------------------------------------------------
# Main function
# ---------------------------------------------------------------------------

def run_folder_summary_step(
    src_file_paths: list,          # list of source file absolute paths
    src_sheet: str = "",           # fixed sheet name (used in "folder" mode)
    header_row: int = 1,           # header row in source sheet (1-based)

    # ── Source mode ────────────────────────────────────────────────────────
    # "folder" : many files × one fixed sheet  (default / backward-compat)
    # "sheets" : one file  × many sheets
    # "both"   : many files × many sheets per file
    source_mode: str = "folder",

    # Sheet selection (used in "sheets" and "both" modes)
    sheet_selection: str = "all",   # all | contains | starts_with | ends_with | regex | list
    sheet_pattern: str = "",        # pattern string for contains/starts_with/etc.
    sheet_list: list = None,        # explicit list of sheet names (for "list" selection)
    sheet_exclude: str = "",        # comma-sep substrings — matching sheets are excluded

    # ── Variables ──────────────────────────────────────────────────────────
    variables: list = None,
    # Each variable dict:
    # { "name": "A", "label": "Gross Sub", "search_text": "Gross Subscription",
    #   "search_col": "B", "value_col": "D",
    #   "match_type": "icontains",   # exact|iexact|contains|icontains|starts_with|ends_with|regex
    #   "multi_match": "first",      # first|last|sum|avg|max|min
    #   "row_offset": 0, "on_missing": "zero", "missing_default": 0, "include": True }

    # ── Formulas ───────────────────────────────────────────────────────────
    formulas: list = None,
    # Each formula dict:
    # { "name": "Net Amount", "expression": "A - B + C", "on_error": "flag" }

    # ── Target ─────────────────────────────────────────────────────────────
    tgt_file: str = "",
    tgt_sheet: str = "Summary",
    tgt_header_row: int = 1,

    filename_col: str = "File Name",  # identifier column header; auto-labelled by mode
    sheet_name_col: str = "",          # optional separate column for the sheet name (empty = disabled)
    append_mode: bool = True,

    include_var_cols: bool = True,
    include_status_col: bool = True,
    status_col_name: str = "Status",

    on_missing_default=None,
) -> tuple:
    """
    Process each (file, sheet) task, extract variables, evaluate formulas,
    write one consolidated row per task into the target workbook.

    Returns (success: bool, message: str)
    """
    variables = variables or []
    formulas = formulas or []

    if not src_file_paths:
        return False, "No source files provided"
    if not tgt_file:
        return False, "No target file specified"
    if str(tgt_file).lower().endswith(".xlsb"):
        return False, "Cannot write to .xlsb target — use a .xlsx file"

    # ── Build task list ────────────────────────────────────────────────────
    tasks = _build_tasks(
        source_mode, src_file_paths, src_sheet,
        sheet_selection, sheet_pattern, sheet_list or [], sheet_exclude,
    )
    if not tasks:
        return False, (
            "No tasks found. "
            + ("No sheets matched the selection criteria." if source_mode != "folder"
               else f"No source files resolved to sheet '{src_sheet}'.")
        )

    rows_written = 0
    errors = []

    # ── Load or create target workbook ─────────────────────────────────────
    try:
        if os.path.exists(tgt_file):
            wb_tgt = wb_cache.load(tgt_file)
        else:
            wb_tgt = openpyxl.Workbook()
            wb_tgt.active.title = tgt_sheet or "Summary"
    except Exception as e:
        return False, f"Cannot open target file: {e}"

    # Get or create target sheet
    if tgt_sheet and tgt_sheet in wb_tgt.sheetnames:
        ws_tgt = wb_tgt[tgt_sheet]
    elif wb_tgt.sheetnames:
        if tgt_sheet:
            ws_tgt = wb_tgt.create_sheet(tgt_sheet)
        else:
            ws_tgt = wb_tgt.active
    else:
        ws_tgt = wb_tgt.active
        if tgt_sheet:
            ws_tgt.title = tgt_sheet

    # ── Build / verify target header row ──────────────────────────────────
    out_headers = []
    if filename_col:
        out_headers.append(filename_col)
    if sheet_name_col:
        out_headers.append(sheet_name_col)
    if include_var_cols:
        for v in variables:
            if v.get("include", True):
                out_headers.append(v.get("label") or v.get("name") or "Var")
    for f in formulas:
        out_headers.append(f.get("name") or "Formula")
    if include_status_col:
        out_headers.append(status_col_name or "Status")

    header_r = int(tgt_header_row or 1)
    existing_headers = [
        ws_tgt.cell(row=header_r, column=c).value
        for c in range(1, ws_tgt.max_column + 1)
    ]
    existing_headers_clean = [str(h).strip() if h is not None else "" for h in existing_headers]

    def _tgt_col_for(header_name: str) -> int:
        name = (header_name or "").strip()
        if name in existing_headers_clean:
            return existing_headers_clean.index(name) + 1
        new_col = len(existing_headers_clean) + 1
        existing_headers_clean.append(name)
        ws_tgt.cell(row=header_r, column=new_col).value = name
        return new_col

    col_map = {h: _tgt_col_for(h) for h in out_headers}

    # ── Determine next write row ───────────────────────────────────────────
    if append_mode:
        last_row = header_r
        for row in ws_tgt.iter_rows(min_row=header_r + 1, max_col=ws_tgt.max_column):
            if any(cell.value is not None for cell in row):
                last_row = row[0].row
        next_write_row = last_row + 1
    else:
        for row in ws_tgt.iter_rows(min_row=header_r + 1, max_row=ws_tgt.max_row):
            for cell in row:
                cell.value = None
        next_write_row = header_r + 1

    # ── Process each task ─────────────────────────────────────────────────
    for (src_path, sheet_name, identifier) in tasks:
        row_notes = []
        var_values = {}
        found_flags = {}

        try:
            df = wb_cache.read_excel(src_path, sheet_name=sheet_name, header=header_row - 1)
            df.columns = [str(c).strip() for c in df.columns]
        except Exception as e:
            errors.append(f"{identifier}: cannot read sheet '{sheet_name}': {e}")
            write_row = next_write_row
            if filename_col and filename_col in col_map:
                ws_tgt.cell(row=write_row, column=col_map[filename_col]).value = identifier
            if sheet_name_col and sheet_name_col in col_map:
                ws_tgt.cell(row=write_row, column=col_map[sheet_name_col]).value = sheet_name
            if include_status_col and status_col_name in col_map:
                ws_tgt.cell(row=write_row, column=col_map[status_col_name]).value = \
                    f"❌ Read error: {e}"
            next_write_row += 1
            continue

        # Extract variables
        for v in variables:
            vname = v.get("name") or "V"
            search_text = v.get("search_text") or ""
            search_col = _col_idx(v.get("search_col") or "A")
            value_col = _col_idx(v.get("value_col") or "B")
            match_type = v.get("match_type") or "icontains"
            multi_match = v.get("multi_match") or "first"
            row_offset = int(v.get("row_offset") or 0)
            on_missing = (v.get("on_missing") or "zero").lower()

            val, found, note = _extract_variable(
                df, search_col, value_col,
                search_text, match_type, multi_match, row_offset, header_row,
            )
            found_flags[vname] = found

            if not found:
                row_notes.append(f"{vname}: {note}")
                if on_missing == "zero":
                    val = 0
                elif on_missing in ("blank", "flag"):
                    val = None
                elif on_missing == "error":
                    errors.append(f"{identifier}: required variable '{vname}' not found — {note}")
                    val = None
                elif on_missing == "default_value":
                    val = v.get("missing_default")
                    if val is None:
                        val = on_missing_default
                else:
                    val = 0

            var_values[vname] = val

        # Evaluate formulas
        formula_results = {}
        for f in formulas:
            fname = f.get("name") or "Formula"
            expression = f.get("expression") or ""
            on_error = (f.get("on_error") or "flag").lower()

            result, err_msg = _safe_eval(expression, var_values)
            if err_msg:
                row_notes.append(f"{fname}: {err_msg}")
                if on_error == "zero":
                    result = 0
                elif on_error == "blank":
                    result = None
                else:
                    result = f"⚠️ {err_msg}"
            formula_results[fname] = result

        # Status
        missing_vars = [n for n, f in found_flags.items() if not f]
        status_text = (f"⚠️ Missing: {', '.join(missing_vars)}" if missing_vars else "✅ OK")
        if row_notes:
            status_text += " | " + "; ".join(row_notes)

        # Write row
        write_row = next_write_row

        if filename_col and filename_col in col_map:
            ws_tgt.cell(row=write_row, column=col_map[filename_col]).value = identifier

        if sheet_name_col and sheet_name_col in col_map:
            ws_tgt.cell(row=write_row, column=col_map[sheet_name_col]).value = sheet_name

        if include_var_cols:
            for v in variables:
                if not v.get("include", True):
                    continue
                vname = v.get("name") or "V"
                vlabel = v.get("label") or vname
                if vlabel in col_map:
                    ws_tgt.cell(row=write_row, column=col_map[vlabel]).value = var_values.get(vname)

        for f in formulas:
            fname = f.get("name") or "Formula"
            if fname in col_map:
                ws_tgt.cell(row=write_row, column=col_map[fname]).value = formula_results.get(fname)

        if include_status_col and status_col_name in col_map:
            ws_tgt.cell(row=write_row, column=col_map[status_col_name]).value = status_text

        next_write_row += 1
        rows_written += 1

    # ── Save ──────────────────────────────────────────────────────────────
    try:
        wb_cache.save(wb_tgt, tgt_file)
    except Exception as e:
        return False, f"Wrote {rows_written} rows but failed to save target: {e}"

    mode_label = {"folder": "files", "sheets": "sheets", "both": "file-sheet pairs"}.get(
        (source_mode or "folder").lower(), "items"
    )
    if errors:
        return rows_written > 0, (
            f"Folder Summary: {rows_written}/{len(tasks)} {mode_label} OK, "
            f"{len(errors)} error(s): "
            + "; ".join(errors[:3])
            + (" ..." if len(errors) > 3 else "")
        )
    return True, (
        f"Folder Summary: {rows_written} row(s) written "
        f"({mode_label}) → '{tgt_sheet}'"
    )
