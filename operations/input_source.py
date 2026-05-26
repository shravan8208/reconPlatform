"""
Input Source / File Collector operation.
Resolves files from various sources (single path, folder scans, date-based rules)
and registers them under named classes so downstream steps can reference them by name.
"""

import os
import re
from datetime import date, timedelta
from pathlib import Path
from typing import Optional

import openpyxl
import pandas as pd

# ---------------------------------------------------------------------------
# Token resolution
# ---------------------------------------------------------------------------

_MONTH_NAMES = [
    "", "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]
_MONTH_SHORT = [
    "", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
]


def _resolve_tokens(pattern: str, ref: Optional[date] = None) -> str:
    """Replace date/time tokens in a filename pattern with actual values."""
    if not pattern:
        return pattern
    today = ref or date.today()
    first_of_month = today.replace(day=1)
    prev_month_last = first_of_month - timedelta(days=1)
    prev_month_first = prev_month_last.replace(day=1)

    replacements = {
        "{current_month}": _MONTH_SHORT[today.month],
        "{current_month_full}": _MONTH_NAMES[today.month],
        "{current_month_num}": f"{today.month:02d}",
        "{previous_month}": _MONTH_SHORT[prev_month_last.month],
        "{previous_month_full}": _MONTH_NAMES[prev_month_last.month],
        "{previous_month_num}": f"{prev_month_last.month:02d}",
        "{current_year}": str(today.year),
        "{prev_year}": str(today.year - 1),
        "{today}": today.strftime("%Y-%m-%d"),
        "{today_ddmmyyyy}": today.strftime("%d%m%Y"),
        "{today_mmddyyyy}": today.strftime("%m%d%Y"),
        "{yesterday}": (today - timedelta(days=1)).strftime("%Y-%m-%d"),
    }
    result = pattern
    for token, value in replacements.items():
        result = result.replace(token, value)
    return result


# ---------------------------------------------------------------------------
# File scanning helpers
# ---------------------------------------------------------------------------

_EXCEL_EXTS = {".xlsx", ".xls", ".xlsm", ".xlsb"}
_CSV_EXTS   = {".csv", ".txt"}

# Display value used in the UI for "pick any Excel format"
ALL_EXCEL_LABEL = "All Excel (.xlsx, .xlsb, .xls, .xlsm)"


def _matches_extension(filename: str, ext_filter: str) -> bool:
    """
    Returns True if filename matches ext_filter.

    ext_filter can be:
      "*"                                    → any file
      ".xlsx"                                → single extension
      ".xlsx,.xlsb,.xls"                     → comma-separated list
      "All Excel (.xlsx, .xlsb, .xls, .xlsm)"  → all Excel formats (UI label)
      "excel"                                → shorthand for all Excel formats
    """
    if not ext_filter or ext_filter.strip() == "*":
        return True

    _, ext = os.path.splitext(filename.lower())
    flt = ext_filter.strip().lower()

    # "All Excel" shorthand (matches the UI selectbox label or bare keyword)
    if flt.startswith("all excel") or flt == "excel":
        return ext in _EXCEL_EXTS

    # Comma-separated list of extensions
    parts = [p.strip() for p in ext_filter.split(",") if p.strip()]
    wanted = set()
    for p in parts:
        p = p.lower()
        if p.startswith("all excel") or p == "excel":
            wanted |= _EXCEL_EXTS
        elif p == "*":
            return True
        else:
            wanted.add(p if p.startswith(".") else "." + p)

    return ext in wanted


def _scan_folder_mode_b(folder: str, extension: str = ALL_EXCEL_LABEL) -> list[str]:
    """Return all files in folder matching the extension."""
    folder_path = Path(folder)
    if not folder_path.is_dir():
        return []
    matches = []
    for f in sorted(folder_path.iterdir()):
        if f.is_file() and _matches_extension(f.name, extension):
            matches.append(str(f))
    return matches


def _scan_folder_mode_c(
    folder: str,
    contains: str = "",
    starts_with: str = "",
    ends_with: str = "",
    extension: str = ALL_EXCEL_LABEL,
    excludes: str = "",
) -> list[str]:
    """Return files matching include rules, excluding blacklisted substrings."""
    folder_path = Path(folder)
    if not folder_path.is_dir():
        return []

    exclude_terms = [e.strip().lower() for e in (excludes or "").split(",") if e.strip()]
    matches = []

    for f in sorted(folder_path.iterdir()):
        if not f.is_file():
            continue
        name = f.stem  # without extension
        name_lower = name.lower()
        full_lower = f.name.lower()

        if not _matches_extension(f.name, extension):
            continue
        if contains and contains.lower() not in full_lower:
            continue
        if starts_with and not full_lower.startswith(starts_with.lower()):
            continue
        if ends_with:
            stem_lower = os.path.splitext(full_lower)[0]
            if not stem_lower.endswith(ends_with.lower().lstrip("_").lstrip("-")):
                # Also check the full name without ext
                if not name_lower.endswith(ends_with.lower()):
                    continue
        if any(ex in full_lower for ex in exclude_terms):
            continue

        matches.append(str(f))

    return matches


def _scan_folder_mode_d(
    folder: str,
    date_pattern: str,
    date_pick: str = "first",
    extension: str = ALL_EXCEL_LABEL,
) -> list[str]:
    """
    Scan folder with a tokenized filename pattern.
    date_pick: 'first' = first match alphabetically, 'latest_modified' = newest mtime
    """
    resolved_pattern = _resolve_tokens(date_pattern).lower()
    folder_path = Path(folder)
    if not folder_path.is_dir():
        return []

    matches = []
    for f in folder_path.iterdir():
        if not f.is_file():
            continue
        if not _matches_extension(f.name, extension):
            continue
        if resolved_pattern in f.name.lower():
            matches.append(f)

    if not matches:
        return []

    if date_pick == "latest_modified":
        matches.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    else:
        matches.sort(key=lambda f: f.name.lower())

    return [str(f) for f in matches]


# ---------------------------------------------------------------------------
# Sheet selection
# ---------------------------------------------------------------------------

def _get_sheet_names(file_path: str) -> list[str]:
    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".xlsb":
            # openpyxl cannot read .xlsb — use pyxlsb instead
            try:
                import pyxlsb
                with pyxlsb.open_workbook(file_path) as wb:
                    return list(wb.sheets)
            except ImportError:
                # pyxlsb not installed — fall back to pandas which auto-selects engine
                df = pd.read_excel(file_path, engine="pyxlsb", nrows=0)
                return [str(df.columns.name or "Sheet1")]
        if ext in (".xlsx", ".xlsm", ".xls"):
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            names = list(wb.sheetnames)
            wb.close()
            return names
        if ext in (".csv", ".txt"):
            return ["Sheet1"]
    except Exception:
        pass
    return []


def _sheet_has_headers(file_path: str, sheet_name: str, required_headers: list[str]) -> bool:
    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext in (".xlsx", ".xlsm", ".xlsb", ".xls"):
            df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=1)
        else:
            df = pd.read_csv(file_path, nrows=1)
        cols_lower = {c.strip().lower() for c in df.columns}
        return all(h.strip().lower() in cols_lower for h in required_headers)
    except Exception:
        return False


def _month_score(sheet_name: str) -> int:
    """Return month number (1-12) if sheet name contains a month, else 0."""
    s = sheet_name.lower()
    for i, m in enumerate(_MONTH_SHORT[1:], 1):
        if m.lower() in s:
            return i
    for i, m in enumerate(_MONTH_NAMES[1:], 1):
        if m.lower() in s:
            return i
    return 0


def resolve_sheet(file_path: str, sheet_cfg: dict) -> tuple[str, str]:
    """
    Resolve which sheet to use given sheet_cfg.
    Returns (sheet_name, status_msg).
    sheet_cfg keys:
      sheet_mode: first | exact | contains | latest_month | detect_headers
      sheet_exact: str
      sheet_contains: str
      sheet_headers: comma-sep header names
    """
    mode = (sheet_cfg.get("sheet_mode") or "first").lower()
    sheets = _get_sheet_names(file_path)

    if not sheets:
        return "", "Could not read sheet list"

    if mode == "first":
        return sheets[0], f"First sheet: '{sheets[0]}'"

    if mode == "exact":
        name = (sheet_cfg.get("sheet_exact") or "").strip()
        if name in sheets:
            return name, f"Exact match: '{name}'"
        return "", f"Sheet '{name}' not found. Available: {sheets}"

    if mode == "contains":
        term = (sheet_cfg.get("sheet_contains") or "").strip().lower()
        matched = [s for s in sheets if term in s.lower()]
        if matched:
            return matched[0], f"Contains '{term}': '{matched[0]}'"
        return "", f"No sheet contains '{term}'. Available: {sheets}"

    if mode == "latest_month":
        today = date.today()
        # Score sheets by how close their month is to current month
        scored = [(s, _month_score(s)) for s in sheets]
        scored = [(s, sc) for s, sc in scored if sc > 0]
        if not scored:
            return sheets[0], f"No month sheet found, using first: '{sheets[0]}'"
        # Prefer current month, then previous
        cur = today.month
        def closeness(sc):
            diff = (cur - sc) % 12
            return diff
        scored.sort(key=lambda x: closeness(x[1]))
        return scored[0][0], f"Latest month sheet: '{scored[0][0]}'"

    if mode == "detect_headers":
        headers_raw = sheet_cfg.get("sheet_headers") or ""
        required = [h.strip() for h in headers_raw.split(",") if h.strip()]
        if not required:
            return sheets[0], "No headers specified, using first sheet"
        for s in sheets:
            if _sheet_has_headers(file_path, s, required):
                return s, f"Detected by headers in '{s}'"
        return "", f"No sheet has all headers: {required}"

    return sheets[0], f"Unknown sheet mode '{mode}', using first: '{sheets[0]}'"


# ---------------------------------------------------------------------------
# File class resolution
# ---------------------------------------------------------------------------

def resolve_file_class(class_cfg: dict) -> dict:
    """
    Resolve a single file class config to:
      {class_name, mode, file_path, sheet_name, status, error}
    """
    class_name = (class_cfg.get("class_name") or "").strip()
    mode = (class_cfg.get("mode") or "A").upper()
    extension = class_cfg.get("rule_extension") or class_cfg.get("extension") or ALL_EXCEL_LABEL

    file_path = ""
    candidates = []
    error = ""

    if mode == "A":
        fp = (class_cfg.get("file_path") or "").strip()
        if not fp:
            error = "No file path provided"
        elif not os.path.exists(fp):
            error = f"File not found: {fp}"
        else:
            file_path = fp

    elif mode == "B":
        folder = (class_cfg.get("folder_path") or "").strip()
        if not folder:
            error = "No folder path provided"
        elif not os.path.isdir(folder):
            error = f"Folder not found: {folder}"
        else:
            candidates = _scan_folder_mode_b(folder, extension)
            if not candidates:
                error = f"No {extension} files found in {folder}"
            else:
                file_path = candidates[0]

    elif mode == "C":
        folder = (class_cfg.get("folder_path") or "").strip()
        if not folder:
            error = "No folder path provided"
        elif not os.path.isdir(folder):
            error = f"Folder not found: {folder}"
        else:
            candidates = _scan_folder_mode_c(
                folder,
                contains=class_cfg.get("rule_contains", ""),
                starts_with=class_cfg.get("rule_starts_with", ""),
                ends_with=class_cfg.get("rule_ends_with", ""),
                extension=extension,
                excludes=class_cfg.get("rule_excludes", ""),
            )
            if not candidates:
                error = "No files matched the rules"
            else:
                file_path = candidates[0]

    elif mode == "D":
        folder = (class_cfg.get("folder_path") or "").strip()
        pattern = (class_cfg.get("date_pattern") or "").strip()
        if not folder:
            error = "No folder path provided"
        elif not os.path.isdir(folder):
            error = f"Folder not found: {folder}"
        elif not pattern:
            error = "No date pattern provided"
        else:
            date_pick = class_cfg.get("date_pick", "first")
            candidates = _scan_folder_mode_d(folder, pattern, date_pick, extension)
            if not candidates:
                resolved = _resolve_tokens(pattern)
                error = f"No files found matching '{resolved}' in {folder}"
            else:
                file_path = candidates[0]

    else:
        error = f"Unknown mode: {mode}"

    # Resolve sheet
    sheet_name = ""
    sheet_msg = ""
    if file_path:
        sheet_name, sheet_msg = resolve_sheet(file_path, class_cfg)
        if not sheet_name:
            error = sheet_msg

    status = "found" if file_path and sheet_name else "missing"

    return {
        "class_name": class_name,
        "mode": mode,
        "file_path": file_path,
        "sheet_name": sheet_name,
        "status": status,
        "error": error,
        "sheet_msg": sheet_msg,
        "candidates": candidates,
        "resolved_pattern": _resolve_tokens(class_cfg.get("date_pattern", "")) if mode == "D" else "",
    }


# ---------------------------------------------------------------------------
# Step runner (called by executor)
# ---------------------------------------------------------------------------

def run_input_source_step(file_classes: list[dict]) -> tuple[bool, str]:
    """
    Resolve all file classes and register them in session state.

    For single-file classes (Mode A, or folder modes that matched exactly 1 file):
      uploaded_file_paths[name] = "/path/to/file.xlsx"

    For multi-file classes (folder modes with N > 1 files):
      uploaded_file_paths[name]  = "/path/to/first_file.xlsx"  (compat for single-file steps)
      uploaded_file_groups[name] = ["/path/file1.xlsx", "/path/file2.xlsx", ...]

    normalize_import reads uploaded_file_groups when available so it processes every file.

    Must run in main thread (modifies st.session_state).
    """
    import streamlit as st

    if not isinstance(st.session_state.get("uploaded_file_paths"), dict):
        st.session_state.uploaded_file_paths = {}
    if not isinstance(st.session_state.get("uploaded_file_groups"), dict):
        st.session_state.uploaded_file_groups = {}

    registered = []
    errors = []

    for cls_cfg in file_classes:
        name = cls_cfg.get("class_name") or f"File_{len(registered)+1}"

        # Mode E — file already registered by the upload panel before execution
        if cls_cfg.get("mode") == "E":
            if name in st.session_state.uploaded_file_paths:
                registered.append(f"{name} (user-uploaded)")
            else:
                errors.append(f"{name}: file not uploaded yet")
            continue

        result = resolve_file_class(cls_cfg)
        name = result["class_name"] or name

        if result["status"] == "found":
            # Always register the first (representative) file for single-file step compat
            st.session_state.uploaded_file_paths[name] = result["file_path"]

            # For folder-mode classes register the full list
            all_candidates = result.get("candidates") or []
            if len(all_candidates) > 1:
                st.session_state.uploaded_file_groups[name] = all_candidates
            elif name in st.session_state.uploaded_file_groups:
                # Clean up stale group if the class is now single-file
                del st.session_state.uploaded_file_groups[name]

            registered.append(name)
        else:
            errors.append(f"{name}: {result['error']}")

    if errors:
        msg = f"Registered {len(registered)}, errors: {'; '.join(errors)}"
        return len(registered) > 0, msg

    return True, f"Registered {len(registered)} file class(es): {', '.join(registered)}"
