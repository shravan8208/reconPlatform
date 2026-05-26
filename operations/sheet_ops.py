"""
Sheet-level operations: delete, rename, reorder sheets within a workbook.
Pure openpyxl implementation — NO win32com.
"""

from __future__ import annotations

import re
from typing import List, Optional, Tuple

import openpyxl

from core import wb_cache


# ---------------------------------------------------------------------------
# Delete Sheets
# ---------------------------------------------------------------------------

def run_delete_sheets_step(
    file_path: str,
    mode: str,
    sheet_names: Optional[List[str]] = None,
    pattern: str = "",
    pattern_type: str = "contains",
    case_sensitive: bool = False,
    protect_last: bool = True,
) -> Tuple[bool, str]:
    """
    Delete one or more sheets from a workbook.

    Parameters
    ----------
    file_path    : path to the Excel workbook
    mode         : one of
                     "named"      – delete the sheets listed in sheet_names
                     "keep_only"  – keep only the sheets listed in sheet_names,
                                    delete everything else
                     "pattern"    – delete sheets whose names match pattern
                     "empty"      – delete sheets that contain no data
    sheet_names  : list of sheet names (used by "named" and "keep_only" modes)
    pattern      : text for pattern matching (used by "pattern" mode)
    pattern_type : "contains" | "starts_with" | "ends_with" | "exact" | "regex"
                   (used by "pattern" mode)
    case_sensitive: whether pattern matching is case-sensitive (default False)
    protect_last : if True, refuse to delete the last remaining sheet
                   (Excel itself blocks this — openpyxl will error otherwise)

    Returns
    -------
    (success: bool, message: str)
    """
    try:
        wb = wb_cache.load(file_path)
        all_sheets = wb.sheetnames[:]          # snapshot before any deletion

        if not all_sheets:
            return False, "Workbook has no sheets."

        # ----------------------------------------------------------------
        # Determine which sheets to delete
        # ----------------------------------------------------------------
        to_delete: List[str] = []

        if mode == "named":
            requested = sheet_names or []
            for s in requested:
                if s in all_sheets:
                    to_delete.append(s)
                # silently skip sheets that no longer exist

        elif mode == "keep_only":
            keep = set(sheet_names or [])
            to_delete = [s for s in all_sheets if s not in keep]

        elif mode == "pattern":
            if not pattern:
                return False, "Pattern is required for 'pattern' mode."
            to_delete = _sheets_matching_pattern(
                all_sheets, pattern, pattern_type, case_sensitive
            )

        elif mode == "empty":
            to_delete = _find_empty_sheets(wb, all_sheets)

        else:
            return False, f"Unknown mode: '{mode}'. Use named / keep_only / pattern / empty."

        if not to_delete:
            return True, "No sheets matched the deletion criteria — nothing deleted."

        # ----------------------------------------------------------------
        # Safety: protect the last remaining sheet
        # ----------------------------------------------------------------
        surviving = [s for s in all_sheets if s not in to_delete]
        if protect_last and len(surviving) == 0:
            # Would delete every sheet — refuse
            return False, (
                f"Deletion would remove ALL {len(all_sheets)} sheet(s). "
                "A workbook must keep at least one sheet. "
                "Use 'keep_only' mode and select at least one sheet to retain."
            )

        # ----------------------------------------------------------------
        # Delete
        # ----------------------------------------------------------------
        deleted: List[str] = []
        skipped: List[str] = []

        for sheet_name in to_delete:
            if sheet_name not in wb.sheetnames:
                skipped.append(sheet_name)
                continue
            del wb[sheet_name]
            deleted.append(sheet_name)

        wb_cache.save(wb, file_path)
        wb.close()

        msg_parts = [f"Deleted {len(deleted)} sheet(s): {', '.join(deleted)}"]
        if skipped:
            msg_parts.append(f"Skipped (not found): {', '.join(skipped)}")
        msg_parts.append(f"Remaining sheets: {', '.join(wb_cache.load(file_path).sheetnames)}")
        return True, " | ".join(msg_parts)

    except Exception as exc:
        return False, f"Error deleting sheets: {exc}"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _sheets_matching_pattern(
    sheets: List[str],
    pattern: str,
    pattern_type: str,
    case_sensitive: bool,
) -> List[str]:
    """Return sheet names that match the given pattern."""
    matched = []
    for name in sheets:
        cmp_name = name if case_sensitive else name.lower()
        cmp_pat  = pattern if case_sensitive else pattern.lower()

        if pattern_type == "contains":
            hit = cmp_pat in cmp_name
        elif pattern_type == "starts_with":
            hit = cmp_name.startswith(cmp_pat)
        elif pattern_type == "ends_with":
            hit = cmp_name.endswith(cmp_pat)
        elif pattern_type == "exact":
            hit = cmp_name == cmp_pat
        elif pattern_type == "regex":
            flags = 0 if case_sensitive else re.IGNORECASE
            try:
                hit = bool(re.search(pattern, name, flags))
            except re.error:
                hit = False
        else:
            hit = False

        if hit:
            matched.append(name)
    return matched


def _find_empty_sheets(wb: openpyxl.Workbook, sheet_names: List[str]) -> List[str]:
    """
    Return sheet names that are completely empty
    (max_row is None or 0, or every cell in the used range is None/blank).
    """
    empty = []
    for name in sheet_names:
        ws = wb[name]
        if ws.max_row is None or ws.max_row == 0:
            empty.append(name)
            continue
        # Check if every cell in the used range is blank
        has_data = False
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None and str(cell.value).strip() != "":
                    has_data = True
                    break
            if has_data:
                break
        if not has_data:
            empty.append(name)
    return empty


# ---------------------------------------------------------------------------
# Utility: list all sheet names from a file (used by UI)
# ---------------------------------------------------------------------------

def get_sheet_names(file_path: str) -> List[str]:
    """Return all sheet names from an Excel file (uses cache when active)."""
    try:
        wb = wb_cache.load(file_path, data_only=True)
        names = wb.sheetnames[:]
        wb.close()
        return names
    except Exception:
        return []
