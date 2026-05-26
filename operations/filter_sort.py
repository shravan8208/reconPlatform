"""
Filter and Sort operations using pandas.
Pure Python/pandas implementation - NO win32com!
"""

import pandas as pd
from copy import copy
from openpyxl.utils.dataframe import dataframe_to_rows
from core import wb_cache
import openpyxl
from openpyxl.utils import column_index_from_string


def _copy_cell(src, dst):
    """Copy cell value + all styling from src to dst."""
    dst.value = src.value
    if src.has_style:
        dst._style = copy(src._style)
    dst.number_format = src.number_format
    dst.font = copy(src.font)
    dst.fill = copy(src.fill)
    dst.border = copy(src.border)
    dst.alignment = copy(src.alignment)
    dst.protection = copy(src.protection)
    dst.comment = src.comment


def run_filter_step(
    file_path,
    sheet_name,
    col_letter,
    filter_value=None,
    start_row=2,
    end_row='last',
    remove_empty=False,
    header_row=1,
):
    """
    Filter rows based on column value OR remove rows where a column is empty.
    
    FAST implementation: copies only matching rows to a new sheet (O(n)) instead of
    deleting rows one-by-one (O(n²)). Preserves formulas and formatting.
    
    Args:
        file_path: Path to Excel/CSV file
        sheet_name: Sheet name (Excel only)
        col_letter: Column letter to filter on (A, B, C...)
        filter_value: Value to filter for (ignored when remove_empty=True)
        start_row: Start row (1-indexed) [currently informational; not used for filtering]
        end_row: End row or 'last'      [currently informational; not used for filtering]
        remove_empty: If True, remove entire rows where selected column is blank/NaN
        header_row: Header row number (1-indexed) for Excel
    
    Returns:
        Tuple of (success: bool, message: str)
    """
    try:
        lower = str(file_path).lower()
        
        # ========================
        # CSV / TXT: use pandas
        # ========================
        if lower.endswith(".csv") or lower.endswith(".txt"):
            df = pd.read_csv(file_path)
            col_idx = column_index_from_string(col_letter) - 1
            col_name = df.columns[col_idx]
            original_rows = len(df)
            
            if remove_empty:
                ser = df[col_name]
                empty_mask = ser.isna() | (ser.astype(str).str.strip() == "")
                df_filtered = df[~empty_mask].copy()
                action_msg = f"Removed {int(empty_mask.sum())} row(s) where {col_name} is empty"
            else:
                df_filtered = df[df[col_name] == filter_value].copy()
                action_msg = f"Filtered {original_rows} rows to {len(df_filtered)} rows where {col_name}={filter_value}"
            
            df_filtered.to_csv(file_path, index=False)
            return True, action_msg
        
        # ========================
        # Excel: FAST in-place filter (copy matching rows to temp sheet)
        # ========================
        # Load workbook with formulas (for copying)
        wb = wb_cache.load(file_path, data_only=False)
        if sheet_name not in wb.sheetnames:
            wb.close()
            return False, f"Sheet not found: {sheet_name}"

        # Also load with data_only=True to get CALCULATED VALUES for filtering
        # This is needed to filter on columns that contain formulas
        wb_data = wb_cache.load(file_path, data_only=True)
        ws_data = wb_data[sheet_name] if sheet_name in wb_data.sheetnames else wb_data.active
        
        ws_old = wb[sheet_name]
        col_idx_1 = column_index_from_string(col_letter)
        hdr_row = int(header_row)
        first_data_row = hdr_row + 1
        last_row = ws_old.max_row
        max_col = ws_old.max_column
        
        # Decide which rows to KEEP (using calculated values for comparison)
        rows_to_keep: list[int] = []
        removed_count = 0
        
        for r in range(first_data_row, last_row + 1):
            # Get the CALCULATED value (not formula text) for filtering
            v = ws_data.cell(row=r, column=col_idx_1).value
            
            # If no cached value, fall back to formula cell value
            if v is None:
                v = ws_old.cell(row=r, column=col_idx_1).value
            
            if remove_empty:
                is_empty = (v is None) or (str(v).strip() == "")
                if is_empty:
                    removed_count += 1
                else:
                    rows_to_keep.append(r)
            else:
                # Compare with filter_value (handle type differences)
                if v == filter_value or str(v) == str(filter_value):
                    rows_to_keep.append(r)
                else:
                    removed_count += 1
        
        wb_data.close()
        
        # Create a new sheet with a temp name
        temp_name = f"__filter_temp_{sheet_name}"[:31]
        if temp_name in wb.sheetnames:
            del wb[temp_name]
        ws_new = wb.create_sheet(temp_name)
        
        # Copy column widths
        for col_letter_key, dim in ws_old.column_dimensions.items():
            ws_new.column_dimensions[col_letter_key].width = dim.width
        
        # Copy header rows (1 to header_row)
        for r in range(1, hdr_row + 1):
            for c in range(1, max_col + 1):
                _copy_cell(ws_old.cell(row=r, column=c), ws_new.cell(row=r, column=c))
        
        # Copy matching data rows
        dest_row = hdr_row + 1
        for src_row in rows_to_keep:
            for c in range(1, max_col + 1):
                _copy_cell(ws_old.cell(row=src_row, column=c), ws_new.cell(row=dest_row, column=c))
            dest_row += 1
        
        # Copy merged cell ranges (skip data rows that were removed)
        for merge_range in list(ws_old.merged_cells.ranges):
            # Only copy merges fully within header rows or fully within kept rows
            # For simplicity, copy all merges that overlap with header rows
            if merge_range.min_row <= hdr_row:
                ws_new.merge_cells(str(merge_range))
        
        # Remember sheet position
        old_idx = wb.sheetnames.index(sheet_name)
        
        # Delete old sheet, rename new sheet
        del wb[sheet_name]
        ws_new.title = sheet_name
        
        # Move sheet back to original position
        wb.move_sheet(ws_new, offset=old_idx - wb.sheetnames.index(sheet_name))
        
        wb_cache.save(wb, file_path)
        wb.close()

        kept_count = len(rows_to_keep)
        if remove_empty:
            return True, f"Removed {removed_count} row(s) where column {col_letter} is empty; kept {kept_count} row(s)"
        return True, f"Filtered by column {col_letter}={filter_value}; removed {removed_count} row(s), kept {kept_count}"
        
    except Exception as e:
        return False, f"Error in filter: {str(e)}"


def run_sort_step(file_path, sheet_name, cols, ascending=True, not_null_cols=None, header_row=1):
    """
    Sort DataFrame by columns.
    
    Args:
        file_path: Path to Excel file
        sheet_name: Sheet name
        cols: List of column names to sort by
        ascending: True for ascending, False for descending (can be list for each column)
        not_null_cols: Dict of columns that must not be null (priority sorting)
        header_row: Header row number (1-indexed)
    
    Returns:
        Tuple of (success: bool, message: str)
    """
    try:
        lower = str(file_path).lower()
        if lower.endswith(".csv") or lower.endswith(".txt"):
            # CSV: pandas sort (no formulas to preserve)
            df = pd.read_csv(file_path)
            if isinstance(ascending, bool):
                ascending = [ascending] * len(cols)
            df_sorted = df.sort_values(by=cols, ascending=ascending, na_position='last')
            df_sorted.to_csv(file_path, index=False)
            return True, f"Sorted by {', '.join(cols)}"

        # Excel: sort rows in-place to preserve formulas & formatting
        wb = wb_cache.load(file_path, data_only=False)
        if sheet_name not in wb.sheetnames:
            wb.close()
            return False, f"Sheet not found: {sheet_name}"
        ws = wb[sheet_name]

        hdr_row = int(header_row)
        first_data_row = hdr_row + 1
        last_row = ws.max_row
        max_col = ws.max_column

        # Build header name -> column index mapping from the header row
        header_map = {}
        for c in range(1, max_col + 1):
            v = ws.cell(row=hdr_row, column=c).value
            if v is None:
                continue
            s = str(v).strip()
            if not s:
                continue
            if s not in header_map:
                header_map[s] = c

        sort_col_idxs: list[int] = []
        for name in cols:
            if name not in header_map:
                wb.close()
                return False, f"Sort column not found in header row {hdr_row}: {name}"
            sort_col_idxs.append(header_map[name])

        if isinstance(ascending, bool):
            ascending_list = [ascending] * len(sort_col_idxs)
        else:
            ascending_list = list(ascending)
            if len(ascending_list) != len(sort_col_idxs):
                wb.close()
                return False, "Ascending list length must match number of sort columns"

        def norm(v):
            # keep types when possible; fall back to string to avoid TypeError across mixed types
            if v is None:
                return ""
            return v

        def cell_value(r, c):
            return ws.cell(row=r, column=c).value

        rows = list(range(first_data_row, last_row + 1))

        # Stable multi-key sort: sort by last key first, then earlier keys (tie-break behavior)
        for c_idx, asc in reversed(list(zip(sort_col_idxs, ascending_list))):
            def key_fn(r):
                v = cell_value(r, c_idx)
                if v is None:
                    return None
                if isinstance(v, str) and not v.strip():
                    return None
                return norm(v)

            non_none = []
            none_rows = []
            for r in rows:
                v = key_fn(r)
                (none_rows if v is None else non_none).append((r, v))

            try:
                non_none_sorted = sorted(non_none, key=lambda t: t[1], reverse=not bool(asc))
            except TypeError:
                non_none_sorted = sorted(non_none, key=lambda t: str(t[1]), reverse=not bool(asc))

            rows = [t[0] for t in non_none_sorted] + [t[0] for t in none_rows]

        # Snapshot the row content (values + styles) before rewriting
        def copy_cell(src, dst):
            dst.value = src.value
            if src.has_style:
                dst._style = copy(src._style)
            dst.number_format = src.number_format
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.protection = copy(src.protection)
            dst.comment = src.comment

        snapshot = {}
        for r in rows:
            snapshot[r] = [ws.cell(row=r, column=c) for c in range(1, max_col + 1)]

        dest_r = first_data_row
        for src_r in rows:
            src_cells = snapshot[src_r]
            for c in range(1, max_col + 1):
                src_cell = src_cells[c - 1]
                dst_cell = ws.cell(row=dest_r, column=c)
                copy_cell(src_cell, dst_cell)
            dest_r += 1

        wb_cache.save(wb, file_path)
        wb.close()
        return True, f"Sorted by {', '.join(cols)}"
        
    except Exception as e:
        return False, f"Error in sort: {str(e)}"


