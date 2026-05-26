"""
Import, Append, and Header Mapping operations.
Pure Python/pandas implementation - NO win32com!
"""

import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from core import wb_cache
from operations.filter_utils import apply_filters


def run_import_step(source_path, target_path, src_sheet, tgt_sheet, mapping_lines,
                    start_row=None, src_header_row=1, tgt_header_row=1, append_mode=False,
                    filters=None, filter_combine="AND",
                    filename_col=None, filename_value=None):
    """
    Import data with field mappings from source to target.
    
    Args:
        source_path: Source file path
        target_path: Target file path
        src_sheet: Source sheet name
        tgt_sheet: Target sheet name
        mapping_lines: Mapping string "SourceCol->TargetCol" (one per line) OR list of tuples [(src, tgt), ...]
        start_row: Starting row in target (None = replace or append based on append_mode)
        src_header_row: Source header row (1-indexed)
        tgt_header_row: Target header row (1-indexed)
        append_mode: If True, append imported data after existing rows. If False (default), replace target data.
    
    Returns:
        Tuple of (success: bool, message: str)
    """
    try:
        # .xlsb is read-only — no Python library can write back to this format
        if str(target_path).lower().endswith('.xlsb'):
            return False, (
                "Cannot write to a .xlsb target file — .xlsb is a read-only format. "
                "Please use a .xlsx file as your target. "
                "Tip: register a separate .xlsx working file and import into that instead."
            )

        # Read source
        df_source = wb_cache.read_excel(source_path, sheet_name=src_sheet, header=src_header_row-1)

        # Read target (needed for append mode or to preserve structure)
        df_target = wb_cache.read_excel(target_path, sheet_name=tgt_sheet, header=tgt_header_row-1)
        
        # Strip whitespace from source column names (common issue with messy Excel files)
        df_source.columns = [str(c).strip() for c in df_source.columns]
        df_target.columns = [str(c).strip() for c in df_target.columns]

        # Apply source filters before mapping
        pre_filter_count = len(df_source)
        if filters:
            df_source = apply_filters(df_source, filters, filter_combine or "AND")
            filtered_count = pre_filter_count - len(df_source)
        else:
            filtered_count = 0

        # Build a lookup dict: stripped_name -> original_name for source columns
        src_col_lookup = {str(c).strip(): c for c in df_source.columns}
        
        # Parse mappings (supports both string format and list of tuples)
        mappings = {}
        if isinstance(mapping_lines, list):
            # List of tuples: [(src_col, tgt_col), ...]
            for item in mapping_lines:
                if isinstance(item, (list, tuple)) and len(item) >= 2:
                    src_col = str(item[0]).strip()
                    tgt_col = str(item[1]).strip()
                    if src_col and tgt_col:
                        mappings[src_col] = tgt_col
        else:
            # String format: "SourceCol->TargetCol\n..."
            for line in (mapping_lines or "").strip().split('\n'):
                if '->' in line:
                    parts = line.split('->', 1)
                    src_col = parts[0].strip()
                    tgt_col = parts[1].strip()
                    if src_col and tgt_col:
                        mappings[src_col] = tgt_col
        
        # Create mapped DataFrame (matching with stripped names)
        # This will have source data mapped to target column names
        df_mapped = pd.DataFrame()
        matched_cols = 0
        for src_col, tgt_col in mappings.items():
            # Try exact match first, then stripped match
            if src_col in df_source.columns:
                df_mapped[tgt_col] = df_source[src_col].reset_index(drop=True)
                matched_cols += 1
            elif src_col in src_col_lookup:
                actual_col = src_col_lookup[src_col]
                df_mapped[tgt_col] = df_source[actual_col].reset_index(drop=True)
                matched_cols += 1
        
        if matched_cols == 0:
            return False, f"No columns matched. Source columns: {list(df_source.columns)[:10]}..."

        # Inject source filename column if requested
        if filename_col and filename_value:
            df_mapped[str(filename_col)] = filename_value

        # PRESERVE ALL TARGET COLUMNS: Start with target structure, update only mapped columns
        # This ensures unmapped headers are NOT deleted
        target_cols_original = list(df_target.columns)
        
        # Determine result based on mode
        if append_mode:
            # Append: keep existing target data + add new rows (preserving all columns)
            # Create a new DataFrame with all target columns, fill mapped values
            df_new_rows = pd.DataFrame(columns=target_cols_original)
            for col in target_cols_original:
                if col in df_mapped.columns:
                    df_new_rows[col] = df_mapped[col].values
                else:
                    # Unmapped column stays empty/NaN for new rows
                    df_new_rows[col] = pd.NA
            df_result = pd.concat([df_target, df_new_rows], ignore_index=True)
            mode_msg = f"appended {len(df_mapped)} rows after {len(df_target)} existing rows"
        elif start_row is not None:
            # Insert at specific row (preserving all columns)
            df_new_rows = pd.DataFrame(columns=target_cols_original)
            for col in target_cols_original:
                if col in df_mapped.columns:
                    df_new_rows[col] = df_mapped[col].values
                else:
                    df_new_rows[col] = pd.NA
            df_result = pd.concat([
                df_target.iloc[:start_row-1],
                df_new_rows,
                df_target.iloc[start_row-1:]
            ], ignore_index=True)
            mode_msg = f"inserted {len(df_mapped)} rows at row {start_row}"
        else:
            # Replace mode: Clear existing data rows but KEEP ALL target headers
            # Create new DataFrame with ALL target columns, populate only mapped ones
            df_result = pd.DataFrame(columns=target_cols_original)
            for col in target_cols_original:
                if col in df_mapped.columns:
                    df_result[col] = df_mapped[col].values
                else:
                    # Keep column header but with empty/NaN values for imported rows
                    df_result[col] = pd.NA
            mode_msg = f"imported {len(df_mapped)} rows (replaced data, preserved {len(target_cols_original)} headers)"
        
        # Write back - all original columns are preserved
        with wb_cache.excel_writer(target_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_result.to_excel(writer, sheet_name=tgt_sheet, index=False, startrow=tgt_header_row-1)
        
        filter_note = f", {filtered_count} rows filtered out" if filtered_count else ""
        return True, f"Import: {mode_msg} with {matched_cols}/{len(mappings)} column mappings{filter_note}"
        
    except Exception as e:
        return False, f"Error in import: {str(e)}"


def run_append_step(source_path, target_path, src_sheet, tgt_sheet,
                    filter_col=None, filter_val=None, order_by_col=None, order_ascending=True,
                    column_mapping=None, post_append_col=None, post_append_value=None,
                    src_header_row=1, tgt_header_row=1):
    """
    Append rows from source to target with optional filtering and sorting.
    
    Args:
        source_path: Source file path
        target_path: Target file path
        src_sheet: Source sheet name
        tgt_sheet: Target sheet name
        filter_col: Column to filter on (optional)
        filter_val: Value to filter for
        order_by_col: Column to sort by before append
        order_ascending: Sort order
        column_mapping: Dict of source->target column mappings
        post_append_col: Column to tag after append
        post_append_value: Value to tag with
        src_header_row: Source header row (1-indexed)
        tgt_header_row: Target header row (1-indexed)
    
    Returns:
        Tuple of (success: bool, message: str)
    """
    try:
        if str(target_path).lower().endswith('.xlsb'):
            return False, (
                "Cannot write to a .xlsb target file — .xlsb is read-only. "
                "Use a .xlsx file as the target."
            )

        # Read files
        df_source = wb_cache.read_excel(source_path, sheet_name=src_sheet, header=src_header_row-1)
        df_target = wb_cache.read_excel(target_path, sheet_name=tgt_sheet, header=tgt_header_row-1)

        # Filter source if requested
        if filter_col and filter_val and filter_col in df_source.columns:
            df_source = df_source[df_source[filter_col] == filter_val]
        
        # Sort source if requested
        if order_by_col and order_by_col in df_source.columns:
            df_source = df_source.sort_values(by=order_by_col, ascending=order_ascending)
        
        # Apply column mapping if provided
        if column_mapping:
            df_mapped = pd.DataFrame()
            for src_col, tgt_col in column_mapping.items():
                if src_col in df_source.columns:
                    df_mapped[tgt_col] = df_source[src_col]
            df_source = df_mapped
        
        # Tag rows if requested
        if post_append_col and post_append_value:
            df_source[post_append_col] = post_append_value
        
        # Append
        df_result = pd.concat([df_target, df_source], ignore_index=True)
        
        # Write back
        with wb_cache.excel_writer(target_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_result.to_excel(writer, sheet_name=tgt_sheet, index=False, startrow=tgt_header_row-1)
        
        return True, f"Appended {len(df_source)} rows to {len(df_target)} existing rows"
        
    except Exception as e:
        return False, f"Error in append: {str(e)}"


def run_header_mapping_step(source_path, target_path, src_sheet, tgt_sheet, mapping_lines,
                            start_row=None, src_header_row=1, tgt_header_row=1):
    """
    Map columns from source to target using text configuration.
    
    This is essentially the same as run_import_step with column renaming.
    
    Args:
        source_path: Source file path
        target_path: Target file path
        src_sheet: Source sheet name
        tgt_sheet: Target sheet name
        mapping_lines: Mapping string "OldHeader->NewHeader" (one per line)
        start_row: Starting row in target
        src_header_row: Source header row (1-indexed)
        tgt_header_row: Target header row (1-indexed)
    
    Returns:
        Tuple of (success: bool, message: str)
    """
    return run_import_step(source_path, target_path, src_sheet, tgt_sheet,
                          mapping_lines, start_row, src_header_row, tgt_header_row)


