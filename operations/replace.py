"""
Replace values operations for Excel files.
Uses pandas for efficient find/replace operations.
"""

import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from core import wb_cache


def run_replace_values_step(file_path, sheet_name, rules, include_header=False, header_row=1):
    """
    Find and replace values in Excel file.
    
    Args:
        file_path: Path to Excel file
        sheet_name: Sheet name
        rules: List of dicts with 'old_value', 'new_value', 'column' (optional)
        include_header: Whether to include header row in replacement
        header_row: Header row number (1-indexed)
    
    Returns:
        Tuple of (success: bool, message: str)
    """
    try:
        # Read Excel file into DataFrame
        df = wb_cache.read_excel(file_path, sheet_name=sheet_name, header=header_row-1)
        
        replacements_made = 0
        
        for rule in rules:
            old_val = rule.get('old_value', '')
            new_val = rule.get('new_value', '')
            target_col = rule.get('column', None)  # None = all columns
            
            if target_col and target_col in df.columns:
                # Replace in specific column
                df[target_col] = df[target_col].replace(old_val, new_val)
                replacements_made += (df[target_col] == new_val).sum()
            else:
                # Replace in all columns
                df = df.replace(old_val, new_val)
                replacements_made += 1
        
        # Write back to Excel
        wb = wb_cache.load(file_path)
        
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
        
        ws = wb.create_sheet(sheet_name)
        
        # Write DataFrame to worksheet
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx + header_row - 1, column=c_idx, value=value)
        
        wb_cache.save(wb, file_path)
        wb.close()

        return True, f"Replaced values, {len(rules)} rules applied, ~{replacements_made} changes"
        
    except Exception as e:
        return False, f"Error in replace values: {str(e)}"


def run_replace_single(file_path, sheet_name, old_value, new_value, column=None, 
                       scope='all', start_row=None, end_row=None, header_row=1):
    """
    Single find/replace operation.
    
    Args:
        file_path: Path to Excel file
        sheet_name: Sheet name
        old_value: Value to find
        new_value: Value to replace with
        column: Column name (None = all columns)
        scope: 'all', 'first', or 'range'
        start_row: Start row for range scope
        end_row: End row for range scope
        header_row: Header row number
    
    Returns:
        Tuple of (success: bool, message: str)
    """
    rules = [{
        'old_value': old_value,
        'new_value': new_value,
        'column': column
    }]
    
    return run_replace_values_step(file_path, sheet_name, rules, False, header_row)


