"""
Write to Cell(s) operation.
Write a value to a specific cell or fill an entire column with the same value.
Pure Python/openpyxl implementation - NO win32com!
"""

import openpyxl
from openpyxl.utils import column_index_from_string
from core import wb_cache


def run_write_cell_step(
    file_path: str,
    sheet_name: str,
    mode: str,  # "single_cell" or "column"
    value: str,
    # For single cell mode:
    cell_ref: str = "",  # e.g., "A5", "B10"
    # For column mode:
    column: str = "",  # Column letter, e.g., "A", "B"
    start_row: int = 2,
    end_row: str = "last",  # "last" or a number
    header_row: int = 1,
):
    """
    Write a value to a specific cell or fill a column with the same value.
    
    Args:
        file_path: Path to Excel file
        sheet_name: Sheet name
        mode: "single_cell" or "column"
        value: The value to write
        cell_ref: Cell reference for single cell mode (e.g., "A5")
        column: Column letter for column mode (e.g., "C")
        start_row: Start row for column mode (1-indexed)
        end_row: End row or "last" for column mode
        header_row: Header row number (for reference)
    
    Returns:
        Tuple of (success: bool, message: str)
    """
    try:
        wb = wb_cache.load(file_path)
        
        if sheet_name and sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.active
        
        if mode == "single_cell":
            # Write to a single cell
            if not cell_ref:
                wb.close()
                return False, "Cell reference is required for single cell mode (e.g., A5)"
            
            cell = ws[cell_ref]
            old_value = cell.value
            cell.value = value
            
            wb_cache.save(wb, file_path)
            wb.close()
            
            return True, f"Wrote '{value}' to cell {cell_ref} (was: {old_value})"
        
        elif mode == "column":
            # Write same value to entire column
            if not column:
                wb.close()
                return False, "Column letter is required for column mode (e.g., C)"
            
            col_idx = column_index_from_string(column)
            start_row_num = int(start_row)
            
            if end_row == "last" or end_row == "":
                end_row_num = ws.max_row
            else:
                end_row_num = int(end_row)
            
            count = 0
            for row in range(start_row_num, end_row_num + 1):
                cell = ws.cell(row=row, column=col_idx)
                cell.value = value
                count += 1
            
            wb_cache.save(wb, file_path)
            wb.close()
            
            return True, f"Wrote '{value}' to {count} cells in column {column} (rows {start_row_num}-{end_row_num})"
        
        else:
            wb.close()
            return False, f"Invalid mode: {mode}. Use 'single_cell' or 'column'"
        
    except Exception as e:
        return False, f"Error writing to cell(s): {str(e)}"

