"""
File utilities for reading and managing Excel/CSV files.
No win32com dependency - pure Python with openpyxl/pandas.
"""

import io
import pandas as pd
import openpyxl
from pathlib import Path
from typing import Optional, List


def read_uploaded_file(file_bytes: bytes, filename: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """
    Read uploaded file (CSV or Excel) into a DataFrame.
    
    Args:
        file_bytes: File content as bytes
        filename: Original filename
        sheet_name: Sheet name for Excel files (None = first sheet)
    
    Returns:
        DataFrame with file contents
    """
    if not file_bytes:
        return pd.DataFrame()
    
    bio = io.BytesIO(file_bytes)
    lower = (filename or "").lower()
    
    try:
        if lower.endswith('.csv') or lower.endswith('.txt'):
            return pd.read_csv(bio)
        
        if lower.endswith('.xlsb'):
            engine = "pyxlsb"
            if sheet_name:
                return pd.read_excel(bio, sheet_name=sheet_name, engine=engine)
            else:
                # pyxlsb: get sheet list first
                import pyxlsb
                bio.seek(0)
                with pyxlsb.open_workbook(bio) as wb:
                    sheet = wb.sheets[0] if wb.sheets else None
                bio.seek(0)
                return pd.read_excel(bio, sheet_name=sheet, engine=engine)
        if lower.endswith('.xls') or lower.endswith('.xlsx') or lower.endswith('.xlsm'):
            if sheet_name:
                return pd.read_excel(bio, sheet_name=sheet_name)
            else:
                x = pd.ExcelFile(bio)
                sheet = x.sheet_names[0] if x.sheet_names else None
                return pd.read_excel(bio, sheet_name=sheet)
    except Exception:
        try:
            return pd.read_csv(bio)
        except Exception:
            return pd.DataFrame()
    
    return pd.DataFrame()


def get_sheet_names_for_file(file_path: str) -> List[str]:
    """
    Get list of sheet names from an Excel file.
    Supports .xlsx, .xlsm, .xls (via openpyxl/xlrd) and .xlsb (via pyxlsb).

    Args:
        file_path: Path to Excel file

    Returns:
        List of sheet names
    """
    lower = (file_path or "").lower()
    try:
        if lower.endswith(".xlsb"):
            import pyxlsb
            with pyxlsb.open_workbook(file_path) as wb:
                return list(wb.sheets)
        elif lower.endswith(('.xls', '.xlsx', '.xlsm')):
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            names = list(wb.sheetnames)
            wb.close()
            return names
    except Exception:
        pass
    return []


def resolve_file_path(label: str, uploaded_files: List[str]) -> str:
    """
    Resolve a file label to actual file path.
    
    Args:
        label: File label or path
        uploaded_files: List of uploaded file paths
    
    Returns:
        Resolved file path or empty string
    """
    import os
    
    # Try exact match first
    for f in uploaded_files:
        if os.path.basename(f) == label or f == label:
            return f
    
    # Try case-insensitive match
    label_lower = label.lower()
    for f in uploaded_files:
        if os.path.basename(f).lower() == label_lower or f.lower() == label_lower:
            return f
    
    # Try partial match
    for f in uploaded_files:
        f_basename = os.path.basename(f)
        if label in f_basename or f_basename in label:
            return f
    
    # If it's an actual path that exists, return it
    if os.path.exists(label):
        return label
    
    return ""


def col_letter_to_index(col_letter: str) -> int:
    """
    Convert Excel column letter to index (A=1, B=2, etc.).
    
    Args:
        col_letter: Column letter (e.g., 'A', 'AB')
    
    Returns:
        Column index (1-based)
    """
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(col_letter)


def is_xls_file(path: Path) -> bool:
    """
    Check if file is old XLS format (not XLSX).
    
    Args:
        path: File path
    
    Returns:
        True if XLS format
    """
    return str(path).lower().endswith('.xls') and not str(path).lower().endswith('.xlsx')


def verify_and_preview_excel(file_path: str) -> tuple:
    """
    Verify Excel file can be opened and get preview info.
    
    Args:
        file_path: Path to Excel file
    
    Returns:
        Tuple of (success: bool, message: str, sheet_names: list)
    """
    try:
        if str(file_path).lower().endswith(".xlsb"):
            import pyxlsb
            with pyxlsb.open_workbook(file_path) as wb:
                sheets = list(wb.sheets)
            return True, f"File valid (.xlsb). Sheets: {', '.join(sheets)}", sheets
        wb = openpyxl.load_workbook(file_path, read_only=True)
        sheets = list(wb.sheetnames)
        wb.close()
        return True, f"File valid. Sheets: {', '.join(sheets)}", sheets
    except Exception as e:
        return False, f"Error opening file: {str(e)}", []


