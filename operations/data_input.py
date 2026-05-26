"""
Data Input operation - Load data as the first step in a workflow.
This is a NEW operation not in MAIN7.PY.
"""

import pandas as pd
import openpyxl
from pathlib import Path


def run_data_input_step(file_path, sheet_name=None, header_row=1, output_path=None):
    """
    Load data from a file as the first step in a workflow.
    This operation reads the file and optionally saves it to a new location.
    
    Args:
        file_path: Path to input file (CSV or Excel)
        sheet_name: Sheet name (for Excel files, None = first sheet)
        header_row: Header row number (1-indexed)
        output_path: Optional output path (if None, modifies in place)
    
    Returns:
        Tuple of (success: bool, message: str, dataframe: pd.DataFrame or None)
    """
    try:
        file_path_obj = Path(file_path)
        
        if not file_path_obj.exists():
            return False, f"File not found: {file_path}", None
        
        # Read file based on extension
        if file_path.lower().endswith('.csv'):
            df = pd.read_csv(file_path)
        elif file_path.lower().endswith(('.xlsx', '.xls')):
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row-1)
            else:
                df = pd.read_excel(file_path, header=header_row-1)
        else:
            return False, "Unsupported file format (use CSV or Excel)", None
        
        # Optionally save to output path
        if output_path:
            if output_path.lower().endswith('.csv'):
                df.to_csv(output_path, index=False)
            elif output_path.lower().endswith(('.xlsx', '.xls')):
                df.to_excel(output_path, index=False, engine='openpyxl')
            else:
                return False, "Unsupported output format", None
            
            msg = f"Loaded {len(df)} rows, {len(df.columns)} columns and saved to {output_path}"
        else:
            msg = f"Loaded {len(df)} rows, {len(df.columns)} columns"
        
        return True, msg, df
        
    except Exception as e:
        return False, f"Error loading data: {str(e)}", None


def get_file_info(file_path):
    """
    Get information about a data file.
    
    Args:
        file_path: Path to file
    
    Returns:
        Dict with file information
    """
    try:
        file_path_obj = Path(file_path)
        
        if not file_path_obj.exists():
            return {'exists': False, 'error': 'File not found'}
        
        info = {
            'exists': True,
            'name': file_path_obj.name,
            'size': file_path_obj.stat().st_size,
            'extension': file_path_obj.suffix
        }
        
        # Get sheet names for Excel files
        if file_path.lower().endswith(('.xlsx', '.xls')):
            wb = openpyxl.load_workbook(file_path, read_only=True)
            info['sheets'] = wb.sheetnames
            wb.close()
        
        # Get row/column count
        if file_path.lower().endswith('.csv'):
            df = pd.read_csv(file_path, nrows=1)
            info['columns'] = len(df.columns)
        elif file_path.lower().endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file_path, nrows=1)
            info['columns'] = len(df.columns)
        
        return info
        
    except Exception as e:
        return {'exists': False, 'error': str(e)}


