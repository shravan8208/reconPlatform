"""
Excel operations for Recon Engine.
Each module handles a specific type of Excel manipulation.
ALL operations are win32com-free!
"""

# Formula operations
from .formula import (
    run_formula_step,
    apply_multiple_formulas_to_xls,
    run_convert_to_values_step,
    adjust_formula_for_row,
    clear_vlookup_cache,
    run_formula_broadcast_step,
)

# Copy/Paste operations
from .copy_paste import run_copy_paste_step

# Replace operations
from .replace import (
    run_replace_values_step,
    run_replace_single
)

# Data manipulation
from .data_manipulation import (
    run_forward_fill_step,
    run_insert_delete_step,
    run_delete_columns_step,
    run_clear_columns_data_step,
    run_delete_by_condition_step,
    run_advanced_delete_step,
    preview_advanced_delete,
)

# VLOOKUP operations
from .vlookup import (
    vlookup_single,
    run_vlookup_step,
    run_vlookup_explicit_step
)

# Advanced VLOOKUP
from .advance_vlookup import run_advance_vlookup_step, run_advance_vlookup_v2

# Filter and Sort
from .filter_sort import (
    run_filter_step,
    run_sort_step
)

# Import/Append/Header Mapping
from .import_append import (
    run_import_step,
    run_append_step,
    run_header_mapping_step
)

# Conditional operations
from .conditional import run_conditional_write_step

# SUMIFS operations
from .sumifs import run_sumifs_step

# Merge cells
from .merge_cells import (
    run_merge_cells_step,
    run_unmerge_cells_step
)

# Data Input
from .data_input import (
    run_data_input_step,
    get_file_info
)

# Write to Cell(s)
from .write_cell import run_write_cell_step

# Generic operations (flexible building blocks)
from .generic_ops import (
    run_quota_update_step,
    run_group_rollup_step,
)

# Sheet operations (delete / manage sheets)
from .sheet_ops import run_delete_sheets_step, get_sheet_names

# Input Source / File Collector
from .input_source import run_input_source_step, resolve_file_class, _resolve_tokens

# Folder Summary (multi-file extraction + consolidation)
from .folder_summary import run_folder_summary_step

# Filter utilities (shared by import and delete-by-condition)
from .filter_utils import apply_filters, OPERATORS, NEEDS_VALUE

# Formatting & Beautification
from .formatter import run_format_step

# Split Export Engine
from .split_export import (
    run_split_export_step,
    SPLIT_MODES,
    SPLIT_MODE_KEYS,
    SPLIT_MODE_LABELS,
)

# Unpivot (Wide → Long)
from .unpivot import run_unpivot_step

# Sheet Updater
from .sheet_updater import (
    run_sheet_updater_step,
    SHEET_MATCH_MODES,
    SHEET_MATCH_KEYS,
    SHEET_MATCH_LABELS,
    PARTICULARS_MATCH_MODES,
    PARTICULARS_MATCH_KEYS,
    PARTICULARS_MATCH_LABELS,
)

# Pivot Engine
from .pivot_engine import (
    run_pivot_step,
    detect_columns,
    save_pivot_template,
    load_pivot_template,
    delete_pivot_template,
    list_pivot_templates,
    AGGFUNC_KEYS,
    AGGFUNC_LABELS,
)

# Normalize Import (per-class header mapping → working file)
from .normalize_import import (
    run_normalize_import_step,
    detect_headers,
    apply_header_mapping,
    load_mapping_from_file,
    load_mapping_from_bytes,
    load_all_templates,
    save_template,
    delete_template,
)

__all__ = [
    # Formula
    'run_formula_step',
    'apply_multiple_formulas_to_xls',
    'run_convert_to_values_step',
    'adjust_formula_for_row',
    'run_formula_broadcast_step',
    
    # Copy/Paste
    'run_copy_paste_step',
    
    # Replace
    'run_replace_values_step',
    'run_replace_single',
    
    # Data Manipulation
    'run_forward_fill_step',
    'run_insert_delete_step',
    'run_delete_columns_step',
    'run_clear_columns_data_step',
    'run_advanced_delete_step',
    'preview_advanced_delete',
    
    # VLOOKUP
    'vlookup_single',
    'run_vlookup_step',
    'run_vlookup_explicit_step',
    'run_advance_vlookup_step',
    
    # Filter/Sort
    'run_filter_step',
    'run_sort_step',
    
    # Import/Append
    'run_import_step',
    'run_append_step',
    'run_header_mapping_step',
    
    # Conditional
    'run_conditional_write_step',
    
    # SUMIFS
    'run_sumifs_step',
    
    # Merge Cells
    'run_merge_cells_step',
    'run_unmerge_cells_step',
    
    # Data Input
    'run_data_input_step',
    'get_file_info',
    
    # Write to Cell(s)
    'run_write_cell_step',

    # Generic Ops
    'run_quota_update_step',
    'run_group_rollup_step',

    # Sheet Ops
    'run_delete_sheets_step',
    'get_sheet_names',

    # Input Source
    'run_input_source_step',
    'resolve_file_class',
    '_resolve_tokens',

    # Format & Beautify
    'run_format_step',

    # Split Export Engine
    'run_split_export_step',
    'SPLIT_MODES',
    'SPLIT_MODE_KEYS',
    'SPLIT_MODE_LABELS',

    # Unpivot
    'run_unpivot_step',

    # Sheet Updater
    'run_sheet_updater_step',
    'SHEET_MATCH_MODES',
    'SHEET_MATCH_KEYS',
    'SHEET_MATCH_LABELS',
    'PARTICULARS_MATCH_MODES',
    'PARTICULARS_MATCH_KEYS',
    'PARTICULARS_MATCH_LABELS',

    # Pivot Engine
    'run_pivot_step',
    'detect_columns',
    'save_pivot_template',
    'load_pivot_template',
    'delete_pivot_template',
    'list_pivot_templates',
    'AGGFUNC_KEYS',
    'AGGFUNC_LABELS',

    # Folder Summary
    'run_folder_summary_step',

    # Normalize Import
    'run_normalize_import_step',
    'detect_headers',
    'apply_header_mapping',
    'load_mapping_from_file',
    'load_mapping_from_bytes',
    'load_all_templates',
    'save_template',
    'delete_template',
]
