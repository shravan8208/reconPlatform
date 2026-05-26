"""
Core utilities for Recon Engine.
"""

from core import wb_cache

from .file_utils import (
    read_uploaded_file, 
    get_sheet_names_for_file, 
    resolve_file_path,
    col_letter_to_index,
    is_xls_file,
    verify_and_preview_excel
)

from .workflow_manager import (
    load_workflows_from_file,
    save_workflow_to_file,
    collect_current_workflow_state,
    apply_loaded_workflow_to_session
)

from .session_manager import (
    safe_rerun, 
    set_module_mode, 
    save_module,
    initialize_session_state
)

# Alias for convenience
init_session_state = initialize_session_state

__all__ = [
    # File utilities
    'read_uploaded_file',
    'get_sheet_names_for_file',
    'resolve_file_path',
    'col_letter_to_index',
    'is_xls_file',
    'verify_and_preview_excel',
    
    # Workflow management
    'load_workflows_from_file',
    'save_workflow_to_file',
    'collect_current_workflow_state',
    'apply_loaded_workflow_to_session',
    
    # Session management
    'safe_rerun',
    'set_module_mode',
    'save_module',
    'initialize_session_state',
    'init_session_state',  # Alias

    # Workbook cache
    'wb_cache',
]
