"""
Session state management utilities for Streamlit.
"""

import streamlit as st


def safe_rerun():
    """
    Trigger a Streamlit rerun in a backwards-compatible way.
    
    Tries st.rerun() first (newer versions), falls back to
    st.experimental_rerun(), or mutates session state as last resort.
    """
    try:
        # Try modern st.rerun() first (Streamlit 1.27+)
        if hasattr(st, 'rerun'):
            st.rerun()
            return
    except Exception:
        pass
    
    try:
        # Try experimental_rerun for older versions
        rerun_fn = getattr(st, "experimental_rerun", None)
        if callable(rerun_fn):
            rerun_fn()
            return
    except Exception:
        pass

    # Fallback: toggle a session_state counter to provoke rerun
    st.session_state['_safe_rerun_counter'] = st.session_state.get('_safe_rerun_counter', 0) + 1


def set_module_mode(mode: str):
    """
    Set the module mode in session state.
    
    Args:
        mode: Module mode string
    """
    st.session_state.module_mode = mode


def save_module():
    """
    Save a new module to session state.
    """
    name = st.session_state.get("new_module_name", "").strip()
    if not name:
        st.warning("Please enter a module name before saving.")
        return
    
    if 'modules' not in st.session_state:
        st.session_state.modules = []
    
    if name in st.session_state.modules:
        st.warning(f"Module '{name}' already exists.")
        return
    
    st.session_state.modules.append(name)
    st.session_state.selected_module = name
    st.session_state.new_module_name = ""
    st.success(f"Module '{name}' saved and selected.")


def initialize_session_state():
    """
    Initialize common session state variables.
    """
    if 'workflow_steps' not in st.session_state:
        st.session_state.workflow_steps = []
    
    if 'uploaded_file_paths' not in st.session_state:
        # Must be a dict: {display_name: absolute_path}
        # (Older versions mistakenly used a list; keep dict to match the UI.)
        st.session_state.uploaded_file_paths = {}
    
    if 'workflows' not in st.session_state:
        st.session_state.workflows = {}
    
    if 'main_file_select' not in st.session_state:
        st.session_state.main_file_select = ""
    
    if 'main_sheet_name' not in st.session_state:
        st.session_state.main_sheet_name = ""

