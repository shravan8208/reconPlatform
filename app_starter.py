"""
Recon Engine - Complete App with ALL 18 Operations
NO win32com dependency! Cross-platform compatible.
"""

import streamlit as st
import pandas as pd
import json
import os
import re
import shutil
import uuid
import threading
import concurrent.futures as cf
import openpyxl
from pathlib import Path
from datetime import datetime

# Import ALL refactored operations
from operations import (
    # Formula operations
    run_formula_step,
    run_convert_to_values_step,
    run_formula_broadcast_step,
    run_unpivot_step,
    
    # Copy/Paste
    run_copy_paste_step,
    
    # Replace
    run_replace_values_step,
    run_replace_single,
    
    # Data Manipulation
    run_forward_fill_step,
    run_insert_delete_step,
    run_delete_columns_step,
    run_clear_columns_data_step,
    
    # VLOOKUP
    run_vlookup_step,
    run_vlookup_explicit_step,
    run_advance_vlookup_step,
    
    # Filter/Sort
    run_filter_step,
    run_sort_step,
    
    # Import/Append
    run_import_step,
    run_append_step,
    run_header_mapping_step,
    
    # Conditional
    run_conditional_write_step,
    
    # SUMIFS
    run_sumifs_step,
    
    # Merge Cells
    run_merge_cells_step,
    run_unmerge_cells_step,
    
    # Data Input
    run_data_input_step,
    get_file_info,
    
    # Write to Cell(s)
    run_write_cell_step,

    # Generic Ops
    run_quota_update_step,
    run_group_rollup_step,

    # Input Source
    run_input_source_step,
    resolve_file_class,
    _resolve_tokens,

    # Folder Summary
    run_folder_summary_step,

    # Normalize Import
    run_normalize_import_step,
    detect_headers,
    apply_header_mapping,
    load_mapping_from_bytes,
    load_all_templates,
    save_template,
    delete_template,

    # Filter utilities
    apply_filters,
    OPERATORS,
    NEEDS_VALUE,

    # Delete by condition / Advanced Delete
    run_delete_by_condition_step,
    run_advanced_delete_step,
    preview_advanced_delete,

    # Format & Beautify
    run_format_step,

    # Pivot
    run_pivot_step,
    detect_columns,
    save_pivot_template,
    load_pivot_template,
    delete_pivot_template,
    list_pivot_templates,
    AGGFUNC_KEYS,
    AGGFUNC_LABELS,

    # Split Export
    run_split_export_step,
    SPLIT_MODES,
    SPLIT_MODE_KEYS,
    SPLIT_MODE_LABELS,

    # Sheet Row Inserter
    run_sheet_row_inserter_step,
    ANCHOR_MATCH_MODES,
    ANCHOR_MATCH_KEYS,
    ANCHOR_MATCH_LABELS,
    INSERT_POSITIONS,
    INSERT_POSITION_KEYS,
    INSERT_POSITION_LABELS,

    # Sheet Updater
    run_sheet_updater_step,
    SHEET_MATCH_MODES,
    SHEET_MATCH_KEYS,
    SHEET_MATCH_LABELS,
    PARTICULARS_MATCH_MODES,
    PARTICULARS_MATCH_KEYS,
    PARTICULARS_MATCH_LABELS,

    # Sheet Ops
    run_delete_sheets_step,
    get_sheet_names,
)

# Import core utilities
from core import (
    load_workflows_from_file,
    save_workflow_to_file,
    init_session_state
)
from core.file_utils import get_sheet_names_for_file

# Page configuration
st.set_page_config(
    page_title="PKF Proserve | Reconciliation Engine",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize session state
init_session_state()

if 'uploaded_file_paths'not in st.session_state or not isinstance(st.session_state.uploaded_file_paths, dict):
    # Safety: older state could be a list; normalize to dict.
    st.session_state.uploaded_file_paths = {}
if 'workflow_steps'not in st.session_state:
    st.session_state.workflow_steps = []
if 'insert_at'not in st.session_state:
    # When set (0-based index), the next added step will be inserted at this position.
    st.session_state.insert_at = None
if 'editing_step_idx'not in st.session_state:
    # Which step (0-based index) is currently being edited in the Workflow tab.
    st.session_state.editing_step_idx = None
if 'editing_step_data'not in st.session_state:
    # The full step dict being edited (for visual editing in original form)
    st.session_state.editing_step_data = None
if '_workflow_notice'not in st.session_state:
    st.session_state._workflow_notice = ""
if 'default_export_path'not in st.session_state:
    # User-specified default export folder or file path.
    st.session_state.default_export_path = ""

# Ensure every step has a stable ID for dependency management
def ensure_step_ids():
    steps = st.session_state.get("workflow_steps", []) or []
    changed = False
    for s in steps:
        if not isinstance(s, dict):
            continue
        if not s.get("id"):
            s["id"] = uuid.uuid4().hex[:12]
            changed = True
        if "depends_on"not in s:
            s["depends_on"] = []
            changed = True
        if "lane"not in s:
            s["lane"] = ""
            changed = True
    if changed:
        st.session_state.workflow_steps = steps

ensure_step_ids()

# ── PKF Proserve Professional White Theme ────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

/* ── Base & Background ── */
html, body, [data-testid="stAppViewContainer"], .stApp {
    background-color: #ffffff !important;
    color: #1a1a1a !important;
    font-family: 'IBM Plex Sans', -apple-system, sans-serif !important;
    font-size: 13px;
}
[data-testid="stAppViewBlockContainer"] {
    padding-top: 0 !important;
}

/* ── Sidebar ── */
section[data-testid="stSidebar"] {
    background-color: #f7f7f7 !important;
    border-right: 1px solid #e0e0e0 !important;
}
section[data-testid="stSidebar"] * {
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-size: 12px !important;
    color: #333333 !important;
}
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {
    color: #1a1a1a !important;
    font-size: 10px !important;
    text-transform: uppercase;
    letter-spacing: 0.15em;
    font-weight: 600;
    border-bottom: 1px solid #e0e0e0;
    padding-bottom: 4px;
    margin-bottom: 8px;
}

/* ── Typography ── */
h1, h2, h3, h4 {
    font-family: 'IBM Plex Sans', sans-serif !important;
    color: #1a1a1a !important;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    font-weight: 600;
    margin-bottom: 6px;
}
h1 { font-size: 17px !important; border-bottom: 2px solid #c8392b; padding-bottom: 6px; }
h2 { font-size: 13px !important; border-bottom: 1px solid #e0e0e0; padding-bottom: 4px; }
h3 { font-size: 11px !important; color: #444444 !important; }
p, label, .stMarkdown { color: #333333 !important; font-size: 12px !important; }
code {
    color: #c8392b !important;
    background: #fff5f5 !important;
    border: 1px solid #f5c6c6;
    padding: 1px 5px;
    border-radius: 2px;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 11px !important;
}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
    background-color: #ffffff !important;
    border-bottom: 2px solid #e0e0e0 !important;
    gap: 0px;
}
.stTabs [data-baseweb="tab"] {
    background-color: transparent !important;
    color: #888888 !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-size: 10px !important;
    font-weight: 600;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    padding: 10px 22px !important;
    border-radius: 0 !important;
    border: none !important;
}
.stTabs [data-baseweb="tab"]:hover {
    color: #1a1a1a !important;
    background-color: #f5f5f5 !important;
}
.stTabs [aria-selected="true"] {
    color: #c8392b !important;
    background-color: #ffffff !important;
    border-bottom: 2px solid #c8392b !important;
}
.stTabs [data-baseweb="tab-panel"] {
    background-color: #ffffff !important;
    padding-top: 16px !important;
}

/* ── Buttons ── */
.stButton > button {
    background-color: #ffffff !important;
    color: #1a1a1a !important;
    border: 1px solid #cccccc !important;
    border-radius: 2px !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-size: 11px !important;
    font-weight: 500;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    width: 100%;
    padding: 6px 14px !important;
    transition: all 0.12s ease;
}
.stButton > button:hover {
    background-color: #1a1a1a !important;
    color: #ffffff !important;
    border-color: #1a1a1a !important;
}
.stButton > button[kind="primary"] {
    background-color: #c8392b !important;
    color: #ffffff !important;
    border: none !important;
    font-weight: 600 !important;
}
.stButton > button[kind="primary"]:hover {
    background-color: #a02d22 !important;
}

/* ── Inputs ── */
.stTextInput > div > div > input,
.stNumberInput > div > div > input,
.stTextArea > div > div > textarea {
    background-color: #ffffff !important;
    color: #1a1a1a !important;
    border: 1px solid #cccccc !important;
    border-radius: 2px !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-size: 12px !important;
}
.stTextInput > div > div > input:focus,
.stNumberInput > div > div > input:focus,
.stTextArea > div > div > textarea:focus {
    border-color: #c8392b !important;
    box-shadow: 0 0 0 1px #c8392b !important;
}
.stSelectbox > div > div,
.stMultiSelect > div > div {
    background-color: #ffffff !important;
    border: 1px solid #cccccc !important;
    border-radius: 2px !important;
    color: #1a1a1a !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-size: 12px !important;
}
/* Dropdown list */
[data-baseweb="popover"], [data-baseweb="menu"] {
    background-color: #ffffff !important;
    border: 1px solid #cccccc !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.10) !important;
}
[data-baseweb="option"] {
    background-color: #ffffff !important;
    color: #1a1a1a !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-size: 12px !important;
}
[data-baseweb="option"]:hover {
    background-color: #f5f5f5 !important;
    color: #c8392b !important;
}

/* ── Checkboxes & Radio ── */
.stCheckbox > label, .stRadio > label {
    color: #333333 !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-size: 12px !important;
}
.stRadio [data-baseweb="radio"] span { border-color: #c8392b !important; }
.stCheckbox [data-baseweb="checkbox"] span { border-color: #c8392b !important; }

/* ── Expander ── */
.streamlit-expanderHeader {
    background-color: #f7f7f7 !important;
    color: #444444 !important;
    border: 1px solid #e0e0e0 !important;
    border-radius: 2px !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-size: 11px !important;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.08em;
}
.streamlit-expanderHeader:hover { color: #c8392b !important; background-color: #f0f0f0 !important; }
.streamlit-expanderContent {
    background-color: #fafafa !important;
    border: 1px solid #e0e0e0 !important;
    border-top: none !important;
}

/* ── Alert boxes ── */
.stAlert {
    border-radius: 2px !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-size: 12px !important;
}
div[data-baseweb="notification"] {
    background-color: #fffbf9 !important;
    border-left: 3px solid #c8392b !important;
    color: #1a1a1a !important;
    border-radius: 0 !important;
}
.stSuccess { background-color: #f0faf5 !important; border-left: 3px solid #1a7a4a !important; }
.stWarning { background-color: #fffbf0 !important; border-left: 3px solid #b8860b !important; }
.stError   { background-color: #fff5f5 !important; border-left: 3px solid #c8392b !important; }
.stInfo    { background-color: #f0f5ff !important; border-left: 3px solid #1a4fa0 !important; }

/* ── Dividers ── */
hr { border: none !important; border-top: 1px solid #e8e8e8 !important; margin: 14px 0 !important; }

/* ── Progress bar ── */
.stProgress > div > div > div { background-color: #c8392b !important; }

/* ── Download button ── */
.stDownloadButton > button {
    background-color: #ffffff !important;
    color: #333333 !important;
    border: 1px solid #cccccc !important;
    border-radius: 2px !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-size: 11px !important;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.06em;
}
.stDownloadButton > button:hover {
    background-color: #1a1a1a !important;
    color: #ffffff !important;
}

/* ── Dataframes / Tables ── */
.stDataFrame, .dataframe {
    background-color: #ffffff !important;
    border: 1px solid #e0e0e0 !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 11px !important;
}
.stDataFrame thead th {
    background-color: #f5f5f5 !important;
    color: #1a1a1a !important;
    border-bottom: 2px solid #1a1a1a !important;
    font-size: 10px !important;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    font-weight: 600;
}
.stDataFrame tbody tr:hover { background-color: #fafafa !important; }
.stDataFrame tbody td { border-color: #f0f0f0 !important; }

/* ── JSON viewer ── */
.stJson { background-color: #fafafa !important; border: 1px solid #e0e0e0 !important; border-radius: 2px !important; }

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: #f5f5f5; }
::-webkit-scrollbar-thumb { background: #cccccc; border-radius: 2px; }
::-webkit-scrollbar-thumb:hover { background: #c8392b; }

/* ── PKF Header strip ── */
.pkf-header {
    background: #1a1a1a;
    border-bottom: 3px solid #c8392b;
    padding: 10px 24px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    font-family: 'IBM Plex Sans', sans-serif;
    letter-spacing: 0.06em;
    margin-bottom: 0;
}
.pkf-brand {
    color: #ffffff;
    font-size: 15px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.1em;
}
.pkf-product {
    color: #aaaaaa;
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: 0.18em;
    margin-top: 2px;
}
.pkf-dev {
    color: #666666;
    font-size: 9px;
    text-transform: uppercase;
    letter-spacing: 0.2em;
    text-align: right;
}
.pkf-clock {
    color: #888888;
    font-size: 10px;
    font-family: 'IBM Plex Mono', monospace;
    font-variant-numeric: tabular-nums;
    text-align: right;
    margin-top: 2px;
}

/* ── Sub-header ticker ── */
.pkf-ticker {
    background: #f5f5f5;
    border-bottom: 1px solid #e0e0e0;
    padding: 4px 24px;
    font-family: 'IBM Plex Sans', sans-serif;
    font-size: 9px;
    color: #999999;
    letter-spacing: 0.2em;
    text-transform: uppercase;
    margin-bottom: 12px;
}

/* ── Section label chips ── */
.pkf-section {
    display: inline-block;
    background: #1a1a1a;
    color: #ffffff;
    font-family: 'IBM Plex Sans', sans-serif;
    font-size: 9px;
    font-weight: 600;
    letter-spacing: 0.2em;
    text-transform: uppercase;
    padding: 3px 10px;
    margin-bottom: 10px;
    border-radius: 1px;
}

/* ── Misc ── */
.stCaption {
    color: #999999 !important;
    font-size: 10px !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
}
.big-font { font-size: 13px !important; font-weight: 600; color: #c8392b; }
[data-testid="stMarkdownContainer"] p { font-size: 12px !important; color: #333333 !important; }
</style>
""", unsafe_allow_html=True)


def main():
    # ── Flush pending edit populate ──────────────────────────────────────────
    # Must run BEFORE any widget is instantiated. When "Edit"is clicked,
    # populate_form_from_step() cannot run in the same render cycle (widgets
    # are already instantiated). Instead, the data is stored here and applied
    # at the very top of the next render so widgets pick up pre-set values.
    if "_pending_edit_populate"in st.session_state:
        _pep = st.session_state["_pending_edit_populate"]
        del st.session_state["_pending_edit_populate"]
        populate_form_from_step(_pep[0], _pep[1])
    # ─────────────────────────────────────────────────────────────────────────

    # ── Brand header ─────────────────────────────────────────────────────────
    st.markdown("""
<div class="pkf-header">
    <div>
        <div class="pkf-brand">PKF Proserve Pvt Ltd</div>
        <div class="pkf-product">Reconciliation Workflow Engine</div>
    </div>
    <div style="text-align:right;">
        <div class="pkf-dev">Developed by Cap Corporate</div>
        <div class="pkf-clock" id="pkf-clock">--</div>
    </div>
</div>
<div class="pkf-ticker">
    RECON ENGINE &nbsp;|&nbsp; ALL OPERATIONS &nbsp;|&nbsp; CROSS-PLATFORM &nbsp;|&nbsp; NO WIN32COM
</div>
<script>
(function tick() {
    var el = document.getElementById('pkf-clock');
    if (el) {
        var now = new Date();
        el.textContent = now.toLocaleTimeString('en-GB', {hour12: false})
            + '  ' + now.toLocaleDateString('en-GB');
    }
    setTimeout(tick, 1000);
})();
</script>
""", unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.markdown("""
<div style="padding: 8px 0 4px 0; border-bottom: 1px solid #e0e0e0; margin-bottom: 8px;">
    <div style="color:#1a1a1a; font-size:11px; font-weight:700; letter-spacing:0.15em; text-transform:uppercase;">
        PKF Proserve Pvt Ltd
    </div>
    <div style="color:#444; font-size:9px; letter-spacing:0.2em; text-transform:uppercase; margin-top:2px;">
        Reconciliation Engine
    </div>
</div>
""", unsafe_allow_html=True)
        st.markdown('<div class="pkf-section">Registered Files</div>', unsafe_allow_html=True)
        st.caption("Files are edited in-place. Use the Add Files step to register absolute local paths.")

        with st.expander("Export settings", expanded=False):
            st.session_state.default_export_path = st.text_input(
                "Default export folder (or file path)",
                value=st.session_state.get("default_export_path", ""),
                help="If a step has Export enabled but no per-step path, this default is used. "
                     "Examples: `/Users/you/Exports/` or `/Users/you/Exports/output.xlsx`",
                key="sidebar_default_export_path",
            )
        
        # Show uploaded files
        if st.session_state.uploaded_file_paths:
            st.markdown('<div class="pkf-section" style="margin-top:8px;">Registered Files</div>', unsafe_allow_html=True)
            for fname, fpath in st.session_state.uploaded_file_paths.items():
                st.text(f" {fname}")
                with st.expander(f"Path / Download: {fname}", expanded=False):
                    st.code(fpath)
                    try:
                        with open(fpath, "rb") as f:
                            data = f.read()
                        lower = (fname or "").lower()
                        if lower.endswith(".csv"):
                            mime = "text/csv"
                        elif lower.endswith(".xlsx"):
                            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        elif lower.endswith(".xls"):
                            mime = "application/vnd.ms-excel"
                        else:
                            mime = "application/octet-stream"
                        st.download_button(
                            label=f"Download current {fname}",
                            data=data,
                            file_name=f"{fname}",
                            mime=mime,
                            key=f"dl_{fname}",
                            use_container_width=True,
                        )
                    except Exception as e:
                        st.error(f"Could not read file: {e}")
        else:
            st.info("No files registered yet. Add an **Add File(s)** step first.")
        
        st.divider()

        # Workflow Management
        st.markdown('<div class="pkf-section">Workflow Library</div>', unsafe_allow_html=True)

        # Save
        workflow_name = st.text_input("Workflow name", placeholder="e.g. Monthly Recon")
        if st.button("Save Workflow", use_container_width=True):
            if workflow_name and st.session_state.workflow_steps:
                data = {
                    "steps": st.session_state.workflow_steps,
                    "uploaded_file_paths": dict(st.session_state.uploaded_file_paths),
                    "saved_at": datetime.utcnow().isoformat()
                }
                success, error = save_workflow_to_file(workflow_name, data)
                if success:
                    st.success(f"Saved: {workflow_name}")
                else:
                    st.error(error)
            else:
                st.warning("Enter a name and add at least one step.")

        # Load
        workflows = load_workflows_from_file()
        if workflows:
            selected = st.selectbox("Select workflow", [""] + list(workflows.keys()))
            if st.button("Load Workflow", use_container_width=True) and selected:
                st.session_state.workflow_steps = workflows[selected].get('steps', [])
                uploaded = workflows[selected].get("uploaded_file_paths")
                if isinstance(uploaded, dict):
                    st.session_state.uploaded_file_paths = uploaded
                ensure_step_ids()
                st.success(f"Loaded: {selected}")
                st.rerun()

        # Clear
        st.divider()
        if st.button("Clear Workflow", use_container_width=True):
            st.session_state.workflow_steps = []
            st.success("Workflow cleared.")
            st.rerun()

        # Footer
        st.markdown("""
<div style="position:fixed; bottom:0; left:0; width:240px; background:#ffffff;
     border-top:1px solid #e0e0e0; padding:6px 12px; z-index:9999;">
    <div style="color:#aaaaaa; font-size:8px; letter-spacing:0.15em; text-transform:uppercase;
         font-family:'IBM Plex Sans',sans-serif; line-height:1.6;">
        PKF Proserve Pvt Ltd<br>
        <span style="color:#cccccc;">Developed by Cap Corporate</span>
    </div>
</div>
""", unsafe_allow_html=True)
    
    # Main tabs
    tabs = st.tabs([
        "Add Steps",
        "Workflow",
        "Execute",
        "Status"
    ])
    
    # TAB 1: Add Steps
    with tabs[0]:
        render_add_steps_tab()
    
    # TAB 2: Current Workflow
    with tabs[1]:
        render_workflow_tab()
    
    # TAB 3: Execute
    with tabs[2]:
        render_execute_tab()
    
    # TAB 4: Status
    with tabs[3]:
        render_status_tab()


def _restore_style(pfx: str, i: int, style: dict):
    """Restore style editor session-state keys for rule i under prefix pfx."""
    key = f"{pfx}_{i}"
    st.session_state[f"{key}_bold"]         = bool(style.get("bold", False))
    st.session_state[f"{key}_italic"]       = bool(style.get("italic", False))
    st.session_state[f"{key}_underline"]    = bool(style.get("underline", False))
    st.session_state[f"{key}_strike"]       = bool(style.get("strike", False))
    st.session_state[f"{key}_font_size"]    = int(style.get("font_size", 0) or 0)
    fc = style.get("font_color", "")
    st.session_state[f"{key}_font_color"]   = ("#" + fc.lstrip("#")) if fc else "#000000"
    bg = style.get("bg_color", "")
    st.session_state[f"{key}_bg_color"]     = ("#" + bg.lstrip("#")) if bg else "#FFFFFF"
    st.session_state[f"{key}_border_style"] = style.get("border_style", "none")
    st.session_state[f"{key}_border_sides"] = style.get("border_sides", "all")
    st.session_state[f"{key}_h_align"]      = style.get("h_align", "")
    st.session_state[f"{key}_wrap_text"]    = bool(style.get("wrap_text", False))


def populate_form_from_step(step_type: str, config: dict):
    """Populate form session state keys from a step's config for visual editing."""
    cfg = config or {}
    
    # Common patterns: file, sheet, column, header_row, export
    key_maps = {
        "formula": [("file", "formula_file"), ("sheet", "formula_sheet"), ("formula", "formula_expr"),
                    ("column", "formula_col"), ("start", "formula_start"), ("convert", "formula_convert"),
                    ("column_header_row", "formula_col_header_row")],
        "copy_paste": [("file", "copy_paste_file"), ("sheet", "copy_paste_sheet"), ("src_col", "copy_paste_src"),
                       ("paste_type", "copy_paste_type"), ("header_row", "copy_paste_header")],
        "convert_values": [("file", "convert_values_file"), ("sheet", "convert_values_sheet"), ("column", "convert_values_col"),
                          ("start", "convert_values_start")],
        "forward_fill": [("file", "forward_fill_file"), ("sheet", "forward_fill_sheet"), ("column", "forward_fill_col")],
        "replace": [("file", "replace_file"), ("sheet", "replace_sheet"), ("old_val", "replace_old"),
                   ("new_val", "replace_new"), ("column", "replace_col")],
        "conditional": [("file", "conditional_file"), ("sheet", "conditional_sheet")],
        "format": [("file", "format_file"), ("sheet", "format_sheet")],
        "pivot":  [("src_file", "pivot_src_file"), ("src_sheet", "pivot_src_sheet"),
                   ("tgt_file", "pivot_tgt_file")],
        "split_export": [("src_file", "se_src_file"), ("src_sheet", "se_src_sheet"),
                         ("tgt_file", "se_tgt_file")],
        "sheet_row_inserter": [("src_file", "sri_src_file"), ("src_sheet", "sri_src_sheet"),
                               ("tgt_file", "sri_tgt_file")],
        "sheet_updater": [("src_file", "su_src_file"), ("src_sheet", "su_src_sheet"),
                          ("tgt_file", "su_tgt_file")],
        "formula_broadcast": [("file", "fb_file")],
        "unpivot": [("src_file", "up_src_file"), ("src_sheet", "up_src_sheet"),
                    ("tgt_file", "up_tgt_file")],
        "advance_vlookup": [("file", "adv_vlookup_file"), ("sheet", "adv_vlookup_sheet"), ("prefix", "adv_vlookup_prefix"),
                           ("src_col", "adv_vlookup_src_col"), ("lookup_file", "adv_vlookup_lookup_file"),
                           ("lookup_sheet", "adv_vlookup_lookup_sheet"), ("key_col", "adv_vlookup_key_col"),
                           ("return_col", "adv_vlookup_return_col"), ("output_col", "adv_vlookup_output_col")],
        "sumifs": [("file", "sumifs_file"), ("sheet", "sumifs_sheet"), ("src_workbook", "sumifs_src_file"),
                  ("src_sheet", "sumifs_src_sheet"), ("sum_range", "sumifs_sum_range"), ("output_col", "sumifs_output_col"),
                  ("start", "sumifs_start")],
        "append": [("src_file", "append_src_file"), ("src_sheet", "append_src_sheet"), 
                  ("tgt_file", "append_tgt_file"), ("tgt_sheet", "append_tgt_sheet")],
        "header_map": [("src_file", "header_map_src_file"), ("src_sheet", "header_map_src_sheet"),
                      ("tgt_file", "header_map_tgt_file"), ("tgt_sheet", "header_map_tgt_sheet"), ("mapping", "header_map_mapping")],
        "filter": [("file", "filter_file"), ("sheet", "filter_sheet"), ("column", "filter_col"),
                  ("value", "filter_value"), ("remove_empty", "filter_remove_empty"), ("header_row", "filter_header_row")],
        "insert": [("file", "insert_file"), ("sheet", "insert_sheet"), ("axis", "insert_axis"), ("start", "insert_start")],
        "delete": [("file", "delete_file"), ("sheet", "delete_sheet"), ("axis", "delete_axis"), ("start", "delete_start")],
        "merge_cells": [("file", "merge_cells_file"), ("sheet", "merge_cells_sheet"), ("range", "merge_cells_range")],
        "data_input": [("file", "data_input_file"), ("sheet", "data_input_sheet")],
        "quota_update": [],
        "group_rollup": [],
    }
    
    if step_type in key_maps:
        for cfg_key, ss_key in key_maps[step_type]:
            if cfg_key in cfg:
                st.session_state[ss_key] = cfg[cfg_key]
    
    # Special handling for import (has mapping list)
    if step_type == "import":
        _src_file   = cfg.get("src_file", "")
        _src_sheet  = cfg.get("src_sheet", "")
        _tgt_file   = cfg.get("tgt_file", "")
        _tgt_sheet  = cfg.get("tgt_sheet", "")
        _src_hrow   = int(cfg.get("src_header_row", 1) or 1)
        _tgt_hrow   = int(cfg.get("tgt_header_row", 1) or 1)

        st.session_state["import_src_file"]       = _src_file
        st.session_state["import_src_sheet"]      = _src_sheet
        st.session_state["import_tgt_file"]       = _tgt_file
        st.session_state["import_tgt_sheet"]      = _tgt_sheet
        st.session_state["import_src_header_row"] = _src_hrow
        st.session_state["import_tgt_header_row"] = _tgt_hrow
        st.session_state["import_append_mode"]    = cfg.get("append_mode", False)

        # ── Pre-set the ctx guard so the form does NOT wipe mappings on render ──
        # The form computes: ctx = (src_file, src_sheet, src_hrow, tgt_file, tgt_sheet, tgt_hrow)
        # If import_mapping_ctx already equals that tuple, the wipe branch is skipped.
        st.session_state["import_mapping_ctx"] = (
            _src_file, _src_sheet, _src_hrow,
            _tgt_file, _tgt_sheet, _tgt_hrow,
        )

        # ── Restore column mappings ───────────────────────────────────────────
        mapping = cfg.get("mapping", [])
        if isinstance(mapping, list):
            st.session_state["import_mappings"] = [
                (m[0], m[1]) if isinstance(m, (list, tuple)) else (m, m)
                for m in mapping
            ]
        else:
            st.session_state["import_manual_mapping"] = str(mapping)

        # ── Restore append mode & filename col ───────────────────────────────
        st.session_state["import_append_mode"]       = bool(cfg.get("append_mode", False))
        fn_col = cfg.get("filename_col") or ""
        st.session_state["import_add_filename_col"]  = bool(fn_col)
        st.session_state["import_filename_col_name"] = fn_col

        # ── Restore filter conditions ─────────────────────────────────────────
        # filter_builder uses keys: import_filter_fc, import_filter_combine,
        #   import_filter_cond_{i}_col / _op / _val
        from operations.filter_utils import OPERATORS as _FU_OPS
        _fu_op_labels = list(_FU_OPS.values())
        _fu_op_keys   = list(_FU_OPS.keys())
        _combine_opts = ["AND  (all must match)", "OR  (any can match)"]

        saved_filters  = cfg.get("filters") or []
        saved_combine  = (cfg.get("filter_combine") or "AND").upper()

        st.session_state["import_filter_fc"] = max(1, len(saved_filters))
        st.session_state["import_filter_combine"] = (
            _combine_opts[0] if saved_combine == "AND" else _combine_opts[1]
        )
        for _i, _f in enumerate(saved_filters):
            st.session_state[f"import_filter_cond_{_i}_col"] = _f.get("column", "")
            _op_k = _f.get("operator", "eq")
            st.session_state[f"import_filter_cond_{_i}_op"] = (
                _op_k if _op_k in _fu_op_keys else _fu_op_keys[0]
            )
            st.session_state[f"import_filter_cond_{_i}_val"] = str(_f.get("value", "") or "")
    
    # Special handling for sort (has sort_cols list)
    if step_type == "sort":
        st.session_state["sort_file"] = cfg.get("file", "")
        st.session_state["sort_sheet"] = cfg.get("sheet", "")
        st.session_state["sort_header_row"] = cfg.get("header_row", 1)
        cols = cfg.get("sort_cols", [])
        asc_list = cfg.get("sort_ascending", [True] * len(cols))
        st.session_state["sort_builder_keys"] = [{"col": c, "asc": a} for c, a in zip(cols, asc_list)]
    
    # Special handling for vlookup (has multiple columns)
    if step_type == "vlookup":
        st.session_state["vlex_lv_file"] = cfg.get("lookup_value_file", "")
        st.session_state["vlex_lv_sheet"] = cfg.get("lookup_value_sheet", "")
        st.session_state["vlex_lv_hdr"] = cfg.get("lookup_value_header_row", 1)
        lv_col = cfg.get("lookup_value_column", "")
        st.session_state["vlex_lv_cols"] = [lv_col] if isinstance(lv_col, str) else list(lv_col)
        
        st.session_state["vlex_s_file"] = cfg.get("search_file", "")
        st.session_state["vlex_s_sheet"] = cfg.get("search_sheet", "")
        st.session_state["vlex_s_hdr"] = cfg.get("search_header_row", 1)
        s_col = cfg.get("search_column", "")
        st.session_state["vlex_s_cols"] = [s_col] if isinstance(s_col, str) else list(s_col)
        
        st.session_state["vlex_r_file"] = cfg.get("return_file", "")
        st.session_state["vlex_r_sheet"] = cfg.get("return_sheet", "")
        st.session_state["vlex_r_hdr"] = cfg.get("return_header_row", 1)
        st.session_state["vlex_r_col"] = cfg.get("return_column", "")
        st.session_state["vlex_not_found"] = cfg.get("not_found_value", "Not found")
        
        st.session_state["vlex_d_file"] = cfg.get("return_to_file", "")
        st.session_state["vlex_d_sheet"] = cfg.get("return_to_sheet", "")
        st.session_state["vlex_d_hdr"] = cfg.get("return_to_header_row", 1)
        st.session_state["vlex_d_col"] = cfg.get("return_to_column", "")

        # Optional: fallback chain
        fb = (cfg.get("fallback_lookups") or [])
        st.session_state["vlex_chain_enabled"] = bool(fb)
        if fb:
            fb0 = fb[0] or {}
            st.session_state["vlex2_s_file"] = fb0.get("search_file", "")
            st.session_state["vlex2_s_sheet"] = fb0.get("search_sheet", "")
            st.session_state["vlex2_s_hdr"] = fb0.get("search_header_row", 1)
            fb_s_cols = fb0.get("search_columns") or fb0.get("search_column") or []
            if isinstance(fb_s_cols, str):
                st.session_state["vlex2_s_cols"] = [fb_s_cols]
            else:
                st.session_state["vlex2_s_cols"] = list(fb_s_cols)
            st.session_state["vlex2_r_file"] = fb0.get("return_file", "")
            st.session_state["vlex2_r_sheet"] = fb0.get("return_sheet", "")
            st.session_state["vlex2_r_hdr"] = fb0.get("return_header_row", 1)
            st.session_state["vlex2_r_col"] = fb0.get("return_column", "")

    # Special handling for quota_update (has key lists)
    if step_type == "quota_update":
        st.session_state["quota_src_file"] = cfg.get("source_file", "")
        st.session_state["quota_src_sheet"] = cfg.get("source_sheet", "")
        st.session_state["quota_src_hdr"] = cfg.get("source_header_row", 1)
        st.session_state["quota_filter_col"] = cfg.get("source_filter_column", "")
        st.session_state["quota_filter_op"] = cfg.get("source_filter_operator", "equals")
        st.session_state["quota_filter_val"] = cfg.get("source_filter_value", "")
        st.session_state["quota_src_keys"] = list(cfg.get("source_key_columns") or [])
        st.session_state["quota_src_norms"] = list(cfg.get("source_key_norms") or [])

        st.session_state["quota_tgt_file"] = cfg.get("target_file", "")
        st.session_state["quota_tgt_sheet"] = cfg.get("target_sheet", "")
        st.session_state["quota_tgt_hdr"] = cfg.get("target_header_row", 1)
        st.session_state["quota_tgt_keys"] = list(cfg.get("target_key_columns") or [])
        st.session_state["quota_tgt_norms"] = list(cfg.get("target_key_norms") or [])
        st.session_state["quota_write_col"] = cfg.get("write_column", "")
        st.session_state["quota_write_val"] = cfg.get("write_value", "")

    # Special handling for group_rollup (has key list)
    if step_type == "group_rollup":
        st.session_state["roll_file"] = cfg.get("file", "")
        st.session_state["roll_sheet"] = cfg.get("sheet", "")
        st.session_state["roll_hdr"] = cfg.get("header_row", 1)
        st.session_state["roll_type_col"] = cfg.get("type_column", "")
        st.session_state["roll_amount_col"] = cfg.get("amount_column", "")
        st.session_state["roll_from_type"] = cfg.get("from_type_value", "")
        st.session_state["roll_into_type"] = cfg.get("into_type_value", "")
        st.session_state["roll_include_from"] = bool(cfg.get("include_from_in_sum", True))
        st.session_state["roll_delete_dup_into"] = bool(cfg.get("delete_duplicate_into_rows", True))
        st.session_state["roll_delete_from"] = bool(cfg.get("delete_from_rows", False))
        st.session_state["roll_keys"] = list(cfg.get("key_columns") or [])

    # Special handling for copy_paste (supports multiple targets)
    if step_type == "copy_paste":
        tgts = cfg.get("tgt_cols")
        if not tgts:
            t = cfg.get("tgt_col")
            tgts = [t] if t else []
        if "copy_paste_tgt_cols"not in st.session_state:
            st.session_state.copy_paste_tgt_cols = []
        st.session_state.copy_paste_tgt_cols = list(tgts)

    # Special handling for delete (supports deleting specific columns)
    if step_type == "delete":
        if "delete_cols"not in st.session_state:
            st.session_state.delete_cols = []
        # Restore advanced delete state
        if cfg.get("row_mode") == "advanced_condition":
            from operations.filter_utils import OPERATORS as _ADV_OPS
            _adv_op_keys   = list(_ADV_OPS.keys())
            _adv_op_labels = list(_ADV_OPS.values())
            st.session_state["adv_del_scope"]        = cfg.get("sheet_scope", "single")
            st.session_state["adv_del_sheet"]        = cfg.get("sheet", "")
            st.session_state["adv_del_selected_sheets"] = cfg.get("selected_sheets") or []
            st.session_state["adv_del_include_sheets_raw"] = ", ".join(cfg.get("include_sheets") or [])
            st.session_state["adv_del_exclude_sheets_raw"] = ", ".join(cfg.get("exclude_sheets") or [])
            st.session_state["adv_del_header_row"]   = int(cfg.get("header_row", 1) or 1)
            st.session_state["adv_del_combine"]      = cfg.get("filter_combine", "AND")
            st.session_state["adv_del_save_log"]     = bool(cfg.get("save_log", True))
            st.session_state["adv_del_log_path"]     = cfg.get("log_path", "")
            _adv_filts = cfg.get("filters") or []
            st.session_state["adv_del_n_conds"] = max(1, len(_adv_filts))
            for _ai, _af in enumerate(_adv_filts):
                st.session_state[f"adv_del_cond_{_ai}_col"] = _af.get("column", "")
                _aop = _af.get("operator", "eq")
                st.session_state[f"adv_del_cond_{_ai}_op"] = _aop
                st.session_state[f"adv_del_cond_{_ai}_op_label"] = _adv_op_labels[_adv_op_keys.index(_aop)] if _aop in _adv_op_keys else _adv_op_labels[0]
                st.session_state[f"adv_del_cond_{_ai}_val"] = _af.get("value", "")
        if str(cfg.get("axis", "")).lower() == "column"and cfg.get("columns"):
            st.session_state["delete_header_row"] = int(cfg.get("header_row", 1) or 1)
            st.session_state.delete_cols = list(cfg.get("columns") or [])
            st.session_state["delete_col_action"] = (
                "Clear data under header (keep column)"
                if cfg.get("action") == "clear_data"
                else "Delete the column(s) بالكامل (remove column)"
            )
            if cfg.get("action") == "clear_data":
                st.session_state["delete_clear_end_last"] = str(cfg.get("end_row", "last")).lower() == "last"
                if not st.session_state["delete_clear_end_last"]:
                    try:
                        st.session_state["delete_clear_end_num"] = int(cfg.get("end_row"))
                    except Exception:
                        st.session_state["delete_clear_end_num"] = 100
    
    # Special handling for conditional
    if step_type == "conditional":
        from operations.conditional import OPERATORS
        _OP_LABELS = list(OPERATORS.values())
        _OP_KEYS   = list(OPERATORS.keys())
        st.session_state["conditional_file"]       = cfg.get("file", "")
        st.session_state["conditional_sheet"]      = cfg.get("sheet", "")
        st.session_state["conditional_col_mode"]   = cfg.get("col_mode", "letter")
        st.session_state["conditional_header_row"] = int(cfg.get("header_row", 1) or 1)
        # Scope
        _cscope = cfg.get("scope", "single")
        st.session_state["conditional_scope"] = _cscope
        st.session_state["conditional_include_sheets"] = ", ".join(cfg.get("include_sheets") or [])
        st.session_state["conditional_exclude_sheets"] = ", ".join(cfg.get("exclude_sheets") or [])

        # Normalise to rules list (handle both old single-rule and new multi-rule configs)
        rules = cfg.get("rules")
        if not rules:
            # Legacy single-rule config
            rules = [{
                "cond_col":   cfg.get("cond_col", ""),
                "operator":   cfg.get("operator", "equals"),
                "cond_val":   cfg.get("cond_val", ""),
                "target_col": cfg.get("target_col", ""),
                "write_val":  cfg.get("write_val", ""),
            }]

        st.session_state["cond_n_rules"] = max(1, len(rules))
        for i, r in enumerate(rules):
            st.session_state[f"cond_rule_{i}_cond_col"]      = r.get("cond_col", "")
            _op_key = r.get("operator", "equals")
            _op_lbl = _OP_LABELS[_OP_KEYS.index(_op_key)] if _op_key in _OP_KEYS else _OP_LABELS[0]
            st.session_state[f"cond_rule_{i}_operator_label"]= _op_lbl
            st.session_state[f"cond_rule_{i}_cond_val"]      = r.get("cond_val", "")
            st.session_state[f"cond_rule_{i}_target_col"]    = r.get("target_col", "")
            st.session_state[f"cond_rule_{i}_write_val"]     = r.get("write_val", "")

    # Special handling for pivot step
    if step_type == "pivot":
        st.session_state["pivot_src_file"]      = cfg.get("src_file", "")
        st.session_state["pivot_src_sheet"]     = cfg.get("src_sheet", "")
        st.session_state["pivot_src_header_row"]= int(cfg.get("src_header_row", 1) or 1)
        st.session_state["pivot_tgt_file"]      = cfg.get("tgt_file", "")
        st.session_state["pivot_tgt_sheet"]     = cfg.get("tgt_sheet", "Pivot")
        st.session_state["pivot_tgt_start_row"] = int(cfg.get("tgt_start_row", 1) or 1)
        st.session_state["pivot_tgt_start_col"] = cfg.get("tgt_start_col", "A")
        st.session_state["pivot_row_fields"]    = cfg.get("row_fields", [])
        st.session_state["pivot_col_fields"]    = cfg.get("col_fields", [])
        st.session_state["pivot_grand_rows"]    = bool(cfg.get("grand_total_rows", False))
        st.session_state["pivot_grand_cols"]    = bool(cfg.get("grand_total_cols", False))
        st.session_state["pivot_overwrite_sheet"]= bool(cfg.get("overwrite_sheet", True))
        st.session_state["pivot_apply_fmt"]     = bool(cfg.get("apply_formatting", True))
        st.session_state["pivot_sort_field"]    = cfg.get("sort_field", "")
        st.session_state["pivot_sort_asc"]      = bool(cfg.get("sort_ascending", True))
        _fv = cfg.get("fill_value", 0)
        st.session_state["pivot_fill_value_str"]= str(_fv)
        st.session_state["pivot_header_bg"]     = "#" + str(cfg.get("header_bg_color", "1F3864")).lstrip("#")
        st.session_state["pivot_header_fg"]     = "#" + str(cfg.get("header_font_color", "FFFFFF")).lstrip("#")
        st.session_state["pivot_total_bg"]      = "#" + str(cfg.get("total_bg_color", "D9E1F2")).lstrip("#")
        vfs = cfg.get("value_fields", [])
        st.session_state["pivot_n_vf"] = max(1, len(vfs))
        for _i, _vf in enumerate(vfs):
            st.session_state[f"pivot_vf_{_i}_col"]   = _vf.get("column", "")
            st.session_state[f"pivot_vf_{_i}_agg"]   = _vf.get("aggfunc", "sum")
            st.session_state[f"pivot_vf_{_i}_label"] = _vf.get("label", "")
        # Restore filter conditions
        _pvt_filters = cfg.get("filter_conditions") or []
        st.session_state["pivot_filter_fc"]      = max(1, len(_pvt_filters))
        st.session_state["pivot_filter_combine"] = (
            "AND  (all must match)" if cfg.get("filter_combine", "AND") == "AND"
            else "OR  (any can match)"
        )
        for _i, _f in enumerate(_pvt_filters):
            st.session_state[f"pivot_filter_cond_{_i}_col"] = _f.get("column", "")
            st.session_state[f"pivot_filter_cond_{_i}_op"]  = _f.get("operator", "eq")
            st.session_state[f"pivot_filter_cond_{_i}_val"] = str(_f.get("value", "") or "")

    # Special handling for format step
    if step_type == "format":
        st.session_state["format_file"]     = cfg.get("file", "")
        st.session_state["format_sheet"]    = cfg.get("sheet", "")
        st.session_state["format_mode_key"] = cfg.get("format_mode", "static")
        st.session_state["format_col_mode"] = cfg.get("col_mode", "letter")
        st.session_state["format_header_row"] = int(cfg.get("header_row", 1) or 1)

        fmt_mode = cfg.get("format_mode", "static")
        if fmt_mode == "static":
            rules = cfg.get("static_rules") or []
            pfx   = "fmt_s"
            st.session_state["format_n_rules"] = max(1, len(rules))
            for i, r in enumerate(rules):
                st.session_state[f"{pfx}_{i}_tgt_type"] = r.get("target_type", "row")
                st.session_state[f"{pfx}_{i}_tgt_spec"] = r.get("target_spec", "")
                _restore_style(pfx, i, r.get("style") or {})
        else:
            rules = cfg.get("cond_rules") or []
            pfx   = "fmt_c"
            st.session_state["format_n_rules"] = max(1, len(rules))
            for i, r in enumerate(rules):
                st.session_state[f"{pfx}_{i}_check_col"]   = r.get("check_col", "")
                st.session_state[f"{pfx}_{i}_operator"]    = r.get("operator", "gt")
                st.session_state[f"{pfx}_{i}_value"]       = str(r.get("value", "") or "")
                st.session_state[f"{pfx}_{i}_value2"]      = str(r.get("value2", "") or "")
                st.session_state[f"{pfx}_{i}_apply_to"]    = r.get("apply_to", "row")
                st.session_state[f"{pfx}_{i}_target_cols"] = r.get("target_cols", "")
                _restore_style(pfx, i, r.get("style") or {})

    # Special handling for folder_summary
    if step_type == "folder_summary":
        st.session_state["fs_src_file"] = cfg.get("src_file", "")
        st.session_state["fs_src_sheet"] = cfg.get("src_sheet", "")
        st.session_state["fs_src_hrow"] = int(cfg.get("header_row", 1) or 1)
        st.session_state["fs_source_mode"] = cfg.get("source_mode", "folder")
        st.session_state["fs_sheet_selection"] = cfg.get("sheet_selection", "all")
        st.session_state["fs_sheet_pattern"] = cfg.get("sheet_pattern", "")
        st.session_state["fs_sheet_list_raw"] = ", ".join(cfg.get("sheet_list") or [])
        st.session_state["fs_sheet_exclude"] = cfg.get("sheet_exclude", "")
        st.session_state["fs_tgt_file"] = cfg.get("tgt_file", "")
        st.session_state["fs_tgt_sheet"] = cfg.get("tgt_sheet", "Summary")
        st.session_state["fs_tgt_hrow"] = int(cfg.get("tgt_header_row", 1) or 1)
        st.session_state["fs_append_mode"] = bool(cfg.get("append_mode", True))
        st.session_state["fs_filename_col"] = cfg.get("filename_col", "File Name")
        st.session_state["fs_sheet_name_col"] = cfg.get("sheet_name_col", "")
        st.session_state["fs_incl_vars"] = bool(cfg.get("include_var_cols", True))
        st.session_state["fs_incl_status"] = bool(cfg.get("include_status_col", True))
        st.session_state["fs_status_col_name"] = cfg.get("status_col_name", "Status")
        variables = cfg.get("variables") or []
        formulas = cfg.get("formulas") or []
        st.session_state["fs_var_count"] = max(1, len(variables))
        st.session_state["fs_formula_count"] = max(1, len(formulas))
        for i, v in enumerate(variables):
            st.session_state[f"fs_var_{i}_name"] = v.get("name", "")
            st.session_state[f"fs_var_{i}_search_text"] = v.get("search_text", "")
            st.session_state[f"fs_var_{i}_label"] = v.get("label", "")
            st.session_state[f"fs_var_{i}_search_col"] = v.get("search_col", "A")
            st.session_state[f"fs_var_{i}_value_col"] = v.get("value_col", "D")
            st.session_state[f"fs_var_{i}_match_type"] = v.get("match_type", "icontains")
            st.session_state[f"fs_var_{i}_multi_match"] = v.get("multi_match", "first")
            st.session_state[f"fs_var_{i}_row_offset"] = int(v.get("row_offset", 0) or 0)
            st.session_state[f"fs_var_{i}_on_missing"] = v.get("on_missing", "zero")
            st.session_state[f"fs_var_{i}_default"] = str(v.get("missing_default", "") or "")
            st.session_state[f"fs_var_{i}_include"] = bool(v.get("include", True))
        for i, f in enumerate(formulas):
            st.session_state[f"fs_formula_{i}_name"] = f.get("name", "")
            st.session_state[f"fs_formula_{i}_expr"] = f.get("expression", "")
            st.session_state[f"fs_formula_{i}_on_error"] = f.get("on_error", "flag")

    # Special handling for split_export
    if step_type == "split_export":
        st.session_state["se_src_file"]        = cfg.get("src_file", "")
        st.session_state["se_src_sheet"]       = cfg.get("src_sheet", "")
        st.session_state["se_src_header_row"]  = int(cfg.get("src_header_row", 1) or 1)
        _mode = cfg.get("split_mode", "single")
        st.session_state["se_split_mode_idx"]  = SPLIT_MODE_KEYS.index(_mode) if _mode in SPLIT_MODE_KEYS else 0
        st.session_state["se_split_col"]       = cfg.get("split_col", "")
        st.session_state["se_split_col2"]      = cfg.get("split_col2", "")
        st.session_state["se_tgt_file"]        = cfg.get("tgt_file", "")
        st.session_state["se_tgt_sheet"]       = cfg.get("tgt_sheet", "Export")
        st.session_state["se_output_folder"]   = cfg.get("output_folder", "")
        st.session_state["se_file_tmpl"]       = cfg.get("file_name_template", "{value}.xlsx")
        st.session_state["se_sheet_tmpl"]      = cfg.get("sheet_name_template", "{value}")
        st.session_state["se_sheet_tmpl2"]     = cfg.get("sheet_name_template2", "{value}")
        st.session_state["se_include_cols"]    = ", ".join(cfg.get("include_columns") or [])
        st.session_state["se_exclude_cols"]    = ", ".join(cfg.get("exclude_columns") or [])
        st.session_state["se_include_header"]  = bool(cfg.get("include_header", True))
        st.session_state["se_overwrite"]       = bool(cfg.get("overwrite", True))
        st.session_state["se_max_rows"]        = int(cfg.get("max_rows_per_sheet", 0) or 0)
        st.session_state["se_add_summary"]     = bool(cfg.get("add_summary_sheet", False))
        st.session_state["se_summary_name"]    = cfg.get("summary_sheet_name", "_Summary")
        st.session_state["se_apply_fmt"]       = bool(cfg.get("apply_formatting", True))
        st.session_state["se_header_bg"]       = "#" + str(cfg.get("header_bg_color", "1F3864")).lstrip("#")
        st.session_state["se_header_fg"]       = "#" + str(cfg.get("header_font_color", "FFFFFF")).lstrip("#")
        # Sort columns
        sc = cfg.get("sort_cols") or []
        st.session_state["se_sort_cols"] = ", ".join(sc)
        _asc = cfg.get("sort_ascending", True)
        st.session_state["se_sort_asc"] = "Ascending" if _asc else "Descending"
        # Filters
        _se_filters = cfg.get("filter_conditions") or []
        st.session_state["se_filter_fc"]      = max(1, len(_se_filters))
        st.session_state["se_filter_combine"] = (
            "AND  (all must match)" if cfg.get("filter_combine", "AND") == "AND"
            else "OR  (any can match)"
        )
        for _i, _f in enumerate(_se_filters):
            st.session_state[f"se_filter_cond_{_i}_col"] = _f.get("column", "")
            st.session_state[f"se_filter_cond_{_i}_op"]  = _f.get("operator", "eq")
            st.session_state[f"se_filter_cond_{_i}_val"] = str(_f.get("value", "") or "")

    # Special handling for unpivot
    if step_type == "unpivot":
        st.session_state["up_src_file"]       = cfg.get("src_file", "")
        st.session_state["up_src_sheet"]      = cfg.get("src_sheet", "")
        st.session_state["up_src_header_row"] = int(cfg.get("src_header_row", 1) or 1)
        st.session_state["up_id_cols"]        = cfg.get("id_cols") or []
        st.session_state["up_value_cols"]     = cfg.get("value_cols") or []
        st.session_state["up_var_name"]       = cfg.get("var_name", "Category")
        st.session_state["up_value_name"]     = cfg.get("value_name", "Value")
        st.session_state["up_tgt_file"]       = cfg.get("tgt_file", "")
        st.session_state["up_tgt_sheet"]      = cfg.get("tgt_sheet", "Unpivot")
        st.session_state["up_append_mode"]    = bool(cfg.get("append_mode", False))
        st.session_state["up_drop_na"]        = bool(cfg.get("drop_na", True))

    # Special handling for formula_broadcast
    if step_type == "formula_broadcast":
        st.session_state["fb_scope"]          = cfg.get("scope", "all_sheets")
        st.session_state["fb_formula"]        = cfg.get("formula", "")
        st.session_state["fb_cell"]           = cfg.get("cell", "")
        st.session_state["fb_file"]           = cfg.get("file", "")
        st.session_state["fb_include_sheets"] = cfg.get("include_sheets", "")
        st.session_state["fb_exclude_sheets"] = cfg.get("exclude_sheets", "")
        st.session_state["fb_sheet_name"]     = cfg.get("sheet_name", "")

    # Special handling for write_cell (scope / include / exclude)
    if step_type == "write_cell":
        st.session_state["write_cell_scope"]          = cfg.get("scope", "single")
        st.session_state["write_cell_include_sheets"] = ", ".join(cfg.get("include_sheets") or [])
        st.session_state["write_cell_exclude_sheets"] = ", ".join(cfg.get("exclude_sheets") or [])
        st.session_state["write_cell_ref"]            = cfg.get("cell_ref", "A1")
        st.session_state["write_cell_column"]         = cfg.get("column", "A")

    # Special handling for sheet_row_inserter
    if step_type == "sheet_row_inserter":
        st.session_state["sri_src_file"]           = cfg.get("src_file", "")
        st.session_state["sri_src_sheet"]          = cfg.get("src_sheet", "")
        st.session_state["sri_col_category"]       = cfg.get("src_col_category", "")
        st.session_state["sri_col_particulars"]    = cfg.get("src_col_particulars", "")
        st.session_state["sri_col_value"]          = cfg.get("src_col_value", "")
        st.session_state["sri_tgt_file"]           = cfg.get("tgt_file", "")
        _sm = cfg.get("sheet_match_mode", "cat_in_sheet")
        st.session_state["sri_sheet_match_idx"]    = SHEET_MATCH_KEYS.index(_sm) if _sm in SHEET_MATCH_KEYS else 0
        st.session_state["sri_anchor_col"]         = cfg.get("anchor_col", "")
        st.session_state["sri_anchor_mode"]        = cfg.get("anchor_mode", "particulars")
        st.session_state["sri_anchor_value"]       = cfg.get("anchor_value", "")
        _am = cfg.get("anchor_match", "exact")
        st.session_state["sri_anchor_match_idx"]   = ANCHOR_MATCH_KEYS.index(_am) if _am in ANCHOR_MATCH_KEYS else 0
        _ip = cfg.get("insert_position", "after_last")
        st.session_state["sri_insert_pos_idx"]     = INSERT_POSITION_KEYS.index(_ip) if _ip in INSERT_POSITION_KEYS else 2
        st.session_state["sri_tgt_col_particulars"] = cfg.get("tgt_col_particulars", "")
        st.session_state["sri_tgt_col_value"]      = cfg.get("tgt_col_value", "")
        st.session_state["sri_tgt_header_row"]     = int(cfg.get("tgt_header_row", 1) or 1)
        _ecm = cfg.get("extra_col_map") or []
        st.session_state["sri_extra_map_count"]    = max(len(_ecm), 0)
        for _i, _m in enumerate(_ecm):
            st.session_state[f"sri_extra_{_i}_src"] = _m.get("src_col", "")
            st.session_state[f"sri_extra_{_i}_tgt"] = _m.get("tgt_col", "")

    # Special handling for sheet_updater
    if step_type == "sheet_updater":
        st.session_state["su_src_file"]           = cfg.get("src_file", "")
        st.session_state["su_src_sheet"]          = cfg.get("src_sheet", "")
        st.session_state["su_col_category"]       = cfg.get("src_col_category", "")
        st.session_state["su_col_particulars"]    = cfg.get("src_col_particulars", "")
        st.session_state["su_col_value"]          = cfg.get("src_col_value", "")
        st.session_state["su_tgt_file"]           = cfg.get("tgt_file", "")
        _sm = cfg.get("sheet_match_mode", "cat_in_sheet")
        st.session_state["su_sheet_match_idx"]    = SHEET_MATCH_KEYS.index(_sm) if _sm in SHEET_MATCH_KEYS else 0
        st.session_state["su_tgt_col_lookup"]     = cfg.get("tgt_col_lookup", "")
        st.session_state["su_tgt_col_write"]      = cfg.get("tgt_col_write", "")
        st.session_state["su_tgt_header_row"]     = int(cfg.get("tgt_header_row", 1) or 1)
        st.session_state["su_case_sensitive"]     = bool(cfg.get("case_sensitive", False))
        # Particulars matching
        _pm = cfg.get("particulars_match_mode", "iexact")
        st.session_state["su_part_match_idx"]     = PARTICULARS_MATCH_KEYS.index(_pm) if _pm in PARTICULARS_MATCH_KEYS else 0
        _rules = cfg.get("particulars_rules") or []
        st.session_state["su_advanced_mode"]      = bool(_rules)
        st.session_state["su_part_rules_count"]   = max(len(_rules), 1)
        for _i, _r in enumerate(_rules):
            st.session_state[f"su_part_rule_{_i}_particular"]  = _r.get("particular", "")
            _rm = _r.get("match_mode", "iexact")
            st.session_state[f"su_part_rule_{_i}_match_mode"]  = PARTICULARS_MATCH_KEYS.index(_rm) if _rm in PARTICULARS_MATCH_KEYS else 0

    # Special handling for add_files
    if step_type == "add_files":
        files = cfg.get("files", [])
        paths = [f.get("path", "") for f in files if f.get("path")]
        st.session_state["add_files_paths"] = "\n".join(paths)
        st.session_state["add_files_confirm"] = True  # Already confirmed before


def render_add_steps_tab():
    """Render the Add Steps tab with all 18 operations"""
    
    # Check if we're in edit mode
    editing_idx = st.session_state.get("editing_step_idx")
    editing_data = st.session_state.get("editing_step_data")
    
    if editing_idx is not None and editing_data:
        st.markdown(f'<h2 style="color:#1a1a1a;font-family:IBM Plex Sans,sans-serif;text-transform:uppercase;letter-spacing:.1em;font-size:13px;font-weight:600;border-bottom:2px solid #c8392b;padding-bottom:4px;margin-bottom:12px;">Editing Step {editing_idx + 1}</h2>', unsafe_allow_html=True)
        st.warning(f"Editing: {editing_data.get('name', 'Unknown step')} — Make changes and click Save Changes")
        c1, c2 = st.columns([1, 3])
        with c1:
            if st.button("Cancel Edit", key="cancel_visual_edit", use_container_width=True):
                st.session_state.editing_step_idx = None
                st.session_state.editing_step_data = None
                st.session_state.pop("adding_step", None)
                st.rerun()
    else:
        st.markdown('<h2 style="color:#1a1a1a;font-family:IBM Plex Sans,sans-serif;text-transform:uppercase;letter-spacing:.1em;font-size:13px;font-weight:600;border-bottom:2px solid #c8392b;padding-bottom:4px;margin-bottom:12px;">Add Operations to Workflow</h2>', unsafe_allow_html=True)

    # Persistent feedback when a step is added (survives reruns)
    if st.session_state.get("_last_added_step_name"):
        st.success(f"Added step: {st.session_state.get('_last_added_step_name')}")
        st.session_state["_last_added_step_name"] = ""

    # Insertion hint (set from Workflow tab "Insert after"buttons)
    if st.session_state.get("_workflow_notice") and editing_idx is None:
        st.info(st.session_state.get("_workflow_notice"))
        c1, c2 = st.columns([1, 2])
        with c1:
            if st.button("Clear insert target", key="clear_insert_target", use_container_width=True):
                st.session_state.insert_at = None
                st.session_state._workflow_notice = ""
                st.rerun()
        with c2:
            st.caption("If you leave this set, the **next** step you add will be inserted at that position (not appended).")

    st.caption(f"Current workflow steps: {len(st.session_state.workflow_steps)}")
    
    col1, col2, col3, col4 = st.columns(4)

    # Column 1: Data Transformation
    with col1:
        st.markdown('<div class="pkf-section">Transform</div>', unsafe_allow_html=True)
        if st.button("Formula", use_container_width=True):
            st.session_state.adding_step = "formula"
        if st.button("Broadcast Formula", use_container_width=True,
                     help="Write a formula into one fixed cell across all sheets in a workbook, or across all files in a group"):
            st.session_state.adding_step = "formula_broadcast"
        if st.button("Copy/Paste", use_container_width=True):
            st.session_state.adding_step = "copy_paste"
        if st.button("Convert Values", use_container_width=True):
            st.session_state.adding_step = "convert_values"
        if st.button("Forward Fill", use_container_width=True):
            st.session_state.adding_step = "forward_fill"
        if st.button("Replace", use_container_width=True):
            st.session_state.adding_step = "replace"
        if st.button("Conditional", use_container_width=True):
            st.session_state.adding_step = "conditional"
        if st.button("Write Cell(s)", use_container_width=True):
            st.session_state.adding_step = "write_cell"
    
    # Column 2: Data Merging
    with col2:
        st.markdown('<div class="pkf-section">Merge / Lookup</div>', unsafe_allow_html=True)
        if st.button("VLOOKUP", use_container_width=True):
            st.session_state.adding_step = "vlookup"
        if st.button("Advanced VLOOKUP", use_container_width=True):
            st.session_state.adding_step = "advance_vlookup"
        if st.button("Quota Update (Generic)", use_container_width=True):
            st.session_state.adding_step = "quota_update"
        if st.button("Group Rollup (Generic)", use_container_width=True):
            st.session_state.adding_step = "group_rollup"
        if st.button("SUMIFS", use_container_width=True):
            st.session_state.adding_step = "sumifs"
        if st.button("Append", use_container_width=True):
            st.session_state.adding_step = "append"
        if st.button("Import", use_container_width=True):
            st.session_state.adding_step = "import"
        if st.button("Header Map", use_container_width=True):
            st.session_state.adding_step = "header_map"
    
    # Column 3: Structural
    with col3:
        st.markdown('<div class="pkf-section">Structure</div>', unsafe_allow_html=True)
        if st.button("Filter", use_container_width=True):
            st.session_state.adding_step = "filter"
        if st.button("Sort", use_container_width=True):
            st.session_state.adding_step = "sort"
        if st.button("Insert R/C", use_container_width=True):
            st.session_state.adding_step = "insert"
        if st.button("Delete R/C", use_container_width=True):
            st.session_state.adding_step = "delete"
        if st.button("Merge Cells", use_container_width=True):
            st.session_state.adding_step = "merge_cells"
        if st.button("Delete Sheets", use_container_width=True,
                     help="Remove sheets from a workbook by name, pattern, keep-only, or delete empty sheets"):
            st.session_state.adding_step = "delete_sheets"
        if st.button("Format & Beautify", use_container_width=True,
                     help="Apply bold, colors, borders, conditional highlights to any range or column"):
            st.session_state.adding_step = "format"
        if st.button("Pivot Table", use_container_width=True,
                     help="Build a pivot table from source data — sums/counts/averages by row and column groupings"):
            st.session_state.adding_step = "pivot"
        if st.button("Unpivot / Wide→Long", use_container_width=True,
                     help="Convert wide-format table (columns = categories) into long/database format (one row per category) using pandas melt"):
            st.session_state.adding_step = "unpivot"
        if st.button("Split Export", use_container_width=True,
                     help="Slice data into one file/multiple sheets, multiple files, or two-level file+sheet splits by column values"):
            st.session_state.adding_step = "split_export"
        if st.button("Sheet Updater", use_container_width=True,
                     help="Push values from a master (category/particulars/value) sheet into matching sheets of a target workbook"):
            st.session_state.adding_step = "sheet_updater"
        if st.button("Sheet Row Inserter", use_container_width=True,
                     help="Insert new rows at a specific position inside a multi-sheet workbook — finds an anchor row and inserts above/below it"):
            st.session_state.adding_step = "sheet_row_inserter"
    
    # Column 4: Data Input
    with col4:
        st.markdown('<div class="pkf-section">Input</div>', unsafe_allow_html=True)
        if st.button("Input Source", use_container_width=True, help="Flexible file collector: single path, folder scan, filename rules, or dynamic date patterns"):
            st.session_state.adding_step = "input_source"
            st.session_state.pop("is_classes", None)  # reset form state
        if st.button("Normalize & Import", use_container_width=True, help="Map input headers to standard names and append to working file"):
            st.session_state.adding_step = "normalize_import"
        if st.button("Folder Summary", use_container_width=True, help="Extract labeled values from many Excel files and consolidate into one summary row per file"):
            st.session_state.adding_step = "folder_summary"
        if st.button("Add File(s)", use_container_width=True):
            st.session_state.adding_step = "add_files"
        if st.button("Data Input", use_container_width=True):
            st.session_state.adding_step = "data_input"
    
    st.divider()
    
    # Show form for selected operation
    if 'adding_step'in st.session_state:
        step_type = st.session_state.adding_step
        
        if step_type == "formula":
            render_formula_form()
        elif step_type == "formula_broadcast":
            render_formula_broadcast_form()
        elif step_type == "copy_paste":
            render_copy_paste_form()
        elif step_type == "convert_values":
            render_convert_values_form()
        elif step_type == "forward_fill":
            render_forward_fill_form()
        elif step_type == "replace":
            render_replace_form()
        elif step_type == "conditional":
            render_conditional_form()
        elif step_type == "write_cell":
            render_write_cell_form()
        elif step_type == "vlookup":
            render_vlookup_form()
        elif step_type == "advance_vlookup":
            render_advance_vlookup_form()
        elif step_type == "quota_update":
            render_quota_update_form()
        elif step_type == "group_rollup":
            render_group_rollup_form()
        elif step_type == "sumifs":
            render_sumifs_form()
        elif step_type == "append":
            render_append_form()
        elif step_type == "import":
            render_import_form()
        elif step_type == "header_map":
            render_header_map_form()
        elif step_type == "filter":
            render_filter_form()
        elif step_type == "sort":
            render_sort_form()
        elif step_type == "insert":
            render_insert_form()
        elif step_type == "delete":
            render_delete_form()
        elif step_type == "merge_cells":
            render_merge_cells_form()
        elif step_type == "delete_sheets":
            render_delete_sheets_form()
        elif step_type == "format":
            render_format_form()
        elif step_type == "pivot":
            render_pivot_form()
        elif step_type == "split_export":
            render_split_export_form()
        elif step_type == "sheet_updater":
            render_sheet_updater_form()
        elif step_type == "sheet_row_inserter":
            render_sheet_row_inserter_form()
        elif step_type == "unpivot":
            render_unpivot_form()
        elif step_type == "input_source":
            render_input_source_form()
        elif step_type == "normalize_import":
            render_normalize_import_form()
        elif step_type == "folder_summary":
            render_folder_summary_form()
        elif step_type == "add_files":
            render_add_files_form()
        elif step_type == "data_input":
            render_data_input_form()


# =============================================================================
# FORM RENDERERS FOR ALL 18 OPERATIONS
# =============================================================================

def render_formula_form():
    st.markdown('<div class="pkf-section">Formula Operation</div>', unsafe_allow_html=True)
    file = st.selectbox("File", get_files(), key="formula_file", help="Register files via **Add File(s)** step first.")
    sheet = sheet_selectbox("Sheet", file, key="formula_sheet")
    formula = st.text_input("Formula", "=A2+B2", key="formula_expr")

    # Warn if formula uses Excel 365-only functions (TEXTBEFORE, TEXTAFTER, etc.)
    _MODERN_FUNCS = ["TEXTBEFORE", "TEXTAFTER", "TEXTJOIN", "XLOOKUP", "UNIQUE",
                     "SORT", "SORTBY", "FILTER", "SEQUENCE", "RANDARRAY", "XMATCH", "LET", "LAMBDA"]
    _formula_upper = (formula or "").upper()
    _found_modern = [f for f in _MODERN_FUNCS if re.search(rf'\b{f}\s*\(', _formula_upper)]
    if _found_modern:
        st.warning(
            f" **Excel version compatibility:** This formula uses "
            f"**{', '.join(_found_modern)}** — available only in Excel 365 (v2206+).  \n"
            f"Older Excel versions will show `#NAME?` and Excel may add `@` before the function name.  \n"
            f" **Fix:** Check **Convert to Values** below — the app will compute the result "
            f"in Python and write the final value directly, so no Excel formula is needed."
        )

    # Target column — choose by letter OR by header name
    col_mode = st.radio(
        "Target column specified by",
        ["Column letter (e.g. C)", "Header name"],
        key="formula_col_mode",
        horizontal=True,
    )

    header_row_for_col = 1  # default, only meaningful in header-name mode
    if col_mode == "Column letter (e.g. C)":
        column = st.text_input("Column Letter", "C", key="formula_col")
    else:
        header_row_for_col = st.number_input(
            "Header row", min_value=1, value=1, step=1, key="formula_col_header_row",
            help="Which row contains column headers? Used to build the dropdown below."
        )
        col_headers = get_headers_for_file_sheet(file, sheet, int(header_row_for_col))
        if col_headers:
            column = st.selectbox(
                "Target column header", options=col_headers, key="formula_col_header_name"
            )
        else:
            st.warning("No headers found — make sure the file is registered and the header row is correct.")
            column = st.text_input("Target column header (type manually)", "", key="formula_col_header_name")

    start = st.number_input("Start Row", 1, value=2, key="formula_start")
    end = end_row_input("End Row", key_prefix="formula_end", default_last=True, default_number=10)
    convert = st.checkbox("Convert to Values", key="formula_convert")
    export = export_options_ui("formula")

    if st.button("Add Step", key="formula_add"):
        if not column:
            st.warning("Please specify a target column.")
            return
        use_header = col_mode != "Column letter (e.g. C)"
        resolved_col = column  # default: already a letter
        if use_header:
            # Resolve header name → column letter right now using openpyxl
            file_path_now = get_file_path(file)
            if file_path_now:
                try:
                    import openpyxl as _ox
                    from openpyxl.utils import get_column_letter as _gcl
                    _wb = _ox.load_workbook(file_path_now, read_only=True, data_only=True)
                    _ws = _wb[sheet] if sheet and sheet in _wb.sheetnames else _wb.active
                    _hrow = int(header_row_for_col)
                    for _c in range(1, _ws.max_column + 1):
                        _v = _ws.cell(row=_hrow, column=_c).value
                        if _v is not None and str(_v).strip() == str(column).strip():
                            resolved_col = _gcl(_c)
                            break
                    _wb.close()
                except Exception as _e:
                    st.warning(f"Could not resolve header to column letter: {_e}. Storing header name as fallback.")
            if resolved_col == column:
                st.warning(f"Header '{column}'not found in row {header_row_for_col} — stored as-is. Verify header row is correct.")
        add_step("formula", f"Formula: {formula} → {column} ({resolved_col})"if use_header else f"Formula: {formula} → {column}", {
            "file": file, "sheet": sheet, "formula": formula,
            "column": resolved_col,
            "start": start, "end": end, "convert": convert,
            "export": export,
        })


def render_formula_broadcast_form():
    """Broadcast Formula — write a formula to a fixed cell across all sheets or all files."""
    st.markdown('<div class="pkf-section">Broadcast Formula</div>', unsafe_allow_html=True)
    st.caption(
        "Write **one formula** into **one fixed cell** everywhere — across every sheet in a workbook, "
        "or across the same sheet in every file in a group."
    )

    # ── Scope ─────────────────────────────────────────────────────────────────
    _scope_opts   = ["All Sheets in one Workbook", "One Sheet across All Files"]
    _scope_keys   = ["all_sheets", "all_files"]
    _saved_scope  = st.session_state.get("fb_scope", "all_sheets")
    _scope_idx    = _scope_keys.index(_saved_scope) if _saved_scope in _scope_keys else 0
    scope_label   = st.selectbox(
        "Scope",
        _scope_opts,
        index=_scope_idx,
        key="fb_scope_label",
        help="**All Sheets**: same cell in every sheet of one file.  "
             "**All Files**: same cell+sheet in every file registered under a file-group label.",
    )
    scope = _scope_keys[_scope_opts.index(scope_label)]

    st.divider()

    # ── Formula + Cell ────────────────────────────────────────────────────────
    st.markdown("**What to write**")
    fc1, fc2 = st.columns([4, 1])
    with fc1:
        fb_formula = st.text_input(
            "Formula",
            value=st.session_state.get("fb_formula", ""),
            key="fb_formula",
            placeholder="e.g.  =SUM(A1:A9)  or  =TOTAL",
        )
    with fc2:
        fb_cell = st.text_input(
            "Target Cell",
            value=st.session_state.get("fb_cell", ""),
            key="fb_cell",
            placeholder="e.g.  A10",
        )

    st.divider()

    # ── File selection ────────────────────────────────────────────────────────
    if scope == "all_sheets":
        st.markdown("**Workbook**")
        fb_file = st.selectbox("File", get_files(), key="fb_file",
                               help="The workbook whose sheets will all receive the formula")

        st.markdown("**Sheet Filters  *(optional)***")
        fc3, fc4 = st.columns(2)
        with fc3:
            fb_include = st.text_input(
                "Include only these sheets  *(comma-separated, empty = all)*",
                value=st.session_state.get("fb_include_sheets", ""),
                key="fb_include_sheets",
                placeholder="e.g.  Jan, Feb, Mar",
            )
        with fc4:
            fb_exclude = st.text_input(
                "Exclude these sheets  *(comma-separated)*",
                value=st.session_state.get("fb_exclude_sheets", ""),
                key="fb_exclude_sheets",
                placeholder="e.g.  Summary, Template",
            )
        fb_sheet_name = ""

    else:  # all_files
        st.markdown("**File Group**")
        fb_file = st.selectbox(
            "File / Group Label",
            get_files(),
            key="fb_file",
            help="Select a file label — if it was collected as a group (folder scan), "
                 "the formula is written into every file in the group",
        )
        fb_sheet_name = st.text_input(
            "Sheet Name  *(must exist in each file)*",
            value=st.session_state.get("fb_sheet_name", ""),
            key="fb_sheet_name",
            placeholder="e.g.  Data  or  Sheet1",
        )
        fb_include = ""
        fb_exclude = ""

    export = st.checkbox("Export after run", value=True, key="fb_export")

    # ── Save step ─────────────────────────────────────────────────────────────
    if st.button("➕ Add Broadcast Formula Step", type="primary", use_container_width=True):
        if not fb_formula or not fb_formula.strip():
            st.error("Please enter a formula.")
            return
        if not fb_cell or not fb_cell.strip():
            st.error("Please enter a target cell (e.g. A10).")
            return
        if not fb_file:
            st.error("Please select a file.")
            return
        if scope == "all_files" and not fb_sheet_name.strip():
            st.error("Please enter the sheet name for 'All Files' mode.")
            return

        _scope_short = "all-sheets" if scope == "all_sheets" else "all-files"
        if scope == "all_sheets":
            _label = f"Broadcast Formula  [{_scope_short}]  {fb_formula.strip()} → {fb_cell.strip().upper()}"
        else:
            _label = f"Broadcast Formula  [{_scope_short}]  {fb_formula.strip()} → {fb_sheet_name}!{fb_cell.strip().upper()}"

        add_step("formula_broadcast", _label, {
            "scope":          scope,
            "formula":        fb_formula.strip(),
            "cell":           fb_cell.strip().upper(),
            "file":           fb_file,
            "sheet_name":     fb_sheet_name.strip(),
            "include_sheets": fb_include.strip(),
            "exclude_sheets": fb_exclude.strip(),
            "export":         export,
        })


def render_copy_paste_form():
    st.markdown('<div class="pkf-section">Copy/Paste Operation</div>', unsafe_allow_html=True)
    file = st.selectbox("File", get_files(), key="copy_paste_file")
    sheet = sheet_selectbox("Sheet", file, key="copy_paste_sheet")
    paste_type = st.selectbox("Type", ["All", "Values Only"], key="copy_paste_type")
    header_row = st.number_input("Header Row", 1, value=1, key="copy_paste_header")
    export = export_options_ui("copy_paste")

    # Reset target list when context changes
    ctx = (file or "", sheet or "", int(header_row))
    if st.session_state.get("copy_paste_ctx") != ctx:
        st.session_state.copy_paste_ctx = ctx
        st.session_state.copy_paste_tgt_cols = []

    if "copy_paste_tgt_cols"not in st.session_state:
        st.session_state.copy_paste_tgt_cols = []

    headers = get_headers_for_file_sheet(file, sheet, int(header_row))

    # Source column
    st.markdown('<div class="pkf-section">Source</div>', unsafe_allow_html=True)
    if headers:
        src_col = st.selectbox("Source Column (header)", options=[""] + headers, key="copy_paste_src")
        st.caption("Tip: you can select a header name; letter inputs also work if you prefer.")
    else:
        src_col = st.text_input("Source Column (letter or header)", "A", key="copy_paste_src")

    # Target columns (addable)
    st.markdown('<div class="pkf-section">Targets (paste into 1 or more columns)</div>', unsafe_allow_html=True)
    if headers:
        c1, c2 = st.columns([5, 1])
        with c1:
            tgt_to_add = st.selectbox("Add Target Column", options=[""] + headers, key="copy_paste_tgt_add")
        with c2:
            st.write("")
            if st.button("", key="copy_paste_tgt_add_btn", disabled=not tgt_to_add, use_container_width=True):
                if tgt_to_add not in st.session_state.copy_paste_tgt_cols:
                    st.session_state.copy_paste_tgt_cols.append(tgt_to_add)
                    st.rerun()
    else:
        c1, c2 = st.columns([5, 1])
        with c1:
            tgt_to_add = st.text_input("Add Target Column (letter or header)", "B", key="copy_paste_tgt_add_text")
        with c2:
            st.write("")
            if st.button("", key="copy_paste_tgt_add_btn_text", disabled=not tgt_to_add, use_container_width=True):
                tgt_to_add = str(tgt_to_add).strip()
                if tgt_to_add and tgt_to_add not in st.session_state.copy_paste_tgt_cols:
                    st.session_state.copy_paste_tgt_cols.append(tgt_to_add)
                    st.rerun()

    if st.session_state.copy_paste_tgt_cols:
        st.markdown("**Target columns:**")
        for idx, col in enumerate(st.session_state.copy_paste_tgt_cols):
            r1, r2 = st.columns([8, 1])
            with r1:
                st.text(f" {idx+1}. {col}")
            with r2:
                if st.button("", key=f"copy_paste_tgt_del_{idx}", use_container_width=True):
                    st.session_state.copy_paste_tgt_cols.pop(idx)
                    st.rerun()
    else:
        st.warning("Add at least one Target Column.")

    if st.button("Add Step", key="copy_paste_add"):
        tgt_cols = st.session_state.get("copy_paste_tgt_cols", []) or []
        if not src_col:
            st.warning("Please select a Source Column.")
            return
        if not tgt_cols:
            st.warning("Please add at least one Target Column.")
            return

        tgt_desc = ", ".join(tgt_cols)
        cfg = {
            "file": file, "sheet": sheet, "src_col": src_col,
            "paste_type": paste_type, "header_row": header_row,
            "export": export,
        }
        # Backward compatible: keep tgt_col if single, otherwise store tgt_cols
        if len(tgt_cols) == 1:
            cfg["tgt_col"] = tgt_cols[0]
        else:
            cfg["tgt_cols"] = tgt_cols

        add_step("copy_paste", f"Copy {src_col} → {tgt_desc}", cfg)


def render_convert_values_form():
    st.markdown('<div class="pkf-section">Convert to Values</div>', unsafe_allow_html=True)
    file = st.selectbox("File", get_files(), key="convert_values_file")
    sheet = sheet_selectbox("Sheet", file, key="convert_values_sheet")
    column = st.text_input("Column", "C", key="convert_values_col")
    start = st.number_input("Start Row", 1, value=2, key="convert_values_start")
    end = end_row_input("End Row", key_prefix="convert_values_end", default_last=True, default_number=10)
    export = export_options_ui("convert_values")

    if st.button("Add Step", key="convert_values_add"):
        add_step("convert_values", f"Convert {column} to Values", {
            "file": file, "sheet": sheet, "column": column,
            "start": start, "end": end,
            "export": export,
        })


def render_forward_fill_form():
    st.markdown('<div class="pkf-section">Forward Fill</div>', unsafe_allow_html=True)
    file = st.selectbox("File", get_files(), key="forward_fill_file")
    sheet = sheet_selectbox("Sheet", file, key="forward_fill_sheet")
    column = st.text_input("Column", "A", key="forward_fill_col")
    export = export_options_ui("forward_fill")

    if st.button("Add Step", key="forward_fill_add"):
        add_step("forward_fill", f"Fill {column}", {
            "file": file, "sheet": sheet, "column": column,
            "export": export,
        })


def render_replace_form():
    st.markdown('<div class="pkf-section">Replace Values</div>', unsafe_allow_html=True)
    file = st.selectbox("File", get_files(), key="replace_file")
    sheet = sheet_selectbox("Sheet", file, key="replace_sheet")
    old_val = st.text_input("Old Value", key="replace_old")
    new_val = st.text_input("New Value", key="replace_new")
    column = st.text_input("Column (blank = all)", "", key="replace_col")
    export = export_options_ui("replace")

    if st.button("Add Step", key="replace_add"):
        add_step("replace", f"Replace {old_val} → {new_val}", {
            "file": file, "sheet": sheet, "old_val": old_val,
            "new_val": new_val, "column": column,
            "export": export,
        })


def render_conditional_form():
    from operations.conditional import OPERATORS, NO_VALUE_OPERATORS

    OPERATOR_LABELS = list(OPERATORS.values())
    OPERATOR_KEYS   = list(OPERATORS.keys())

    st.markdown('<div class="pkf-section">Conditional Write</div>', unsafe_allow_html=True)
    st.caption(
        "Define one or more rules. Each rule checks a condition column and writes a "
        "value into a target column for every matching row — all applied in a single pass."
    )

    file = st.selectbox("File", get_files(), key="conditional_file")

    # ── Scope: single sheet vs all sheets ────────────────────────────────────
    SCOPE_OPTS = ["Single Sheet", "All Sheets in Workbook"]
    SCOPE_KEYS = ["single", "all_sheets"]
    _saved_scope = st.session_state.get("conditional_scope", "single")
    _saved_scope_i = SCOPE_KEYS.index(_saved_scope) if _saved_scope in SCOPE_KEYS else 0
    _scope_label = st.radio(
        "Apply to",
        SCOPE_OPTS, index=_saved_scope_i,
        key="conditional_scope_radio", horizontal=True,
        help=(
            "**Single Sheet** — apply rules to one sheet (classic behaviour).\n\n"
            "**All Sheets** — run the same rules on every sheet in the workbook. "
            "Optionally restrict which sheets are processed with include/exclude filters."
        ),
    )
    cond_scope = SCOPE_KEYS[SCOPE_OPTS.index(_scope_label)]

    if cond_scope == "single":
        sheet = sheet_selectbox("Sheet", file, key="conditional_sheet")
        cond_include_sheets = []
        cond_exclude_sheets = []
    else:
        sheet = None  # not used for all_sheets
        st.caption(
            "All sheets will be processed. "
            "Optionally enter comma-separated sheet names to **include** (only those) "
            "or **exclude** (skip those)."
        )
        _inc_raw = st.text_input(
            "Include only these sheets *(comma-separated, leave blank for all)*",
            value=st.session_state.get("conditional_include_sheets", ""),
            key="conditional_include_sheets",
            placeholder="e.g. Sheet1, Sheet2",
        )
        _exc_raw = st.text_input(
            "Exclude these sheets *(comma-separated)*",
            value=st.session_state.get("conditional_exclude_sheets", ""),
            key="conditional_exclude_sheets",
            placeholder="e.g. Summary, Template",
        )
        cond_include_sheets = [s.strip() for s in _inc_raw.split(",") if s.strip()]
        cond_exclude_sheets = [s.strip() for s in _exc_raw.split(",") if s.strip()]

    st.divider()

    # ── Column identification mode (shared across all rules) ──────────────────
    COL_MODES     = ["Column Letter / Number", "Column Header Name"]
    COL_MODE_KEYS = ["letter", "header"]
    _saved_mode   = st.session_state.get("conditional_col_mode", "letter")
    _saved_mode_i = COL_MODE_KEYS.index(_saved_mode) if _saved_mode in COL_MODE_KEYS else 0

    col_mode_label = st.radio(
        "Identify columns by",
        COL_MODES, index=_saved_mode_i,
        key="conditional_col_mode_radio", horizontal=True,
        help=(
            "**Letter / Number** — A, B, C … or 1, 2, 3 …  \n"
            "**Header Name** — exact column header text (case-insensitive)."
        ),
    )
    col_mode = COL_MODE_KEYS[COL_MODES.index(col_mode_label)]

    if col_mode == "header":
        header_row = st.number_input(
            "Header row (1-based)",
            min_value=1,
            value=int(st.session_state.get("conditional_header_row", 1) or 1),
            key="conditional_header_row",
        )
    else:
        header_row = 1

    _ph_col = "Header name" if col_mode == "header" else "Col (A / 1)"
    _ph_val = "Value  (blank = empty)"

    st.divider()

    # ── Rule count management ─────────────────────────────────────────────────
    if "cond_n_rules" not in st.session_state:
        st.session_state["cond_n_rules"] = 1

    btn_a, btn_r = st.columns([1, 1])
    with btn_a:
        if st.button("Add Rule", key="cond_add_rule"):
            st.session_state["cond_n_rules"] += 1
    with btn_r:
        if st.button("Remove Last Rule", key="cond_rem_rule") and st.session_state["cond_n_rules"] > 1:
            st.session_state["cond_n_rules"] -= 1

    n_rules = st.session_state["cond_n_rules"]

    # ── Rule table header ─────────────────────────────────────────────────────
    hc = st.columns([2, 2, 2, 2, 2, 1])
    for lbl, col in zip(
        ["Condition Column", "Operator", "Condition Value", "Target Column", "Write Value", ""],
        hc
    ):
        col.markdown(f"**{lbl}**")

    # ── Rule rows ─────────────────────────────────────────────────────────────
    for i in range(n_rules):
        rc = st.columns([2, 2, 2, 2, 2, 1])

        rc[0].text_input(
            f"Cond Col {i+1}", key=f"cond_rule_{i}_cond_col",
            placeholder=_ph_col, label_visibility="collapsed",
        )

        _saved_op_i = OPERATOR_KEYS.index(
            st.session_state.get(f"cond_rule_{i}_operator", "equals")
        ) if st.session_state.get(f"cond_rule_{i}_operator", "equals") in OPERATOR_KEYS else 0
        rc[1].selectbox(
            f"Operator {i+1}", OPERATOR_LABELS, index=_saved_op_i,
            key=f"cond_rule_{i}_operator_label", label_visibility="collapsed",
        )

        _op_key = OPERATOR_KEYS[OPERATOR_LABELS.index(
            st.session_state.get(f"cond_rule_{i}_operator_label", OPERATOR_LABELS[0])
        )]
        _needs_val = _op_key not in NO_VALUE_OPERATORS
        rc[2].text_input(
            f"Cond Val {i+1}", key=f"cond_rule_{i}_cond_val",
            placeholder=_ph_val if _needs_val else "(not needed)",
            label_visibility="collapsed",
            disabled=not _needs_val,
        )

        rc[3].text_input(
            f"Target Col {i+1}", key=f"cond_rule_{i}_target_col",
            placeholder=_ph_col, label_visibility="collapsed",
        )

        rc[4].text_input(
            f"Write Val {i+1}", key=f"cond_rule_{i}_write_val",
            placeholder="Value to write", label_visibility="collapsed",
        )

        # Per-rule row count preview
        if rc[5].button("?", key=f"cond_prev_{i}", help="Preview matching rows for this rule"):
            fp_prev = get_file_path(file)
            if fp_prev and sheet:
                try:
                    from operations.conditional import _resolve_column, _build_mask
                    from core import wb_cache as _wbc
                    _df = _wbc.read_excel(fp_prev, sheet_name=sheet, header=int(header_row) - 1)
                    _df.columns = [str(c).strip() for c in _df.columns]
                    _cc = st.session_state.get(f"cond_rule_{i}_cond_col", "")
                    _op = OPERATOR_KEYS[OPERATOR_LABELS.index(
                        st.session_state.get(f"cond_rule_{i}_operator_label", OPERATOR_LABELS[0])
                    )]
                    _cv = st.session_state.get(f"cond_rule_{i}_cond_val", "") or ""
                    if _cc:
                        _col_r = _resolve_column(_df, _cc, col_mode, int(header_row))
                        _mask  = _build_mask(_df[_col_r], _op, _cv)
                        st.info(f"Rule {i+1}: {int(_mask.sum())} / {len(_df)} rows match.")
                    else:
                        st.warning(f"Rule {i+1}: specify a condition column first.")
                except Exception as _ex:
                    st.warning(f"Rule {i+1} preview error: {_ex}")

    st.divider()
    export = export_options_ui("conditional")

    if st.button("Add Step", key="conditional_add", type="primary"):
        # Collect rules
        rules = []
        for i in range(n_rules):
            _cc  = str(st.session_state.get(f"cond_rule_{i}_cond_col",  "") or "").strip()
            _tc  = str(st.session_state.get(f"cond_rule_{i}_target_col","") or "").strip()
            _op  = OPERATOR_KEYS[OPERATOR_LABELS.index(
                st.session_state.get(f"cond_rule_{i}_operator_label", OPERATOR_LABELS[0])
            )]
            _cv  = str(st.session_state.get(f"cond_rule_{i}_cond_val",  "") or "").strip()
            _wv  = st.session_state.get(f"cond_rule_{i}_write_val", "") or ""
            if not _cc and not _tc:
                continue   # skip completely empty rows silently
            if not _cc:
                st.error(f"Rule {i+1}: condition column is required.")
                return
            if not _tc:
                st.error(f"Rule {i+1}: target column is required.")
                return
            rules.append({
                "cond_col":   _cc,
                "operator":   _op,
                "cond_val":   _cv,
                "target_col": _tc,
                "write_val":  _wv,
            })

        if not rules:
            st.error("Add at least one rule.")
            return

        # Build human-readable step label
        _sheet_label = sheet if cond_scope == "single" else "all sheets"
        if len(rules) == 1:
            r = rules[0]
            op_d   = OPERATORS.get(r["operator"], r["operator"])
            val_p  = "" if r["operator"] in NO_VALUE_OPERATORS else f" '{r['cond_val']}'"
            label  = f"If [{r['cond_col']}] {op_d}{val_p}  →  [{r['target_col']}] = '{r['write_val']}' ({_sheet_label})"
        else:
            label = f"Conditional Write — {len(rules)} rules on {_sheet_label}"

        add_step("conditional", label, {
            "file":            file,
            "sheet":           sheet or "",
            "col_mode":        col_mode,
            "header_row":      int(header_row),
            "rules":           rules,
            "scope":           cond_scope,
            "include_sheets":  cond_include_sheets,
            "exclude_sheets":  cond_exclude_sheets,
            "export":          export,
        })


def _render_style_editor(prefix: str):
    """
    Render a compact style editor and return a style dict.
    prefix is used as a unique key namespace.
    """
    from operations.formatter import BORDER_STYLE_OPTS, BORDER_SIDE_OPTS, ALIGN_OPTS

    st.markdown("**Style**")
    c1, c2, c3, c4 = st.columns(4)
    bold      = c1.checkbox("Bold",        key=f"{prefix}_bold",      value=st.session_state.get(f"{prefix}_bold", False))
    italic    = c2.checkbox("Italic",      key=f"{prefix}_italic",    value=st.session_state.get(f"{prefix}_italic", False))
    underline = c3.checkbox("Underline",   key=f"{prefix}_underline", value=st.session_state.get(f"{prefix}_underline", False))
    strike    = c4.checkbox("Strike",      key=f"{prefix}_strike",    value=st.session_state.get(f"{prefix}_strike", False))

    sc1, sc2, sc3 = st.columns([1, 2, 2])
    font_size  = sc1.number_input("Font Size", min_value=0, max_value=72,
                                   value=int(st.session_state.get(f"{prefix}_font_size", 0) or 0),
                                   key=f"{prefix}_font_size",
                                   help="0 = keep existing size")
    font_color = sc2.color_picker("Font Color",
                                   value=st.session_state.get(f"{prefix}_font_color", "#000000"),
                                   key=f"{prefix}_font_color")
    bg_color   = sc3.color_picker("Background Color",
                                   value=st.session_state.get(f"{prefix}_bg_color", "#FFFFFF"),
                                   key=f"{prefix}_bg_color")

    bc1, bc2, bc3 = st.columns([2, 2, 2])
    border_style_keys   = list(BORDER_STYLE_OPTS.keys())
    _saved_bs = st.session_state.get(f"{prefix}_border_style", "none")
    _bs_i     = border_style_keys.index(_saved_bs) if _saved_bs in border_style_keys else 0
    border_style = bc1.selectbox("Border Style", border_style_keys, index=_bs_i,
                                  key=f"{prefix}_border_style")

    _saved_bsd = st.session_state.get(f"{prefix}_border_sides", "all")
    _bsd_i     = BORDER_SIDE_OPTS.index(_saved_bsd) if _saved_bsd in BORDER_SIDE_OPTS else 0
    border_sides = bc2.selectbox("Border Sides",
                                  ["all", "outer", "top", "bottom", "left", "right", "top_bottom"],
                                  index=_bsd_i,
                                  key=f"{prefix}_border_sides")

    align_keys   = ["", "left", "center", "right", "justify"]
    _saved_al    = st.session_state.get(f"{prefix}_h_align", "")
    _al_i        = align_keys.index(_saved_al) if _saved_al in align_keys else 0
    h_align   = bc3.selectbox("Alignment", align_keys, index=_al_i,
                               key=f"{prefix}_h_align",
                               format_func=lambda x: x.title() if x else "(keep existing)")

    wrap_text = st.checkbox("Wrap Text", key=f"{prefix}_wrap_text",
                             value=st.session_state.get(f"{prefix}_wrap_text", False))

    # Build style dict — only include non-default values so we don't clobber existing styles
    style = {}
    if bold:      style["bold"]      = True
    if italic:    style["italic"]    = True
    if underline: style["underline"] = True
    if strike:    style["strike"]    = True
    if font_size and font_size > 0:
        style["font_size"] = font_size
    # Colors: only include if user picked something other than the neutral defaults
    if font_color and font_color.upper() not in ("#000000", "#000000FF"):
        style["font_color"] = font_color.lstrip("#")
    if bg_color and bg_color.upper() not in ("#FFFFFF", "#FFFFFFFF"):
        style["bg_color"] = bg_color.lstrip("#")
    if border_style and border_style != "none":
        style["border_style"] = border_style
        style["border_sides"] = border_sides
    if h_align:
        style["h_align"] = h_align
    if wrap_text:
        style["wrap_text"] = True

    return style


def render_format_form():
    from operations.formatter import COND_OPERATORS, COND_NO_VALUE, COND_TWO_VALUE, APPLY_TO_OPTS

    COND_OP_KEYS   = list(COND_OPERATORS.keys())
    COND_OP_LABELS = list(COND_OPERATORS.values())
    APPLY_KEYS     = list(APPLY_TO_OPTS.keys())
    APPLY_LABELS   = list(APPLY_TO_OPTS.values())

    st.markdown('<div class="pkf-section">Format & Beautify</div>', unsafe_allow_html=True)
    st.caption(
        "**Static** mode: apply styles to fixed rows, columns, or cell ranges (headers, templates).  "
        "**Conditional** mode: apply styles based on data values — e.g., red fill when amount < 0."
    )

    file  = st.selectbox("File",  get_files(), key="format_file")
    sheet = sheet_selectbox("Sheet", file, key="format_sheet")

    st.divider()

    # ── Mode ─────────────────────────────────────────────────────────────────
    MODE_OPTS = ["Static (fixed ranges)", "Conditional (based on data)"]
    MODE_KEYS = ["static", "conditional"]
    _saved_mode = st.session_state.get("format_mode_key", "static")
    _mode_i     = MODE_KEYS.index(_saved_mode) if _saved_mode in MODE_KEYS else 0
    mode_label  = st.radio("Mode", MODE_OPTS, index=_mode_i,
                            key="format_mode_radio", horizontal=True)
    format_mode = MODE_KEYS[MODE_OPTS.index(mode_label)]

    # ── Column mode + header row (conditional only) ──────────────────────────
    if format_mode == "conditional":
        CM_OPTS = ["Column Letter / Number", "Column Header Name"]
        CM_KEYS = ["letter", "header"]
        _saved_cm = st.session_state.get("format_col_mode", "letter")
        _cm_i = CM_KEYS.index(_saved_cm) if _saved_cm in CM_KEYS else 0
        col_mode_lbl = st.radio("Identify columns by", CM_OPTS, index=_cm_i,
                                 key="format_col_mode_radio", horizontal=True)
        col_mode   = CM_KEYS[CM_OPTS.index(col_mode_lbl)]
        header_row = st.number_input("Header row (1-based)", min_value=1,
                                      value=int(st.session_state.get("format_header_row", 1) or 1),
                                      key="format_header_row")
    else:
        col_mode   = "letter"
        header_row = 1

    st.divider()

    # ── Rule count ────────────────────────────────────────────────────────────
    if "format_n_rules" not in st.session_state:
        st.session_state["format_n_rules"] = 1

    ra, rb = st.columns([1, 1])
    with ra:
        if st.button("Add Rule", key="fmt_add_rule"):
            st.session_state["format_n_rules"] += 1
    with rb:
        if st.button("Remove Last Rule", key="fmt_rem_rule") and st.session_state["format_n_rules"] > 1:
            st.session_state["format_n_rules"] -= 1

    n_rules = st.session_state["format_n_rules"]

    # ── Rule builder ──────────────────────────────────────────────────────────
    all_rules = []

    if format_mode == "static":
        TGT_OPTS = ["Row(s)", "Column(s)", "Cell Range"]
        TGT_KEYS = ["row", "col", "range"]

        for i in range(n_rules):
            with st.expander(f"Rule {i + 1}", expanded=(i == 0)):
                t1, t2 = st.columns([1, 2])
                _saved_tt = st.session_state.get(f"fmt_s_{i}_tgt_type", "row")
                _tt_i = TGT_KEYS.index(_saved_tt) if _saved_tt in TGT_KEYS else 0
                tgt_type_lbl = t1.selectbox(
                    "Target Type", TGT_OPTS, index=_tt_i,
                    key=f"fmt_s_{i}_tgt_type_lbl",
                )
                tgt_type = TGT_KEYS[TGT_OPTS.index(tgt_type_lbl)]

                _ph = {
                    "row":   "e.g.  1   or   1,2,3",
                    "col":   "e.g.  A   or   A,C,F",
                    "range": "e.g.  A1:Z5",
                }[tgt_type]

                tgt_spec = t2.text_input(
                    "Target Spec", placeholder=_ph,
                    value=st.session_state.get(f"fmt_s_{i}_tgt_spec", ""),
                    key=f"fmt_s_{i}_tgt_spec",
                )

                style = _render_style_editor(f"fmt_s_{i}")
                all_rules.append({
                    "target_type": tgt_type,
                    "target_spec": tgt_spec,
                    "style":       style,
                })

    else:  # conditional
        _ph_col = "Header name" if col_mode == "header" else "Col (A / 1)"

        for i in range(n_rules):
            with st.expander(f"Rule {i + 1}", expanded=(i == 0)):
                r1c1, r1c2, r1c3 = st.columns([2, 2, 2])

                check_col = r1c1.text_input(
                    "Check Column", placeholder=_ph_col,
                    value=st.session_state.get(f"fmt_c_{i}_check_col", ""),
                    key=f"fmt_c_{i}_check_col",
                )

                _saved_op  = st.session_state.get(f"fmt_c_{i}_operator", "gt")
                _op_i      = COND_OP_KEYS.index(_saved_op) if _saved_op in COND_OP_KEYS else 0
                op_lbl     = r1c2.selectbox(
                    "Operator", COND_OP_LABELS, index=_op_i,
                    key=f"fmt_c_{i}_op_lbl",
                )
                op_key = COND_OP_KEYS[COND_OP_LABELS.index(op_lbl)]

                needs_val  = op_key not in COND_NO_VALUE
                needs_val2 = op_key in COND_TWO_VALUE

                value = r1c3.text_input(
                    "Value", placeholder="Comparison value" if needs_val else "(not needed)",
                    value=st.session_state.get(f"fmt_c_{i}_value", ""),
                    key=f"fmt_c_{i}_value",
                    disabled=not needs_val,
                )

                if needs_val2:
                    value2 = st.text_input(
                        "Value 2 (upper bound for between)",
                        value=st.session_state.get(f"fmt_c_{i}_value2", ""),
                        key=f"fmt_c_{i}_value2",
                        placeholder="Upper bound",
                    )
                else:
                    value2 = st.session_state.get(f"fmt_c_{i}_value2", "")

                r2c1, r2c2 = st.columns([2, 3])
                _saved_at  = st.session_state.get(f"fmt_c_{i}_apply_to", "row")
                _at_i      = APPLY_KEYS.index(_saved_at) if _saved_at in APPLY_KEYS else 0
                at_lbl     = r2c1.selectbox(
                    "Apply Formatting To", APPLY_LABELS, index=_at_i,
                    key=f"fmt_c_{i}_apply_to_lbl",
                )
                apply_to = APPLY_KEYS[APPLY_LABELS.index(at_lbl)]

                target_cols = ""
                if apply_to == "cols":
                    target_cols = r2c2.text_input(
                        "Target Columns (comma-sep)",
                        placeholder="e.g.  A,C,F  or header names",
                        value=st.session_state.get(f"fmt_c_{i}_target_cols", ""),
                        key=f"fmt_c_{i}_target_cols",
                    )

                style = _render_style_editor(f"fmt_c_{i}")
                all_rules.append({
                    "check_col":   check_col,
                    "operator":    op_key,
                    "value":       value,
                    "value2":      value2,
                    "apply_to":    apply_to,
                    "target_cols": target_cols,
                    "style":       style,
                })

    st.divider()
    export = export_options_ui("format")

    if st.button("Add Step", key="format_add", type="primary"):
        # Validate
        if not file:
            st.error("Select a file.")
            return
        if not sheet:
            st.error("Select a sheet.")
            return

        valid_rules = []
        if format_mode == "static":
            for idx, r in enumerate(all_rules, 1):
                if not r.get("target_spec", "").strip():
                    st.error(f"Rule {idx}: target spec is required.")
                    return
                if not r.get("style"):
                    st.warning(f"Rule {idx}: no style selected — rule will have no visual effect.")
                valid_rules.append(r)
        else:
            for idx, r in enumerate(all_rules, 1):
                if not r.get("check_col", "").strip():
                    st.error(f"Rule {idx}: check column is required.")
                    return
                if not r.get("style"):
                    st.warning(f"Rule {idx}: no style selected — rule will have no visual effect.")
                valid_rules.append(r)

        if not valid_rules:
            st.error("Add at least one rule.")
            return

        # Build human-readable label
        if format_mode == "static":
            if len(valid_rules) == 1:
                r = valid_rules[0]
                label = f"Format {r['target_type'].upper()} {r['target_spec']} on {sheet}"
            else:
                label = f"Format {len(valid_rules)} static rules on {sheet}"
        else:
            if len(valid_rules) == 1:
                r = valid_rules[0]
                op_lbl = COND_OPERATORS.get(r["operator"], r["operator"])
                val_s  = "" if r["operator"] in COND_NO_VALUE else f" {r['value']}"
                label  = f"Format: [{r['check_col']}] {op_lbl}{val_s} → {APPLY_TO_OPTS.get(r['apply_to'], r['apply_to'])} on {sheet}"
            else:
                label = f"Conditional Format {len(valid_rules)} rules on {sheet}"

        cfg_key = "static_rules" if format_mode == "static" else "cond_rules"
        add_step("format", label, {
            "file":        file,
            "sheet":       sheet,
            "format_mode": format_mode,
            "col_mode":    col_mode,
            "header_row":  int(header_row),
            cfg_key:       valid_rules,
            "export":      export,
        })


def render_pivot_form():
    """Full Pivot Table Builder form."""

    st.markdown('<div class="pkf-section">Pivot Table Builder</div>', unsafe_allow_html=True)
    st.caption(
        "Configure row groupings, column groupings, and value aggregations.  "
        "The pivot result is written as structured data to an output sheet — "
        "re-runnable, template-saveable, and works on any size data."
    )

    # ── Template load ─────────────────────────────────────────────────────
    saved_templates = list_pivot_templates()
    if saved_templates:
        with st.expander("Load Saved Template", expanded=False):
            t1, t2 = st.columns([3, 1])
            tpl_choice = t1.selectbox("Template", [""] + saved_templates, key="pivot_tpl_load_choice",
                                       label_visibility="collapsed")
            if t2.button("Load", key="pivot_tpl_load_btn") and tpl_choice:
                ok, tpl_cfg = load_pivot_template(tpl_choice)
                if ok:
                    _pivot_apply_template(tpl_cfg)
                    st.success(f"Template '{tpl_choice}' loaded.")
                    st.rerun()
                else:
                    st.error(tpl_cfg)

    st.divider()

    # ── Source ───────────────────────────────────────────────────────────
    st.markdown("**Source Data**")
    p1, p2, p3 = st.columns([3, 3, 1])
    with p1:
        src_file = st.selectbox("Source File", get_files(), key="pivot_src_file")
    with p2:
        src_sheet = sheet_selectbox("Source Sheet", src_file, key="pivot_src_sheet")
    with p3:
        src_hrow = st.number_input("Header Row", min_value=1,
                                    value=int(st.session_state.get("pivot_src_header_row", 1) or 1),
                                    key="pivot_src_header_row")

    # Auto-detect columns when source changes
    detect_ctx = (src_file or "", src_sheet or "", int(src_hrow))
    if st.session_state.get("pivot_detect_ctx") != detect_ctx:
        st.session_state["pivot_detect_ctx"] = detect_ctx
        st.session_state["pivot_col_info"] = {}
        if src_file and src_sheet:
            fp = get_file_path(src_file)
            if fp:
                st.session_state["pivot_col_info"] = detect_columns(fp, src_sheet, int(src_hrow))

    col_info: dict = st.session_state.get("pivot_col_info", {})
    all_cols    = list(col_info.keys())
    num_cols    = [c for c, m in col_info.items() if m["is_numeric"]]
    text_cols   = [c for c, m in col_info.items() if not m["is_numeric"]]

    if col_info:
        with st.expander(f"Detected Columns  ({len(all_cols)} total — "
                         f"{len(num_cols)} numeric, {len(text_cols)} text/categorical)", expanded=False):
            ci1, ci2 = st.columns(2)
            ci1.markdown("**Numeric** (suggested for values)")
            for c in num_cols:
                m = col_info[c]
                ci1.markdown(f"- `{c}` — {m['n_unique']} unique  ·  sample: {', '.join(m['sample'])}")
            ci2.markdown("**Text / Categorical** (suggested for rows/columns)")
            for c in text_cols:
                m = col_info[c]
                ci2.markdown(f"- `{c}` — {m['n_unique']} unique  ·  sample: {', '.join(m['sample'])}")

    st.divider()

    # ── Row / Column fields ───────────────────────────────────────────────
    st.markdown("**Field Configuration**")
    fc1, fc2 = st.columns(2)

    # Row fields (multiselect)
    _saved_rf = st.session_state.get("pivot_row_fields", [])
    _rf_valid = [v for v in _saved_rf if v in all_cols] if all_cols else _saved_rf
    row_fields = fc1.multiselect(
        "Row Fields  (group rows by these columns)",
        options=all_cols or _rf_valid,
        default=_rf_valid,
        key="pivot_row_fields",
        help="These become the rows of your pivot table. Multi-select for nested groupings.",
    )

    # Column pivot fields (optional multiselect)
    _saved_cf = st.session_state.get("pivot_col_fields", [])
    _cf_valid = [v for v in _saved_cf if v in all_cols and v not in row_fields] if all_cols else _saved_cf
    col_pivot_opts = [c for c in all_cols if c not in row_fields]
    col_fields = fc2.multiselect(
        "Column Fields  (optional — splits values across columns)",
        options=col_pivot_opts or _cf_valid,
        default=_cf_valid,
        key="pivot_col_fields",
        help="Each unique value in this column becomes a separate column in the pivot. Best for low-cardinality fields like Month, Quarter, Category.",
    )

    st.divider()

    # ── Value fields (dynamic list) ───────────────────────────────────────
    st.markdown("**Value Fields & Aggregations**")

    if "pivot_n_vf" not in st.session_state:
        st.session_state["pivot_n_vf"] = 1

    va, vb = st.columns([1, 1])
    with va:
        if st.button("Add Value Field", key="pivot_add_vf"):
            st.session_state["pivot_n_vf"] += 1
    with vb:
        if st.button("Remove Last", key="pivot_rem_vf") and st.session_state["pivot_n_vf"] > 1:
            st.session_state["pivot_n_vf"] -= 1

    n_vf = st.session_state["pivot_n_vf"]
    vf_cols = st.columns([3, 2, 3, 1])
    for lbl, col in zip(["Column", "Aggregation", "Output Label (optional)", ""], vf_cols):
        col.markdown(f"**{lbl}**")

    value_fields_list = []
    for i in range(n_vf):
        vc = st.columns([3, 2, 3, 1])

        _vf_col_opts = (num_cols if num_cols else all_cols) or [""]
        _saved_vfc = st.session_state.get(f"pivot_vf_{i}_col", "")
        _vfc_i = _vf_col_opts.index(_saved_vfc) if _saved_vfc in _vf_col_opts else 0
        if _vf_col_opts:
            chosen_col = vc[0].selectbox(
                f"VF Col {i}", _vf_col_opts, index=_vfc_i,
                key=f"pivot_vf_{i}_col", label_visibility="collapsed",
            )
        else:
            chosen_col = vc[0].text_input(
                f"VF Col {i}", value=_saved_vfc,
                key=f"pivot_vf_{i}_col", placeholder="Column name",
                label_visibility="collapsed",
            )

        agg_labels = [AGGFUNC_LABELS[k] for k in AGGFUNC_KEYS]
        _saved_agg = st.session_state.get(f"pivot_vf_{i}_agg", "sum")
        _agg_i = AGGFUNC_KEYS.index(_saved_agg) if _saved_agg in AGGFUNC_KEYS else 0
        chosen_agg_lbl = vc[1].selectbox(
            f"VF Agg {i}", agg_labels, index=_agg_i,
            key=f"pivot_vf_{i}_agg_lbl", label_visibility="collapsed",
        )
        chosen_agg = AGGFUNC_KEYS[agg_labels.index(chosen_agg_lbl)]

        chosen_label = vc[2].text_input(
            f"VF Label {i}",
            value=st.session_state.get(f"pivot_vf_{i}_label", ""),
            key=f"pivot_vf_{i}_label",
            placeholder=f"{chosen_col} ({chosen_agg_lbl})" if chosen_col else "e.g. Total Sales",
            label_visibility="collapsed",
        )

        if vc[3].button("", key=f"pivot_vf_{i}_del", help="Remove this field"):
            st.session_state["pivot_n_vf"] = max(1, n_vf - 1)
            for sfx in ("_col", "_agg", "_agg_lbl", "_label"):
                st.session_state.pop(f"pivot_vf_{i}{sfx}", None)
            st.rerun()

        value_fields_list.append({
            "column":  chosen_col,
            "aggfunc": chosen_agg,
            "label":   chosen_label.strip(),
        })

    st.divider()

    # ── Filters (optional) ────────────────────────────────────────────────
    with st.expander("Pre-Pivot Filters  (optional — filter source rows before pivoting)", expanded=False):
        pivot_filters, pivot_combine = render_filter_builder("pivot_filter", all_cols)
        if pivot_filters:
            st.info(f"**Active:** {len(pivot_filters)} filter condition(s) · combine: **{pivot_combine}**")

    st.divider()

    # ── Output settings ───────────────────────────────────────────────────
    st.markdown("**Output Settings**")
    o1, o2 = st.columns(2)

    tgt_file = o1.selectbox("Target File", get_files(),
                             index=_file_index(st.session_state.get("pivot_tgt_file", src_file)),
                             key="pivot_tgt_file")
    tgt_sheet_val = o2.text_input("Target Sheet Name",
                                   value=st.session_state.get("pivot_tgt_sheet", "Pivot"),
                                   key="pivot_tgt_sheet",
                                   placeholder="e.g. Pivot, Summary")

    oc1, oc2, oc3 = st.columns([1, 1, 2])
    tgt_start_row = oc1.number_input("Start Row", min_value=1,
                                      value=int(st.session_state.get("pivot_tgt_start_row", 1) or 1),
                                      key="pivot_tgt_start_row")
    tgt_start_col = oc2.text_input("Start Column", value=st.session_state.get("pivot_tgt_start_col", "A"),
                                    key="pivot_tgt_start_col", placeholder="A")
    fill_value_str = oc3.text_input("Fill empty cells with",
                                     value=str(st.session_state.get("pivot_fill_value", "0")),
                                     key="pivot_fill_value_str",
                                     help="Value for missing intersections (default 0)")

    ob1, ob2, ob3, ob4 = st.columns(4)
    grand_rows    = ob1.checkbox("Grand Total Row",    key="pivot_grand_rows",
                                  value=bool(st.session_state.get("pivot_grand_rows", False)))
    grand_cols    = ob2.checkbox("Grand Total Column", key="pivot_grand_cols",
                                  value=bool(st.session_state.get("pivot_grand_cols", False)))
    overwrite_sht = ob3.checkbox("Overwrite Sheet",    key="pivot_overwrite_sheet",
                                  value=bool(st.session_state.get("pivot_overwrite_sheet", True)),
                                  help="Recreate the target sheet from scratch (recommended)")
    apply_fmt     = ob4.checkbox("Apply Formatting",   key="pivot_apply_fmt",
                                  value=bool(st.session_state.get("pivot_apply_fmt", True)),
                                  help="Bold headers, colored totals, auto column widths")

    # Sort options
    sort_opts = [""] + (list(pivot_df_preview_cols(row_fields, col_fields, value_fields_list)) or [])
    _saved_sf = st.session_state.get("pivot_sort_field", "")
    _sf_i     = sort_opts.index(_saved_sf) if _saved_sf in sort_opts else 0
    so1, so2 = st.columns([3, 1])
    sort_field_val = so1.selectbox("Sort output by", sort_opts, index=_sf_i,
                                    key="pivot_sort_field",
                                    format_func=lambda x: x if x else "(no sort)")
    sort_asc = so2.checkbox("Ascending", key="pivot_sort_asc",
                             value=bool(st.session_state.get("pivot_sort_asc", True)))

    # Formatting colors (shown when apply_fmt is checked)
    if apply_fmt:
        fc1c, fc2c, fc3c = st.columns(3)
        header_bg  = fc1c.color_picker("Header Background",
                                        value=st.session_state.get("pivot_header_bg", "#1F3864"),
                                        key="pivot_header_bg")
        header_fg  = fc2c.color_picker("Header Font",
                                        value=st.session_state.get("pivot_header_fg", "#FFFFFF"),
                                        key="pivot_header_fg")
        total_bg   = fc3c.color_picker("Total Row Color",
                                        value=st.session_state.get("pivot_total_bg", "#D9E1F2"),
                                        key="pivot_total_bg")
    else:
        header_bg = st.session_state.get("pivot_header_bg", "#1F3864")
        header_fg = st.session_state.get("pivot_header_fg", "#FFFFFF")
        total_bg  = st.session_state.get("pivot_total_bg", "#D9E1F2")

    st.divider()

    # ── Template save ─────────────────────────────────────────────────────
    with st.expander("Save as Template  (reuse this pivot config on future data)", expanded=False):
        ts1, ts2 = st.columns([3, 1])
        tpl_save_name = ts1.text_input("Template Name", key="pivot_tpl_save_name",
                                        placeholder="e.g. Monthly Sales Pivot")
        if ts2.button("Save Template", key="pivot_tpl_save_btn"):
            if tpl_save_name.strip():
                _tpl_cfg = {
                    "row_fields": row_fields,
                    "col_fields": col_fields,
                    "value_fields": value_fields_list,
                    "tgt_sheet": tgt_sheet_val,
                    "tgt_start_row": int(tgt_start_row),
                    "tgt_start_col": tgt_start_col,
                    "grand_total_rows": grand_rows,
                    "grand_total_cols": grand_cols,
                    "sort_field": sort_field_val,
                    "sort_ascending": sort_asc,
                    "apply_formatting": apply_fmt,
                    "header_bg_color": header_bg.lstrip("#"),
                    "header_font_color": header_fg.lstrip("#"),
                    "total_bg_color": total_bg.lstrip("#"),
                }
                ok, msg = save_pivot_template(tpl_save_name.strip(), _tpl_cfg)
                st.success(msg) if ok else st.error(msg)
            else:
                st.warning("Enter a template name.")

        # Delete saved templates
        if saved_templates:
            dt1, dt2 = st.columns([3, 1])
            del_tpl = dt1.selectbox("Delete template", [""] + saved_templates, key="pivot_tpl_del_choice",
                                     label_visibility="collapsed")
            if dt2.button("Delete", key="pivot_tpl_del_btn") and del_tpl:
                ok, msg = delete_pivot_template(del_tpl)
                st.success(msg) if ok else st.error(msg)
                st.rerun()

    export = export_options_ui("pivot")

    # ── Add Step ──────────────────────────────────────────────────────────
    if st.button("Add Step", key="pivot_add", type="primary"):
        if not src_file:
            st.error("Select a source file.")
            return
        if not src_sheet:
            st.error("Select a source sheet.")
            return
        if not row_fields:
            st.error("Select at least one Row Field.")
            return
        vf_clean = [v for v in value_fields_list if v.get("column", "").strip()]
        if not vf_clean:
            st.error("Add at least one Value Field.")
            return
        if not tgt_file:
            st.error("Select a target file.")
            return
        if not tgt_sheet_val.strip():
            st.error("Enter a target sheet name.")
            return

        # Parse fill value
        try:
            fill_val = float(fill_value_str) if fill_value_str.strip() else 0
            if fill_val == int(fill_val):
                fill_val = int(fill_val)
        except ValueError:
            fill_val = fill_value_str  # keep as string (e.g. "-")

        # Human-readable label
        rf_str = ", ".join(row_fields[:3])
        if len(row_fields) > 3:
            rf_str += f" +{len(row_fields)-3}"
        vf_str = ", ".join(
            f"{v['column']} ({AGGFUNC_LABELS.get(v['aggfunc'], v['aggfunc'])})"
            for v in vf_clean[:2]
        )
        if len(vf_clean) > 2:
            vf_str += f" +{len(vf_clean)-2}"
        label = f"Pivot: {rf_str}  |  {vf_str}  →  {tgt_sheet_val}"

        add_step("pivot", label, {
            "src_file":         src_file,
            "src_sheet":        src_sheet,
            "src_header_row":   int(src_hrow),
            "tgt_file":         tgt_file,
            "tgt_sheet":        tgt_sheet_val.strip(),
            "tgt_start_row":    int(tgt_start_row),
            "tgt_start_col":    tgt_start_col.strip().upper() or "A",
            "overwrite_sheet":  overwrite_sht,
            "row_fields":       row_fields,
            "col_fields":       col_fields,
            "value_fields":     vf_clean,
            "filter_conditions": pivot_filters if pivot_filters else [],
            "filter_combine":   pivot_combine,
            "fill_value":       fill_val,
            "grand_total_rows": grand_rows,
            "grand_total_cols": grand_cols,
            "sort_field":       sort_field_val or "",
            "sort_ascending":   sort_asc,
            "apply_formatting": apply_fmt,
            "header_bg_color":  header_bg.lstrip("#"),
            "header_font_color": header_fg.lstrip("#"),
            "total_bg_color":   total_bg.lstrip("#"),
            "export":           export,
        })


def pivot_df_preview_cols(row_fields, col_fields, value_fields_list):
    """Return plausible output column names for the sort dropdown."""
    cols = list(row_fields or [])
    for vf in (value_fields_list or []):
        c = vf.get("column", "")
        if c:
            lbl = vf.get("label") or c
            cols.append(lbl)
    return cols


def _file_index(label):
    """Return the index of a file label in get_files(), defaulting to 0."""
    files = get_files()
    try:
        return files.index(label)
    except (ValueError, TypeError):
        return 0


def _pivot_apply_template(tpl: dict):
    """Push template values into session state for the pivot form."""
    st.session_state["pivot_row_fields"] = tpl.get("row_fields", [])
    st.session_state["pivot_col_fields"] = tpl.get("col_fields", [])
    vfs = tpl.get("value_fields", [])
    st.session_state["pivot_n_vf"] = max(1, len(vfs))
    for i, vf in enumerate(vfs):
        st.session_state[f"pivot_vf_{i}_col"]   = vf.get("column", "")
        st.session_state[f"pivot_vf_{i}_agg"]   = vf.get("aggfunc", "sum")
        st.session_state[f"pivot_vf_{i}_label"] = vf.get("label", "")
    st.session_state["pivot_tgt_sheet"]       = tpl.get("tgt_sheet", "Pivot")
    st.session_state["pivot_tgt_start_row"]   = tpl.get("tgt_start_row", 1)
    st.session_state["pivot_tgt_start_col"]   = tpl.get("tgt_start_col", "A")
    st.session_state["pivot_grand_rows"]      = tpl.get("grand_total_rows", False)
    st.session_state["pivot_grand_cols"]      = tpl.get("grand_total_cols", False)
    st.session_state["pivot_sort_field"]      = tpl.get("sort_field", "")
    st.session_state["pivot_sort_asc"]        = tpl.get("sort_ascending", True)
    st.session_state["pivot_apply_fmt"]       = tpl.get("apply_formatting", True)
    st.session_state["pivot_header_bg"]       = "#" + tpl.get("header_bg_color", "1F3864")
    st.session_state["pivot_header_fg"]       = "#" + tpl.get("header_font_color", "FFFFFF")
    st.session_state["pivot_total_bg"]        = "#" + tpl.get("total_bg_color", "D9E1F2")


def render_write_cell_form():
    st.markdown('<div class="pkf-section">Write to Cell(s)</div>', unsafe_allow_html=True)
    st.caption("Write a value to a specific cell OR fill an entire column with the same value.")

    file = st.selectbox("File", get_files(), key="write_cell_file")

    # ── Scope ──────────────────────────────────────────────────────────────────
    SCOPE_LABELS = ["Single Sheet", "All Sheets in Workbook"]
    SCOPE_KEYS   = ["single", "all_sheets"]
    _wc_scope_i  = SCOPE_KEYS.index(st.session_state.get("write_cell_scope", "single"))
    _wc_scope_lbl = st.radio(
        "Write to",
        SCOPE_LABELS, index=_wc_scope_i,
        key="write_cell_scope_radio", horizontal=True,
        help=(
            "**Single Sheet** — write to one sheet (classic).\n\n"
            "**All Sheets** — write the same cell/column into every sheet in the workbook. "
            "Optionally restrict which sheets are included."
        ),
    )
    wc_scope = SCOPE_KEYS[SCOPE_LABELS.index(_wc_scope_lbl)]

    if wc_scope == "single":
        sheet = sheet_selectbox("Sheet", file, key="write_cell_sheet")
        wc_include_sheets = []
        wc_exclude_sheets = []
    else:
        sheet = None
        st.caption(
            "All sheets will be updated. "
            "Optionally enter comma-separated sheet names to **include** (only those) "
            "or **exclude** (skip those)."
        )
        _wc_inc = st.text_input(
            "Include only *(comma-sep, blank = all)*",
            value=st.session_state.get("write_cell_include_sheets", ""),
            key="write_cell_include_sheets",
            placeholder="Sheet1, Sheet2",
        )
        _wc_exc = st.text_input(
            "Exclude *(comma-sep)*",
            value=st.session_state.get("write_cell_exclude_sheets", ""),
            key="write_cell_exclude_sheets",
            placeholder="Summary, Template",
        )
        wc_include_sheets = [s.strip() for s in _wc_inc.split(",") if s.strip()]
        wc_exclude_sheets = [s.strip() for s in _wc_exc.split(",") if s.strip()]

    st.divider()

    mode = st.radio(
        "Mode",
        ["Single Cell", "Column (same value for all rows)"],
        key="write_cell_mode",
        horizontal=True,
    )

    value = st.text_input("Value to write", "", key="write_cell_value")

    if mode == "Single Cell":
        cell_ref = st.text_input(
            "Cell Reference *(e.g., A5, B10)*",
            value=st.session_state.get("write_cell_ref", "A1"),
            key="write_cell_ref",
        )
        column = ""
        start_row = 2
        end_row = "last"
        header_row = 1
        if wc_scope == "all_sheets":
            st.caption(
                f"📝  Cell **{cell_ref or '…'}** will be written with `{value or '…'}` "
                f"in every target sheet."
            )
    else:
        cell_ref = ""
        column = st.text_input(
            "Column Letter *(e.g., A, B, C)*",
            value=st.session_state.get("write_cell_column", "A"),
            key="write_cell_column",
        )
        header_row = st.number_input("Header Row", min_value=1, value=1, key="write_cell_header")
        start_row = st.number_input("Start Row", min_value=1, value=2, key="write_cell_start")
        end_row = end_row_input("End Row", key_prefix="write_cell_end", default_last=True, default_number=100)
        if wc_scope == "all_sheets":
            st.caption(
                f"📝  Column **{column or '…'}** rows {start_row}–{end_row if end_row != 'last' else 'last'} "
                f"will be filled with `{value or '…'}` in every target sheet."
            )

    export = export_options_ui("write_cell")

    if st.button("Add Step", key="write_cell_add"):
        if not file:
            st.error("Select a file.")
            return
        mode_key = "single_cell" if mode == "Single Cell" else "column"

        if wc_scope == "all_sheets":
            _scope_label = "all sheets"
            if wc_include_sheets:
                _scope_label = f"{len(wc_include_sheets)} sheet(s)"
        else:
            _scope_label = sheet or "(active)"

        if mode == "Single Cell":
            if not cell_ref:
                st.error("Enter a cell reference (e.g. A5).")
                return
            step_name = f"Write '{value}' → {cell_ref} ({_scope_label})"
        else:
            if not column:
                st.error("Enter a column letter.")
                return
            step_name = f"Write '{value}' → col {column} ({_scope_label})"

        add_step("write_cell", step_name, {
            "file":           file,
            "sheet":          sheet or "",
            "mode":           mode_key,
            "value":          value,
            "cell_ref":       cell_ref,
            "column":         column,
            "start_row":      int(start_row),
            "end_row":        end_row,
            "header_row":     int(header_row),
            "scope":          wc_scope,
            "include_sheets": wc_include_sheets,
            "exclude_sheets": wc_exclude_sheets,
            "export":         export,
        })


def render_quota_update_form():
    st.markdown('<div class="pkf-section">Quota Update (Generic)</div>', unsafe_allow_html=True)
    st.caption("Build quotas from a filtered SOURCE table, then apply a capped write into a TARGET table by key.")

    st.markdown('<div class="pkf-section">Source (build quotas)</div>', unsafe_allow_html=True)
    src_file = st.selectbox("Source File", get_files(), key="quota_src_file")
    src_sheet = sheet_selectbox("Source Sheet", src_file, key="quota_src_sheet")
    src_header_row = st.number_input("Source Header Row", min_value=1, value=1, step=1, key="quota_src_hdr")

    src_headers = get_headers_for_file_sheet(src_file, src_sheet, int(src_header_row))
    filter_col = st.selectbox("Filter column", options=[""] + src_headers, key="quota_filter_col") if src_headers else st.text_input("Filter column", "", key="quota_filter_col")
    filter_op = st.selectbox("Filter operator", ["equals", "not_equals", "contains", "starts_with", "ends_with"], key="quota_filter_op")
    filter_val = st.text_input("Filter value", "", key="quota_filter_val")

    # Source key columns (addable)
    st.markdown('<div class="pkf-section">Source key columns (to count)</div>', unsafe_allow_html=True)
    src_ctx = (src_file or "", src_sheet or "", int(src_header_row))
    if st.session_state.get("quota_src_key_ctx") != src_ctx:
        st.session_state.quota_src_key_ctx = src_ctx
        st.session_state.quota_src_keys = []
        st.session_state.quota_src_norms = []
    if "quota_src_keys"not in st.session_state:
        st.session_state.quota_src_keys = []
    if "quota_src_norms"not in st.session_state:
        st.session_state.quota_src_norms = []

    c1, c2 = st.columns([5, 1])
    with c1:
        src_key_to_add = st.selectbox("Add source key column", options=[""] + src_headers, key="quota_src_key_add") if src_headers else st.text_input("Add source key column", "", key="quota_src_key_add")
    with c2:
        st.write("")
        if st.button("", key="quota_src_key_add_btn", disabled=not src_key_to_add, use_container_width=True):
            if src_key_to_add not in st.session_state.quota_src_keys:
                st.session_state.quota_src_keys.append(src_key_to_add)
                st.session_state.quota_src_norms.append("text")
                st.rerun()

    if st.session_state.quota_src_keys:
        for i, col in enumerate(st.session_state.quota_src_keys):
            r1, r2, r3 = st.columns([5, 3, 1])
            with r1:
                st.text(f"{i+1}. {col}")
            with r2:
                st.session_state.quota_src_norms[i] = st.selectbox(
                    "Normalize",
                    ["text", "upper_text", "int_string"],
                    index=["text", "upper_text", "int_string"].index(st.session_state.quota_src_norms[i]) if st.session_state.quota_src_norms[i] in ["text","upper_text","int_string"] else 0,
                    key=f"quota_src_norm_{i}",
                )
            with r3:
                if st.button("", key=f"quota_src_key_del_{i}", use_container_width=True):
                    st.session_state.quota_src_keys.pop(i)
                    st.session_state.quota_src_norms.pop(i)
                    st.rerun()
    else:
        st.warning("Add at least one source key column.")

    st.markdown('<div class="pkf-section">Target (apply capped write)</div>', unsafe_allow_html=True)
    tgt_file = st.selectbox("Target File", get_files(), key="quota_tgt_file")
    tgt_sheet = sheet_selectbox("Target Sheet", tgt_file, key="quota_tgt_sheet")
    tgt_header_row = st.number_input("Target Header Row", min_value=1, value=1, step=1, key="quota_tgt_hdr")
    tgt_end_row = end_row_input("Target end row", key_prefix="quota_tgt_end", default_last=True, default_number=100)

    tgt_headers = get_headers_for_file_sheet(tgt_file, tgt_sheet, int(tgt_header_row))

    # Target key columns (addable, must align with source keys)
    st.markdown('<div class="pkf-section">Target key columns (must match source key count/order)</div>', unsafe_allow_html=True)
    tgt_ctx = (tgt_file or "", tgt_sheet or "", int(tgt_header_row))
    if st.session_state.get("quota_tgt_key_ctx") != tgt_ctx:
        st.session_state.quota_tgt_key_ctx = tgt_ctx
        st.session_state.quota_tgt_keys = []
        st.session_state.quota_tgt_norms = []
    if "quota_tgt_keys"not in st.session_state:
        st.session_state.quota_tgt_keys = []
    if "quota_tgt_norms"not in st.session_state:
        st.session_state.quota_tgt_norms = []

    c1, c2 = st.columns([5, 1])
    with c1:
        tgt_key_to_add = st.selectbox("Add target key column", options=[""] + tgt_headers, key="quota_tgt_key_add") if tgt_headers else st.text_input("Add target key column", "", key="quota_tgt_key_add")
    with c2:
        st.write("")
        if st.button("", key="quota_tgt_key_add_btn", disabled=not tgt_key_to_add, use_container_width=True):
            if tgt_key_to_add not in st.session_state.quota_tgt_keys:
                st.session_state.quota_tgt_keys.append(tgt_key_to_add)
                st.session_state.quota_tgt_norms.append("text")
                st.rerun()

    if st.session_state.quota_tgt_keys:
        for i, col in enumerate(st.session_state.quota_tgt_keys):
            r1, r2, r3 = st.columns([5, 3, 1])
            with r1:
                st.text(f"{i+1}. {col}")
            with r2:
                st.session_state.quota_tgt_norms[i] = st.selectbox(
                    "Normalize",
                    ["text", "upper_text", "int_string"],
                    index=["text", "upper_text", "int_string"].index(st.session_state.quota_tgt_norms[i]) if st.session_state.quota_tgt_norms[i] in ["text","upper_text","int_string"] else 0,
                    key=f"quota_tgt_norm_{i}",
                )
            with r3:
                if st.button("", key=f"quota_tgt_key_del_{i}", use_container_width=True):
                    st.session_state.quota_tgt_keys.pop(i)
                    st.session_state.quota_tgt_norms.pop(i)
                    st.rerun()
    else:
        st.warning("Add at least one target key column.")

    write_col = st.selectbox("Write column", options=[""] + tgt_headers, key="quota_write_col") if tgt_headers else st.text_input("Write column", "", key="quota_write_col")
    write_val = st.text_input("Write value", "", key="quota_write_val")

    export = export_options_ui("quota_update")

    if st.button("Add Step", key="quota_add"):
        src_keys = st.session_state.get("quota_src_keys", []) or []
        tgt_keys = st.session_state.get("quota_tgt_keys", []) or []
        if not src_keys or not tgt_keys:
            st.warning("Add key columns on both source and target.")
            return
        if len(src_keys) != len(tgt_keys):
            st.warning("Source and target key columns must have the same count.")
            return
        if not filter_col or filter_val == "":
            st.warning("Select filter column and enter filter value.")
            return
        if not write_col or write_val == "":
            st.warning("Select write column and enter write value.")
            return

        add_step("quota_update", f"Quota Update → write '{write_val}'", {
            "source_file": src_file,
            "source_sheet": src_sheet,
            "source_header_row": int(src_header_row),
            "source_filter_column": filter_col,
            "source_filter_operator": filter_op,
            "source_filter_value": filter_val,
            "source_key_columns": src_keys,
            "source_key_norms": st.session_state.get("quota_src_norms", []) or [],
            "target_file": tgt_file,
            "target_sheet": tgt_sheet,
            "target_header_row": int(tgt_header_row),
            "target_end_row": tgt_end_row,
            "target_key_columns": tgt_keys,
            "target_key_norms": st.session_state.get("quota_tgt_norms", []) or [],
            "write_column": write_col,
            "write_value": write_val,
            "export": export,
        })


def render_group_rollup_form():
    st.markdown('<div class="pkf-section">Group Rollup (Generic)</div>', unsafe_allow_html=True)
    st.caption("Roll up amounts across row types within groups and optionally delete duplicates.")

    file = st.selectbox("File", get_files(), key="roll_file")
    sheet = sheet_selectbox("Sheet", file, key="roll_sheet")
    header_row = st.number_input("Header Row", min_value=1, value=1, step=1, key="roll_hdr")
    end_row = end_row_input("End row", key_prefix="roll_end", default_last=True, default_number=100)

    headers = get_headers_for_file_sheet(file, sheet, int(header_row))
    type_col = st.selectbox("Type column", options=[""] + headers, key="roll_type_col") if headers else st.text_input("Type column", "", key="roll_type_col")
    amount_col = st.selectbox("Amount column", options=[""] + headers, key="roll_amount_col") if headers else st.text_input("Amount column", "", key="roll_amount_col")

    from_type = st.text_input("From type value (e.g., SLB)", "Security Lending & Borrowing", key="roll_from_type")
    into_type = st.text_input("Into type value (e.g., EQUITY)", "EQUITY", key="roll_into_type")

    include_from = st.checkbox("Include FROM amount in the merged sum", value=True, key="roll_include_from")
    delete_dup_into = st.checkbox("Delete duplicate INTO rows (keep first)", value=True, key="roll_delete_dup_into")
    delete_from_rows = st.checkbox("Delete FROM rows (optional)", value=False, key="roll_delete_from")

    # Key columns
    st.markdown('<div class="pkf-section">Group key columns</div>', unsafe_allow_html=True)
    ctx = (file or "", sheet or "", int(header_row))
    if st.session_state.get("roll_key_ctx") != ctx:
        st.session_state.roll_key_ctx = ctx
        st.session_state.roll_keys = []
    if "roll_keys"not in st.session_state:
        st.session_state.roll_keys = []

    c1, c2 = st.columns([5, 1])
    with c1:
        key_to_add = st.selectbox("Add key column", options=[""] + headers, key="roll_key_add") if headers else st.text_input("Add key column", "", key="roll_key_add")
    with c2:
        st.write("")
        if st.button("", key="roll_key_add_btn", disabled=not key_to_add, use_container_width=True):
            if key_to_add not in st.session_state.roll_keys:
                st.session_state.roll_keys.append(key_to_add)
                st.rerun()

    if st.session_state.roll_keys:
        for i, col in enumerate(st.session_state.roll_keys):
            r1, r2 = st.columns([8, 1])
            with r1:
                st.text(f"{i+1}. {col}")
            with r2:
                if st.button("", key=f"roll_key_del_{i}", use_container_width=True):
                    st.session_state.roll_keys.pop(i)
                    st.rerun()
    else:
        st.warning("Add at least one key column.")

    export = export_options_ui("group_rollup")
    if st.button("Add Step", key="roll_add"):
        keys = st.session_state.get("roll_keys", []) or []
        if not keys or not type_col or not amount_col:
            st.warning("Select key columns, type column, and amount column.")
            return
        add_step("group_rollup", f"Group Rollup: {from_type} → {into_type}", {
            "file": file,
            "sheet": sheet,
            "header_row": int(header_row),
            "end_row": end_row,
            "key_columns": keys,
            "type_column": type_col,
            "amount_column": amount_col,
            "from_type_value": from_type,
            "into_type_value": into_type,
            "include_from_in_sum": bool(include_from),
            "delete_duplicate_into_rows": bool(delete_dup_into),
            "delete_from_rows": bool(delete_from_rows),
            "export": export,
        })


def render_vlookup_form():
    st.markdown('<div class="pkf-section">VLOOKUP (Explicit)</div>', unsafe_allow_html=True)
    st.caption("Configure: Lookup Value → Search/Return table → write back to destination.")
    st.info("**Composite key**: Add multiple columns to Lookup Value AND Search Value to match on multiple columns (e.g., Account + Date).")

    # 1) Lookup Value (supports multiple columns)
    st.markdown('<div class="pkf-section">1) Lookup Value (from your data)</div>', unsafe_allow_html=True)
    lv_file = st.selectbox("Lookup Value - File", get_files(), key="vlex_lv_file")
    lv_sheet = sheet_selectbox("Lookup Value - Sheet", lv_file, key="vlex_lv_sheet")
    lv_header_row = st.number_input("Lookup Value - Header row", min_value=1, value=1, step=1, key="vlex_lv_hdr")
    
    # Reset lookup columns when context changes
    lv_ctx = (lv_file or "", lv_sheet or "", int(lv_header_row))
    if st.session_state.get("vlex_lv_ctx") != lv_ctx:
        st.session_state.vlex_lv_ctx = lv_ctx
        st.session_state.vlex_lv_cols = []
    if "vlex_lv_cols"not in st.session_state:
        st.session_state.vlex_lv_cols = []
    
    lv_headers = get_headers_for_file_sheet(lv_file, lv_sheet, int(lv_header_row))
    
    if lv_headers:
        c1, c2 = st.columns([5, 1])
        with c1:
            lv_col_to_add = st.selectbox("Add Lookup Column", options=[""] + lv_headers, index=0, key="vlex_lv_add_col")
        with c2:
            st.write("")
            if st.button("", key="vlex_lv_add_btn", disabled=not lv_col_to_add, use_container_width=True):
                if lv_col_to_add not in st.session_state.vlex_lv_cols:
                    st.session_state.vlex_lv_cols.append(lv_col_to_add)
                    st.rerun()
        
        if st.session_state.vlex_lv_cols:
            st.markdown("**Lookup columns (in order):**")
            for idx, col in enumerate(st.session_state.vlex_lv_cols):
                r1, r2 = st.columns([8, 1])
                with r1:
                    st.text(f" {idx+1}. {col}")
                with r2:
                    if st.button("", key=f"vlex_lv_del_{idx}", use_container_width=True):
                        st.session_state.vlex_lv_cols.pop(idx)
                        st.rerun()
    else:
        st.warning("Select file/sheet/header row to see columns.")

    # 2) Search Value (supports multiple columns - must match Lookup columns count)
    st.markdown('<div class="pkf-section">2) Search Value (key column(s) in lookup table)</div>', unsafe_allow_html=True)
    s_file = st.selectbox("Search Value - File", get_files(), key="vlex_s_file")
    s_sheet = sheet_selectbox("Search Value - Sheet", s_file, key="vlex_s_sheet")
    s_header_row = st.number_input("Search Value - Header row", min_value=1, value=1, step=1, key="vlex_s_hdr")
    
    # Reset search columns when context changes
    s_ctx = (s_file or "", s_sheet or "", int(s_header_row))
    if st.session_state.get("vlex_s_ctx") != s_ctx:
        st.session_state.vlex_s_ctx = s_ctx
        st.session_state.vlex_s_cols = []
    if "vlex_s_cols"not in st.session_state:
        st.session_state.vlex_s_cols = []
    
    s_headers = get_headers_for_file_sheet(s_file, s_sheet, int(s_header_row))
    
    lv_col_count = len(st.session_state.get("vlex_lv_cols", []))
    s_col_count = len(st.session_state.get("vlex_s_cols", []))
    
    if s_headers:
        if lv_col_count > 0:
            st.caption(f"Add **{lv_col_count}** search column(s) to match the {lv_col_count} lookup column(s).")
        c1, c2 = st.columns([5, 1])
        with c1:
            s_col_to_add = st.selectbox("Add Search Column", options=[""] + s_headers, index=0, key="vlex_s_add_col")
        with c2:
            st.write("")
            if st.button("", key="vlex_s_add_btn", disabled=not s_col_to_add, use_container_width=True):
                if s_col_to_add not in st.session_state.vlex_s_cols:
                    st.session_state.vlex_s_cols.append(s_col_to_add)
                    st.rerun()
        
        if st.session_state.vlex_s_cols:
            st.markdown("**Search columns (in order):**")
            for idx, col in enumerate(st.session_state.vlex_s_cols):
                r1, r2 = st.columns([8, 1])
                with r1:
                    st.text(f" {idx+1}. {col}")
                with r2:
                    if st.button("", key=f"vlex_s_del_{idx}", use_container_width=True):
                        st.session_state.vlex_s_cols.pop(idx)
                        st.rerun()
    else:
        st.warning("Select file/sheet/header row to see columns.")
    
    # Show matching mode info
    if lv_col_count > 0 and s_col_count > 0:
        if lv_col_count == 1 and s_col_count >= 1:
            st.success(f"Mode: **OR search** — lookup value will be searched in ANY of the {s_col_count} search column(s)")
        elif lv_col_count == s_col_count:
            st.success(f"Mode: **Composite key (AND)** — all {lv_col_count} column(s) must match")
        else:
            st.warning(f"{lv_col_count} lookup + {s_col_count} search: Use 1 lookup + N search (OR), or N lookup = N search (AND)")

    # 3) Return Value
    st.markdown('<div class="pkf-section">3) Return Value (value column to return)</div>', unsafe_allow_html=True)
    r_file = st.selectbox("Return Value - File", get_files(), key="vlex_r_file")
    r_sheet = sheet_selectbox("Return Value - Sheet", r_file, key="vlex_r_sheet")
    r_header_row = st.number_input("Return Value - Header row", min_value=1, value=1, step=1, key="vlex_r_hdr")
    r_col = column_selectbox("Return Value - Column", r_file, r_sheet, int(r_header_row), key="vlex_r_col")
    not_found = st.text_input("If not found, write this message", value="Not found", key="vlex_not_found")

    # Optional: Chain another VLOOKUP if not found
    st.markdown("---")
    chain_enabled = st.checkbox("If NOT found, run another VLOOKUP (Fallback #2)", value=False, key="vlex_chain_enabled")

    fallback_lookups = []
    if chain_enabled:
        st.markdown('<div class="pkf-section">Fallback #2 (only for rows where #1 is not found)</div>', unsafe_allow_html=True)

        # Search Value #2 (supports multiple columns)
        s2_file = st.selectbox("Fallback Search - File", get_files(), key="vlex2_s_file")
        s2_sheet = sheet_selectbox("Fallback Search - Sheet", s2_file, key="vlex2_s_sheet")
        s2_header_row = st.number_input("Fallback Search - Header row", min_value=1, value=1, step=1, key="vlex2_s_hdr")

        s2_ctx = (s2_file or "", s2_sheet or "", int(s2_header_row))
        if st.session_state.get("vlex2_s_ctx") != s2_ctx:
            st.session_state.vlex2_s_ctx = s2_ctx
            st.session_state.vlex2_s_cols = []
        if "vlex2_s_cols"not in st.session_state:
            st.session_state.vlex2_s_cols = []

        s2_headers = get_headers_for_file_sheet(s2_file, s2_sheet, int(s2_header_row))
        if s2_headers:
            c1, c2 = st.columns([5, 1])
            with c1:
                s2_col_to_add = st.selectbox("Add Fallback Search Column", options=[""] + s2_headers, index=0, key="vlex2_s_add_col")
            with c2:
                st.write("")
                if st.button("", key="vlex2_s_add_btn", disabled=not s2_col_to_add, use_container_width=True):
                    if s2_col_to_add not in st.session_state.vlex2_s_cols:
                        st.session_state.vlex2_s_cols.append(s2_col_to_add)
                        st.rerun()

            if st.session_state.vlex2_s_cols:
                st.markdown("**Fallback search columns (in order):**")
                for idx, col in enumerate(st.session_state.vlex2_s_cols):
                    r1, r2 = st.columns([8, 1])
                    with r1:
                        st.text(f" {idx+1}. {col}")
                    with r2:
                        if st.button("", key=f"vlex2_s_del_{idx}", use_container_width=True):
                            st.session_state.vlex2_s_cols.pop(idx)
                            st.rerun()
        else:
            st.warning("Select fallback search file/sheet/header row to see columns.")

        # Return Value #2
        st.markdown('<div class="pkf-section">Fallback Return Value (column to return)</div>', unsafe_allow_html=True)
        r2_file = st.selectbox("Fallback Return - File", get_files(), key="vlex2_r_file")
        r2_sheet = sheet_selectbox("Fallback Return - Sheet", r2_file, key="vlex2_r_sheet")
        r2_header_row = st.number_input("Fallback Return - Header row", min_value=1, value=1, step=1, key="vlex2_r_hdr")
        r2_col = column_selectbox("Fallback Return - Column", r2_file, r2_sheet, int(r2_header_row), key="vlex2_r_col")

        s2_cols = st.session_state.get("vlex2_s_cols", [])
        if s2_cols and r2_col:
            s2_col_val = s2_cols[0] if len(s2_cols) == 1 else s2_cols
            fallback_lookups.append({
                "search_file": s2_file,
                "search_sheet": s2_sheet,
                "search_header_row": int(s2_header_row),
                "search_columns": s2_col_val,
                "return_file": r2_file,
                "return_sheet": r2_sheet,
                "return_header_row": int(r2_header_row),
                "return_column": r2_col,
            })

    # 4) Return To (destination)
    st.markdown('<div class="pkf-section">4) Return To (destination)</div>', unsafe_allow_html=True)
    d_file = st.selectbox("Return To - File", get_files(), key="vlex_d_file")
    d_sheet = sheet_selectbox("Return To - Sheet", d_file, key="vlex_d_sheet")
    d_header_row = st.number_input("Return To - Header row", min_value=1, value=1, step=1, key="vlex_d_hdr")
    d_col = column_selectbox("Return To - Column", d_file, d_sheet, int(d_header_row), key="vlex_d_col")
    d_end_row = end_row_input("End Row", key_prefix="vlex_end", default_last=True, default_number=10)
    export = export_options_ui("vlookup")

    if st.button("Add Step", key="vlex_add"):
        lv_cols = st.session_state.get("vlex_lv_cols", [])
        s_cols = st.session_state.get("vlex_s_cols", [])
        
        if not lv_cols:
            st.warning("Please add at least one Lookup Value column.")
            return
        if not s_cols:
            st.warning("Please add at least one Search Value column.")
            return
        # Valid modes: 1 lookup + N search (OR), or N lookup = N search (AND)
        valid = (len(lv_cols) == 1) or (len(lv_cols) == len(s_cols))
        if not valid:
            st.error(f"Invalid: {len(lv_cols)} lookup + {len(s_cols)} search. Use 1 lookup + N search (OR), or equal counts (AND).")
            return
        
        # Single column -> string, multiple -> list
        lv_col_val = lv_cols[0] if len(lv_cols) == 1 else lv_cols
        s_col_val = s_cols[0] if len(s_cols) == 1 else s_cols
        
        col_desc = '+'.join(lv_cols) if len(lv_cols) > 1 else lv_cols[0]
        add_step("vlookup", f"VLOOKUP: {col_desc} → {d_col}", {
            "mode": "explicit",
            "lookup_value_file": lv_file,
            "lookup_value_sheet": lv_sheet,
            "lookup_value_header_row": int(lv_header_row),
            "lookup_value_column": lv_col_val,

            "search_file": s_file,
            "search_sheet": s_sheet,
            "search_header_row": int(s_header_row),
            "search_column": s_col_val,

            "return_file": r_file,
            "return_sheet": r_sheet,
            "return_header_row": int(r_header_row),
            "return_column": r_col,

            "return_to_file": d_file,
            "return_to_sheet": d_sheet,
            "return_to_header_row": int(d_header_row),
            "return_to_column": d_col,

            "not_found_value": not_found,
            "end_row": d_end_row,
            "fallback_lookups": fallback_lookups,
            "export": export,
        })


def render_advance_vlookup_form():
    st.markdown('<div class="pkf-section">Advanced VLOOKUP with Condition</div>', unsafe_allow_html=True)
    st.caption("Full VLOOKUP with IF condition builder - only runs VLOOKUP when condition is met.")
    
    # Condition operators for the builder
    CONDITION_OPERATORS = [
        ("equals", "= (Equals)"),
        ("not_equals", "<> (Not Equals)"),
        ("greater_than", "> (Greater Than)"),
        ("less_than", "< (Less Than)"),
        ("greater_equal", ">= (Greater Than or Equal)"),
        ("less_equal", "<= (Less Than or Equal)"),
        ("starts_with", "Starts With"),
        ("not_starts_with", "Does Not Start With"),
        ("ends_with", "Ends With"),
        ("not_ends_with", "Does Not End With"),
        ("contains", "Contains"),
        ("not_contains", "Does Not Contain"),
        ("is_blank", "Is Blank"),
        ("is_not_blank", "Is Not Blank"),
    ]
    
    # =========================================================================
    # 1) CONDITION BUILDER (Required)
    # =========================================================================
    st.markdown("---")
    st.markdown('<div class="pkf-section">IF Condition</div>', unsafe_allow_html=True)
    st.info("The VLOOKUP will **only** run for rows where this condition is TRUE. Otherwise, the ELSE action applies.")
    
    # Condition source file/sheet
    cond_file = st.selectbox("Condition - File", get_files(), key="advl_cond_file")
    cond_sheet = sheet_selectbox("Condition - Sheet", cond_file, key="advl_cond_sheet")
    cond_header_row = st.number_input("Condition - Header Row", min_value=1, value=1, key="advl_cond_hdr")
    
    cond_headers = get_headers_for_file_sheet(cond_file, cond_sheet, int(cond_header_row))
    
    c1, c2, c3 = st.columns([3, 2, 3])
    with c1:
        if cond_headers:
            cond_column = st.selectbox("IF Column", options=[""] + cond_headers, key="advl_cond_col")
        else:
            cond_column = st.text_input("IF Column", "", key="advl_cond_col")
    
    with c2:
        operator_options = [op[1] for op in CONDITION_OPERATORS]
        operator_idx = st.selectbox("Operator", options=range(len(operator_options)), 
                                     format_func=lambda x: operator_options[x], key="advl_cond_op")
        cond_operator = CONDITION_OPERATORS[operator_idx][0]
    
    with c3:
        # Some operators don't need a value (is_blank, is_not_blank)
        if cond_operator in ["is_blank", "is_not_blank"]:
            cond_value = ""
            st.text_input("Value", value="(not needed)", disabled=True, key="advl_cond_val_disabled")
        else:
            cond_value = st.text_input("Value", "", key="advl_cond_val")
    
    # Show condition preview
    if cond_column:
        op_display = dict(CONDITION_OPERATORS).get(cond_operator, cond_operator)
        if cond_operator in ["is_blank", "is_not_blank"]:
            st.success(f"Condition: **IF** `{cond_column}` **{op_display}** → Do VLOOKUP")
        else:
            st.success(f"Condition: **IF** `{cond_column}` **{op_display}** `{cond_value}` → Do VLOOKUP")
    
    # =========================================================================
    # 2) VLOOKUP Configuration (same as regular VLOOKUP)
    # =========================================================================
    st.markdown("---")
    st.markdown('<div class="pkf-section">VLOOKUP (When Condition is TRUE)</div>', unsafe_allow_html=True)
    
    # --- Lookup Value ---
    st.markdown('<div class="pkf-section">1) Lookup Value (from your data)</div>', unsafe_allow_html=True)
    lv_file = st.selectbox("Lookup Value - File", get_files(), key="advl_lv_file")
    lv_sheet = sheet_selectbox("Lookup Value - Sheet", lv_file, key="advl_lv_sheet")
    lv_header_row = st.number_input("Lookup Value - Header Row", min_value=1, value=1, key="advl_lv_hdr")
    
    lv_ctx = (lv_file or "", lv_sheet or "", int(lv_header_row))
    if st.session_state.get("advl_lv_ctx") != lv_ctx:
        st.session_state.advl_lv_ctx = lv_ctx
        st.session_state.advl_lv_cols = []
    if "advl_lv_cols"not in st.session_state:
        st.session_state.advl_lv_cols = []
    
    lv_headers = get_headers_for_file_sheet(lv_file, lv_sheet, int(lv_header_row))
    
    if lv_headers:
        c1, c2 = st.columns([5, 1])
        with c1:
            lv_col_to_add = st.selectbox("Add Lookup Column", options=[""] + lv_headers, index=0, key="advl_lv_add_col")
        with c2:
            st.write("")
            if st.button("", key="advl_lv_add_btn", disabled=not lv_col_to_add, use_container_width=True):
                if lv_col_to_add not in st.session_state.advl_lv_cols:
                    st.session_state.advl_lv_cols.append(lv_col_to_add)
                    st.rerun()
        
        if st.session_state.advl_lv_cols:
            st.markdown("**Lookup columns:**")
            for idx, col in enumerate(st.session_state.advl_lv_cols):
                r1, r2 = st.columns([8, 1])
                with r1:
                    st.text(f" {idx+1}. {col}")
                with r2:
                    if st.button("", key=f"advl_lv_del_{idx}", use_container_width=True):
                        st.session_state.advl_lv_cols.pop(idx)
                        st.rerun()
    else:
        st.warning("Select file/sheet/header row to see columns.")
    
    # --- Search Value ---
    st.markdown('<div class="pkf-section">2) Search Value (key column(s) in lookup table)</div>', unsafe_allow_html=True)
    s_file = st.selectbox("Search - File", get_files(), key="advl_s_file")
    s_sheet = sheet_selectbox("Search - Sheet", s_file, key="advl_s_sheet")
    s_header_row = st.number_input("Search - Header Row", min_value=1, value=1, key="advl_s_hdr")
    
    s_ctx = (s_file or "", s_sheet or "", int(s_header_row))
    if st.session_state.get("advl_s_ctx") != s_ctx:
        st.session_state.advl_s_ctx = s_ctx
        st.session_state.advl_s_cols = []
    if "advl_s_cols"not in st.session_state:
        st.session_state.advl_s_cols = []
    
    s_headers = get_headers_for_file_sheet(s_file, s_sheet, int(s_header_row))
    
    if s_headers:
        c1, c2 = st.columns([5, 1])
        with c1:
            s_col_to_add = st.selectbox("Add Search Column", options=[""] + s_headers, index=0, key="advl_s_add_col")
        with c2:
            st.write("")
            if st.button("", key="advl_s_add_btn", disabled=not s_col_to_add, use_container_width=True):
                if s_col_to_add not in st.session_state.advl_s_cols:
                    st.session_state.advl_s_cols.append(s_col_to_add)
                    st.rerun()
        
        if st.session_state.advl_s_cols:
            st.markdown("**Search columns:**")
            for idx, col in enumerate(st.session_state.advl_s_cols):
                r1, r2 = st.columns([8, 1])
                with r1:
                    st.text(f" {idx+1}. {col}")
                with r2:
                    if st.button("", key=f"advl_s_del_{idx}", use_container_width=True):
                        st.session_state.advl_s_cols.pop(idx)
                        st.rerun()
    else:
        st.warning("Select file/sheet/header row to see columns.")
    
    # --- Return Value ---
    st.markdown('<div class="pkf-section">3) Return Value (value column to return)</div>', unsafe_allow_html=True)
    r_file = st.selectbox("Return - File", get_files(), key="advl_r_file")
    r_sheet = sheet_selectbox("Return - Sheet", r_file, key="advl_r_sheet")
    r_header_row = st.number_input("Return - Header Row", min_value=1, value=1, key="advl_r_hdr")
    r_col = column_selectbox("Return - Column", r_file, r_sheet, int(r_header_row), key="advl_r_col")
    not_found_value = st.text_input("If not found, write", value="Not found", key="advl_not_found")
    
    # --- Return To ---
    st.markdown('<div class="pkf-section">4) Return To (destination)</div>', unsafe_allow_html=True)
    d_file = st.selectbox("Return To - File", get_files(), key="advl_d_file")
    d_sheet = sheet_selectbox("Return To - Sheet", d_file, key="advl_d_sheet")
    d_header_row = st.number_input("Return To - Header Row", min_value=1, value=1, key="advl_d_hdr")
    d_col = column_selectbox("Return To - Column", d_file, d_sheet, int(d_header_row), key="advl_d_col")
    d_end_row = end_row_input("End Row", key_prefix="advl_end", default_last=True, default_number=10)
    
    # =========================================================================
    # 3) ELSE Action (When Condition is FALSE)
    # =========================================================================
    st.markdown("---")
    st.markdown('<div class="pkf-section">ELSE Action (When Condition is FALSE)</div>', unsafe_allow_html=True)
    
    else_action = st.radio(
        "When condition is NOT met:",
        ["Copy from another column", "Write a fixed value", "Leave unchanged"],
        key="advl_else_action",
        horizontal=True
    )
    
    else_source_file = ""
    else_source_sheet = ""
    else_source_col = ""
    else_fixed_value = ""
    
    if else_action == "Copy from another column":
        c1, c2, c3 = st.columns(3)
        with c1:
            else_source_file = st.selectbox("ELSE - Source File", get_files(), key="advl_else_file")
        with c2:
            else_source_sheet = sheet_selectbox("ELSE - Sheet", else_source_file, key="advl_else_sheet")
        with c3:
            else_source_col = column_selectbox("ELSE - Column", else_source_file, else_source_sheet, 
                                                int(d_header_row), key="advl_else_col")
    elif else_action == "Write a fixed value":
        else_fixed_value = st.text_input("ELSE - Fixed Value to write", "", key="advl_else_fixed")
    
    # =========================================================================
    # Export and Add Step
    # =========================================================================
    st.markdown("---")
    export = export_options_ui("advance_vlookup")
    
    if st.button("Add Step", key="advl_add", type="primary"):
        lv_cols = st.session_state.get("advl_lv_cols", [])
        s_cols = st.session_state.get("advl_s_cols", [])
        
        # Validation
        errors = []
        if not cond_column:
            errors.append("Select a condition column")
        if cond_operator not in ["is_blank", "is_not_blank"] and not cond_value:
            errors.append("Enter a condition value")
        if not lv_cols:
            errors.append("Add at least one Lookup Value column")
        if not s_cols:
            errors.append("Add at least one Search Value column")
        if not r_col:
            errors.append("Select a Return Value column")
        if not d_col:
            errors.append("Select a Return To column")
        
        if errors:
            for err in errors:
                st.warning(f"{err}")
            return
        
        # Build step name
        op_display = dict(CONDITION_OPERATORS).get(cond_operator, cond_operator)
        if cond_operator in ["is_blank", "is_not_blank"]:
            step_name = f"Adv VLOOKUP (IF {cond_column} {op_display})"
        else:
            step_name = f"Adv VLOOKUP (IF {cond_column} {op_display} {cond_value})"
        
        add_step("advance_vlookup", step_name, {
            # Condition
            "condition_enabled": True,
            "condition_file": cond_file,
            "condition_sheet": cond_sheet,
            "condition_header_row": int(cond_header_row),
            "condition_column": cond_column,
            "condition_operator": cond_operator,
            "condition_value": cond_value,
            # Lookup Value
            "lookup_value_file": lv_file,
            "lookup_value_sheet": lv_sheet,
            "lookup_value_header_row": int(lv_header_row),
            "lookup_value_columns": lv_cols,
            # Search Value
            "search_file": s_file,
            "search_sheet": s_sheet,
            "search_header_row": int(s_header_row),
            "search_columns": s_cols,
            # Return Value
            "return_file": r_file,
            "return_sheet": r_sheet,
            "return_header_row": int(r_header_row),
            "return_column": r_col,
            "not_found_value": not_found_value,
            # Return To
            "destination_file": d_file,
            "destination_sheet": d_sheet,
            "destination_header_row": int(d_header_row),
            "destination_column": d_col,
            "end_row": d_end_row,
            # ELSE action
            "else_action": else_action,
            "else_source_file": else_source_file,
            "else_source_sheet": else_source_sheet,
            "else_source_column": else_source_col,
            "else_fixed_value": else_fixed_value,
            # Export
            "export": export,
        })


def render_sumifs_form():
    st.markdown('<div class="pkf-section">SUMIFS  *(group-sum from source → write to target)*</div>', unsafe_allow_html=True)

    # ── Target (the sheet to write results into) ──────────────────────────────
    st.markdown('<div class="pkf-section">Target  *(where the result goes)*</div>', unsafe_allow_html=True)
    tgt_file  = st.selectbox("Target File",  get_files(), key="sumifs_file")
    tgt_sheet = sheet_selectbox("Target Sheet", tgt_file, key="sumifs_sheet")
    tgt_hrow  = st.number_input("Target Header Row", min_value=1, value=1, step=1, key="sumifs_tgt_hrow")
    tgt_headers = get_headers_for_file_sheet(tgt_file, tgt_sheet, int(tgt_hrow))

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Lookup column in Target** *(key to match on)*")
        if tgt_headers:
            tgt_lookup_sel = st.selectbox("Target lookup column", options=tgt_headers, key="sumifs_tgt_lookup_hdr")
            # resolve to column letter
            tgt_lookup_col = get_column_letter_for_header(tgt_file, tgt_sheet, int(tgt_hrow), tgt_lookup_sel)
            st.caption(f"→ column **{tgt_lookup_col}**")
        else:
            tgt_lookup_col = st.text_input("Target lookup column (letter)", "K", key="sumifs_tgt_lookup_col")
    with col2:
        st.markdown("**Output column in Target** *(where sum is written)*")
        if tgt_headers:
            out_sel = st.selectbox("Output column", options=tgt_headers, key="sumifs_out_hdr")
            output_col = get_column_letter_for_header(tgt_file, tgt_sheet, int(tgt_hrow), out_sel)
            st.caption(f"→ column **{output_col}**")
        else:
            output_col = st.text_input("Output column (letter)", "L", key="sumifs_out_col")

    st.divider()

    # ── Source (the sheet to read/group from) ─────────────────────────────────
    st.markdown('<div class="pkf-section">Source  *(exhaustive detail sheet)*</div>', unsafe_allow_html=True)
    src_file  = st.selectbox("Source File",  get_files(), key="sumifs_src_file")
    src_sheet = sheet_selectbox("Source Sheet", src_file, key="sumifs_src_sheet")
    src_hrow  = st.number_input("Source Header Row", min_value=1, value=1, step=1, key="sumifs_src_hrow")
    src_headers = get_headers_for_file_sheet(src_file, src_sheet, int(src_hrow))

    col3, col4 = st.columns(2)
    with col3:
        st.markdown("**Lookup column in Source** *(key to match on)*")
        if src_headers:
            src_lookup_sel = st.selectbox("Source lookup column", options=src_headers, key="sumifs_src_lookup_hdr")
            src_lookup_col = get_column_letter_for_header(src_file, src_sheet, int(src_hrow), src_lookup_sel)
            st.caption(f"→ column **{src_lookup_col}**")
        else:
            src_lookup_col = st.text_input("Source lookup column (letter)", "D", key="sumifs_src_lookup_col")
    with col4:
        st.markdown("**Sum column in Source** *(values to add up)*")
        if src_headers:
            sum_sel = st.selectbox("Sum column", options=src_headers, key="sumifs_sum_hdr")
            sum_col = get_column_letter_for_header(src_file, src_sheet, int(src_hrow), sum_sel)
            st.caption(f"→ column **{sum_col}**")
        else:
            sum_col = st.text_input("Sum column (letter or number)", "B", key="sumifs_sum_col")

    export = export_options_ui("sumifs")

    if st.button("Add Step", key="sumifs_add"):
        if not tgt_file or not src_file:
            st.warning("Select both Target and Source files.")
            return
        label = (
            f"SUMIFS: {src_file}[{src_sheet}] col {src_lookup_col}→{sum_col} "
            f"→ {tgt_file}[{tgt_sheet}] col {tgt_lookup_col} → {output_col}"
        )
        add_step("sumifs", label, {
            "file":           tgt_file,
            "sheet":          tgt_sheet,
            "tgt_header_row": int(tgt_hrow),
            "tgt_lookup_col": tgt_lookup_col,
            "output_col":     output_col,
            "src_workbook":   src_file,
            "src_sheet":      src_sheet,
            "src_header_row": int(src_hrow),
            "src_lookup_col": src_lookup_col,
            "sum_col":        sum_col,
            "export":         export,
        })


def render_append_form():
    st.markdown('<div class="pkf-section">Append Rows</div>', unsafe_allow_html=True)
    src_file = st.selectbox("Source File", get_files(), key="append_src_file")
    src_sheet = sheet_selectbox("Source Sheet", src_file, key="append_src_sheet")
    tgt_file = st.selectbox("Target File", get_files(), key="append_tgt_file")
    tgt_sheet = sheet_selectbox("Target Sheet", tgt_file, key="append_tgt_sheet")
    export = export_options_ui("append")

    if st.button("Add Step", key="append_add"):
        add_step("append", f"Append {src_file} → {tgt_file}", {
            "src_file": src_file, "src_sheet": src_sheet,
            "tgt_file": tgt_file, "tgt_sheet": tgt_sheet,
            "export": export,
        })


def render_import_form():
    st.markdown('<div class="pkf-section">Import</div>', unsafe_allow_html=True)
    
    # Source
    st.markdown('<div class="pkf-section">Source</div>', unsafe_allow_html=True)
    src_file = st.selectbox("Source File", get_files(), key="import_src_file")
    src_sheet = sheet_selectbox("Source Sheet", src_file, key="import_src_sheet")
    src_header_row = st.number_input("Source header row", min_value=1, value=1, key="import_src_header_row")
    
    # Target
    st.markdown('<div class="pkf-section">Target</div>', unsafe_allow_html=True)
    tgt_file = st.selectbox("Target File", get_files(), key="import_tgt_file")
    # Warn immediately if the chosen target is .xlsb (read-only format)
    _tgt_path_check = get_file_path(tgt_file) if tgt_file else ""
    if _tgt_path_check and _tgt_path_check.lower().endswith(".xlsb"):
        st.error(
            " **Target file is .xlsb** — this format is read-only and cannot be written to.  \n"
            "Please select a **.xlsx** file as the target.  \n"
            "The .xlsb file can only be used as a **Source** to read data from."
        )
    tgt_sheet = sheet_selectbox("Target Sheet", tgt_file, key="import_tgt_sheet")
    tgt_header_row = st.number_input("Target header row", min_value=1, value=1, key="import_tgt_header_row")

    # Get headers for dropdowns
    src_headers = get_headers_for_file_sheet(src_file, src_sheet, int(src_header_row))
    tgt_headers = get_headers_for_file_sheet(tgt_file, tgt_sheet, int(tgt_header_row))
    
    # Reset mappings when context changes
    ctx = (src_file or "", src_sheet or "", int(src_header_row), tgt_file or "", tgt_sheet or "", int(tgt_header_row))
    if st.session_state.get("import_mapping_ctx") != ctx:
        st.session_state.import_mapping_ctx = ctx
        st.session_state.import_mappings = []
    
    if "import_mappings"not in st.session_state or not isinstance(st.session_state.import_mappings, list):
        st.session_state.import_mappings = []
    
    # Column mapping UI
    st.markdown('<div class="pkf-section">Column Mapping</div>', unsafe_allow_html=True)
    st.caption("Map source columns → target columns. Only mapped columns will be imported.")
    
    if src_headers and tgt_headers:
        c1, c2, c3 = st.columns([4, 4, 2])
        with c1:
            new_src_col = st.selectbox("Source Column", options=[""] + src_headers, index=0, key="import_new_src_col")
        with c2:
            new_tgt_col = st.selectbox("Target Column", options=[""] + tgt_headers, index=0, key="import_new_tgt_col")
        with c3:
            st.write("")  # spacer
            if st.button("Add", key="import_add_mapping", use_container_width=True, disabled=(not new_src_col or not new_tgt_col)):
                existing_srcs = [m[0] for m in st.session_state.import_mappings]
                if new_src_col in existing_srcs:
                    st.warning(f"Source column '{new_src_col}'already mapped")
                else:
                    st.session_state.import_mappings.append((new_src_col, new_tgt_col))
                    st.rerun()
        
        # Show current mappings
        if st.session_state.import_mappings:
            st.markdown("**Current mappings:**")
            for idx, (s, t) in enumerate(st.session_state.import_mappings):
                r1, r2, r3 = st.columns([5, 5, 1])
                with r1:
                    st.text(f"{s}")
                with r2:
                    st.text(f"→ {t}")
                with r3:
                    if st.button("", key=f"import_del_map_{idx}", use_container_width=True):
                        st.session_state.import_mappings.pop(idx)
                        st.rerun()
            
            if st.button("Clear all mappings", key="import_clear_mappings"):
                st.session_state.import_mappings = []
                st.rerun()
        else:
            st.info("No mappings added yet. Select source and target columns above.")
    else:
        st.warning("Select source and target file/sheet/header row to see available columns.")
        # Fallback to text input
        st.text_area("Or enter mappings manually (SourceCol->TargetCol per line)", "", key="import_manual_mapping")
    
    # Filter source data
    st.markdown('<div class="pkf-section">Filter Source Data  *(optional)*</div>', unsafe_allow_html=True)
    st.caption("Only rows matching ALL (or ANY) conditions will be imported. Leave empty to import all rows.")
    import_filters, import_combine = render_filter_builder("import_filter", src_headers)
    if import_filters:
        st.info(f"**Active:** {len(import_filters)} filter condition(s) · combine: **{import_combine}**")

    # Append mode option
    st.markdown('<div class="pkf-section">Import Mode</div>', unsafe_allow_html=True)
    append_mode = st.checkbox(
        "**Append after last row** (keep existing target data, add imported rows after)",
        value=False,
        key="import_append_mode",
        help="If checked, imported data will be added after existing rows in the target. If unchecked, target data will be replaced."
    )

    # Source filename column option
    st.markdown('<div class="pkf-section">Source Filename Column  *(optional)*</div>', unsafe_allow_html=True)
    add_filename_col = st.checkbox(
        "Add a column with the source filename for each row",
        value=False,
        key="import_add_filename_col",
        help="Useful when importing from multiple files — each row will record which file it came from."
    )
    filename_col_name = ""
    if add_filename_col:
        filename_col_name = st.text_input(
            "Column name for source filename",
            value="file_name",
            key="import_filename_col_name",
            help="This column will be added to the target with the source file's basename (e.g. Report_Jan.xlsx)."
        )

    export = export_options_ui("import")

    if st.button("Add Step", key="import_add"):
        # Collect mappings
        mappings = st.session_state.get("import_mappings", [])
        manual = st.session_state.get("import_manual_mapping", "")

        if not mappings and not manual.strip():
            st.warning("Please add at least one column mapping.")
            return

        # Convert mappings to list of tuples format (or string if manual)
        if mappings:
            mapping_data = mappings  # list of (src, tgt) tuples
        else:
            mapping_data = manual  # string format

        active_filters = import_filters  # returned directly by render_filter_builder above
        active_combine = import_combine
        filter_note = f", {len(active_filters)} filter(s)"if active_filters else ""
        mode_label = " [APPEND]"if append_mode else ""
        fn_col = filename_col_name.strip() if add_filename_col and filename_col_name.strip() else None
        fn_note = f", +filename col '{fn_col}'"if fn_col else ""
        add_step("import", f"Import {src_file} → {tgt_file} ({len(mappings) if mappings else '?'} cols){mode_label}{filter_note}{fn_note}", {
            "src_file": src_file, "src_sheet": src_sheet,
            "tgt_file": tgt_file, "tgt_sheet": tgt_sheet,
            "src_header_row": int(src_header_row),
            "tgt_header_row": int(tgt_header_row),
            "mapping": mapping_data,
            "append_mode": append_mode,
            "filters": active_filters,
            "filter_combine": active_combine,
            "filename_col": fn_col,
            "export": export,
        })


def render_header_map_form():
    st.markdown('<div class="pkf-section">Header Mapping</div>', unsafe_allow_html=True)
    src_file = st.selectbox("Source File", get_files(), key="header_map_src_file")
    src_sheet = sheet_selectbox("Source Sheet", src_file, key="header_map_src_sheet")
    tgt_file = st.selectbox("Target File", get_files(), key="header_map_tgt_file")
    tgt_sheet = sheet_selectbox("Target Sheet", tgt_file, key="header_map_tgt_sheet")
    mapping = st.text_area("Mapping", "OldHeader->NewHeader", key="header_map_mapping")
    export = export_options_ui("header_map")

    if st.button("Add Step", key="header_map_add"):
        add_step("header_map", f"Map Headers", {
            "src_file": src_file, "src_sheet": src_sheet,
            "tgt_file": tgt_file, "tgt_sheet": tgt_sheet, "mapping": mapping,
            "export": export,
        })


def render_filter_form():
    st.markdown('<div class="pkf-section">Filter</div>', unsafe_allow_html=True)
    file = st.selectbox("File", get_files(), key="filter_file")
    sheet = sheet_selectbox("Sheet", file, key="filter_sheet")
    column = st.text_input("Column", "A", key="filter_col")
    remove_empty = st.checkbox("Remove rows where this column is empty", value=False, key="filter_remove_empty")
    value = ""
    if not remove_empty:
        value = st.text_input("Filter Value", "Active", key="filter_value")
    header_row = st.number_input("Header Row", min_value=1, value=1, step=1, key="filter_header_row")
    export = export_options_ui("filter")

    if st.button("Add Step", key="filter_add"):
        if remove_empty:
            name = f"Filter: remove empty rows in {column}"
        else:
            name = f"Filter: {column}={value}"
        add_step("filter", name, {
            "file": file,
            "sheet": sheet,
            "column": column,
            "value": value,
            "remove_empty": bool(remove_empty),
            "header_row": int(header_row),
            "export": export,
        })


def render_sort_form():
    st.markdown('<div class="pkf-section">Sort</div>', unsafe_allow_html=True)
    file = st.selectbox("File", get_files(), key="sort_file")
    sheet = sheet_selectbox("Sheet", file, key="sort_sheet")
    header_row = st.number_input("Header Row", min_value=1, value=1, step=1, key="sort_header_row")

    # Reset builder when context changes (file/sheet/header_row)
    ctx = (file or "", sheet or "", int(header_row))
    if st.session_state.get("sort_builder_ctx") != ctx:
        st.session_state.sort_builder_ctx = ctx
        st.session_state.sort_builder_keys = []

    if "sort_builder_keys"not in st.session_state or not isinstance(st.session_state.sort_builder_keys, list):
        st.session_state.sort_builder_keys = []

    headers = get_headers_for_file_sheet(file, sheet, int(header_row))

    c1, c2, c3 = st.columns([3, 2, 1])
    with c1:
        if headers:
            col_to_add = st.selectbox("Add sort column", options=[""] + headers, index=0, key="sort_add_col")
        else:
            st.selectbox("Add sort column", options=["(Select file/sheet/header row)"], index=0, key="sort_add_col__disabled", disabled=True)
            col_to_add = ""
    with c2:
        direction = st.selectbox("Direction", options=["Ascending", "Descending"], index=0, key="sort_add_dir")
    with c3:
        if st.button("Add", key="sort_add_btn", use_container_width=True, disabled=(not col_to_add)):
            existing = [x.get("col") for x in st.session_state.sort_builder_keys]
            if col_to_add in existing:
                st.warning("Column already added to sort order.")
            else:
                st.session_state.sort_builder_keys.append({
                    "col": col_to_add,
                    "asc": (direction == "Ascending")
                })
                st.rerun()

    st.markdown("**Sort order (tie-break priority)**")
    if not st.session_state.sort_builder_keys:
        st.caption("Add one or more columns above. The first column is the primary sort, then second breaks ties, etc.")
    else:
        for idx, item in enumerate(st.session_state.sort_builder_keys):
            col = item.get("col", "")
            asc = bool(item.get("asc", True))
            label = "ASC"if asc else "DESC"
            r1, r2, r3, r4 = st.columns([6, 1, 1, 1])
            with r1:
                st.write(f"{idx+1}. {col} ({label})")
            with r2:
                if st.button("", key=f"sort_key_up_{idx}", disabled=(idx == 0), use_container_width=True):
                    keys = st.session_state.sort_builder_keys
                    keys[idx - 1], keys[idx] = keys[idx], keys[idx - 1]
                    st.session_state.sort_builder_keys = keys
                    st.rerun()
            with r3:
                if st.button("", key=f"sort_key_down_{idx}", disabled=(idx == len(st.session_state.sort_builder_keys) - 1), use_container_width=True):
                    keys = st.session_state.sort_builder_keys
                    keys[idx + 1], keys[idx] = keys[idx], keys[idx + 1]
                    st.session_state.sort_builder_keys = keys
                    st.rerun()
            with r4:
                if st.button("", key=f"sort_key_del_{idx}", use_container_width=True):
                    keys = st.session_state.sort_builder_keys
                    keys.pop(idx)
                    st.session_state.sort_builder_keys = keys
                    st.rerun()

        if st.button("Clear sort columns", key="sort_clear_cols"):
            st.session_state.sort_builder_keys = []
            st.rerun()

    export = export_options_ui("sort")

    if st.button("Add Step", key="sort_add"):
        keys = st.session_state.sort_builder_keys or []
        if not keys:
            st.warning("Please add at least one sort column.")
            return
        cols = [k["col"] for k in keys if k.get("col")]
        asc_list = [bool(k.get("asc", True)) for k in keys if k.get("col")]
        arrows = ["↑"if a else "↓"for a in asc_list]
        name_bits = [f"{c}{arrows[i]}"for i, c in enumerate(cols)]
        add_step("sort", f"Sort: " + ", ".join(name_bits), {
            "file": file,
            "sheet": sheet,
            "header_row": int(header_row),
            "sort_cols": cols,
            "sort_ascending": asc_list,
            "export": export,
        })


def render_insert_form():
    st.markdown('<div class="pkf-section">Insert Rows/Columns</div>', unsafe_allow_html=True)
    file = st.selectbox("File", get_files(), key="insert_file")
    sheet = sheet_selectbox("Sheet", file, key="insert_sheet")
    axis = st.selectbox("Type", ["Row", "Column"], key="insert_axis")
    start_idx = st.number_input("Start Index", 1, value=1, key="insert_start")
    end_idx = end_index_input("End Index", key_prefix="insert_end", default_last=False, default_number=int(start_idx))
    export = export_options_ui("insert")

    if st.button("Add Step", key="insert_add"):
        add_step("insert", f"Insert {axis}(s) at {start_idx}", {
            "file": file, "sheet": sheet, "axis": axis,
            "start": start_idx, "end": end_idx,
            "export": export,
        })


def _render_advanced_delete_form(file: str, export):
    """Advanced Delete sub-form — multi-sheet delete by condition with preview & log."""
    from operations.filter_utils import OPERATORS as ADV_OPS, NEEDS_VALUE as ADV_NEEDS_VAL

    ADV_OP_KEYS   = list(ADV_OPS.keys())
    ADV_OP_LABELS = list(ADV_OPS.values())

    st.markdown("---")
    st.markdown("#### 🗑️ Advanced Delete")
    st.caption(
        "Define conditions and choose which sheets to scan. "
        "A **preview** shows matching row counts before anything is deleted. "
        "A **delete log** is saved automatically so you can review what was removed."
    )

    # ── Sheet scope ───────────────────────────────────────────────────────────
    SCOPE_LABELS = ["Current Sheet Only", "Selected Sheets", "All Sheets in Workbook"]
    SCOPE_KEYS   = ["single", "selected_sheets", "all_sheets"]
    _ss_idx = SCOPE_KEYS.index(st.session_state.get("adv_del_scope", "single"))
    _scope_lbl = st.radio(
        "Apply to", SCOPE_LABELS, index=_ss_idx,
        key="adv_del_scope_radio", horizontal=True,
    )
    adv_scope = SCOPE_KEYS[SCOPE_LABELS.index(_scope_lbl)]

    adv_sheet        = ""
    adv_sel_sheets   = []
    adv_inc_sheets   = []
    adv_exc_sheets   = []

    if adv_scope == "single":
        adv_sheet = sheet_selectbox("Sheet", file, key="adv_del_sheet")

    elif adv_scope == "selected_sheets":
        _fp_adv = get_file_path(file) if file else None
        _all_sheets_adv = []
        if _fp_adv:
            try:
                from core import wb_cache as _adv_wbc
                _adv_wb = _adv_wbc.load(_fp_adv)
                _all_sheets_adv = _adv_wb.sheetnames
            except Exception:
                pass
        adv_sel_sheets = st.multiselect(
            "Select Sheets to process",
            options=_all_sheets_adv,
            default=st.session_state.get("adv_del_selected_sheets", []),
            key="adv_del_selected_sheets",
        )

    else:  # all_sheets
        st.caption("All sheets will be processed. Optionally enter comma-separated names to **include** (only those) or **exclude** (skip those).")
        _inc_raw = st.text_input(
            "Include only *(comma-sep, blank = all)*",
            value=st.session_state.get("adv_del_include_sheets_raw", ""),
            key="adv_del_include_sheets_raw",
            placeholder="Sheet1, Sheet2",
        )
        _exc_raw = st.text_input(
            "Exclude *(comma-sep)*",
            value=st.session_state.get("adv_del_exclude_sheets_raw", ""),
            key="adv_del_exclude_sheets_raw",
            placeholder="Summary, Template",
        )
        adv_inc_sheets = [s.strip() for s in _inc_raw.split(",") if s.strip()]
        adv_exc_sheets = [s.strip() for s in _exc_raw.split(",") if s.strip()]

    # ── Header row ────────────────────────────────────────────────────────────
    adv_header_row = int(st.number_input(
        "Header row", min_value=1,
        value=int(st.session_state.get("adv_del_header_row", 1) or 1),
        key="adv_del_header_row",
    ))

    # ── Condition builder ─────────────────────────────────────────────────────
    # Resolve headers for the reference sheet (single or first selected)
    _ref_sheet = adv_sheet or (adv_sel_sheets[0] if adv_sel_sheets else "")
    _adv_row_headers = get_headers_for_file_sheet(file, _ref_sheet, adv_header_row) if _ref_sheet else []

    st.markdown("**Conditions** — rows matching these conditions will be deleted")
    st.caption(
        "Supports: equals, contains, blank/not blank, greater/less than, date before/after/equal. "
        "Use column header names or letters (A, B, C …)."
    )

    if "adv_del_n_conds" not in st.session_state:
        st.session_state["adv_del_n_conds"] = 1
    _ba, _br = st.columns([1, 1])
    with _ba:
        if st.button("Add Condition", key="adv_del_add_cond"):
            st.session_state["adv_del_n_conds"] += 1
    with _br:
        if st.button("Remove Last", key="adv_del_rem_cond") and st.session_state["adv_del_n_conds"] > 1:
            st.session_state["adv_del_n_conds"] -= 1

    _n_conds = st.session_state["adv_del_n_conds"]

    # AND / OR combine
    adv_combine = st.radio(
        "Combine conditions with",
        ["AND", "OR"], index=0,
        key="adv_del_combine", horizontal=True,
        help="AND → row must match ALL conditions.  OR → row matches ANY condition.",
    )

    # Column / Operator / Value header
    _hc = st.columns([3, 3, 3, 1])
    for _lbl, _col in zip(["Column", "Condition", "Value", ""], _hc):
        _col.markdown(f"**{_lbl}**")

    adv_filters = []
    for _i in range(_n_conds):
        _rc = st.columns([3, 3, 3, 1])
        if _adv_row_headers:
            _saved_col = st.session_state.get(f"adv_del_cond_{_i}_col", "")
            _col_opts  = [""] + _adv_row_headers
            _col_idx   = _col_opts.index(_saved_col) if _saved_col in _col_opts else 0
            _c_val = _rc[0].selectbox(
                f"Col {_i+1}", _col_opts, index=_col_idx,
                key=f"adv_del_cond_{_i}_col", label_visibility="collapsed",
            )
        else:
            _c_val = _rc[0].text_input(
                f"Col {_i+1}",
                value=st.session_state.get(f"adv_del_cond_{_i}_col", ""),
                key=f"adv_del_cond_{_i}_col", label_visibility="collapsed",
                placeholder="Header or A/B/C",
            )
        _saved_op_key = st.session_state.get(f"adv_del_cond_{_i}_op", "eq")
        _saved_op_idx = ADV_OP_KEYS.index(_saved_op_key) if _saved_op_key in ADV_OP_KEYS else 0
        _op_lbl = _rc[1].selectbox(
            f"Op {_i+1}", ADV_OP_LABELS, index=_saved_op_idx,
            key=f"adv_del_cond_{_i}_op_label", label_visibility="collapsed",
        )
        _op_key     = ADV_OP_KEYS[ADV_OP_LABELS.index(_op_lbl)]
        _needs_val  = _op_key in ADV_NEEDS_VAL
        _date_hint  = _op_key in {"date_before", "date_after", "date_on"}
        _v_val = _rc[2].text_input(
            f"Val {_i+1}",
            value=st.session_state.get(f"adv_del_cond_{_i}_val", ""),
            key=f"adv_del_cond_{_i}_val", label_visibility="collapsed",
            placeholder=("DD-MM-YYYY" if _date_hint else ("value" if _needs_val else "(not needed)")),
            disabled=not _needs_val,
        )
        # Delete this condition row
        if _n_conds > 1 and _rc[3].button("✕", key=f"adv_del_rem_{_i}"):
            for _j in range(_i, _n_conds - 1):
                st.session_state[f"adv_del_cond_{_j}_col"]      = st.session_state.get(f"adv_del_cond_{_j+1}_col", "")
                st.session_state[f"adv_del_cond_{_j}_op"]       = st.session_state.get(f"adv_del_cond_{_j+1}_op", "eq")
                st.session_state[f"adv_del_cond_{_j}_op_label"] = st.session_state.get(f"adv_del_cond_{_j+1}_op_label", ADV_OP_LABELS[0])
                st.session_state[f"adv_del_cond_{_j}_val"]      = st.session_state.get(f"adv_del_cond_{_j+1}_val", "")
            st.session_state["adv_del_n_conds"] -= 1
            st.rerun()

        if _c_val and _c_val.strip():
            adv_filters.append({
                "column":   _c_val.strip(),
                "operator": _op_key,
                "value":    _v_val.strip() if _needs_val else "",
            })

    # ── Log options ───────────────────────────────────────────────────────────
    st.divider()
    adv_save_log = st.checkbox(
        "Save delete log  *(Excel file listing every deleted row)*",
        value=bool(st.session_state.get("adv_del_save_log", True)),
        key="adv_del_save_log",
    )
    adv_log_path = ""
    if adv_save_log:
        adv_log_path = st.text_input(
            "Log file path *(leave blank for auto — saved in `delete_logs/` beside the source file)*",
            value=st.session_state.get("adv_del_log_path", ""),
            key="adv_del_log_path",
            placeholder="e.g. /Users/me/logs/delete_log.xlsx",
        )

    # ── Preview ───────────────────────────────────────────────────────────────
    st.divider()
    if st.button("🔍 Preview — count matching rows", key="adv_del_preview"):
        _fp = get_file_path(file) if file else None
        if not _fp:
            st.warning("Select a file first.")
        elif not adv_filters:
            st.warning("Add at least one condition.")
        else:
            with st.spinner("Scanning…"):
                _prev = preview_advanced_delete(
                    file_path=_fp,
                    sheet_scope=adv_scope,
                    sheet_name=adv_sheet,
                    selected_sheets=adv_sel_sheets,
                    include_sheets=adv_inc_sheets,
                    exclude_sheets=adv_exc_sheets,
                    filters=adv_filters,
                    filter_combine=adv_combine,
                    header_row=adv_header_row,
                )
            if _prev.get("error"):
                st.error(f"Preview error: {_prev['error']}")
            else:
                _total = _prev["total"]
                _by_sheet = _prev["by_sheet"]
                _n_sheets = len([s for s, n in _by_sheet.items() if n > 0])
                st.warning(
                    f"⚠️  **{_total} matching row(s)** found across **{_n_sheets} sheet(s)**. "
                    f"Adding this step will delete them permanently."
                )
                if _by_sheet:
                    import pandas as _pd_prev
                    _tbl = _pd_prev.DataFrame(
                        [{"Sheet": s, "Matching Rows": n} for s, n in _by_sheet.items()]
                    )
                    st.dataframe(_tbl, use_container_width=True, hide_index=True)

    # ── Add Step ──────────────────────────────────────────────────────────────
    if st.button("➕ Add Advanced Delete Step", key="adv_del_add", type="primary"):
        if not file:
            st.error("Select a file.")
            return
        if not adv_filters:
            st.error("Add at least one condition.")
            return
        if adv_scope == "selected_sheets" and not adv_sel_sheets:
            st.error("Select at least one sheet.")
            return

        _scope_short = {"single": "1 sheet", "selected_sheets": f"{len(adv_sel_sheets)} sheets", "all_sheets": "all sheets"}.get(adv_scope, adv_scope)
        _cond_short  = f"{len(adv_filters)} condition(s)"
        _label       = f"Advanced Delete — {_cond_short}, {_scope_short}"

        add_step("delete", _label, {
            "file":             file,
            "sheet":            adv_sheet or "",
            "row_mode":         "advanced_condition",
            "sheet_scope":      adv_scope,
            "selected_sheets":  adv_sel_sheets,
            "include_sheets":   adv_inc_sheets,
            "exclude_sheets":   adv_exc_sheets,
            "filters":          adv_filters,
            "filter_combine":   adv_combine,
            "header_row":       adv_header_row,
            "save_log":         adv_save_log,
            "log_path":         adv_log_path,
            "export":           export,
        })


def render_delete_form():
    st.markdown('<div class="pkf-section">Delete Rows/Columns</div>', unsafe_allow_html=True)
    file = st.selectbox("File", get_files(), key="delete_file")
    sheet = sheet_selectbox("Sheet", file, key="delete_sheet")
    axis = st.selectbox("Type", ["Row", "Column"], key="delete_axis")
    export = export_options_ui("delete")

    if axis == "Row":
        row_mode = st.radio(
            "Delete rows by",
            ["Row number range", "Condition (delete rows where…)", "Advanced Delete *(multi-sheet + log)*"],
            key="delete_row_mode",
            horizontal=True,
        )

        if row_mode == "Condition (delete rows where…)":
            header_row = st.number_input("Header row", min_value=1, value=1, key="delete_cond_hrow")
            row_headers = get_headers_for_file_sheet(file, sheet, int(header_row))
            st.caption("Define one or more conditions. Any row matching the combined condition will be **permanently deleted**.")
            del_filters, del_combine = render_filter_builder("delete_cond", row_headers)

            if st.button("Add Step", key="delete_add_cond"):
                if not file:
                    st.warning("Select a file.")
                    return
                if not del_filters:
                    st.warning("Add at least one condition.")
                    return
                add_step("delete", f"Delete rows by condition ({len(del_filters)} rule(s))", {
                    "file": file, "sheet": sheet,
                    "row_mode": "condition",
                    "filters": del_filters,
                    "filter_combine": del_combine,
                    "header_row": int(header_row),
                    "export": export,
                })
            return  # early return — rest of form not needed

        if row_mode == "Advanced Delete *(multi-sheet + log)*":
            _render_advanced_delete_form(file, export)
            return  # early return

        start_idx = st.number_input("Start Index", 1, value=1, key="delete_start")
        end_idx = end_index_input("End Index", key_prefix="delete_end", default_last=False, default_number=int(start_idx))
        cols_to_delete = []
        header_row = 1
    else:
        # Column deletion: ask which column(s) to delete
        header_row = st.number_input("Header Row (for column names)", min_value=1, value=1, step=1, key="delete_header_row")
        col_action = st.radio(
            "Column action",
            ["Delete the column(s) بالكامل (remove column)", "Clear data under header (keep column)"],
            key="delete_col_action",
            horizontal=True,
        )
        clear_only = (col_action == "Clear data under header (keep column)")
        if clear_only:
            end_row = end_row_input("Clear until row", key_prefix="delete_clear_end", default_last=True, default_number=100)
        else:
            end_row = "last"

        # Reset selection when context changes
        ctx = (file or "", sheet or "", int(header_row), str(col_action))
        if st.session_state.get("delete_col_ctx") != ctx:
            st.session_state.delete_col_ctx = ctx
            st.session_state.delete_cols = []
        if "delete_cols"not in st.session_state:
            st.session_state.delete_cols = []

        headers = get_headers_for_file_sheet(file, sheet, int(header_row))

        st.markdown('<div class="pkf-section">Columns to delete</div>', unsafe_allow_html=True)
        if headers:
            c1, c2 = st.columns([5, 1])
            with c1:
                col_to_add = st.selectbox("Add Column", options=[""] + headers, key="delete_col_add")
            with c2:
                st.write("")
                if st.button("", key="delete_col_add_btn", disabled=not col_to_add, use_container_width=True):
                    if col_to_add not in st.session_state.delete_cols:
                        st.session_state.delete_cols.append(col_to_add)
                        st.rerun()
        else:
            c1, c2 = st.columns([5, 1])
            with c1:
                col_to_add = st.text_input("Add Column (letter or header)", "", key="delete_col_add_text")
            with c2:
                st.write("")
                if st.button("", key="delete_col_add_btn_text", disabled=not col_to_add, use_container_width=True):
                    col_to_add = str(col_to_add).strip()
                    if col_to_add and col_to_add not in st.session_state.delete_cols:
                        st.session_state.delete_cols.append(col_to_add)
                        st.rerun()

        if st.session_state.delete_cols:
            for idx, c in enumerate(st.session_state.delete_cols):
                r1, r2 = st.columns([8, 1])
                with r1:
                    st.text(f" {idx+1}. {c}")
                with r2:
                    if st.button("", key=f"delete_col_del_{idx}", use_container_width=True):
                        st.session_state.delete_cols.pop(idx)
                        st.rerun()
        else:
            st.warning("Add at least one column to delete.")

        cols_to_delete = st.session_state.get("delete_cols", []) or []
        start_idx = 1
        end_idx = 1

    if st.button("Add Step", key="delete_add"):
        if axis == "Column":
            if not cols_to_delete:
                st.warning("Please add at least one column to delete.")
                return
            if clear_only:
                add_step("delete", f"Clear Column Data: {', '.join(cols_to_delete)}", {
                    "file": file, "sheet": sheet, "axis": axis,
                    "action": "clear_data",
                    "header_row": int(header_row),
                    "columns": cols_to_delete,
                    "end_row": end_row,
                    "export": export,
                })
            else:
                add_step("delete", f"Delete Column(s): {', '.join(cols_to_delete)}", {
                    "file": file, "sheet": sheet, "axis": axis,
                    "action": "delete_columns",
                    "header_row": int(header_row),
                    "columns": cols_to_delete,
                    "export": export,
                })
        else:
            add_step("delete", f"Delete {axis}(s) at {start_idx}", {
                "file": file, "sheet": sheet, "axis": axis,
                "start": start_idx, "end": end_idx,
                "export": export,
            })


def render_merge_cells_form():
    st.markdown('<div class="pkf-section">Merge Cells</div>', unsafe_allow_html=True)
    file = st.selectbox("File", get_files(), key="merge_cells_file")
    sheet = sheet_selectbox("Sheet", file, key="merge_cells_sheet")
    range_str = st.text_input("Range", "A1:C1", key="merge_cells_range")
    export = export_options_ui("merge_cells")

    if st.button("Add Step", key="merge_cells_add"):
        add_step("merge_cells", f"Merge {range_str}", {
            "file": file, "sheet": sheet, "range": range_str,
            "export": export,
        })


def render_delete_sheets_form():
    st.markdown('<div class="pkf-section">Delete Sheets</div>', unsafe_allow_html=True)
    st.caption("Remove one or more sheets from a workbook. Choose how to identify the sheets to delete.")

    file = st.selectbox("File", get_files(), key="ds_file",
                        help="Select the workbook whose sheets you want to manage.")
    export = export_options_ui("ds")

    # Load available sheets for this file
    file_path = get_file_path(file) if file else None
    available_sheets: list[str] = []
    if file_path:
        try:
            available_sheets = get_sheet_names(file_path)
        except Exception:
            available_sheets = []

    MODE_LABELS = {
        "named":     "Delete specific sheets (choose by name)",
        "keep_only": "Keep only these sheets (delete the rest)",
        "pattern":   "Delete sheets matching a name pattern",
        "empty":     "Delete all empty sheets",
    }
    mode_label = st.radio(
        "Deletion mode",
        list(MODE_LABELS.values()),
        key="ds_mode_label",
        help=(
            "**Named** — pick exact sheets to remove.\n\n"
            "**Keep only** — inverse: pick sheets to KEEP, delete everything else.\n\n"
            "**Pattern** — delete sheets whose names contain/start with/end with a string or regex.\n\n"
            "**Empty** — automatically find and remove sheets with no data."
        ),
    )
    mode = {v: k for k, v in MODE_LABELS.items()}[mode_label]

    sheet_names: list[str] = []

    # ── Mode: named ──────────────────────────────────────────────────────────
    if mode == "named":
        if available_sheets:
            sheet_names = st.multiselect(
                "Sheets to DELETE",
                options=available_sheets,
                key="ds_named_sheets",
                help="Select one or more sheets. They will be permanently removed.",
            )
            if sheet_names:
                st.warning(f"Will delete: **{', '.join(sheet_names)}**")
        else:
            sheet_names_raw = st.text_area(
                "Sheet names (one per line)",
                key="ds_named_manual",
                placeholder="Sheet1\nTemp\nHelper",
            )
            sheet_names = [s.strip() for s in sheet_names_raw.splitlines() if s.strip()]

    # ── Mode: keep_only ──────────────────────────────────────────────────────
    elif mode == "keep_only":
        if available_sheets:
            sheet_names = st.multiselect(
                "Sheets to KEEP",
                options=available_sheets,
                key="ds_keep_sheets",
                help="All other sheets will be deleted. At least one sheet must be kept.",
            )
            delete_preview = [s for s in available_sheets if s not in sheet_names]
            if delete_preview:
                st.warning(f"Will delete: **{', '.join(delete_preview)}**")
            elif sheet_names:
                st.success("No sheets will be deleted.")
        else:
            keep_raw = st.text_area(
                "Sheets to KEEP (one per line)",
                key="ds_keep_manual",
                placeholder="FinalOutput\nSummary",
            )
            sheet_names = [s.strip() for s in keep_raw.splitlines() if s.strip()]

    # ── Mode: pattern ─────────────────────────────────────────────────────────
    elif mode == "pattern":
        col_pat, col_type = st.columns([3, 2])
        with col_pat:
            pattern = st.text_input(
                "Pattern",
                key="ds_pattern",
                placeholder="e.g.  temp_  or  ^Sheet\\d+$",
                help="The text or regex to match against sheet names.",
            )
        with col_type:
            pattern_type = st.selectbox(
                "Match type",
                options=["contains", "starts_with", "ends_with", "exact", "regex"],
                key="ds_pattern_type",
                format_func=lambda x: {
                    "contains": "Contains",
                    "starts_with": "Starts with",
                    "ends_with": "Ends with",
                    "exact": "Exact match",
                    "regex": "Regex",
                }[x],
            )
        case_sensitive = st.checkbox("Case-sensitive", value=False, key="ds_case")

        # Live preview of matching sheets
        if pattern and available_sheets:
            from operations.sheet_ops import _sheets_matching_pattern
            preview_matches = _sheets_matching_pattern(
                available_sheets, pattern, pattern_type, case_sensitive
            )
            if preview_matches:
                st.warning(f"Matching sheets (will be deleted): **{', '.join(preview_matches)}**")
            else:
                st.info("No sheets currently match this pattern.")

    # ── Mode: empty ───────────────────────────────────────────────────────────
    elif mode == "empty":
        if available_sheets and file_path:
            from operations.sheet_ops import _find_empty_sheets, get_sheet_names as _gsn
            try:
                import openpyxl as _xl
                from core import wb_cache as _wbc
                _wb = _wbc.load(file_path)
                empty_preview = _find_empty_sheets(_wb, available_sheets)
                _wb.close()
                if empty_preview:
                    st.warning(f"Empty sheets found (will be deleted): **{', '.join(empty_preview)}**")
                else:
                    st.success("No empty sheets found in this file.")
            except Exception:
                st.info("Open the file to preview empty sheets.")

    # ── Add Step button ───────────────────────────────────────────────────────
    if st.button("Add Step", key="ds_add"):
        if not file:
            st.warning("Select a file first.")
            return

        cfg: dict = {
            "file": file,
            "mode": mode,
            "protect_last": True,
            "export": export,
        }

        if mode in ("named", "keep_only"):
            if not sheet_names:
                st.warning("Select at least one sheet.")
                return
            cfg["sheet_names"] = sheet_names
            verb = "Delete" if mode == "named" else "Keep only"
            step_name = f"{verb}: {', '.join(sheet_names[:3])}{'…' if len(sheet_names) > 3 else ''}"

        elif mode == "pattern":
            if not st.session_state.get("ds_pattern", "").strip():
                st.warning("Enter a pattern.")
                return
            cfg["pattern"]      = st.session_state["ds_pattern"].strip()
            cfg["pattern_type"] = st.session_state.get("ds_pattern_type", "contains")
            cfg["case_sensitive"] = bool(st.session_state.get("ds_case", False))
            step_name = f"Delete sheets ({cfg['pattern_type']}: '{cfg['pattern']}')"

        else:  # empty
            step_name = "Delete empty sheets"

        add_step("delete_sheets", step_name, cfg)


def render_data_input_form():
    st.markdown('<div class="pkf-section">Data Input</div>', unsafe_allow_html=True)
    file = st.selectbox("File", get_files(), key="data_input_file", help="Data Input expects a registered file.")
    sheet = sheet_selectbox("Sheet (for Excel)", file, key="data_input_sheet")

    if st.button("Add Step", key="data_input_add"):
        add_step("data_input", f"Load: {file}", {
            "file": file, "sheet": sheet
        })


def _is_class_key(i, field):
    return f"is_cls_{i}_{field}"


def _get_class_cfg(i):
    """Read one file class config from session_state widget keys."""
    def g(field, default=""):
        return st.session_state.get(_is_class_key(i, field), default)

    return {
        "class_name":     g("name"),
        "mode":           g("mode", "A"),
        "file_path":      g("file_path"),
        "folder_path":    g("folder_path"),
        "rule_contains":  g("rule_contains"),
        "rule_starts_with": g("rule_starts_with"),
        "rule_ends_with": g("rule_ends_with"),
        "rule_extension": g("rule_extension", "All Excel (.xlsx, .xlsb, .xls, .xlsm)"),
        "rule_excludes":  g("rule_excludes"),
        "date_pattern":   g("date_pattern"),
        "date_pick":      g("date_pick", "first"),
        "sheet_mode":     g("sheet_mode", "first"),
        "sheet_exact":    g("sheet_exact"),
        "sheet_contains": g("sheet_contains"),
        "sheet_headers":  g("sheet_headers"),
        # Mode E — user upload
        "upload_accept":  g("upload_accept", "Excel (.xlsx / .xls / .xlsm / .xlsb)"),
        "upload_multi":   g("upload_multi", False),
    }


def render_input_source_form():
    """Flexible file collector step: single path, folder, filename rules, or date patterns."""
    st.markdown('<div class="pkf-section">Input Source / File Collector</div>', unsafe_allow_html=True)
    st.caption(
        "Define one or more **file classes** (e.g. *Bank Statement*, *ERP Ledger*). "
        "Later steps can reference files by their class name instead of raw paths."
    )

    # Maintain list of classes in session state
    if "is_classes"not in st.session_state:
        st.session_state.is_classes = [{}]  # start with one empty class

    num_classes = len(st.session_state.is_classes)

    col_add, col_remove = st.columns([1, 1])
    with col_add:
        if st.button("Add File Class", key="is_add_cls"):
            st.session_state.is_classes.append({})
            st.rerun()
    with col_remove:
        if num_classes > 1 and st.button("Remove Last Class", key="is_rm_cls"):
            st.session_state.is_classes.pop()
            st.rerun()

    st.divider()

    MODE_OPTIONS = [
        "A — Single File",
        "B — Folder: All Excel Files",
        "C — Folder: Filename Rules",
        "D — Folder: Dynamic Date Rules",
        "E — User Upload at Execution Time",
    ]
    SHEET_MODES = [
        "first — Use First Sheet",
        "exact — Exact Sheet Name",
        "contains — Sheet Name Contains",
        "latest_month — Latest Month Sheet",
        "detect_headers — Detect by Column Headers",
    ]
    EXT_OPTIONS = [
        "All Excel (.xlsx, .xlsb, .xls, .xlsm)",   # default — picks up every Excel format
        ".xlsx",
        ".xlsb",
        ".xls",
        ".xlsm",
        ".csv",
        ".txt",
        "*",
    ]

    for i in range(len(st.session_state.is_classes)):
        with st.expander(
            f"File Class {i+1}: {st.session_state.get(_is_class_key(i,'name')) or '(unnamed)'}",
            expanded=True,
        ):
            # --- Class name ---
            st.text_input(
                "Class Name  (used to reference this file in later steps)",
                key=_is_class_key(i, "name"),
                placeholder="e.g. Bank Statement",
            )

            # --- Mode ---
            mode_label = st.selectbox(
                "Source Mode",
                options=MODE_OPTIONS,
                key=_is_class_key(i, "mode_label"),
                index=0,
                help="How to locate the file(s)",
            )
            mode = mode_label[0]  # "A", "B", "C", or "D"
            st.session_state[_is_class_key(i, "mode")] = mode

            # --- Mode-specific inputs ---
            if mode == "E":
                st.info(
                    " **User Upload** — the workflow executor will prompt for this file "
                    "before the run starts. No path is needed now."
                )
                st.selectbox(
                    "Accepted file types",
                    options=["Excel (.xlsx / .xls / .xlsm / .xlsb)", "CSV (.csv)", "Any"],
                    key=_is_class_key(i, "upload_accept"),
                )
                st.checkbox(
                    "Allow multiple files (user can upload more than one)",
                    key=_is_class_key(i, "upload_multi"),
                    value=False,
                )

            elif mode == "A":
                st.text_input(
                    "Absolute File Path",
                    key=_is_class_key(i, "file_path"),
                    placeholder="/Users/you/Documents/bank_may2026.xlsx",
                )

            elif mode in ("B", "C", "D"):
                st.text_input(
                    "Folder Path",
                    key=_is_class_key(i, "folder_path"),
                    placeholder="/Users/you/Documents/uploads/",
                )
                st.selectbox(
                    "File Extension",
                    options=EXT_OPTIONS,
                    key=_is_class_key(i, "rule_extension"),
                )

                if mode == "C":
                    st.markdown("**Filename Rules** (all conditions must be satisfied):")
                    c1, c2 = st.columns(2)
                    with c1:
                        st.text_input("Name contains", key=_is_class_key(i, "rule_contains"), placeholder="Bank")
                        st.text_input("Name starts with", key=_is_class_key(i, "rule_starts_with"), placeholder="HDFC")
                    with c2:
                        st.text_input("Name ends with", key=_is_class_key(i, "rule_ends_with"), placeholder="_May2026")
                        st.text_input(
                            "Exclude if name contains (comma-sep)",
                            key=_is_class_key(i, "rule_excludes"),
                            placeholder="old, backup, draft",
                        )

                elif mode == "D":
                    st.markdown("**Dynamic Date Pattern**")
                    st.caption(
                        "Tokens: `{current_month}` `{current_month_full}` `{current_month_num}` "
                        "`{previous_month}` `{previous_month_full}` `{previous_month_num}` "
                        "`{current_year}` `{prev_year}` `{today}` `{yesterday}`"
                    )
                    pattern = st.text_input(
                        "Filename pattern (tokens will be replaced at run-time)",
                        key=_is_class_key(i, "date_pattern"),
                        placeholder="BankStatement_{current_month}_{current_year}",
                    )
                    if pattern:
                        from operations.input_source import _resolve_tokens
                        st.caption(f"Preview today: **{_resolve_tokens(pattern)}**")
                    st.radio(
                        "When multiple files match, use:",
                        options=["first — First alphabetically", "latest_modified — Latest modified"],
                        key=_is_class_key(i, "date_pick_label"),
                        horizontal=True,
                    )
                    date_pick_label = st.session_state.get(_is_class_key(i, "date_pick_label"), "first — First alphabetically")
                    st.session_state[_is_class_key(i, "date_pick")] = date_pick_label.split(" — ")[0]

            st.divider()
            # --- Sheet selection ---
            st.markdown("**Sheet Selection**")
            sheet_mode_label = st.selectbox(
                "How to pick the sheet",
                options=SHEET_MODES,
                key=_is_class_key(i, "sheet_mode_label"),
            )
            sheet_mode = sheet_mode_label.split(" — ")[0]
            st.session_state[_is_class_key(i, "sheet_mode")] = sheet_mode

            if sheet_mode == "exact":
                st.text_input("Exact sheet name", key=_is_class_key(i, "sheet_exact"), placeholder="Sheet1")
            elif sheet_mode == "contains":
                st.text_input("Sheet name contains", key=_is_class_key(i, "sheet_contains"), placeholder="Bank")
            elif sheet_mode == "detect_headers":
                st.text_input(
                    "Required column headers (comma-separated)",
                    key=_is_class_key(i, "sheet_headers"),
                    placeholder="Date, Narration, Debit, Credit, Balance",
                )

    st.divider()

    # --- Preview / Validate ---
    if st.button("Validate — Preview Which Files Will Be Loaded", key="is_validate"):
        cfgs = [_get_class_cfg(i) for i in range(len(st.session_state.is_classes))]
        rows = []
        all_ok = True
        for cfg in cfgs:
            if cfg.get("mode") == "E":
                rows.append({
                    "Class": cfg["class_name"] or "(unnamed)",
                    "Mode": "E — User Upload",
                    "Status": "",
                    "File Found": "Provided at run time",
                    "Sheet": cfg.get("sheet_exact") or cfg.get("sheet_mode") or "first",
                    "Note": "User will upload before execution",
                })
                continue
            result = resolve_file_class(cfg)
            status_icon = ""if result["status"] == "found"else ""
            if result["status"] != "found":
                all_ok = False
            rows.append({
                "Class": result["class_name"] or "(unnamed)",
                "Mode": result["mode"],
                "Status": status_icon,
                "File Found": os.path.basename(result["file_path"]) if result["file_path"] else "—",
                "Sheet": result["sheet_name"] or "—",
                "Note": result["error"] or result["sheet_msg"],
            })
            if result.get("candidates") and len(result["candidates"]) > 1:
                st.caption(
                    f"Class '{result['class_name']}': {len(result['candidates'])} files matched, "
                    f"using first: `{os.path.basename(result['candidates'][0])}`"
                )
        import pandas as pd
        st.dataframe(pd.DataFrame(rows), use_container_width=True)
        if all_ok:
            st.success("All file classes resolved. Ready to add as a workflow step.")
        else:
            st.warning("Some classes could not be resolved. Fix paths/rules before running the workflow.")

    st.divider()

    # --- Add Step button ---
    if st.button("Add Step", key="is_add_step"):
        cfgs = [_get_class_cfg(i) for i in range(len(st.session_state.is_classes))]

        # Validate class names
        names = [c["class_name"] for c in cfgs if c["class_name"]]
        if not names:
            st.warning("Please give at least one file class a name.")
            return

        # Register files immediately for current session dropdowns (best-effort)
        # Mode-E (user upload) classes have no path yet — skip registration now
        for cfg in cfgs:
            if cfg.get("mode") == "E":
                continue
            result = resolve_file_class(cfg)
            if result["file_path"] and result["class_name"]:
                try:
                    register_file_path(result["file_path"], display_name=result["class_name"])
                    if result["sheet_name"]:
                        # Store associated sheet so dropdowns can pre-select it
                        if "file_sheet_hints"not in st.session_state:
                            st.session_state.file_sheet_hints = {}
                        st.session_state.file_sheet_hints[result["class_name"]] = result["sheet_name"]
                except Exception:
                    pass

        step_names = ", ".join(names)
        add_step(
            "input_source",
            f"Input Source: {step_names}",
            {"file_classes": cfgs},
        )


def render_normalize_import_form():
    """
    Normalize Headers & Import step.
    Maps a registered file class's messy headers to standard names,
    then appends the clean data to a target working file.
    """
    st.markdown('<div class="pkf-section">Normalize Headers & Import</div>', unsafe_allow_html=True)
    st.caption(
        "Pick a source file class, define how its headers should be renamed, "
        "then append the standardized rows to a working file."
    )

    registered = get_files()  # ["", "Bank Statement", "ERP Ledger", ...]

    # ── Source ──────────────────────────────────────────────────────────────
    st.markdown('<div class="pkf-section">1 · Source File Class</div>', unsafe_allow_html=True)
    src_file = st.selectbox(
        "Source file class",
        options=registered,
        key="ni_src_file",
        help="Select a file class registered via Input Source or Add File(s).",
    )
    src_sheet = sheet_selectbox("Source sheet", src_file, key="ni_src_sheet")
    src_header_row = st.number_input("Header row in source", min_value=1, value=1, key="ni_src_hrow")

    # Auto-detect button
    if st.button("Detect Headers", key="ni_detect") and src_file:
        src_path = get_file_path(src_file)
        headers = detect_headers(src_path, src_sheet, int(src_header_row))
        st.session_state["ni_detected_headers"] = headers

    detected = st.session_state.get("ni_detected_headers", [])
    if detected:
        st.success(f"Detected {len(detected)} headers: {', '.join(detected[:10])}{'…'if len(detected) > 10 else ''}")

    st.divider()

    # ── Mapping configuration ────────────────────────────────────────────────
    st.markdown('<div class="pkf-section">2 · Header Mapping</div>', unsafe_allow_html=True)

    mapping_source = st.radio(
        "Mapping source",
        ["Build manually", "Upload mapping file", "Use saved template"],
        key="ni_map_source",
        horizontal=True,
    )

    mapping_rows = []  # will be populated below

    # ── A: Manual builder ───────────────────────────────────────────────────
    if mapping_source == "Build manually":
        if "ni_rows"not in st.session_state:
            st.session_state["ni_rows"] = 1

        col_add, col_clear = st.columns([1, 1])
        with col_add:
            if st.button("Add Row", key="ni_add_row"):
                st.session_state["ni_rows"] += 1
        with col_clear:
            if st.button("Clear All Rows", key="ni_clear_rows"):
                st.session_state["ni_rows"] = 1
                for k in list(st.session_state.keys()):
                    if k.startswith("ni_row_"):
                        del st.session_state[k]

        n = st.session_state["ni_rows"]
        src_options = [""] + detected if detected else [""]

        header_cols = st.columns([2, 2, 1, 1, 3])
        header_cols[0].markdown("**Source Header**")
        header_cols[1].markdown("**Target Header**")
        header_cols[2].markdown("**Required**")
        header_cols[3].markdown("**Data Type**")
        header_cols[4].markdown("**Alternate Headers** *(comma-sep)*")

        for i in range(n):
            c0, c1, c2, c3, c4 = st.columns([2, 2, 1, 1, 3])
            if detected:
                src_val = st.session_state.get(f"ni_row_{i}_src", "")
                src_idx = src_options.index(src_val) if src_val in src_options else 0
                c0.selectbox("", src_options, index=src_idx, key=f"ni_row_{i}_src", label_visibility="collapsed")
            else:
                c0.text_input("", key=f"ni_row_{i}_src", placeholder="Source Header", label_visibility="collapsed")
            c1.text_input("", key=f"ni_row_{i}_tgt", placeholder="Target Header", label_visibility="collapsed")
            c2.checkbox("", key=f"ni_row_{i}_req", label_visibility="collapsed")
            c3.selectbox("", ["", "text", "number", "date"], key=f"ni_row_{i}_dtype", label_visibility="collapsed")
            c4.text_input("", key=f"ni_row_{i}_alts", placeholder="Date, Value Date, Posting Date", label_visibility="collapsed")

        for i in range(n):
            src = str(st.session_state.get(f"ni_row_{i}_src", "") or "").strip()
            tgt = str(st.session_state.get(f"ni_row_{i}_tgt", "") or "").strip()
            if src and tgt:
                alts_raw = str(st.session_state.get(f"ni_row_{i}_alts", "") or "")
                alts = [a.strip() for a in alts_raw.split(",") if a.strip()]
                mapping_rows.append({
                    "source_header": src,
                    "target_header": tgt,
                    "required": bool(st.session_state.get(f"ni_row_{i}_req", False)),
                    "data_type": str(st.session_state.get(f"ni_row_{i}_dtype", "") or ""),
                    "alternate_headers": alts,
                    "default_value": "",
                })

    # ── B: Upload mapping file ───────────────────────────────────────────────
    elif mapping_source == "Upload mapping file":
        st.caption(
            "Upload an Excel or CSV with columns: **Source Header**, **Target Header**, "
            "Required, Data Type, Alternate Source Headers, Default Value"
        )
        map_file = st.file_uploader(
            "Mapping file (.xlsx / .csv)",
            type=["xlsx", "xls", "csv"],
            key="ni_map_upload",
        )
        if map_file:
            rows, err = load_mapping_from_bytes(map_file.getvalue(), map_file.name)
            if err:
                st.error(f"Could not parse mapping file: {err}")
            else:
                mapping_rows = rows
                st.success(f"Loaded {len(rows)} mapping rule(s) from {map_file.name}")
                st.session_state["ni_uploaded_mapping"] = rows
        elif "ni_uploaded_mapping"in st.session_state:
            mapping_rows = st.session_state["ni_uploaded_mapping"]

    # ── C: Saved template ───────────────────────────────────────────────────
    elif mapping_source == "Use saved template":
        templates = load_all_templates()
        if not templates:
            st.info("No saved templates yet. Build a mapping manually and save it as a template below.")
        else:
            chosen = st.selectbox("Template", [""] + list(templates.keys()), key="ni_tpl_select")
            if chosen:
                tpl = templates[chosen]
                mapping_rows = tpl.get("mapping", [])
                st.success(
                    f"Loaded **{chosen}** — {len(mapping_rows)} rule(s), "
                    f"originally for class: *{tpl.get('source_class', '—')}*"
                )
                col_del, _ = st.columns([1, 3])
                if col_del.button("Delete this template", key="ni_tpl_delete"):
                    delete_template(chosen)
                    st.success(f"Deleted template '{chosen}'")
                    st.rerun()

    # ── Preview / Validate ──────────────────────────────────────────────────
    if mapping_rows and detected:
        st.divider()
        st.markdown('<div class="pkf-section">Preview</div>', unsafe_allow_html=True)
        from operations.normalize_import import apply_header_mapping
        import pandas as pd
        dummy_df = pd.DataFrame(columns=detected)
        _, report = apply_header_mapping(dummy_df, mapping_rows)

        st.dataframe(
            pd.DataFrame(report)[["Source Header", "Detected As", "Target Header", "Required", "Data Type", "Status", "Note"]],
            use_container_width=True,
            hide_index=True,
        )
        missing_req = [r["Source Header"] for r in report if r["Status"] == ""]
        if missing_req:
            st.error(f"Required columns missing from detected headers: {', '.join(missing_req)}")
        else:
            mapped_ok = sum(1 for r in report if r["Status"] == "")
            st.success(f"{mapped_ok}/{len(report)} columns mapped successfully.")
    elif mapping_rows and not detected:
        st.info("Click **Detect Headers** above to preview the mapping against actual file headers.")

    # ── Save as template ────────────────────────────────────────────────────
    if mapping_rows:
        with st.expander("Save as reusable template", expanded=False):
            tpl_name = st.text_input("Template name", key="ni_save_tpl_name", placeholder="HDFC Bank Mapping")
            tpl_class = st.text_input(
                "Source class (for reference)",
                key="ni_save_tpl_class",
                value=src_file or "",
                placeholder="Bank Statement",
            )
            if st.button("Save Template", key="ni_save_tpl_btn"):
                if not tpl_name.strip():
                    st.warning("Enter a template name.")
                else:
                    ok, err = save_template(tpl_name.strip(), tpl_class.strip(), mapping_rows)
                    if ok:
                        st.success(f"Template '{tpl_name}'saved.")
                    else:
                        st.error(f"Save failed: {err}")

    st.divider()

    # ── Target ──────────────────────────────────────────────────────────────
    st.markdown('<div class="pkf-section">3 · Target Working File</div>', unsafe_allow_html=True)
    tgt_file = st.selectbox(
        "Target file (working file to append into)",
        options=registered,
        key="ni_tgt_file",
    )
    tgt_sheet = sheet_selectbox("Target sheet", tgt_file, key="ni_tgt_sheet")
    tgt_header_row = st.number_input("Header row in target", min_value=1, value=1, key="ni_tgt_hrow")
    append_mode = st.checkbox("Append rows (uncheck to replace all data)", value=True, key="ni_append")

    export = export_options_ui("normalize_import")

    # ── Add Step ─────────────────────────────────────────────────────────────
    if st.button("Add Step", key="ni_add_step"):
        if not src_file:
            st.warning("Select a source file class.")
            return
        if not tgt_file:
            st.warning("Select a target working file.")
            return
        if not mapping_rows:
            st.warning("Define at least one mapping row.")
            return
        add_step(
            "normalize_import",
            f"Normalize & Import: {src_file} → {tgt_file}",
            {
                "src_file": src_file,
                "src_sheet": src_sheet,
                "src_header_row": int(src_header_row),
                "mapping": mapping_rows,
                "tgt_file": tgt_file,
                "tgt_sheet": tgt_sheet,
                "tgt_header_row": int(tgt_header_row),
                "append_mode": append_mode,
                "export": export,
            },
        )


def render_folder_summary_form():
    """
    Folder Summary step.
    Three source modes:
      • Folder  — many files × one sheet name
      • Sheets  — one file × many sheets
      • Both    — many files × many sheets per file
    """
    st.markdown('<div class="pkf-section">Folder Summary</div>', unsafe_allow_html=True)
    st.caption(
        "Extract labeled figures by CTRL-F-style search, evaluate BODMAS formulas, "
        "and write one consolidated row per source (file / sheet / file+sheet) into a summary workbook."
    )

    registered = get_files()

    # ── 1 · Source mode + file selection ────────────────────────────────────
    st.markdown('<div class="pkf-section">1 · Source</div>', unsafe_allow_html=True)

    SOURCE_MODE_OPTIONS = [
        "Folder  (many files × one sheet)",
        "Sheets  (one file × many sheets)",
        "Both    (many files × many sheets per file)",
    ]
    SOURCE_MODE_KEYS = ["folder", "sheets", "both"]
    _saved_mode = st.session_state.get("fs_source_mode", "folder")
    _saved_mode_idx = SOURCE_MODE_KEYS.index(_saved_mode) if _saved_mode in SOURCE_MODE_KEYS else 0
    _mode_choice = st.radio(
        "Source mode", SOURCE_MODE_OPTIONS, index=_saved_mode_idx,
        key="fs_source_mode_radio", horizontal=True,
        help=(
            "**Folder**: one sheet name across many files (original behaviour)  \n"
            "**Sheets**: one file, iterate over its sheets  \n"
            "**Both**: iterate files AND sheets per file"
        ),
    )
    source_mode = SOURCE_MODE_KEYS[SOURCE_MODE_OPTIONS.index(_mode_choice)]

    src_file = st.selectbox(
        "Source file class"if source_mode != "sheets"else "Source file (single)",
        options=registered, key="fs_src_file",
        help=(
            "For **Folder/Both**: use a folder-scan class (Input Source Mode B/C/D) so the full file group is processed.  \n"
            "For **Sheets**: pick a single registered file."
        ),
    )
    src_header_row = st.number_input(
        "Header row in source sheet(s)", min_value=1, value=1, key="fs_src_hrow",
        help="Row that contains the column headers in each source sheet.",
    )

    # Sheet name: fixed (folder mode) vs selection criteria (sheets/both mode)
    if source_mode == "folder":
        src_sheet = sheet_selectbox("Source sheet name (same in every file)", src_file, key="fs_src_sheet")
        sheet_selection = "all"
        sheet_pattern = ""
        sheet_list_raw = ""
        sheet_exclude = ""
    else:
        src_sheet = ""   # not used
        SHEET_SEL_OPTS = ["all", "contains", "starts_with", "ends_with", "regex", "list"]
        _ss_default = st.session_state.get("fs_sheet_selection", "all")
        _ss_idx = SHEET_SEL_OPTS.index(_ss_default) if _ss_default in SHEET_SEL_OPTS else 0
        sheet_selection = st.selectbox(
            "Which sheets to include", SHEET_SEL_OPTS, index=_ss_idx, key="fs_sheet_selection",
            help="**all** = every sheet; others filter by name.",
        )
        if sheet_selection == "list":
            sheet_list_raw = st.text_input(
                "Sheet names (comma-separated)", key="fs_sheet_list_raw",
                placeholder="Jan, Feb, Mar",
            )
            sheet_pattern = ""
        elif sheet_selection == "all":
            sheet_list_raw = ""
            sheet_pattern = ""
        else:
            sheet_pattern = st.text_input(
                f"Pattern ({sheet_selection})", key="fs_sheet_pattern",
                placeholder="e.g.  RECON  or  ^APR",
            )
            sheet_list_raw = ""

        sheet_exclude = st.text_input(
            "Exclude sheets containing (comma-sep)",
            key="fs_sheet_exclude",
            placeholder="Summary, Master, Template",
            help="Sheets whose name contains any of these substrings will be skipped.",
        )

        # Live preview button
        if st.button("Preview matching sheets", key="fs_preview_sheets") and src_file:
            from operations.folder_summary import _get_sheet_names, _filter_sheets
            fp = get_file_path(src_file)
            if fp:
                all_sh = _get_sheet_names(fp)
                sl = [s.strip() for s in sheet_list_raw.split(",") if s.strip()] if sheet_selection == "list"else []
                matched = _filter_sheets(all_sh, sheet_selection, sheet_pattern, sl, sheet_exclude)
                st.success(f"**{len(matched)} sheet(s) matched:** {', '.join(matched)}")
            else:
                st.warning("File not found — register it first.")

    st.divider()

    # ── 2 · Variables (CTRL-F extractions) ──────────────────────────────────
    st.markdown('<div class="pkf-section">2 · Variables  *(what to search for)*</div>', unsafe_allow_html=True)
    st.caption(
        "Each variable searches a **particulars column** for a keyword and reads the adjacent value. "
        "The variable **name** (single letter or short word) is used in the formula below."
    )

    MATCH_TYPES = ["icontains", "exact", "iexact", "contains", "starts_with", "ends_with", "regex"]
    MULTI_MATCH = ["first", "last", "sum", "avg", "max", "min"]
    ON_MISSING = ["zero", "blank", "flag", "error", "default_value"]

    if "fs_var_count"not in st.session_state:
        st.session_state["fs_var_count"] = 1

    col_add_v, col_rem_v = st.columns([1, 1])
    with col_add_v:
        if st.button("Add Variable", key="fs_add_var"):
            st.session_state["fs_var_count"] += 1
    with col_rem_v:
        if st.button("Remove Last", key="fs_rem_var") and st.session_state["fs_var_count"] > 1:
            st.session_state["fs_var_count"] -= 1

    n_vars = st.session_state["fs_var_count"]

    # Header row for the variable table
    hc = st.columns([1, 2, 2, 1, 1, 1, 1, 1, 2])
    for lbl, col in zip(
        ["Name*", "Search Text*", "Label (output col)", "Search Col", "Value Col",
         "Match Type", "If Multiple", "On Missing", "Missing Default / Note"],
        hc
    ):
        col.markdown(f"**{lbl}**")

    for i in range(n_vars):
        c = st.columns([1, 2, 2, 1, 1, 1, 1, 1, 2])
        c[0].text_input(f"Name {i+1}", key=f"fs_var_{i}_name", placeholder="A", label_visibility="collapsed")
        c[1].text_input(f"Search {i+1}", key=f"fs_var_{i}_search_text", placeholder="Gross Subscription", label_visibility="collapsed")
        c[2].text_input(f"Label {i+1}", key=f"fs_var_{i}_label", placeholder="Gross Subscription", label_visibility="collapsed")
        c[3].text_input(f"SCol {i+1}", key=f"fs_var_{i}_search_col", value="A", label_visibility="collapsed", help="Column letter or 1-based number for the particulars/label column")
        c[4].text_input(f"VCol {i+1}", key=f"fs_var_{i}_value_col", value="D", label_visibility="collapsed", help="Column letter or 1-based number for the figures")
        _mt_idx = MATCH_TYPES.index(st.session_state.get(f"fs_var_{i}_match_type", "icontains")) if st.session_state.get(f"fs_var_{i}_match_type", "icontains") in MATCH_TYPES else 0
        c[5].selectbox(f"Match {i+1}", MATCH_TYPES, index=_mt_idx, key=f"fs_var_{i}_match_type", label_visibility="collapsed")
        _mm_idx = MULTI_MATCH.index(st.session_state.get(f"fs_var_{i}_multi_match", "first")) if st.session_state.get(f"fs_var_{i}_multi_match", "first") in MULTI_MATCH else 0
        c[6].selectbox(f"Multi {i+1}", MULTI_MATCH, index=_mm_idx, key=f"fs_var_{i}_multi_match", label_visibility="collapsed")
        _om_idx = ON_MISSING.index(st.session_state.get(f"fs_var_{i}_on_missing", "zero")) if st.session_state.get(f"fs_var_{i}_on_missing", "zero") in ON_MISSING else 0
        c[7].selectbox(f"Missing {i+1}", ON_MISSING, index=_om_idx, key=f"fs_var_{i}_on_missing", label_visibility="collapsed")
        c[8].text_input(f"Default {i+1}", key=f"fs_var_{i}_default", placeholder="0 or leave blank", label_visibility="collapsed")

    with st.expander("Advanced: Row Offset & Include column"):
        adv_cols = st.columns([1, 1, 1])
        adv_cols[0].markdown("**Variable**")
        adv_cols[1].markdown("**Row Offset** *(+N = N rows below match)*")
        adv_cols[2].markdown("**Include as output column?**")
        for i in range(n_vars):
            vname = st.session_state.get(f"fs_var_{i}_name", f"V{i+1}") or f"V{i+1}"
            ac = st.columns([1, 1, 1])
            ac[0].markdown(f"`{vname}`")
            ac[1].number_input(f"Offset {i+1}", key=f"fs_var_{i}_row_offset", value=0, step=1, label_visibility="collapsed")
            ac[2].checkbox(f"Include {i+1}", key=f"fs_var_{i}_include", value=True, label_visibility="collapsed")

    st.divider()

    # ── 3 · Formulas (BODMAS) ────────────────────────────────────────────────
    st.markdown('<div class="pkf-section">3 · Formulas  *(BODMAS calculations)*</div>', unsafe_allow_html=True)
    st.caption(
        "Use variable names from section 2. Standard Python math operators: `+  -  *  /  (  )  **  abs()  round()`  "
        "Example: `(A - B - C) / D * 100`"
    )

    if "fs_formula_count"not in st.session_state:
        st.session_state["fs_formula_count"] = 1

    col_add_f, col_rem_f = st.columns([1, 1])
    with col_add_f:
        if st.button("Add Formula", key="fs_add_formula"):
            st.session_state["fs_formula_count"] += 1
    with col_rem_f:
        if st.button("Remove Last", key="fs_rem_formula") and st.session_state["fs_formula_count"] > 1:
            st.session_state["fs_formula_count"] -= 1

    n_formulas = st.session_state["fs_formula_count"]
    fh = st.columns([2, 4, 1])
    fh[0].markdown("**Output Column Name**")
    fh[1].markdown("**Expression**")
    fh[2].markdown("**On Error**")
    for i in range(n_formulas):
        fc = st.columns([2, 4, 1])
        fc[0].text_input(f"Formula name {i+1}", key=f"fs_formula_{i}_name", placeholder="Net Amount", label_visibility="collapsed")
        fc[1].text_input(f"Expression {i+1}", key=f"fs_formula_{i}_expr", placeholder="A - B - C", label_visibility="collapsed")
        _oe_opts = ["flag", "zero", "blank"]
        _oe_idx = _oe_opts.index(st.session_state.get(f"fs_formula_{i}_on_error", "flag")) if st.session_state.get(f"fs_formula_{i}_on_error", "flag") in _oe_opts else 0
        fc[2].selectbox(f"On error {i+1}", _oe_opts, index=_oe_idx, key=f"fs_formula_{i}_on_error", label_visibility="collapsed")

    st.divider()

    # ── 4 · Target ───────────────────────────────────────────────────────────
    st.markdown('<div class="pkf-section">4 · Target (Consolidated Output)</div>', unsafe_allow_html=True)
    tgt_file = st.selectbox("Target file", options=registered, key="fs_tgt_file",
                            help="The summary workbook. Must be a registered .xlsx file.")
    tgt_sheet = st.text_input("Target sheet name", value="Summary", key="fs_tgt_sheet")
    tgt_header_row = st.number_input("Header row in target", min_value=1, value=1, key="fs_tgt_hrow")

    col_a, col_b = st.columns(2)
    with col_a:
        append_mode = st.checkbox("Append rows (keep existing data)", value=True, key="fs_append_mode",
                                  help="If unchecked, existing data rows in the target sheet are cleared before writing.")
    with col_b:
        _id_col_default = {
            "folder": "File Name",
            "sheets": "Sheet Name",
            "both": "Source",
        }.get(source_mode, "File Name")
        _id_col_help = {
            "folder": "Column header for the source filename.",
            "sheets": "Column header for the sheet name.",
            "both": "Column header showing 'filename — sheet'.",
        }.get(source_mode, "Identifier column.")
        filename_col = st.text_input(
            "Identifier column header", value=_id_col_default,
            key="fs_filename_col", help=_id_col_help + "Leave blank to skip.",
        )

    # Optional sheet name column (only meaningful in sheets/both modes, but always available)
    _sn_col_default = st.session_state.get("fs_sheet_name_col",
        "Sheet Name"if source_mode in ("sheets", "both") else "")
    sheet_name_col = st.text_input(
        "Sheet name column header *(optional)*",
        value=_sn_col_default,
        key="fs_sheet_name_col",
        help=(
            "When set, a separate column with this header will record the **sheet name** for each row. "
            "Useful in **Sheets** and **Both** modes. Leave blank to disable."
        ),
    )

    col_c, col_d = st.columns(2)
    with col_c:
        include_var_cols = st.checkbox("Include individual variable columns", value=True, key="fs_incl_vars",
                                       help="Write each variable's extracted value as a separate column.")
    with col_d:
        include_status_col = st.checkbox("Include status column", value=True, key="fs_incl_status")

    status_col_name = st.text_input("Status column header", value="Status", key="fs_status_col_name",
                                    disabled=not st.session_state.get("fs_incl_status", True))

    st.divider()

    # ── Add Step ──────────────────────────────────────────────────────────────
    if st.button("Add Step", key="fs_add_step", type="primary"):
        if not src_file:
            st.error("Select a source file class.")
            return
        if not tgt_file:
            st.error("Select a target file.")
            return
        if source_mode == "folder"and not src_sheet:
            st.error("Select a source sheet.")
            return

        # Collect variables
        variables = []
        for i in range(n_vars):
            vname = str(st.session_state.get(f"fs_var_{i}_name", "") or "").strip()
            search_text = str(st.session_state.get(f"fs_var_{i}_search_text", "") or "").strip()
            if not vname or not search_text:
                continue
            default_raw = str(st.session_state.get(f"fs_var_{i}_default", "") or "").strip()
            try:
                default_val = float(default_raw) if default_raw else 0
            except ValueError:
                default_val = 0
            variables.append({
                "name": vname,
                "label": str(st.session_state.get(f"fs_var_{i}_label", "") or vname).strip(),
                "search_text": search_text,
                "search_col": str(st.session_state.get(f"fs_var_{i}_search_col", "A") or "A").strip(),
                "value_col": str(st.session_state.get(f"fs_var_{i}_value_col", "D") or "D").strip(),
                "match_type": str(st.session_state.get(f"fs_var_{i}_match_type", "icontains") or "icontains"),
                "multi_match": str(st.session_state.get(f"fs_var_{i}_multi_match", "first") or "first"),
                "row_offset": int(st.session_state.get(f"fs_var_{i}_row_offset", 0) or 0),
                "on_missing": str(st.session_state.get(f"fs_var_{i}_on_missing", "zero") or "zero"),
                "missing_default": default_val,
                "include": bool(st.session_state.get(f"fs_var_{i}_include", True)),
            })

        if not variables:
            st.error("Add at least one variable with a name and search text.")
            return

        # Collect formulas
        formulas = []
        for i in range(n_formulas):
            fname = str(st.session_state.get(f"fs_formula_{i}_name", "") or "").strip()
            expr = str(st.session_state.get(f"fs_formula_{i}_expr", "") or "").strip()
            if not fname or not expr:
                continue
            formulas.append({
                "name": fname,
                "expression": expr,
                "on_error": str(st.session_state.get(f"fs_formula_{i}_on_error", "flag") or "flag"),
            })

        var_names = ", ".join(v["name"] for v in variables)
        formula_names = ", ".join(f["name"] for f in formulas)
        mode_label = {"folder": "files", "sheets": "sheets", "both": "files×sheets"}.get(source_mode, "")
        src_desc = src_sheet if source_mode == "folder"else f"{sheet_selection}({sheet_pattern or sheet_list_raw or 'all'})"
        label_parts = [f"Src: {src_file}", f"{mode_label}: {src_desc}", f"Vars: {var_names}"]
        if formula_names:
            label_parts.append(f"Formulas: {formula_names}")
        label_parts.append(f"→ {tgt_file}")

        sheet_list_parsed = [s.strip() for s in sheet_list_raw.split(",") if s.strip()] \
            if sheet_selection == "list"else []

        add_step(
            "folder_summary",
            "Folder Summary — " + " | ".join(label_parts),
            {
                "src_file": src_file,
                "src_sheet": src_sheet,            # used in folder mode
                "header_row": int(src_header_row),
                "source_mode": source_mode,
                "sheet_selection": sheet_selection,
                "sheet_pattern": sheet_pattern,
                "sheet_list": sheet_list_parsed,
                "sheet_exclude": sheet_exclude,
                "variables": variables,
                "formulas": formulas,
                "tgt_file": tgt_file,
                "tgt_sheet": tgt_sheet,
                "tgt_header_row": int(tgt_header_row),
                "filename_col": filename_col.strip(),
                "sheet_name_col": sheet_name_col.strip(),
                "append_mode": append_mode,
                "include_var_cols": include_var_cols,
                "include_status_col": include_status_col,
                "status_col_name": status_col_name.strip(),
                "on_missing_default": 0,
            },
        )


def render_split_export_form():
    """Split Export — slice a source sheet and export to files/sheets by column values."""
    from core import wb_cache as _wbc

    st.markdown('<div class="pkf-section">Split Export</div>', unsafe_allow_html=True)
    st.caption(
        "Slice your source data and export it — one file, multiple sheets, multiple files, "
        "or a two-level split.  Name files/sheets using column values, dates, and sequence numbers."
    )

    # ── Source ──────────────────────────────────────────────────────────────
    st.markdown("**Source Data  *(the file you are splitting)***")
    s1, s2, s3 = st.columns([3, 3, 1])
    with s1:
        src_file = st.selectbox("Source File", get_files(), key="se_src_file")
    with s2:
        src_sheet = sheet_selectbox("Source Sheet  *(which sheet holds the data to split)*",
                                     src_file, key="se_src_sheet")
    with s3:
        src_hrow = st.number_input("Header Row", min_value=1,
                                    value=int(st.session_state.get("se_src_header_row", 1) or 1),
                                    key="se_src_header_row")

    # Detect columns from source
    _se_detect_ctx = (src_file or "", src_sheet or "", int(src_hrow))
    if st.session_state.get("se_detect_ctx") != _se_detect_ctx:
        st.session_state["se_detect_ctx"] = _se_detect_ctx
        st.session_state["se_col_list"] = []
        if src_file and src_sheet:
            _fp = get_file_path(src_file)
            if _fp:
                try:
                    _df = _wbc.read_excel(_fp, sheet_name=src_sheet, header=int(src_hrow) - 1)
                    import pandas as _pd
                    if isinstance(_df.columns, _pd.MultiIndex):
                        _df.columns = [" ".join(str(p) for p in c
                                                if str(p) not in ("", "nan", "None")).strip()
                                       for c in _df.columns]
                    st.session_state["se_col_list"] = [str(c).strip() for c in _df.columns]
                except Exception:
                    st.session_state["se_col_list"] = []

    _se_cols = st.session_state.get("se_col_list", [])

    st.divider()

    # ── Mode ─────────────────────────────────────────────────────────────────
    st.markdown("**Export Mode**")
    _se_mode_idx = int(st.session_state.get("se_split_mode_idx", 0))
    split_mode = st.selectbox(
        "Mode", SPLIT_MODE_LABELS,
        index=_se_mode_idx,
        key="se_split_mode_label",
    )
    split_mode_key = SPLIT_MODE_KEYS[SPLIT_MODE_LABELS.index(split_mode)]

    # Primary split column (needed for all modes except single)
    split_col = ""
    split_col2 = ""
    if split_mode_key != "single":
        _sc_opts = [""] + _se_cols
        _sc_prev = st.session_state.get("se_split_col", "")
        _sc_idx  = _sc_opts.index(_sc_prev) if _sc_prev in _sc_opts else 0
        split_col = st.selectbox(
            "Primary Split Column  (group by this column)",
            _sc_opts, index=_sc_idx, key="se_split_col",
        )

    if split_mode_key == "files_sheets":
        _sc2_opts = [""] + _se_cols
        _sc2_prev = st.session_state.get("se_split_col2", "")
        _sc2_idx  = _sc2_opts.index(_sc2_prev) if _sc2_prev in _sc2_opts else 0
        split_col2 = st.selectbox(
            "Sheet-Level Split Column  (secondary grouping within each file)",
            _sc2_opts, index=_sc2_idx, key="se_split_col2",
        )

    st.divider()

    # ── Output ───────────────────────────────────────────────────────────────
    st.markdown("**Output Settings**")

    if split_mode_key in ("single", "sheets"):
        if split_mode_key == "sheets":
            st.info(
                "📄  **No target sheet to select.**  "
                "Sheets are created automatically — one per unique value in the split column.  "
                "Name them using the Sheet Name Template below (e.g. `{value}` → sheet named after the category value)."
            )

        # Target file
        tgt_file = st.selectbox(
            "Target File  *(the output .xlsx to write sheets into)*" if split_mode_key == "sheets"
            else "Target File (registered)",
            [""] + get_files(),
            key="se_tgt_file",
        )
        tgt_file_path_direct = st.text_input(
            "Or enter absolute path for output file  (use this when creating a brand-new file not yet registered)",
            value=st.session_state.get("se_tgt_file_path", ""),
            key="se_tgt_file_path",
            placeholder="/Users/you/Documents/output.xlsx",
        )
        if split_mode_key == "single":
            tgt_sheet_val = st.text_input(
                "Target Sheet Name",
                value=st.session_state.get("se_tgt_sheet", "Export"),
                key="se_tgt_sheet",
            )
        else:
            tgt_sheet_val = ""
        output_folder = ""
    else:
        # Multiple files → folder
        tgt_file = ""
        tgt_file_path_direct = ""
        tgt_sheet_val = ""
        output_folder = st.text_input(
            "Output Folder  (files will be created here)",
            value=st.session_state.get("se_output_folder", ""),
            key="se_output_folder",
            placeholder="/Users/you/Documents/Exports/",
        )

    st.markdown("**Naming Templates**")
    st.caption(
        "Available tokens: `{ColumnName}` — value of any source column, "
        "`{value}` — primary split value, `{date}` — today's date, "
        "`{seq}` or `{n}` — sequence number."
    )

    nm1, nm2 = st.columns(2)
    if split_mode_key in ("files", "files_sheets"):
        file_tmpl = nm1.text_input("File Name Template",
                                    value=st.session_state.get("se_file_tmpl", "{value}.xlsx"),
                                    key="se_file_tmpl",
                                    placeholder="{value}.xlsx")
    else:
        file_tmpl = "{value}.xlsx"

    if split_mode_key != "single":
        sheet_tmpl = nm2.text_input("Sheet Name Template  (level 1)",
                                     value=st.session_state.get("se_sheet_tmpl", "{value}"),
                                     key="se_sheet_tmpl",
                                     placeholder="{value}")
    else:
        sheet_tmpl = "{value}"

    sheet_tmpl2 = "{value}"
    if split_mode_key == "files_sheets":
        sheet_tmpl2 = st.text_input("Sheet Name Template  (level 2 — sheet inside each file)",
                                     value=st.session_state.get("se_sheet_tmpl2", "{value}"),
                                     key="se_sheet_tmpl2",
                                     placeholder="{value}")

    st.divider()

    # ── Pre-export Filters ────────────────────────────────────────────────────
    with st.expander("Pre-Export Filters  (optional — narrow rows before splitting)", expanded=False):
        se_filters, se_combine = render_filter_builder("se_filter", _se_cols)

    # ── Column Selection ──────────────────────────────────────────────────────
    with st.expander("Column Selection  (optional — include or exclude columns)", expanded=False):
        cs1, cs2 = st.columns(2)
        incl_raw = cs1.text_input("Include only  (comma-separated column names, blank = all)",
                                   value=st.session_state.get("se_include_cols", ""),
                                   key="se_include_cols",
                                   placeholder="ColA, ColB, ColC")
        excl_raw = cs2.text_input("Exclude  (comma-separated column names)",
                                   value=st.session_state.get("se_exclude_cols", ""),
                                   key="se_exclude_cols",
                                   placeholder="ColX, ColY")
        include_cols = [c.strip() for c in incl_raw.split(",") if c.strip()]
        exclude_cols = [c.strip() for c in excl_raw.split(",") if c.strip()]

    # ── Sorting ───────────────────────────────────────────────────────────────
    with st.expander("Sorting  (optional)", expanded=False):
        srt1, srt2 = st.columns([3, 2])
        sort_cols_raw = srt1.text_input("Sort by  (comma-separated column names)",
                                         value=st.session_state.get("se_sort_cols", ""),
                                         key="se_sort_cols",
                                         placeholder="DateCol, AmountCol")
        sort_dir = srt2.selectbox("Direction",
                                   ["Ascending", "Descending"],
                                   index=0 if st.session_state.get("se_sort_asc", "Ascending") == "Ascending" else 1,
                                   key="se_sort_asc")
        sort_cols = [c.strip() for c in sort_cols_raw.split(",") if c.strip()]
        sort_ascending = (sort_dir == "Ascending")

    st.divider()

    # ── Options ───────────────────────────────────────────────────────────────
    st.markdown("**Options**")
    op1, op2, op3 = st.columns(3)
    include_header = op1.checkbox("Include header row",
                                   value=bool(st.session_state.get("se_include_header", True)),
                                   key="se_include_header")
    overwrite = op2.checkbox("Overwrite existing sheets/files",
                              value=bool(st.session_state.get("se_overwrite", True)),
                              key="se_overwrite")
    max_rows = op3.number_input("Max rows per sheet  (0 = unlimited)",
                                 min_value=0,
                                 value=int(st.session_state.get("se_max_rows", 0) or 0),
                                 step=1000,
                                 key="se_max_rows")

    add_summary = st.checkbox("Add summary/index sheet",
                               value=bool(st.session_state.get("se_add_summary", False)),
                               key="se_add_summary")
    summary_name = "_Summary"
    if add_summary:
        summary_name = st.text_input("Summary sheet name",
                                      value=st.session_state.get("se_summary_name", "_Summary"),
                                      key="se_summary_name")

    # ── Formatting ────────────────────────────────────────────────────────────
    with st.expander("Header Formatting", expanded=False):
        fmtc1, fmtc2 = st.columns(2)
        apply_fmt = fmtc1.checkbox("Apply formatting",
                                    value=bool(st.session_state.get("se_apply_fmt", True)),
                                    key="se_apply_fmt")
        if apply_fmt:
            hbg = fmtc1.color_picker("Header background",
                                      value=st.session_state.get("se_header_bg", "#1F3864"),
                                      key="se_header_bg")
            hfg = fmtc2.color_picker("Header font color",
                                      value=st.session_state.get("se_header_fg", "#FFFFFF"),
                                      key="se_header_fg")
        else:
            hbg = st.session_state.get("se_header_bg", "#1F3864")
            hfg = st.session_state.get("se_header_fg", "#FFFFFF")

    # ── Export option ─────────────────────────────────────────────────────────
    export = st.checkbox("Export after run  (save a non-destructive copy)",
                          value=False, key="se_export_flag")

    # ── Add Step ──────────────────────────────────────────────────────────────
    if st.button("Add Step", key="se_add", type="primary"):
        if not src_file:
            st.error("Select a source file (the file whose data you want to split).")
            return
        if not src_sheet:
            st.error(
                "Select the **Source Sheet** — the sheet inside your source file that contains "
                "the data you want to split.  (This is NOT the output sheet — output sheets are "
                "auto-created from the column values.)"
            )
            return
        if split_mode_key != "single" and not split_col:
            st.error("Select a Primary Split Column.")
            return
        if split_mode_key == "files_sheets" and not split_col2:
            st.error("Select a Sheet-Level Split Column.")
            return
        if split_mode_key in ("single", "sheets"):
            _eff_tgt = (get_file_path(tgt_file) if tgt_file else "") or tgt_file_path_direct.strip()
            if not _eff_tgt:
                st.error("Select a target file or enter an output file path.")
                return
        else:
            if not output_folder.strip():
                st.error("Enter an output folder path.")
                return

        # Build tgt_file value — prefer registered label, fall back to direct path
        _tgt_label = tgt_file or ""
        _tgt_path_val = tgt_file_path_direct.strip() if not _tgt_label else ""

        # Human-readable label
        _mode_short = {
            "single": "→ 1 sheet",
            "sheets": "→ N sheets",
            "files": "→ N files",
            "files_sheets": "→ N files × M sheets",
        }.get(split_mode_key, split_mode_key)
        _label_parts = [f"Split Export ({_mode_short})"]
        if split_col:
            _label_parts.append(f"by {split_col}")
        if split_col2:
            _label_parts.append(f"× {split_col2}")
        _label_parts.append(f"from {src_sheet}")
        _label = "  ".join(_label_parts)

        add_step("split_export", _label, {
            "src_file":             src_file,
            "src_sheet":            src_sheet,
            "src_header_row":       int(src_hrow),
            "split_mode":           split_mode_key,
            "split_col":            split_col,
            "split_col2":           split_col2,
            "tgt_file":             _tgt_label,
            "tgt_file_path":        _tgt_path_val,
            "tgt_sheet":            tgt_sheet_val,
            "output_folder":        output_folder.strip(),
            "file_name_template":   file_tmpl,
            "sheet_name_template":  sheet_tmpl,
            "sheet_name_template2": sheet_tmpl2,
            "filter_conditions":    se_filters if se_filters else [],
            "filter_combine":       se_combine,
            "include_columns":      include_cols,
            "exclude_columns":      exclude_cols,
            "sort_cols":            sort_cols,
            "sort_ascending":       sort_ascending,
            "include_header":       include_header,
            "overwrite":            overwrite,
            "max_rows_per_sheet":   int(max_rows),
            "add_summary_sheet":    add_summary,
            "summary_sheet_name":   summary_name,
            "apply_formatting":     apply_fmt,
            "header_bg_color":      hbg.lstrip("#"),
            "header_font_color":    hfg.lstrip("#"),
            "export":               export,
        })


def render_unpivot_form():
    """Unpivot (Wide → Long) — reshape a wide table into long/database format."""
    from core import wb_cache as _wbc

    st.markdown('<div class="pkf-section">Unpivot / Wide → Long</div>', unsafe_allow_html=True)
    st.caption(
        "Turn a wide table *(one column per category)* into long format *(one row per category)*. "
        "Uses `pandas.melt()` internally."
    )

    # Example callout
    with st.expander("📖 Example", expanded=False):
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**Input (wide)**")
            st.dataframe({"Company": ["ABC Ltd", "XYZ Ltd"],
                          "Rent":    [1000, 2000],
                          "Salary":  [5000, 6000],
                          "Travel":  [700,  900]}, use_container_width=True)
        with col_b:
            st.markdown("**Output (long)**")
            st.dataframe({"Company":  ["ABC Ltd","ABC Ltd","ABC Ltd","XYZ Ltd","XYZ Ltd","XYZ Ltd"],
                          "Category": ["Rent","Salary","Travel","Rent","Salary","Travel"],
                          "Value":    [1000,5000,700,2000,6000,900]}, use_container_width=True)

    st.divider()

    # ── Source ────────────────────────────────────────────────────────────────
    st.markdown("**Source  *(wide-format file)***")
    r1c1, r1c2, r1c3 = st.columns([3, 3, 1])
    with r1c1:
        up_src_file = st.selectbox("Source File", get_files(), key="up_src_file")
    with r1c2:
        up_src_sheet = sheet_selectbox("Source Sheet", up_src_file, key="up_src_sheet")
    with r1c3:
        up_src_hrow = st.number_input(
            "Header Row", min_value=1,
            value=int(st.session_state.get("up_src_header_row", 1) or 1),
            key="up_src_header_row",
        )

    # Auto-detect columns
    _up_ctx = (up_src_file or "", up_src_sheet or "", int(up_src_hrow))
    if st.session_state.get("up_detect_ctx") != _up_ctx:
        st.session_state["up_detect_ctx"] = _up_ctx
        st.session_state["up_col_list"]   = []
        if up_src_file and up_src_sheet:
            _fp = get_file_path(up_src_file)
            if _fp:
                try:
                    _df = _wbc.read_excel(_fp, sheet_name=up_src_sheet,
                                          header=int(up_src_hrow) - 1, nrows=0)
                    st.session_state["up_col_list"] = [str(c).strip() for c in _df.columns]
                except Exception:
                    st.session_state["up_col_list"] = []

    _up_cols = st.session_state.get("up_col_list", [])

    st.divider()

    # ── Column selection ──────────────────────────────────────────────────────
    st.markdown("**Column Selection**")

    if _up_cols:
        # ID columns — multiselect
        _saved_id = [c for c in (st.session_state.get("up_id_cols") or []) if c in _up_cols]
        up_id_cols = st.multiselect(
            "ID Columns  *(columns that identify each row — stay as-is)*",
            options=_up_cols,
            default=_saved_id,
            key="up_id_cols",
            help="e.g. Company, Region, Date — these columns are NOT unpivoted",
        )

        # Value columns — multiselect (default = all non-ID)
        _remaining = [c for c in _up_cols if c not in up_id_cols]
        _saved_val = [c for c in (st.session_state.get("up_value_cols") or []) if c in _remaining]
        up_value_cols = st.multiselect(
            "Value Columns  *(columns to unpivot — leave empty = all remaining)*",
            options=_remaining,
            default=_saved_val if _saved_val else [],
            key="up_value_cols",
            help="e.g. Rent, Salary, Travel — these become rows. Leave empty to unpivot ALL non-ID columns.",
        )
        if not up_value_cols:
            st.info(f"ℹ️ All {len(_remaining)} non-ID column(s) will be unpivoted: `{', '.join(_remaining)}`")
    else:
        st.info("⬆️ Select a source file and sheet above to auto-detect columns.")
        up_id_cols = st.session_state.get("up_id_cols") or []
        up_value_cols = st.session_state.get("up_value_cols") or []

    st.divider()

    # ── Output column names ───────────────────────────────────────────────────
    st.markdown("**Output Column Names**")
    n1, n2 = st.columns(2)
    with n1:
        up_var_name = st.text_input(
            "Category column name",
            value=st.session_state.get("up_var_name", "Category"),
            key="up_var_name",
            help="Name of the new column that holds the original column headers",
        )
    with n2:
        up_value_name = st.text_input(
            "Value column name",
            value=st.session_state.get("up_value_name", "Value"),
            key="up_value_name",
            help="Name of the new column that holds the cell values",
        )

    up_drop_na = st.checkbox(
        "Drop rows where Value is blank/null",
        value=bool(st.session_state.get("up_drop_na", True)),
        key="up_drop_na",
    )

    st.divider()

    # ── Target ────────────────────────────────────────────────────────────────
    st.markdown("**Target  *(where to write the long-format result)***")
    t1, t2 = st.columns([3, 3])
    with t1:
        up_tgt_file = st.selectbox("Target File", get_files(), key="up_tgt_file",
                                   help="Can be the same as the source file — result is written to a new sheet")
    with t2:
        up_tgt_sheet = st.text_input(
            "Target Sheet Name",
            value=st.session_state.get("up_tgt_sheet", "Unpivot"),
            key="up_tgt_sheet",
        )

    export = st.checkbox("Export after run", value=True, key="up_export")

    # ── Save step ─────────────────────────────────────────────────────────────
    if st.button("➕ Add Unpivot Step", type="primary", use_container_width=True):
        if not up_src_file:
            st.error("Please select a Source File.")
            return
        if not up_src_sheet:
            st.error("Please select a Source Sheet.")
            return
        if not up_id_cols:
            st.error("Please select at least one ID Column.")
            return
        if not up_tgt_file:
            st.error("Please select a Target File.")
            return
        if not up_tgt_sheet.strip():
            st.error("Please enter a Target Sheet name.")
            return

        _vc_display = ", ".join(up_value_cols) if up_value_cols else "all non-ID cols"
        _label = (
            f"Unpivot  [{up_src_sheet}]  "
            f"ID: {', '.join(up_id_cols)}  |  "
            f"Values: {_vc_display}  →  {up_var_name}/{up_value_name}"
        )

        add_step("unpivot", _label, {
            "src_file":       up_src_file,
            "src_sheet":      up_src_sheet,
            "src_header_row": int(up_src_hrow),
            "id_cols":        list(up_id_cols),
            "value_cols":     list(up_value_cols),
            "var_name":       up_var_name.strip() or "Category",
            "value_name":     up_value_name.strip() or "Value",
            "tgt_file":       up_tgt_file,
            "tgt_sheet":      up_tgt_sheet.strip(),
            "append_mode":    False,
            "drop_na":        bool(up_drop_na),
            "export":         export,
        })


def render_sheet_row_inserter_form():
    """Sheet Row Inserter — insert rows from a master into specific positions in a multi-sheet workbook."""
    from core import wb_cache as _wbc

    st.markdown('<div class="pkf-section">Sheet Row Inserter</div>', unsafe_allow_html=True)
    st.caption(
        "Insert rows from a master sheet into a multi-sheet target workbook. "
        "For each **Category → Sheet** match, the engine finds an **anchor row** "
        "and physically inserts the master rows before or after it — "
        "no values are overwritten, existing rows shift down."
    )

    # ── Master file ───────────────────────────────────────────────────────────
    st.markdown("**Master File  *(Category | Particulars | Value)***")
    m1, m2 = st.columns([3, 3])
    with m1:
        sri_src_file = st.selectbox("Master File", get_files(), key="sri_src_file")
    with m2:
        sri_src_sheet = sheet_selectbox("Master Sheet", sri_src_file, key="sri_src_sheet")

    # Auto-detect columns
    _sri_ctx = (sri_src_file or "", sri_src_sheet or "")
    if st.session_state.get("sri_detect_ctx") != _sri_ctx:
        st.session_state["sri_detect_ctx"] = _sri_ctx
        st.session_state["sri_col_list"] = []
        if sri_src_file and sri_src_sheet:
            _fp = get_file_path(sri_src_file)
            if _fp:
                try:
                    _df = _wbc.read_excel(_fp, sheet_name=sri_src_sheet, nrows=1)
                    st.session_state["sri_col_list"] = [str(c).strip() for c in _df.columns]
                except Exception:
                    pass

    _sri_cols = st.session_state.get("sri_col_list", [])
    _col_opts = [""] + _sri_cols if _sri_cols else [""]

    st.markdown("**Master Columns**")
    c1, c2, c3 = st.columns(3)
    for _key, _label, _col_widget in [
        ("sri_col_category",    "Category Column *(→ sheet name)*",      c1),
        ("sri_col_particulars", "Particulars Column *(→ anchor / data)*", c2),
        ("sri_col_value",       "Value Column *(data to insert)*",        c3),
    ]:
        with _col_widget:
            if _sri_cols:
                _saved = st.session_state.get(_key, "")
                _idx   = _col_opts.index(_saved) if _saved in _col_opts else 0
                st.selectbox(_label, _col_opts, index=_idx, key=_key)
            else:
                st.text_input(_label, value=st.session_state.get(_key, ""), key=_key,
                              placeholder="Header or A/B/C")

    st.divider()

    # ── Target file ───────────────────────────────────────────────────────────
    st.markdown("**Target File  *(multi-sheet workbook to update)*  ⚠️ rows will be inserted in-place**")
    sri_tgt_file = st.selectbox("Target File", get_files(), key="sri_tgt_file")

    st.markdown("**Sheet Matching**")
    _sm_idx = int(st.session_state.get("sri_sheet_match_idx", 0))
    _sm_lbl = st.selectbox("Match Mode", SHEET_MATCH_LABELS, index=_sm_idx, key="sri_sheet_match_mode_label")
    sri_sheet_match_mode = SHEET_MATCH_KEYS[SHEET_MATCH_LABELS.index(_sm_lbl)]

    st.divider()

    # ── Anchor configuration ──────────────────────────────────────────────────
    st.markdown("**Anchor Row — where to insert**")
    st.caption(
        "The engine searches the target sheet for an **anchor row** and inserts the new rows "
        "at the specified position relative to it."
    )

    sri_anchor_col = st.text_input(
        "Anchor Column *(column in target sheet to search)*",
        value=st.session_state.get("sri_anchor_col", ""),
        key="sri_anchor_col",
        placeholder="e.g.  A  or  Particulars",
    )

    ANCHOR_MODE_LABELS = ["Based on Particulars value *(looks for Particulars in anchor column)*",
                          "Fixed value *(always look for the same text)*"]
    ANCHOR_MODE_KEYS   = ["particulars", "fixed"]
    _am_saved = st.session_state.get("sri_anchor_mode", "particulars")
    _am_idx   = ANCHOR_MODE_KEYS.index(_am_saved) if _am_saved in ANCHOR_MODE_KEYS else 0
    _am_lbl   = st.radio(
        "What to search for",
        ANCHOR_MODE_LABELS, index=_am_idx,
        key="sri_anchor_mode_radio", horizontal=False,
        help=(
            "**Particulars** — for each group of master rows, searches the anchor column "
            "for the Particulars value (e.g. inserts REFUNDS rows next to existing REFUNDS rows).\n\n"
            "**Fixed** — always search for the same text (e.g. always insert before 'GRAND TOTAL')."
        ),
    )
    sri_anchor_mode = ANCHOR_MODE_KEYS[ANCHOR_MODE_LABELS.index(_am_lbl)]

    sri_anchor_value = ""
    if sri_anchor_mode == "fixed":
        sri_anchor_value = st.text_input(
            "Anchor Value *(text to search for in the anchor column)*",
            value=st.session_state.get("sri_anchor_value", ""),
            key="sri_anchor_value",
            placeholder="e.g.  GRAND TOTAL  or  END OF SECTION",
        )

    a1, a2 = st.columns(2)
    with a1:
        _amatch_idx = int(st.session_state.get("sri_anchor_match_idx", 0))
        _amatch_lbl = st.selectbox("Anchor Match Mode", ANCHOR_MATCH_LABELS, index=_amatch_idx, key="sri_anchor_match_label")
        sri_anchor_match = ANCHOR_MATCH_KEYS[ANCHOR_MATCH_LABELS.index(_amatch_lbl)]
    with a2:
        _ipos_idx = int(st.session_state.get("sri_insert_pos_idx", 2))  # default: after_last
        _ipos_lbl = st.selectbox("Insert Position", INSERT_POSITION_LABELS, index=_ipos_idx, key="sri_insert_pos_label")
        sri_insert_position = INSERT_POSITION_KEYS[INSERT_POSITION_LABELS.index(_ipos_lbl)]

    st.divider()

    # ── Column mapping ────────────────────────────────────────────────────────
    st.markdown("**Column Mapping — where to write the values in the inserted rows**")
    st.caption("Enter column letters (A, B…) or header names from the target sheets.")
    w1, w2, w3 = st.columns([3, 3, 1])
    with w1:
        sri_tgt_col_particulars = st.text_input(
            "Particulars → Target Column",
            value=st.session_state.get("sri_tgt_col_particulars", ""),
            key="sri_tgt_col_particulars",
            placeholder="e.g.  A  or  Particulars",
        )
    with w2:
        sri_tgt_col_value = st.text_input(
            "Value → Target Column",
            value=st.session_state.get("sri_tgt_col_value", ""),
            key="sri_tgt_col_value",
            placeholder="e.g.  D  or  Amount",
        )
    with w3:
        sri_tgt_header_row = st.number_input(
            "Header Row",
            min_value=1,
            value=int(st.session_state.get("sri_tgt_header_row", 1) or 1),
            key="sri_tgt_header_row",
        )

    # Extra column mappings
    with st.expander("Additional column mappings *(optional)*", expanded=False):
        _n_extra = int(st.session_state.get("sri_extra_map_count", 0))
        if st.button("➕ Add Mapping", key="sri_extra_add"):
            st.session_state["sri_extra_map_count"] = _n_extra + 1
            st.rerun()
        _n_extra = int(st.session_state.get("sri_extra_map_count", 0))
        _extra_maps = []
        for _i in range(_n_extra):
            _ec = st.columns([4, 4, 1])
            _s = _ec[0].text_input(f"Master col {_i+1}", value=st.session_state.get(f"sri_extra_{_i}_src", ""),
                                   key=f"sri_extra_{_i}_src", placeholder="master header/letter",
                                   label_visibility="collapsed")
            _t = _ec[1].text_input(f"Target col {_i+1}", value=st.session_state.get(f"sri_extra_{_i}_tgt", ""),
                                   key=f"sri_extra_{_i}_tgt", placeholder="target col letter/header",
                                   label_visibility="collapsed")
            if _ec[2].button("✕", key=f"sri_extra_del_{_i}"):
                for _j in range(_i, _n_extra - 1):
                    st.session_state[f"sri_extra_{_j}_src"] = st.session_state.get(f"sri_extra_{_j+1}_src", "")
                    st.session_state[f"sri_extra_{_j}_tgt"] = st.session_state.get(f"sri_extra_{_j+1}_tgt", "")
                st.session_state["sri_extra_map_count"] = _n_extra - 1
                st.rerun()
            if _s.strip() and _t.strip():
                _extra_maps.append({"src_col": _s.strip(), "tgt_col": _t.strip()})

    export = st.checkbox("Export after run", value=True, key="sri_export")

    # ── Add Step ──────────────────────────────────────────────────────────────
    if st.button("➕ Add Sheet Row Inserter Step", type="primary", use_container_width=True):
        sri_col_category    = st.session_state.get("sri_col_category", "")
        sri_col_particulars = st.session_state.get("sri_col_particulars", "")
        sri_col_value       = st.session_state.get("sri_col_value", "")

        if not sri_src_file:
            st.error("Select a Master File."); return
        if not sri_src_sheet:
            st.error("Select a Master Sheet."); return
        if not sri_col_category:
            st.error("Specify the Category column in the master."); return
        if not sri_col_particulars:
            st.error("Specify the Particulars column in the master."); return
        if not sri_col_value:
            st.error("Specify the Value column in the master."); return
        if not sri_tgt_file:
            st.error("Select a Target File."); return
        if not sri_anchor_col:
            st.error("Specify the Anchor Column in the target sheets."); return
        if sri_anchor_mode == "fixed" and not sri_anchor_value:
            st.error("Enter an Anchor Value for Fixed mode."); return
        if not sri_tgt_col_particulars and not sri_tgt_col_value:
            st.error("Specify at least one target write column (Particulars or Value)."); return

        _pos_short = {
            "before_first": "before-first",
            "after_first":  "after-first",
            "after_last":   "after-last",
            "before_last":  "before-last",
        }.get(sri_insert_position, sri_insert_position)
        _anchor_desc = f"'{sri_anchor_value}'" if sri_anchor_mode == "fixed" else "by-particulars"
        _label = (
            f"Row Inserter  [{_anchor_desc} {_pos_short}]  "
            f"{sri_col_category}→sheet / insert {sri_col_particulars}+{sri_col_value}"
        )

        add_step("sheet_row_inserter", _label, {
            "src_file":            sri_src_file,
            "src_sheet":           sri_src_sheet,
            "src_col_category":    sri_col_category,
            "src_col_particulars": sri_col_particulars,
            "src_col_value":       sri_col_value,
            "tgt_file":            sri_tgt_file,
            "sheet_match_mode":    sri_sheet_match_mode,
            "anchor_col":          sri_anchor_col,
            "anchor_mode":         sri_anchor_mode,
            "anchor_value":        sri_anchor_value,
            "anchor_match":        sri_anchor_match,
            "insert_position":     sri_insert_position,
            "tgt_col_particulars": sri_tgt_col_particulars,
            "tgt_col_value":       sri_tgt_col_value,
            "extra_col_map":       _extra_maps,
            "tgt_header_row":      int(sri_tgt_header_row),
            "export":              export,
        })


def render_sheet_updater_form():
    """Sheet Updater — push master (category/particulars/value) values into matching sheets."""
    from core import wb_cache as _wbc

    st.markdown('<div class="pkf-section">Sheet Updater</div>', unsafe_allow_html=True)
    st.caption(
        "Read a master sheet with **Category**, **Particulars**, and **Value** columns. "
        "For each row: match the category to a sheet in the target workbook, find the "
        "Particulars label in a lookup column, and write the Value into the specified column."
    )

    # ── Master (source) ──────────────────────────────────────────────────────
    st.markdown("**Master File  *(the file with Category / Particulars / Value)***")
    m1, m2 = st.columns([3, 3])
    with m1:
        su_src_file = st.selectbox("Master File", get_files(), key="su_src_file")
    with m2:
        su_src_sheet = sheet_selectbox("Master Sheet", su_src_file, key="su_src_sheet")

    # Auto-detect columns from master sheet
    _su_ctx = (su_src_file or "", su_src_sheet or "")
    if st.session_state.get("su_detect_ctx") != _su_ctx:
        st.session_state["su_detect_ctx"] = _su_ctx
        st.session_state["su_col_list"] = []
        if su_src_file and su_src_sheet:
            _fp = get_file_path(su_src_file)
            if _fp:
                try:
                    _df = _wbc.read_excel(_fp, sheet_name=su_src_sheet, nrows=1)
                    st.session_state["su_col_list"] = [str(c).strip() for c in _df.columns]
                except Exception:
                    st.session_state["su_col_list"] = []

    _su_cols = st.session_state.get("su_col_list", [])
    _col_opts = [""] + _su_cols if _su_cols else [""]

    st.markdown("**Master Columns**")
    c1, c2, c3 = st.columns(3)
    with c1:
        if _su_cols:
            _cat_idx = _col_opts.index(st.session_state.get("su_col_category", "")) \
                       if st.session_state.get("su_col_category", "") in _col_opts else 0
            su_col_category = st.selectbox(
                "Category Column *(used to match sheet name)*",
                _col_opts, index=_cat_idx, key="su_col_category"
            )
        else:
            su_col_category = st.text_input(
                "Category Column *(header name or letter)*",
                value=st.session_state.get("su_col_category", ""),
                key="su_col_category"
            )
    with c2:
        if _su_cols:
            _part_idx = _col_opts.index(st.session_state.get("su_col_particulars", "")) \
                        if st.session_state.get("su_col_particulars", "") in _col_opts else 0
            su_col_particulars = st.selectbox(
                "Particulars Column *(row lookup key)*",
                _col_opts, index=_part_idx, key="su_col_particulars"
            )
        else:
            su_col_particulars = st.text_input(
                "Particulars Column *(header name or letter)*",
                value=st.session_state.get("su_col_particulars", ""),
                key="su_col_particulars"
            )
    with c3:
        if _su_cols:
            _val_idx = _col_opts.index(st.session_state.get("su_col_value", "")) \
                       if st.session_state.get("su_col_value", "") in _col_opts else 0
            su_col_value = st.selectbox(
                "Value Column *(what to write)*",
                _col_opts, index=_val_idx, key="su_col_value"
            )
        else:
            su_col_value = st.text_input(
                "Value Column *(header name or letter)*",
                value=st.session_state.get("su_col_value", ""),
                key="su_col_value"
            )

    st.divider()

    # ── Target (output workbook) ──────────────────────────────────────────────
    st.markdown("**Target File  *(multi-sheet workbook to update)*  ⚠️ this file will be modified in-place**")
    su_tgt_file = st.selectbox("Target File", get_files(), key="su_tgt_file")

    st.markdown("**Sheet Matching**")
    st.caption(
        "How the Category value from the master is matched to a sheet name in the target. "
        "Example: category = `HDFC`, sheet name = `HDFC MF Report` → use *'Category value is contained in sheet name'*"
    )
    _sm_idx = int(st.session_state.get("su_sheet_match_idx", 0))
    su_sheet_match_mode_label = st.selectbox(
        "Match Mode",
        SHEET_MATCH_LABELS,
        index=_sm_idx,
        key="su_sheet_match_mode_label",
    )
    su_sheet_match_mode = SHEET_MATCH_KEYS[SHEET_MATCH_LABELS.index(su_sheet_match_mode_label)]

    st.divider()

    # ── Target column configuration ───────────────────────────────────────────
    st.markdown("**Target Sheet Columns**")
    st.caption(
        "These columns must exist in *every* matched target sheet "
        "(use the same column header name across all sheets, or enter a letter like **A**, **B**)."
    )
    t1, t2, t3 = st.columns([3, 3, 1])
    with t1:
        su_tgt_col_lookup = st.text_input(
            "Lookup Column *(search for Particulars here)*",
            value=st.session_state.get("su_tgt_col_lookup", ""),
            key="su_tgt_col_lookup",
            placeholder="e.g.  Particulars  or  A",
        )
    with t2:
        su_tgt_col_write = st.text_input(
            "Write Column *(paste the Value here)*",
            value=st.session_state.get("su_tgt_col_write", ""),
            key="su_tgt_col_write",
            placeholder="e.g.  Amount  or  C",
        )
    with t3:
        su_tgt_header_row = st.number_input(
            "Header Row",
            min_value=1,
            value=int(st.session_state.get("su_tgt_header_row", 1) or 1),
            key="su_tgt_header_row",
        )

    # ── Particulars matching ──────────────────────────────────────────────────
    st.divider()
    st.markdown("**Particulars Matching**")
    st.caption(
        "Choose how Particulars values from the master are matched against the lookup column "
        "in the target sheet. The default *(Exact match)* works for most cases. "
        "Enable **Advanced Mode** to set a different match rule per particular."
    )

    _pm_idx = int(st.session_state.get("su_part_match_idx", 0))
    su_part_match_label = st.selectbox(
        "Default Match Mode  *(applies to all particulars unless overridden below)*",
        PARTICULARS_MATCH_LABELS,
        index=_pm_idx,
        key="su_part_match_mode_label",
    )
    su_particulars_match_mode = PARTICULARS_MATCH_KEYS[PARTICULARS_MATCH_LABELS.index(su_part_match_label)]

    su_advanced_mode = st.toggle(
        "Advanced Mode — set a match rule per particular",
        value=bool(st.session_state.get("su_advanced_mode", False)),
        key="su_advanced_mode",
    )

    su_particulars_rules = []
    if su_advanced_mode:
        st.caption(
            "Add each Particular label and choose how it should be matched. "
            "Rows left blank are ignored."
        )
        _n_rules = int(st.session_state.get("su_part_rules_count", 1))
        rc1, rc2 = st.columns([1, 6])
        with rc2:
            if st.button("➕ Add Row", key="su_add_rule_row"):
                st.session_state["su_part_rules_count"] = _n_rules + 1
                st.rerun()
        _n_rules = int(st.session_state.get("su_part_rules_count", 1))

        for _i in range(_n_rules):
            _cols = st.columns([4, 3, 1])
            with _cols[0]:
                _p = st.text_input(
                    f"Particular #{_i + 1}",
                    value=st.session_state.get(f"su_part_rule_{_i}_particular", ""),
                    key=f"su_part_rule_{_i}_particular",
                    placeholder="e.g. Revenue",
                    label_visibility="collapsed",
                )
            with _cols[1]:
                _rm_idx = int(st.session_state.get(f"su_part_rule_{_i}_match_mode", 0))
                _rm_label = st.selectbox(
                    f"Match #{_i + 1}",
                    PARTICULARS_MATCH_LABELS,
                    index=_rm_idx,
                    key=f"su_part_rule_{_i}_match_mode_label",
                    label_visibility="collapsed",
                )
                _rm_key = PARTICULARS_MATCH_KEYS[PARTICULARS_MATCH_LABELS.index(_rm_label)]
            with _cols[2]:
                if _n_rules > 1 and st.button("✕", key=f"su_del_rule_{_i}"):
                    # Remove this rule by shifting remaining ones down
                    for _j in range(_i, _n_rules - 1):
                        st.session_state[f"su_part_rule_{_j}_particular"]  = st.session_state.get(f"su_part_rule_{_j+1}_particular", "")
                        st.session_state[f"su_part_rule_{_j}_match_mode"]  = st.session_state.get(f"su_part_rule_{_j+1}_match_mode", 0)
                    st.session_state["su_part_rules_count"] = _n_rules - 1
                    st.rerun()
            if _p.strip():
                su_particulars_rules.append({"particular": _p.strip(), "match_mode": _rm_key})

    export = st.checkbox("Export after run", value=True, key="su_export")

    # ── Save step ─────────────────────────────────────────────────────────────
    if st.button("➕ Add Sheet Updater Step", type="primary", use_container_width=True):
        # Validation
        if not su_src_file:
            st.error("Please select a Master File.")
            return
        if not su_src_sheet:
            st.error("Please select a Master Sheet.")
            return
        if not su_col_category:
            st.error("Please specify the Category column in the master.")
            return
        if not su_col_particulars:
            st.error("Please specify the Particulars column in the master.")
            return
        if not su_col_value:
            st.error("Please specify the Value column in the master.")
            return
        if not su_tgt_file:
            st.error("Please select a Target File.")
            return
        if not su_tgt_col_lookup:
            st.error("Please enter the Lookup Column in the target sheet.")
            return
        if not su_tgt_col_write:
            st.error("Please enter the Write Column in the target sheet.")
            return

        _match_short = {
            "cat_in_sheet": "cat⊂sheet",
            "sheet_in_cat": "sheet⊂cat",
            "starts_with":  "starts_with",
            "ends_with":    "ends_with",
            "exact":        "exact",
        }.get(su_sheet_match_mode, su_sheet_match_mode)
        _pm_short = su_particulars_match_mode
        _label = (
            f"Sheet Updater  [{_match_short} / {_pm_short}]  "
            f"{su_col_category}→sheet / {su_col_particulars}→row / {su_col_value}→{su_tgt_col_write}"
        )

        add_step("sheet_updater", _label, {
            "src_file":              su_src_file,
            "src_sheet":             su_src_sheet,
            "src_col_category":      su_col_category,
            "src_col_particulars":   su_col_particulars,
            "src_col_value":         su_col_value,
            "tgt_file":              su_tgt_file,
            "sheet_match_mode":      su_sheet_match_mode,
            "tgt_col_lookup":        su_tgt_col_lookup,
            "tgt_col_write":         su_tgt_col_write,
            "tgt_header_row":        int(su_tgt_header_row),
            "case_sensitive":        False,
            "particulars_match_mode":su_particulars_match_mode,
            "particulars_rules":     su_particulars_rules,
            "export":                export,
        })


def render_add_files_form():
    """
    Workflow step: Add one or more files to the workflow.
    These files can be embedded in the workflow JSON (base64) so loading the workflow restores them.
    """
    st.markdown('<div class="pkf-section">Add File(s) (Workflow Step)</div>', unsafe_allow_html=True)
    st.warning("This app will **NOT** copy to temp. It will edit the files at these paths **in-place**.")
    st.caption("Enter one **absolute** path per line (Mac example: `/Users/you/Documents/file.xlsx`).")

    paths_text = st.text_area("Local file paths", value="", key="add_files_paths")
    confirm_overwrite = st.checkbox("I understand this will overwrite the original file(s) on disk", value=False, key="add_files_confirm")

    paths = [p.strip() for p in (paths_text or "").splitlines() if p.strip()]
    payload_files = [{"name": os.path.basename(p), "path": p} for p in paths]

    # Register immediately for UI dropdowns (always overwrites)
    for p in paths:
        try:
            register_file_path(p, display_name=os.path.basename(p))
        except Exception:
            # don't block typing; we'll validate on Add Step
            pass

    if st.button("Add Step", key="add_files_add"):
        if not payload_files:
            st.warning("Please enter at least one local file path.")
            return
        if not confirm_overwrite:
            st.warning("Please confirm overwrite.")
            return

        # Validate paths now (always overwrites)
        for item in payload_files:
            register_file_path(item["path"], display_name=item["name"])

        add_step(
            "add_files",
            f"Add File(s): {', '.join([pf.get('name','') for pf in payload_files])}",
            {"files": payload_files, "mode": "local_paths"},
        )


# =============================================================================
# WORKFLOW TAB
# =============================================================================

def render_workflow_tab():
    """Display current workflow steps"""
    st.markdown('<h2 style="color:#1a1a1a;font-family:IBM Plex Sans,sans-serif;text-transform:uppercase;letter-spacing:.1em;font-size:13px;font-weight:600;border-bottom:2px solid #c8392b;padding-bottom:4px;margin-bottom:12px;">Current Workflow</h2>', unsafe_allow_html=True)
    
    if not st.session_state.workflow_steps:
        st.info("No steps yet. Add steps from the 'Add Steps'tab.")
        return
    
    ensure_step_ids()
    st.success(f"{len(st.session_state.workflow_steps)} step(s) in workflow")

    # Dependency / parallel lane overview
    with st.expander("Parallel / Dependencies (DAG)", expanded=False):
        st.caption("Define which steps can run in parallel by adding dependencies. "
                   "A step will wait for **all** its upstream steps in `depends_on` to finish.")

        show_graph = st.checkbox("Show dependency graph", value=True, key="wf_show_dag_graph")
        if show_graph:
            # Build a DOT graph
            dot_lines = ["digraph workflow {", 'rankdir="LR";', 'node [shape=box];']
            for idx, s in enumerate(st.session_state.workflow_steps):
                sid = s.get("id")
                label = (s.get("name") or f"Step {idx+1}").replace('"', "'")
                dot_lines.append(f'"{sid}" [label="{idx+1}. {label}"];')
            for s in st.session_state.workflow_steps:
                sid = s.get("id")
                for dep in (s.get("depends_on") or []):
                    dot_lines.append(f'"{dep}" -> "{sid}";')
            dot_lines.append("}")
            st.graphviz_chart("\n".join(dot_lines), use_container_width=True)

        st.caption("Optional: group steps into lanes (for your own organization).")
        lanes = sorted({(s.get("lane") or "").strip() for s in st.session_state.workflow_steps if (s.get("lane") or "").strip()})
        if lanes:
            st.write("Lanes:", ", ".join(lanes))
        else:
            st.write("Lanes: (none)")
    
    for i, step in enumerate(st.session_state.workflow_steps):
        with st.expander(f"**Step {i+1}:** {step.get('name', 'Unknown')}", expanded=False):
            st.caption(f"Step ID: `{step.get('id')}`")

            # Edit dependencies / lane
            other_steps = [
                (s.get("id"), f"Step {j+1}: {s.get('name','')}")
                for j, s in enumerate(st.session_state.workflow_steps)
                if s.get("id") and s.get("id") != step.get("id")
            ]
            id_to_label = {sid: lbl for sid, lbl in other_steps}
            current_dep_ids = [d for d in (step.get("depends_on") or []) if d in id_to_label]

            dep_labels_selected = st.multiselect(
                "Depends on (wait for these steps to finish first)",
                options=[lbl for _, lbl in other_steps],
                default=[id_to_label[d] for d in current_dep_ids],
                key=f"dep_sel_{step.get('id')}",
                help="This enables parallel execution when steps don't depend on each other.",
            )
            # Convert back to IDs
            label_to_id = {lbl: sid for sid, lbl in other_steps}
            new_dep_ids = [label_to_id[lbl] for lbl in dep_labels_selected if lbl in label_to_id]
            step["depends_on"] = new_dep_ids

            lane_val = st.text_input(
                "Lane (optional label for grouping parallel paths)",
                value=step.get("lane", "") or "",
                key=f"lane_{step.get('id')}",
            )
            step["lane"] = lane_val.strip()

            st.json(step)
            
            col1, col2, col3, col4, col5 = st.columns([1, 1, 1, 1, 3])
            with col1:
                if st.button("", key=f"up_{i}", disabled=(i==0)):
                    st.session_state.workflow_steps[i], st.session_state.workflow_steps[i-1] = \
                        st.session_state.workflow_steps[i-1], st.session_state.workflow_steps[i]
                    st.rerun()
            with col2:
                if st.button("", key=f"down_{i}", disabled=(i==len(st.session_state.workflow_steps)-1)):
                    st.session_state.workflow_steps[i], st.session_state.workflow_steps[i+1] = \
                        st.session_state.workflow_steps[i+1], st.session_state.workflow_steps[i]
                    st.rerun()
            with col3:
                if st.button("Insert after", key=f"insert_after_{i}"):
                    st.session_state.insert_at = i + 1
                    st.session_state._workflow_notice = f"Next step will be **inserted after Step {i+1}** (position {i+2})."
                    st.rerun()
            with col4:
                if st.button("Edit", key=f"edit_{i}"):
                    # Visual edit: store data for populate on NEXT render cycle.
                    # We cannot call populate_form_from_step() here because the
                    # form widgets have already been instantiated in this run
                    # (all tabs render on every script execution in Streamlit).
                    step_type = step.get("type", "")
                    step_config = step.get("config", {}) or {}
                    st.session_state.editing_step_idx = i
                    st.session_state.editing_step_data = step
                    st.session_state.adding_step = step_type
                    st.session_state["_pending_edit_populate"] = (step_type, step_config)
                    st.rerun()
            with col5:
                if st.button("", key=f"del_{i}"):
                    st.session_state.workflow_steps.pop(i)
                    # Keep edit/insert indices sane
                    if st.session_state.get("editing_step_idx") == i:
                        st.session_state.editing_step_idx = None
                    if isinstance(st.session_state.get("insert_at"), int) and st.session_state.insert_at is not None:
                        if st.session_state.insert_at > i:
                            st.session_state.insert_at -= 1
                    st.rerun()

            # Visual editing is now in Add Steps tab - show info here
            if st.session_state.get("editing_step_idx") == i and st.session_state.get("editing_step_data"):
                st.info("Go to ** Add Steps** tab to edit this step visually.")
                
                # Fallback: Advanced JSON editor (collapsed)
                with st.expander("Advanced: Edit as JSON", expanded=False):
                    new_name = st.text_input(
                        "Step name",
                        value=step.get("name", ""),
                        key=f"edit_name_{i}",
                    )
                    cfg_str = st.text_area(
                        "Step config (JSON)",
                        value=json.dumps(step.get("config", {}) or {}, indent=2),
                        height=220,
                        key=f"edit_cfg_{i}",
                    )
                    a, b = st.columns(2)
                    with a:
                        if st.button("Save JSON", key=f"save_edit_{i}", use_container_width=True):
                            try:
                                parsed = json.loads(cfg_str or "{}")
                                if not isinstance(parsed, dict):
                                    st.error("Config must be a JSON object (dictionary).")
                                else:
                                    st.session_state.workflow_steps[i]["name"] = new_name
                                    st.session_state.workflow_steps[i]["config"] = parsed
                                    st.session_state.editing_step_idx = None
                                    st.session_state.editing_step_data = None
                                    st.success("Saved changes.")
                                    st.rerun()
                            except Exception as e:
                                st.error(f"Invalid JSON: {e}")
                    with b:
                        if st.button("Cancel", key=f"cancel_edit_{i}", use_container_width=True):
                            st.session_state.editing_step_idx = None
                            st.session_state.editing_step_data = None
                            st.session_state.pop("adding_step", None)
                            st.rerun()


# =============================================================================
# EXECUTE TAB
# =============================================================================

def _collect_upload_requirements(steps: list) -> list[dict]:
    """
    Scan workflow steps for Input Source classes with mode='E'(user upload).
    Returns a list of dicts: [{class_name, accept, multi}, ...]
    """
    reqs = []
    seen = set()
    for step in steps:
        if (step.get("type") or "").lower() != "input_source":
            continue
        for cls in (step.get("config") or {}).get("file_classes", []):
            if cls.get("mode") == "E"and cls.get("class_name"):
                name = cls["class_name"]
                if name not in seen:
                    seen.add(name)
                    reqs.append({
                        "class_name": name,
                        "accept":     cls.get("upload_accept", "Excel (.xlsx / .xls / .xlsm / .xlsb)"),
                        "multi":      bool(cls.get("upload_multi", False)),
                        "sheet_mode": cls.get("sheet_mode", "first"),
                        "sheet_exact": cls.get("sheet_exact", ""),
                    })
    return reqs


def _save_uploaded_file(uploaded_file, class_name: str) -> str:
    """Save a Streamlit UploadedFile to a temp directory and return its path."""
    import tempfile, pathlib
    tmp_dir = pathlib.Path(tempfile.gettempdir()) / "recon_uploads"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    dest = tmp_dir / uploaded_file.name
    dest.write_bytes(uploaded_file.getvalue())
    return str(dest)


def render_execute_tab():
    """Execute workflow"""
    st.markdown('<h2 style="color:#1a1a1a;font-family:IBM Plex Sans,sans-serif;text-transform:uppercase;letter-spacing:.1em;font-size:13px;font-weight:600;border-bottom:2px solid #c8392b;padding-bottom:4px;margin-bottom:12px;">Execute Workflow</h2>', unsafe_allow_html=True)

    if not st.session_state.workflow_steps:
        st.warning("No steps to execute")
        return
    
    total_steps = len(st.session_state.workflow_steps)
    st.info(f"Ready to execute {total_steps} step(s)")

    # ── User-upload requirements ──────────────────────────────────────────────
    upload_reqs = _collect_upload_requirements(st.session_state.workflow_steps)
    uploads_complete = True
    if upload_reqs:
        st.markdown('<div class="pkf-section">Required File Uploads</div>', unsafe_allow_html=True)
        st.caption(
            "The following file classes are marked **User Upload**. "
            "Upload them here before running the workflow."
        )
        _ACCEPT_MAP = {
            "Excel (.xlsx / .xls / .xlsm / .xlsb)": [
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/vnd.ms-excel",
                "application/octet-stream",
            ],
            "CSV (.csv)": ["text/csv", "text/plain"],
            "Any": None,
        }
        if "user_uploaded_files"not in st.session_state:
            st.session_state.user_uploaded_files = {}

        for req in upload_reqs:
            cname = req["class_name"]
            accept_key = req.get("accept", "Excel (.xlsx / .xls / .xlsm / .xlsb)")
            accept_types = _ACCEPT_MAP.get(accept_key)
            already = st.session_state.user_uploaded_files.get(cname)
            with st.container(border=True):
                col_lbl, col_status = st.columns([4, 1])
                with col_lbl:
                    st.markdown(f"**{cname}**")
                with col_status:
                    if already:
                        names = already if isinstance(already, list) else [already]
                        st.success(f"{len(names)} file(s) ready")
                    else:
                        st.warning("Awaiting upload")
                        uploads_complete = False

                uploader_kwargs = dict(
                    label=f"Upload file(s) for **{cname}**",
                    key=f"uploader_{cname}",
                    accept_multiple_files=req.get("multi", False),
                    label_visibility="collapsed",
                )
                if accept_types:
                    uploader_kwargs["type"] = [
                        "xlsx", "xls", "xlsm", "xlsb", "xlsb",
                        "csv", "txt",
                    ] if accept_key == "Any"else (
                        ["xlsx", "xls", "xlsm", "xlsb"] if "Excel"in accept_key
                        else ["csv", "txt"]
                    )
                uploaded = st.file_uploader(**uploader_kwargs)

                if uploaded:
                    files = uploaded if isinstance(uploaded, list) else [uploaded]
                    saved_paths = [_save_uploaded_file(f, cname) for f in files]
                    st.session_state.user_uploaded_files[cname] = saved_paths
                    # Register in the session file map so later steps can find them
                    if req.get("multi") and len(saved_paths) > 1:
                        if "uploaded_file_groups"not in st.session_state:
                            st.session_state.uploaded_file_groups = {}
                        st.session_state.uploaded_file_groups[cname] = saved_paths
                        register_file_path(saved_paths[0], display_name=cname)
                    else:
                        register_file_path(saved_paths[0], display_name=cname)
                    uploads_complete = False  # will re-evaluate on next run
                    st.rerun()

        if not uploads_complete:
            st.error("Upload all required files above before running the workflow.")

        st.divider()
    # ─────────────────────────────────────────────────────────────────────────

    mode = st.radio(
        "Execution mode",
        options=["Sequential", "Parallel (DAG)"],
        index=0,
        horizontal=True,
        help="Parallel mode runs independent steps concurrently and respects dependencies set in the Workflow tab.",
        key="exec_mode",
    )

    st.markdown('<div class="pkf-section">Run range</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        start_step = st.number_input("Start step #", min_value=1, max_value=total_steps, value=1, step=1, key="exec_start_step")
    with c2:
        end_step = st.number_input("End step #", min_value=1, max_value=total_steps, value=total_steps, step=1, key="exec_end_step")

    start_step = int(start_step)
    end_step = int(end_step)
    if start_step > end_step:
        st.warning("Start step must be <= End step")
        return

    run_all = st.button("Run All Steps", type="primary", use_container_width=True,
                        key="run_all_steps", disabled=not uploads_complete)
    run_range = st.button("Run Selected Range", use_container_width=True,
                          key="run_range_steps", disabled=not uploads_complete)
    if run_all:
        if mode == "Parallel (DAG)":
            execute_workflow_dag(1, total_steps)
        else:
            execute_workflow(1, total_steps)
    elif run_range:
        if mode == "Parallel (DAG)":
            execute_workflow_dag(start_step, end_step)
        else:
            execute_workflow(start_step, end_step)

    # -------------------------------------------------------------------------
    # Execution Groups: create N buttons (e.g., 1-10, 11-20, 21-30)
    # -------------------------------------------------------------------------
    if "exec_groups"not in st.session_state:
        # list[dict]: {name, start, end}
        st.session_state.exec_groups = []

    st.markdown("---")
    with st.expander("Execution Groups (create multiple Run buttons)", expanded=False):
        st.caption("Create your own buttons to run step ranges quickly (e.g., 1–10, 11–20, 21–30).")

        # Auto-generate groups by chunk size
        c1, c2, c3 = st.columns([2, 2, 2])
        with c1:
            chunk_size = st.number_input("Auto group size", min_value=1, max_value=max(1, total_steps), value=min(10, total_steps), step=1, key="exec_group_chunk")
        with c2:
            prefix = st.text_input("Group name prefix", value="Steps", key="exec_group_prefix")
        with c3:
            st.write("")
            if st.button("Create groups", use_container_width=True, key="exec_group_autocreate"):
                groups = []
                cs = int(chunk_size)
                pfx = (prefix or "Steps").strip()
                for s in range(1, total_steps + 1, cs):
                    e = min(total_steps, s + cs - 1)
                    groups.append({"name": f"{pfx} {s}-{e}", "start": s, "end": e})
                st.session_state.exec_groups = groups
                st.rerun()

        st.markdown('<div class="pkf-section">Add custom group</div>', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns([4, 2, 2, 2])
        with c1:
            gname = st.text_input("Button label", value="", key="exec_group_name")
        with c2:
            gstart = st.number_input("Start", min_value=1, max_value=total_steps, value=1, step=1, key="exec_group_start")
        with c3:
            gend = st.number_input("End", min_value=1, max_value=total_steps, value=min(total_steps, 10), step=1, key="exec_group_end")
        with c4:
            st.write("")
            if st.button("Add", use_container_width=True, key="exec_group_add"):
                s = int(gstart)
                e = int(gend)
                if s > e:
                    st.warning("Start must be <= End")
                else:
                    label = (gname or f"Steps {s}-{e}").strip()
                    st.session_state.exec_groups.append({"name": label, "start": s, "end": e})
                    st.rerun()

        groups = st.session_state.get("exec_groups", []) or []
        if not groups:
            st.info("No groups yet. Use auto-create or add one manually.")
        else:
            st.markdown('<div class="pkf-section">Run groups</div>', unsafe_allow_html=True)
            # Show as rows of 3 buttons
            per_row = 3
            for row_i in range(0, len(groups), per_row):
                row = groups[row_i:row_i + per_row]
                cols = st.columns(len(row))
                for j, g in enumerate(row):
                    name = str(g.get("name") or "").strip() or "Run"
                    s = int(g.get("start") or 1)
                    e = int(g.get("end") or s)
                    s = max(1, min(total_steps, s))
                    e = max(1, min(total_steps, e))
                    if s > e:
                        s, e = e, s

                    with cols[j]:
                        if st.button(f"{name}\n({s}-{e})", use_container_width=True, key=f"exec_group_run_{row_i+j}"):
                            if mode == "Parallel (DAG)":
                                execute_workflow_dag(s, e)
                            else:
                                execute_workflow(s, e)

                        if st.button("Remove", use_container_width=True, key=f"exec_group_del_{row_i+j}"):
                            try:
                                st.session_state.exec_groups.pop(row_i + j)
                            except Exception:
                                pass
                            st.rerun()


def execute_workflow(start_step_num: int = 1, end_step_num: int | None = None):
    """Execute all steps with NON-DESTRUCTIVE EXPORT support.
    
    When export is enabled for a step:
    - Original file is NOT modified
    - Operation runs on a COPY at the export path
    - Export file becomes available for subsequent steps
    """
    progress = st.progress(0)
    status = st.empty()
    
    total = len(st.session_state.workflow_steps)
    if end_step_num is None:
        end_step_num = total

    start_idx = max(0, int(start_step_num) - 1)
    end_idx = min(total - 1, int(end_step_num) - 1)
    if start_idx > end_idx:
        st.warning("Invalid step range.")
        return

    steps_to_run = list(enumerate(st.session_state.workflow_steps))[start_idx : end_idx + 1]
    range_total = len(steps_to_run)
    
    for j, (i, step) in enumerate(steps_to_run):
        status.text(f"Step {i+1}/{total}: {step.get('name')}")
        
        try:
            # NON-DESTRUCTIVE EXPORT: If export is enabled, copy original to export path
            # and temporarily swap registration so operation runs on the COPY
            swapped_paths, original_mappings = _prepare_non_destructive_export(i + 1, step)
            
            success, msg = execute_step(step)
            
            if swapped_paths:
                # Restore original file registrations and register exports as new files
                _finalize_non_destructive_export(i + 1, swapped_paths, original_mappings)
            
            if success:
                if swapped_paths:
                    st.success(f"Step {i+1}: {msg} (exported only, original unchanged)")
                else:
                    st.success(f"Step {i+1}: {msg} (saved to original)")
            else:
                st.error(f"Step {i+1}: {msg}")
                break
        except Exception as e:
            st.error(f"Step {i+1}: {str(e)}")
            break
        
        progress.progress((j + 1) / range_total)
    
    status.text("Complete!")
    st.balloons()


def execute_workflow_dag(start_step_num: int = 1, end_step_num: int | None = None):
    """
    Execute steps using dependency-aware parallel scheduling (DAG).
    - Respects `step["depends_on"]` (waits for all upstream steps).
    - Runs independent steps in parallel, but avoids concurrent writes to the same output file.
    """
    ensure_step_ids()
    steps_all = st.session_state.get("workflow_steps", []) or []
    total = len(steps_all)
    if total == 0:
        st.warning("No steps to execute.")
        return

    if end_step_num is None:
        end_step_num = total
    start_idx = max(0, int(start_step_num) - 1)
    end_idx = min(total - 1, int(end_step_num) - 1)
    if start_idx > end_idx:
        st.warning("Invalid step range.")
        return

    # Target set: selected range + all upstream dependencies (so it can actually run)
    selected_ids = {steps_all[i].get("id") for i in range(start_idx, end_idx + 1) if steps_all[i].get("id")}
    id_to_step = {s.get("id"): s for s in steps_all if s.get("id")}
    id_to_index = {s.get("id"): i for i, s in enumerate(steps_all) if s.get("id")}

    needed = set(selected_ids)
    stack = list(selected_ids)
    while stack:
        sid = stack.pop()
        s = id_to_step.get(sid)
        if not s:
            continue
        for dep in (s.get("depends_on") or []):
            if dep not in id_to_step:
                st.error(f"Dependency not found for step {id_to_index.get(sid, '?')+1}: missing id {dep}")
                return
            if dep not in needed:
                needed.add(dep)
                stack.append(dep)

    # Build adjacency / indegree for the subgraph
    adj: dict[str, list[str]] = {sid: [] for sid in needed}
    indeg: dict[str, int] = {sid: 0 for sid in needed}
    for sid in needed:
        s = id_to_step[sid]
        for dep in (s.get("depends_on") or []):
            if dep in needed:
                adj[dep].append(sid)
                indeg[sid] += 1

    # Snapshot file map (thread-safe for workers)
    file_map = dict(st.session_state.get("uploaded_file_paths", {}) or {})

    # Precompute write-lock paths for each node (to avoid concurrent writes to same output file)
    def write_paths_for(sid: str) -> set[str]:
        s = id_to_step[sid]
        stype = (s.get("type") or "").lower().strip()
        cfg = s.get("config", {}) or {}
        labels = [x for x in _determine_modified_file_labels(stype, cfg) if x]
        return {str(file_map.get(lab)) for lab in labels if lab and file_map.get(lab)}

    node_paths = {sid: write_paths_for(sid) for sid in needed}

    progress = st.progress(0)
    status = st.empty()
    done = 0
    range_total = len(needed)

    # Scheduler state
    ready = {sid for sid, d in indeg.items() if d == 0}
    running: dict[cf.Future, tuple[str, set[str]]] = {}
    locked_paths: set[str] = set()
    completed: set[str] = set()

    # Use a small pool; many operations are IO-bound but openpyxl/pandas can be CPU-heavy
    max_workers = min(8, max(2, (os.cpu_count() or 4)))
    executor = cf.ThreadPoolExecutor(max_workers=max_workers)

    def run_step_worker(sid: str):
        # IMPORTANT: do not call Streamlit APIs here
        step = id_to_step[sid]
        return execute_step_with_file_map(step, file_map)

    try:
        while len(completed) < len(needed):
            # Prefer running add_files in the main thread (it mutates session_state registrations)
            did_sync = True
            while did_sync:
                did_sync = False
                for sid in sorted(list(ready), key=lambda x: id_to_index.get(x, 10**9)):
                    stype = (id_to_step[sid].get("type") or "").lower().strip()
                    if stype in ("add_files", "input_source"):
                        ready.remove(sid)
                        status.text(f"[DAG] Step {id_to_index.get(sid, '?')+1}: {id_to_step[sid].get('name')}")
                        ok, msg = execute_step(id_to_step[sid])
                        if ok:
                            st.success(f"Step {id_to_index.get(sid, '?')+1}: {msg}")
                            _maybe_export_after_step(id_to_index.get(sid, 0) + 1, id_to_step[sid])
                            # update file map snapshot for downstream workers
                            file_map = dict(st.session_state.get("uploaded_file_paths", {}) or {})
                            # recompute paths for remaining nodes (best-effort)
                            for nid in list(node_paths.keys()):
                                if nid in completed:
                                    continue
                                node_paths[nid] = write_paths_for(nid)
                        else:
                            st.error(f"Step {id_to_index.get(sid, '?')+1}: {msg}")
                            return
                        completed.add(sid)
                        done += 1
                        progress.progress(done / range_total)
                        for nxt in adj.get(sid, []):
                            indeg[nxt] -= 1
                            if indeg[nxt] == 0:
                                ready.add(nxt)
                        did_sync = True
                        break

            # Dispatch as many ready nodes as possible whose file paths aren't locked
            dispatchable = []
            for sid in sorted(list(ready), key=lambda x: id_to_index.get(x, 10**9)):
                paths = node_paths.get(sid) or set()
                if paths and (paths & locked_paths):
                    continue
                dispatchable.append(sid)

            for sid in dispatchable:
                # Remove from ready and lock its paths before submitting
                ready.remove(sid)
                paths = node_paths.get(sid) or set()
                locked_paths |= paths
                
                # NON-DESTRUCTIVE EXPORT: Prepare copy in main thread before worker runs
                step = id_to_step[sid]
                swapped_paths, original_mappings = _prepare_non_destructive_export(id_to_index.get(sid, 0) + 1, step)
                if swapped_paths:
                    # Update file_map snapshot so worker uses export copy
                    file_map.update(st.session_state.get("uploaded_file_paths", {}))
                
                fut = executor.submit(run_step_worker, sid)
                running[fut] = (sid, paths, swapped_paths, original_mappings)

            if not running:
                # No running futures and nothing dispatchable => cycle or lock deadlock
                if ready:
                    st.error("DAG could not make progress. This usually means a dependency cycle.")
                else:
                    st.error("DAG could not make progress.")
                return

            # Wait for any step to complete
            done_futs, _ = cf.wait(running.keys(), return_when=cf.FIRST_COMPLETED)
            for fut in done_futs:
                sid, paths, swapped_paths, original_mappings = running.pop(fut)
                locked_paths -= paths
                idx = id_to_index.get(sid, None)
                step = id_to_step[sid]
                try:
                    ok, msg = fut.result()
                except Exception as e:
                    ok, msg = False, str(e)

                # NON-DESTRUCTIVE EXPORT: Finalize in main thread after worker completes
                if swapped_paths:
                    _finalize_non_destructive_export(idx + 1 if isinstance(idx, int) else 0, swapped_paths, original_mappings)
                    # Update file_map for future workers
                    file_map.update(st.session_state.get("uploaded_file_paths", {}))

                step_num = (idx + 1) if isinstance(idx, int) else "?"
                if ok:
                    if swapped_paths:
                        st.success(f"Step {step_num}: {msg} (exported only, original unchanged)")
                    else:
                        st.success(f"Step {step_num}: {msg} (saved to original)")
                else:
                    st.error(f"Step {step_num}: {msg}")
                    # Cancel remaining
                    for other in list(running.keys()):
                        other.cancel()
                    return

                completed.add(sid)
                done += 1
                progress.progress(done / range_total)

                # Release downstream
                for nxt in adj.get(sid, []):
                    indeg[nxt] -= 1
                    if indeg[nxt] == 0:
                        ready.add(nxt)

        status.text("DAG complete!")
        st.balloons()
    finally:
        executor.shutdown(wait=False, cancel_futures=True)


def _resolve_formula_column(file_path: str, sheet: str, cfg: dict) -> str:
    """Resolve the target column for a formula step.
    If column_is_header is True, scan header row to find the matching column letter."""
    from openpyxl.utils import get_column_letter as _gcl
    col = cfg.get("column", "A")
    if not cfg.get("column_is_header"):
        return col
    try:
        wb = wb_cache.load(file_path)
        ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.active
        header_row = int(cfg.get("column_header_row", 1) or 1)
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=header_row, column=c).value
            if v is not None and str(v).strip() == str(col).strip():
                return _gcl(c)
    except Exception:
        pass
    return col


def execute_step(step):
    """Execute a single step"""
    step_type = step.get('type')
    cfg = step.get('config', {})

    def _registered_labels() -> list[str]:
        m = st.session_state.get("uploaded_file_paths")
        if not isinstance(m, dict):
            return []
        return sorted([str(k) for k in m.keys()])

    def _resolve_required_path(label: str | None, *, field: str) -> tuple[str, str | None]:
        """
        Resolve a workflow file label -> absolute path, with a detailed error if missing.
        """
        if not label:
            return "", f"Missing required file selection for field '{field}'in step type '{step_type}'."
        path = get_file_path(label)
        if path:
            return path, None
        registered = _registered_labels()
        reg_preview = ", ".join([f"'{x}'"for x in registered[:15]])
        more = ""if len(registered) <= 15 else f"(+{len(registered) - 15} more)"
        return "", (
            f"File not found for field '{field}': '{label}'. "
            f"Register it via Add File(s) step first. "
            f"Currently registered: [{reg_preview}]{more}"
        )

    # Steps that don't require a pre-registered target file
    if step_type == "add_files":
        files = cfg.get("files", []) or []
        overwrite = bool(cfg.get("overwrite", True))

        if not files:
            return False, "No files provided in step"

        registered = 0
        for item in files:
            name = item.get("name")
            path = item.get("path")
            if not name:
                continue
            if not path:
                return False, f"File '{name}'has no path. This app runs in local-path-only mode."
            register_file_path(path, display_name=name, overwrite=overwrite)
            registered += 1

        return True, f"Registered {registered} file(s)"

    if step_type == "input_source":
        file_classes = cfg.get("file_classes", []) or []
        if not file_classes:
            return False, "No file classes defined"
        return run_input_source_step(file_classes)

    if step_type == "normalize_import":
        src_label = cfg.get("src_file")
        tgt_label = cfg.get("tgt_file")
        src_path, err = _resolve_required_path(src_label, field="src_file")
        if err:
            return False, err
        tgt_path, err = _resolve_required_path(tgt_label, field="tgt_file")
        if err:
            return False, err
        # Pull full group list when source came from a folder scan
        groups = st.session_state.get("uploaded_file_groups") or {}
        src_paths = groups.get(src_label) if src_label in groups else None
        return run_normalize_import_step(
            src_path=src_path,
            src_sheet=cfg.get("src_sheet"),
            src_header_row=int(cfg.get("src_header_row", 1) or 1),
            mapping=cfg.get("mapping", []) or [],
            tgt_path=tgt_path,
            tgt_sheet=cfg.get("tgt_sheet"),
            tgt_header_row=int(cfg.get("tgt_header_row", 1) or 1),
            append_mode=bool(cfg.get("append_mode", True)),
            src_paths=src_paths,
        )

    if step_type == "pivot":
        src_path, err = _resolve_required_path(cfg.get("src_file"), field="src_file")
        if err: return False, err
        tgt_path, err = _resolve_required_path(cfg.get("tgt_file"), field="tgt_file")
        if err: return False, err
        return run_pivot_step(
            src_file_path=src_path,
            src_sheet=cfg.get("src_sheet"),
            src_header_row=int(cfg.get("src_header_row", 1) or 1),
            tgt_file_path=tgt_path,
            tgt_sheet=cfg.get("tgt_sheet", "Pivot"),
            tgt_start_row=int(cfg.get("tgt_start_row", 1) or 1),
            tgt_start_col=cfg.get("tgt_start_col", "A"),
            overwrite_sheet=bool(cfg.get("overwrite_sheet", True)),
            row_fields=cfg.get("row_fields") or [],
            col_fields=cfg.get("col_fields") or [],
            value_fields=cfg.get("value_fields") or [],
            filter_conditions=cfg.get("filter_conditions") or [],
            filter_combine=cfg.get("filter_combine", "AND"),
            fill_value=cfg.get("fill_value", 0),
            grand_total_rows=bool(cfg.get("grand_total_rows", False)),
            grand_total_cols=bool(cfg.get("grand_total_cols", False)),
            sort_field=cfg.get("sort_field", ""),
            sort_ascending=bool(cfg.get("sort_ascending", True)),
            apply_formatting=bool(cfg.get("apply_formatting", True)),
            header_bg_color=cfg.get("header_bg_color", "1F3864"),
            header_font_color=cfg.get("header_font_color", "FFFFFF"),
            total_bg_color=cfg.get("total_bg_color", "D9E1F2"),
        )

    if step_type == "folder_summary":
        src_label = cfg.get("src_file")
        tgt_label = cfg.get("tgt_file")
        src_path, err = _resolve_required_path(src_label, field="src_file")
        if err:
            return False, err
        tgt_path, err = _resolve_required_path(tgt_label, field="tgt_file")
        if err:
            return False, err
        groups = st.session_state.get("uploaded_file_groups") or {}
        src_paths = groups.get(src_label) if src_label in (groups or {}) else None
        if not src_paths:
            src_paths = [src_path]
        return run_folder_summary_step(
            src_file_paths=src_paths,
            src_sheet=cfg.get("src_sheet", ""),
            header_row=int(cfg.get("header_row", 1) or 1),
            source_mode=cfg.get("source_mode", "folder"),
            sheet_selection=cfg.get("sheet_selection", "all"),
            sheet_pattern=cfg.get("sheet_pattern", ""),
            sheet_list=cfg.get("sheet_list") or [],
            sheet_exclude=cfg.get("sheet_exclude", ""),
            variables=cfg.get("variables") or [],
            formulas=cfg.get("formulas") or [],
            tgt_file=tgt_path,
            tgt_sheet=cfg.get("tgt_sheet", "Summary"),
            tgt_header_row=int(cfg.get("tgt_header_row", 1) or 1),
            filename_col=cfg.get("filename_col", "File Name"),
            sheet_name_col=cfg.get("sheet_name_col", ""),
            append_mode=bool(cfg.get("append_mode", True)),
            include_var_cols=bool(cfg.get("include_var_cols", True)),
            include_status_col=bool(cfg.get("include_status_col", True)),
            status_col_name=cfg.get("status_col_name", "Status"),
            on_missing_default=cfg.get("on_missing_default", 0),
        )

    # Steps that operate on source/target files (no single cfg["file"])
    if step_type in ("import", "append", "header_map"):
        src_label = cfg.get("src_file")
        tgt_label = cfg.get("tgt_file")
        src_path, err = _resolve_required_path(src_label, field="src_file")
        if err:
            return False, err
        tgt_path, err = _resolve_required_path(tgt_label, field="tgt_file")
        if err:
            return False, err

        if step_type == "append":
            return run_append_step(
                src_path, tgt_path, cfg.get("src_sheet"), cfg.get("tgt_sheet")
            )

        if step_type == "import":
            # If the source label came from a folder-scan Input Source,
            # uploaded_file_groups has ALL the files — process every one.
            groups = st.session_state.get("uploaded_file_groups") or {}
            all_src = groups.get(src_label) if src_label in groups else None
            if not all_src:
                all_src = [src_path]

            user_append = bool(cfg.get("append_mode", False))
            errors = []
            last_msg = ""
            for idx, sp in enumerate(all_src):
                # First file: respect user's append_mode.
                # Subsequent files: always append so data accumulates.
                mode = user_append if idx == 0 else True
                fn_col = cfg.get("filename_col") or None
                ok, msg = run_import_step(
                    sp,
                    tgt_path,
                    cfg.get("src_sheet"),
                    cfg.get("tgt_sheet"),
                    cfg.get("mapping"),
                    src_header_row=int(cfg.get("src_header_row", 1) or 1),
                    tgt_header_row=int(cfg.get("tgt_header_row", 1) or 1),
                    append_mode=mode,
                    filters=cfg.get("filters") or [],
                    filter_combine=cfg.get("filter_combine", "AND"),
                    filename_col=fn_col,
                    filename_value=os.path.basename(sp) if fn_col else None,
                )
                if ok:
                    last_msg = msg
                else:
                    errors.append(f"{os.path.basename(sp)}: {msg}")

            if errors:
                partial = len(all_src) - len(errors)
                return partial > 0, f"Import: {partial}/{len(all_src)} files OK. Errors: {'; '.join(errors)}"
            suffix = f"({len(all_src)} files)"if len(all_src) > 1 else ""
            return True, f"{last_msg}{suffix}"

        # header_map
        return run_header_mapping_step(
            src_path, tgt_path, cfg.get("src_sheet"), cfg.get("tgt_sheet"), cfg.get("mapping")
        )

    # Explicit VLOOKUP uses 4 file inputs (no single cfg["file"])
    if step_type == "vlookup"and (cfg.get("mode") == "explicit"or cfg.get("lookup_value_file")):
        lv_label = cfg.get("lookup_value_file")
        s_label = cfg.get("search_file")
        r_label = cfg.get("return_file")
        d_label = cfg.get("return_to_file")

        lv_path, err = _resolve_required_path(lv_label, field="lookup_value_file")
        if err:
            return False, err
        s_path, err = _resolve_required_path(s_label, field="search_file")
        if err:
            return False, err
        r_path, err = _resolve_required_path(r_label, field="return_file")
        if err:
            return False, err
        d_path, err = _resolve_required_path(d_label, field="return_to_file")
        if err:
            return False, err

        # Resolve optional fallback lookups (chain)
        fb_lookups = []
        for fb in cfg.get("fallback_lookups", []) or []:
            fb = fb or {}
            fb_s_label = fb.get("search_file")
            fb_r_label = fb.get("return_file")
            fb_s_path, err = _resolve_required_path(fb_s_label, field="fallback_search_file")
            if err:
                return False, err
            fb_r_path, err = _resolve_required_path(fb_r_label, field="fallback_return_file")
            if err:
                return False, err
            fb_lookups.append({
                "search_file_path": fb_s_path,
                "search_sheet_name": fb.get("search_sheet"),
                "search_header_row": int(fb.get("search_header_row", 1) or 1),
                "search_columns": fb.get("search_columns") or fb.get("search_column"),
                "return_file_path": fb_r_path,
                "return_sheet_name": fb.get("return_sheet"),
                "return_header_row": int(fb.get("return_header_row", 1) or 1),
                "return_column": fb.get("return_column"),
            })

        return run_vlookup_explicit_step(
            lookup_value_file_path=lv_path,
            lookup_value_sheet_name=cfg.get("lookup_value_sheet"),
            lookup_value_header_row=int(cfg.get("lookup_value_header_row", 1) or 1),
            lookup_value_column=cfg.get("lookup_value_column"),
            search_file_path=s_path,
            search_sheet_name=cfg.get("search_sheet"),
            search_header_row=int(cfg.get("search_header_row", 1) or 1),
            search_column=cfg.get("search_column"),
            return_file_path=r_path,
            return_sheet_name=cfg.get("return_sheet"),
            return_header_row=int(cfg.get("return_header_row", 1) or 1),
            return_column=cfg.get("return_column"),
            return_to_file_path=d_path,
            return_to_sheet_name=cfg.get("return_to_sheet"),
            return_to_header_row=int(cfg.get("return_to_header_row", 1) or 1),
            return_to_column=cfg.get("return_to_column"),
            not_found_value=cfg.get("not_found_value", ""),
            end_row=cfg.get("end_row", "last"),
            fallback_lookups=fb_lookups,
        )

    if step_type == "unpivot":
        _up_src_path, _up_err = _resolve_required_path(cfg.get("src_file"), field="src_file")
        if _up_err: return False, _up_err
        _up_tgt_label = cfg.get("tgt_file")
        _up_tgt_p = get_file_path(_up_tgt_label) if _up_tgt_label else ""
        return run_unpivot_step(
            src_file_path=_up_src_path,
            src_sheet=cfg.get("src_sheet", ""),
            src_header_row=int(cfg.get("src_header_row", 1) or 1),
            id_cols=cfg.get("id_cols") or [],
            value_cols=cfg.get("value_cols") or [],
            var_name=cfg.get("var_name", "Category"),
            value_name=cfg.get("value_name", "Value"),
            tgt_file=_up_tgt_p or "",
            tgt_sheet=cfg.get("tgt_sheet", "Unpivot"),
            append_mode=bool(cfg.get("append_mode", False)),
            drop_na=bool(cfg.get("drop_na", True)),
        )

    if step_type == "formula_broadcast":
        _fb_scope = cfg.get("scope", "all_sheets")
        _fb_file_label = cfg.get("file", "")
        _fb_fp = get_file_path(_fb_file_label) if _fb_file_label else ""
        # For all_files: collect every file in the group
        _fb_fps = []
        if _fb_scope == "all_files":
            _grp = st.session_state.get("uploaded_file_groups", {}).get(_fb_file_label)
            if _grp:
                _fb_fps = list(_grp)
            elif _fb_fp:
                _fb_fps = [_fb_fp]
        _incl = [s.strip() for s in (cfg.get("include_sheets") or "").split(",") if s.strip()]
        _excl = [s.strip() for s in (cfg.get("exclude_sheets") or "").split(",") if s.strip()]
        return run_formula_broadcast_step(
            scope=_fb_scope,
            formula=cfg.get("formula", ""),
            cell=cfg.get("cell", ""),
            file_path=_fb_fp or "",
            file_paths=_fb_fps,
            sheet_name=cfg.get("sheet_name", ""),
            include_sheets=_incl,
            exclude_sheets=_excl,
        )

    if step_type == "sheet_updater":
        src_path, err = _resolve_required_path(cfg.get("src_file"), field="src_file")
        if err: return False, err
        _tgt_label = cfg.get("tgt_file")
        _tgt_p = get_file_path(_tgt_label) if _tgt_label else ""
        return run_sheet_updater_step(
            src_file_path=src_path,
            src_sheet=cfg.get("src_sheet", ""),
            src_col_category=cfg.get("src_col_category", ""),
            src_col_particulars=cfg.get("src_col_particulars", ""),
            src_col_value=cfg.get("src_col_value", ""),
            tgt_file=_tgt_p or "",
            sheet_match_mode=cfg.get("sheet_match_mode", "cat_in_sheet"),
            tgt_col_lookup=cfg.get("tgt_col_lookup", ""),
            tgt_col_write=cfg.get("tgt_col_write", ""),
            tgt_header_row=int(cfg.get("tgt_header_row", 1) or 1),
            case_sensitive=bool(cfg.get("case_sensitive", False)),
            particulars_match_mode=cfg.get("particulars_match_mode", "iexact"),
            particulars_rules=cfg.get("particulars_rules") or [],
        )

    if step_type == "sheet_row_inserter":
        src_path, err = _resolve_required_path(cfg.get("src_file"), field="src_file")
        if err: return False, err
        _tgt_label = cfg.get("tgt_file")
        _tgt_p = get_file_path(_tgt_label) if _tgt_label else ""
        return run_sheet_row_inserter_step(
            src_file_path=src_path,
            src_sheet=cfg.get("src_sheet", ""),
            src_col_category=cfg.get("src_col_category", ""),
            src_col_particulars=cfg.get("src_col_particulars", ""),
            src_col_value=cfg.get("src_col_value", ""),
            tgt_file=_tgt_p or "",
            sheet_match_mode=cfg.get("sheet_match_mode", "cat_in_sheet"),
            anchor_col=cfg.get("anchor_col", ""),
            anchor_mode=cfg.get("anchor_mode", "particulars"),
            anchor_value=cfg.get("anchor_value", ""),
            anchor_match=cfg.get("anchor_match", "exact"),
            insert_position=cfg.get("insert_position", "after_last"),
            tgt_col_particulars=cfg.get("tgt_col_particulars", ""),
            tgt_col_value=cfg.get("tgt_col_value", ""),
            extra_col_map=cfg.get("extra_col_map") or [],
            tgt_header_row=int(cfg.get("tgt_header_row", 1) or 1),
        )

    if step_type == "split_export":
        src_path, err = _resolve_required_path(cfg.get("src_file"), field="src_file")
        if err: return False, err
        _tgt_label = cfg.get("tgt_file")
        _tgt_p = get_file_path(_tgt_label) if _tgt_label else ""
        _tgt_resolved = _tgt_p or cfg.get("tgt_file_path", "")
        return run_split_export_step(
            src_file_path=src_path,
            src_sheet=cfg.get("src_sheet", ""),
            src_header_row=int(cfg.get("src_header_row", 1) or 1),
            split_mode=cfg.get("split_mode", "single"),
            split_col=cfg.get("split_col", ""),
            split_col2=cfg.get("split_col2", ""),
            output_folder=cfg.get("output_folder", ""),
            tgt_file=_tgt_resolved,
            tgt_sheet=cfg.get("tgt_sheet", "Export"),
            file_name_template=cfg.get("file_name_template", "{value}.xlsx"),
            sheet_name_template=cfg.get("sheet_name_template", "{value}"),
            sheet_name_template2=cfg.get("sheet_name_template2", "{value}"),
            filter_conditions=cfg.get("filter_conditions") or [],
            filter_combine=cfg.get("filter_combine", "AND"),
            include_columns=cfg.get("include_columns") or [],
            exclude_columns=cfg.get("exclude_columns") or [],
            sort_cols=cfg.get("sort_cols") or [],
            sort_ascending=cfg.get("sort_ascending", True),
            include_header=bool(cfg.get("include_header", True)),
            overwrite=bool(cfg.get("overwrite", True)),
            max_rows_per_sheet=int(cfg.get("max_rows_per_sheet", 0) or 0),
            add_summary_sheet=bool(cfg.get("add_summary_sheet", False)),
            summary_sheet_name=cfg.get("summary_sheet_name", "_Summary"),
            apply_formatting=bool(cfg.get("apply_formatting", True)),
            header_bg_color=cfg.get("header_bg_color", "1F3864"),
            header_font_color=cfg.get("header_font_color", "FFFFFF"),
        )

    # All other steps require file selection / existence
    file_label = cfg.get('file')
    file_path, err = _resolve_required_path(file_label, field="file")
    if err:
        return False, err

    try:
        if step_type == "formula":
            ok, msg = run_formula_step(
                file_path, cfg.get('sheet'), cfg.get('formula'),
                _resolve_formula_column(file_path, cfg.get('sheet'), cfg),
                cfg.get('start'), cfg.get('end'),
                cfg.get('convert', False)
            )
            return ok, f"{msg} (saved to original: {file_path})"
        
        elif step_type == "copy_paste":
            ok, msg = run_copy_paste_step(
                file_path, cfg.get('sheet'), cfg.get('src_col'),
                cfg.get('tgt_cols') if cfg.get('tgt_cols') else cfg.get('tgt_col'),
                cfg.get('paste_type'),
                cfg.get('header_row', 1)
            )
            return ok, f"{msg} (saved to original: {file_path})"
        
        elif step_type == "convert_values":
            ok, msg = run_convert_to_values_step(
                file_path, cfg.get('sheet'), cfg.get('column'),
                cfg.get('start'), cfg.get('end')
            )
            return ok, f"{msg} (saved to original: {file_path})"
        
        elif step_type == "forward_fill":
            ok, msg = run_forward_fill_step(
                file_path, cfg.get('sheet'), cfg.get('column')
            )
            return ok, f"{msg} (saved to original: {file_path})"
        
        elif step_type == "replace":
            ok, msg = run_replace_single(
                file_path, cfg.get('sheet'), cfg.get('old_val'),
                cfg.get('new_val'), cfg.get('column')
            )
            return ok, f"{msg} (saved to original: {file_path})"
        
        elif step_type == "conditional":
            ok, msg = run_conditional_write_step(
                file_path,
                cfg.get("sheet"),
                condition_col=cfg.get("cond_col"),
                condition_value=cfg.get("cond_val", ""),
                target_col=cfg.get("target_col"),
                write_value=cfg.get("write_val", ""),
                operator=cfg.get("operator", "equals"),
                header_row=int(cfg.get("header_row", 1) or 1),
                col_mode=cfg.get("col_mode", "letter"),
                rules=cfg.get("rules") or None,
                scope=cfg.get("scope", "single"),
                include_sheets=cfg.get("include_sheets") or [],
                exclude_sheets=cfg.get("exclude_sheets") or [],
            )
            return ok, f"{msg} (saved to original: {file_path})"

        elif step_type == "write_cell":
            ok, msg = run_write_cell_step(
                file_path,
                cfg.get('sheet'),
                cfg.get('mode', 'single_cell'),
                cfg.get('value', ''),
                cell_ref=cfg.get('cell_ref', ''),
                column=cfg.get('column', ''),
                start_row=int(cfg.get('start_row', 2)),
                end_row=cfg.get('end_row', 'last'),
                header_row=int(cfg.get('header_row', 1)),
                scope=cfg.get('scope', 'single'),
                include_sheets=cfg.get('include_sheets') or [],
                exclude_sheets=cfg.get('exclude_sheets') or [],
            )
            return ok, f"{msg} (saved to original: {file_path})"
        
        elif step_type == "vlookup":
            # Backward-compatible legacy VLOOKUP
            lookup_path = get_file_path(cfg.get('lookup_file'))
            ok, msg = run_vlookup_step(
                file_path, cfg.get('sheet'), cfg.get('main_col'),
                [{'file_path': lookup_path, 'sheet_name': cfg.get('lookup_sheet'),
                  'key_col': cfg.get('key_col'), 'return_col': cfg.get('return_col')}],
                cfg.get('fallback')
            )
            return ok, f"{msg} (saved to original: {file_path})"
        
        elif step_type == "advance_vlookup":
            # Check if this is the new v2 format (has lookup_value_columns)
            if cfg.get('lookup_value_columns'):
                # New v2 format with full dropdown interface
                from operations import run_advance_vlookup_v2
                
                # Resolve file paths in config
                v2_config = dict(cfg)
                v2_config['condition_file'], err = _resolve_required_path(cfg.get('condition_file'), field="condition_file")
                if err:
                    return False, err
                v2_config['lookup_value_file'], err = _resolve_required_path(cfg.get('lookup_value_file'), field="lookup_value_file")
                if err:
                    return False, err
                v2_config['search_file'], err = _resolve_required_path(cfg.get('search_file'), field="search_file")
                if err:
                    return False, err
                v2_config['return_file'], err = _resolve_required_path(cfg.get('return_file'), field="return_file")
                if err:
                    return False, err
                v2_config['destination_file'], err = _resolve_required_path(cfg.get('destination_file'), field="destination_file")
                if err:
                    return False, err
                if cfg.get('else_source_file'):
                    v2_config['else_source_file'], err = _resolve_required_path(cfg.get('else_source_file'), field="else_source_file")
                    if err:
                        return False, err
                
                ok, msg = run_advance_vlookup_v2(v2_config)
                return ok, f"{msg} (saved to original: {v2_config['destination_file']})"
            else:
                # Legacy format with prefix + source column
                lookup_path = get_file_path(cfg.get('lookup_file'))
                ok, msg = run_advance_vlookup_step(
                    file_path, cfg.get('sheet'), cfg.get('prefix'),
                    cfg.get('src_col'), lookup_path, cfg.get('lookup_sheet'),
                    cfg.get('key_col'), cfg.get('return_col'), cfg.get('output_col'),
                    fallback_value=cfg.get('fallback_value'),
                    fallback_col=cfg.get('fallback_col'),
                    target_header_row=int(cfg.get('target_header_row', 1) or 1),
                    lookup_header_row=int(cfg.get('lookup_header_row', 1) or 1),
                    condition_enabled=bool(cfg.get('condition_enabled', False)),
                    condition_column=cfg.get('condition_column'),
                    condition_value=cfg.get('condition_value'),
                    else_source_column=cfg.get('else_source_column'),
                )
                return ok, f"{msg} (saved to original: {file_path})"

        elif step_type == "quota_update":
            src_path, err = _resolve_required_path(cfg.get("source_file"), field="source_file")
            if err:
                return False, err
            tgt_path, err = _resolve_required_path(cfg.get("target_file"), field="target_file")
            if err:
                return False, err
            ok, msg = run_quota_update_step(
                source_path=src_path,
                source_sheet=cfg.get("source_sheet"),
                source_header_row=int(cfg.get("source_header_row", 1) or 1),
                source_filter_column=cfg.get("source_filter_column"),
                source_filter_operator=cfg.get("source_filter_operator", "equals"),
                source_filter_value=cfg.get("source_filter_value", ""),
                source_key_columns=list(cfg.get("source_key_columns") or []),
                source_key_norms=list(cfg.get("source_key_norms") or []),
                target_path=tgt_path,
                target_sheet=cfg.get("target_sheet"),
                target_header_row=int(cfg.get("target_header_row", 1) or 1),
                target_key_columns=list(cfg.get("target_key_columns") or []),
                target_key_norms=list(cfg.get("target_key_norms") or []),
                write_column=cfg.get("write_column"),
                write_value=cfg.get("write_value", ""),
                end_row=cfg.get("target_end_row", "last"),
            )
            return ok, f"{msg} (saved to original: {tgt_path})"

        elif step_type == "group_rollup":
            ok, msg = run_group_rollup_step(
                file_path=file_path,
                sheet_name=cfg.get("sheet"),
                header_row=int(cfg.get("header_row", 1) or 1),
                key_columns=list(cfg.get("key_columns") or []),
                type_column=cfg.get("type_column"),
                amount_column=cfg.get("amount_column"),
                from_type_value=cfg.get("from_type_value", ""),
                into_type_value=cfg.get("into_type_value", ""),
                include_from_in_sum=bool(cfg.get("include_from_in_sum", True)),
                delete_duplicate_into_rows=bool(cfg.get("delete_duplicate_into_rows", True)),
                delete_from_rows=bool(cfg.get("delete_from_rows", False)),
                end_row=cfg.get("end_row", "last"),
            )
            return ok, f"{msg} (saved to original: {file_path})"
        
        elif step_type == "filter":
            ok, msg = run_filter_step(
                file_path,
                cfg.get('sheet'),
                cfg.get('column'),
                cfg.get('value'),
                remove_empty=bool(cfg.get("remove_empty", False)),
                header_row=int(cfg.get("header_row", 1) or 1),
            )
            return ok, f"{msg} (saved to original: {file_path})"
        
        elif step_type == "sort":
            if isinstance(cfg.get("sort_cols"), list) and cfg.get("sort_cols"):
                cols = cfg.get("sort_cols") or []
                ascending = cfg.get("sort_ascending", True)
                header_row = int(cfg.get("header_row", 1) or 1)
            else:
                # Backward compatible: old config used comma-separated string
                cols = [c.strip() for c in (cfg.get('columns', '') or '').split(',') if c.strip()]
                ascending = cfg.get('ascending', True)
                header_row = int(cfg.get("header_row", 1) or 1)
            ok, msg = run_sort_step(
                file_path,
                cfg.get('sheet'),
                cols,
                ascending=ascending,
                header_row=header_row,
            )
            return ok, f"{msg} (saved to original: {file_path})"
        
        elif step_type == "insert":
            ok, msg = run_insert_delete_step(
                file_path, cfg.get('sheet'), "insert", cfg.get('axis'),
                cfg.get('start'), cfg.get('end')
            )
            return ok, f"{msg} (saved to original: {file_path})"
        
        elif step_type == "delete":
            if cfg.get("row_mode") == "advanced_condition":
                ok, msg = run_advanced_delete_step(
                    file_path=file_path,
                    sheet_scope=cfg.get("sheet_scope", "single"),
                    sheet_name=cfg.get("sheet", ""),
                    selected_sheets=cfg.get("selected_sheets") or [],
                    include_sheets=cfg.get("include_sheets") or [],
                    exclude_sheets=cfg.get("exclude_sheets") or [],
                    filters=cfg.get("filters") or [],
                    filter_combine=cfg.get("filter_combine", "AND"),
                    header_row=int(cfg.get("header_row", 1) or 1),
                    save_log=bool(cfg.get("save_log", True)),
                    log_path=cfg.get("log_path", ""),
                )
                return ok, f"{msg} (saved to original: {file_path})"
            if str(cfg.get("axis", "")).lower() == "column"and cfg.get("columns"):
                if cfg.get("action") == "clear_data":
                    ok, msg = run_clear_columns_data_step(
                        file_path,
                        cfg.get("sheet"),
                        cfg.get("columns"),
                        header_row=int(cfg.get("header_row", 1) or 1),
                        end_row=cfg.get("end_row", "last"),
                    )
                else:
                    ok, msg = run_delete_columns_step(
                        file_path,
                        cfg.get("sheet"),
                        cfg.get("columns"),
                        header_row=int(cfg.get("header_row", 1) or 1),
                    )
            else:
                ok, msg = run_insert_delete_step(
                    file_path, cfg.get('sheet'), "delete", cfg.get('axis'),
                    cfg.get('start'), cfg.get('end')
                )
            return ok, f"{msg} (saved to original: {file_path})"
        
        elif step_type == "merge_cells":
            ok, msg = run_merge_cells_step(
                file_path, cfg.get('sheet'), cfg.get('range')
            )
            return ok, f"{msg} (saved to original: {file_path})"

        elif step_type == "format":
            from operations.formatter import run_format_step as _run_fmt
            ok, msg = _run_fmt(
                file_path,
                cfg.get("sheet"),
                format_mode=cfg.get("format_mode", "static"),
                header_row=int(cfg.get("header_row", 1) or 1),
                col_mode=cfg.get("col_mode", "letter"),
                static_rules=cfg.get("static_rules") or [],
                cond_rules=cfg.get("cond_rules") or [],
            )
            return ok, f"{msg} (saved to original: {file_path})"

        elif step_type == "data_input":
            success, msg, df = run_data_input_step(file_path, cfg.get('sheet'))
            return success, msg
        
        elif step_type == "sumifs":
            src_path_si = get_file_path(cfg.get('src_workbook')) or cfg.get('src_workbook') or ""
            ok, msg = run_sumifs_step(
                target_file_path=file_path,
                target_sheet=cfg.get('sheet'),
                source_file_path=src_path_si,
                source_sheet=cfg.get('src_sheet'),
                src_lookup_col=cfg.get('src_lookup_col') or cfg.get('sum_range', 'A'),
                tgt_lookup_col=cfg.get('tgt_lookup_col', 'A'),
                sum_col=cfg.get('sum_col') or cfg.get('sum_range', 'B'),
                output_col=cfg.get('output_col', 'B'),
                src_header_row=int(cfg.get('src_header_row', 1) or 1),
                tgt_header_row=int(cfg.get('tgt_header_row', 1) or 1),
            )
            return ok, f"{msg} (saved to original: {file_path})"
        
        else:
            return False, f"Unknown operation: {step_type}"
    
    except Exception as e:
        return False, f"Error: {str(e)}"


def execute_step_with_file_map(step: dict, file_map: dict) -> tuple[bool, str]:
    """
    Thread-safe step execution: does NOT access Streamlit APIs or st.session_state.
    `file_map` must be a {label -> absolute_path} mapping snapshot from the main thread.
    """
    step_type = (step.get("type") or "").lower().strip()
    cfg = step.get("config", {}) or {}

    def path_for(label: str | None) -> str:
        if not label:
            return ""
        return str(file_map.get(label, "") or "")

    def registered_labels_preview() -> str:
        labels = sorted([str(k) for k in file_map.keys()])
        preview = ", ".join([f"'{x}'"for x in labels[:15]])
        more = ""if len(labels) <= 15 else f"(+{len(labels) - 15} more)"
        return f"[{preview}]{more}"

    def resolve_required_path(label: str | None, *, field: str) -> tuple[str, str | None]:
        if not label:
            return "", f"Missing required file selection for field '{field}'in step type '{step_type}'."
        p = path_for(label)
        if p:
            return p, None
        return "", (
            f"File not found for field '{field}': '{label}'. "
            f"Register it via Add File(s) step first. "
            f"Currently registered: {registered_labels_preview()}"
        )

    # Steps that don't require a pre-registered target file
    if step_type in ("add_files", "input_source"):
        # Must be handled in main thread (updates file_map/session_state)
        return False, f"{step_type} must run in main thread in DAG mode"

    if step_type in ("import", "append", "header_map"):
        src_label = cfg.get("src_file")
        tgt_label = cfg.get("tgt_file")
        src_path, err = resolve_required_path(src_label, field="src_file")
        if err:
            return False, err
        tgt_path, err = resolve_required_path(tgt_label, field="tgt_file")
        if err:
            return False, err

        if step_type == "append":
            return run_append_step(src_path, tgt_path, cfg.get("src_sheet"), cfg.get("tgt_sheet"))

        if step_type == "import":
            return run_import_step(
                src_path,
                tgt_path,
                cfg.get("src_sheet"),
                cfg.get("tgt_sheet"),
                cfg.get("mapping"),
                src_header_row=int(cfg.get("src_header_row", 1) or 1),
                tgt_header_row=int(cfg.get("tgt_header_row", 1) or 1),
                append_mode=bool(cfg.get("append_mode", False)),
                filters=cfg.get("filters") or [],
                filter_combine=cfg.get("filter_combine", "AND"),
            )

        return run_header_mapping_step(src_path, tgt_path, cfg.get("src_sheet"), cfg.get("tgt_sheet"), cfg.get("mapping"))

    if step_type == "vlookup"and (cfg.get("mode") == "explicit"or cfg.get("lookup_value_file")):
        lv_label = cfg.get("lookup_value_file")
        s_label = cfg.get("search_file")
        r_label = cfg.get("return_file")
        d_label = cfg.get("return_to_file")

        lv_path, err = resolve_required_path(lv_label, field="lookup_value_file")
        if err:
            return False, err
        s_path, err = resolve_required_path(s_label, field="search_file")
        if err:
            return False, err
        r_path, err = resolve_required_path(r_label, field="return_file")
        if err:
            return False, err
        d_path, err = resolve_required_path(d_label, field="return_to_file")
        if err:
            return False, err

        # Resolve optional fallback lookups (chain)
        fb_lookups = []
        for fb in cfg.get("fallback_lookups", []) or []:
            fb = fb or {}
            fb_s_label = fb.get("search_file")
            fb_r_label = fb.get("return_file")
            fb_s_path, err = resolve_required_path(fb_s_label, field="fallback_search_file")
            if err:
                return False, err
            fb_r_path, err = resolve_required_path(fb_r_label, field="fallback_return_file")
            if err:
                return False, err
            fb_lookups.append({
                "search_file_path": fb_s_path,
                "search_sheet_name": fb.get("search_sheet"),
                "search_header_row": int(fb.get("search_header_row", 1) or 1),
                "search_columns": fb.get("search_columns") or fb.get("search_column"),
                "return_file_path": fb_r_path,
                "return_sheet_name": fb.get("return_sheet"),
                "return_header_row": int(fb.get("return_header_row", 1) or 1),
                "return_column": fb.get("return_column"),
            })

        return run_vlookup_explicit_step(
            lookup_value_file_path=lv_path,
            lookup_value_sheet_name=cfg.get("lookup_value_sheet"),
            lookup_value_header_row=int(cfg.get("lookup_value_header_row", 1) or 1),
            lookup_value_column=cfg.get("lookup_value_column"),
            search_file_path=s_path,
            search_sheet_name=cfg.get("search_sheet"),
            search_header_row=int(cfg.get("search_header_row", 1) or 1),
            search_column=cfg.get("search_column"),
            return_file_path=r_path,
            return_sheet_name=cfg.get("return_sheet"),
            return_header_row=int(cfg.get("return_header_row", 1) or 1),
            return_column=cfg.get("return_column"),
            return_to_file_path=d_path,
            return_to_sheet_name=cfg.get("return_to_sheet"),
            return_to_header_row=int(cfg.get("return_to_header_row", 1) or 1),
            return_to_column=cfg.get("return_to_column"),
            not_found_value=cfg.get("not_found_value", ""),
            end_row=cfg.get("end_row", "last"),
            fallback_lookups=fb_lookups,
        )

    # Single-file ops
    file_label = cfg.get("file")
    file_path, err = resolve_required_path(file_label, field="file")
    if err:
        return False, err

    if step_type == "formula":
        return run_formula_step(
            file_path, cfg.get("sheet"), cfg.get("formula"),
            _resolve_formula_column(file_path, cfg.get("sheet"), cfg),
            cfg.get("start"), cfg.get("end"),
            cfg.get("convert", False),
        )
    if step_type == "copy_paste":
        return run_copy_paste_step(
            file_path, cfg.get("sheet"), cfg.get("src_col"),
            cfg.get("tgt_cols") if cfg.get("tgt_cols") else cfg.get("tgt_col"),
            cfg.get("paste_type"),
            cfg.get("header_row", 1),
        )
    if step_type == "convert_values":
        return run_convert_to_values_step(file_path, cfg.get("sheet"), cfg.get("column"), cfg.get("start"), cfg.get("end"))
    if step_type == "forward_fill":
        return run_forward_fill_step(file_path, cfg.get("sheet"), cfg.get("column"))
    if step_type == "replace":
        return run_replace_single(file_path, cfg.get("sheet"), cfg.get("old_val"), cfg.get("new_val"), cfg.get("column"))
    if step_type == "conditional":
        return run_conditional_write_step(
            file_path,
            cfg.get("sheet"),
            condition_col=cfg.get("cond_col"),
            condition_value=cfg.get("cond_val", ""),
            target_col=cfg.get("target_col"),
            write_value=cfg.get("write_val", ""),
            operator=cfg.get("operator", "equals"),
            header_row=int(cfg.get("header_row", 1) or 1),
            col_mode=cfg.get("col_mode", "letter"),
            rules=cfg.get("rules") or None,
            scope=cfg.get("scope", "single"),
            include_sheets=cfg.get("include_sheets") or [],
            exclude_sheets=cfg.get("exclude_sheets") or [],
        )
    if step_type == "write_cell":
        return run_write_cell_step(
            file_path, cfg.get("sheet"), cfg.get("mode", "single_cell"), cfg.get("value", ""),
            cell_ref=cfg.get("cell_ref", ""), column=cfg.get("column", ""),
            start_row=int(cfg.get("start_row", 2)), end_row=cfg.get("end_row", "last"),
            header_row=int(cfg.get("header_row", 1)),
            scope=cfg.get("scope", "single"),
            include_sheets=cfg.get("include_sheets") or [],
            exclude_sheets=cfg.get("exclude_sheets") or [],
        )
    if step_type == "filter":
        return run_filter_step(
            file_path,
            cfg.get("sheet"),
            cfg.get("column"),
            cfg.get("value"),
            remove_empty=bool(cfg.get("remove_empty", False)),
            header_row=int(cfg.get("header_row", 1) or 1),
        )
    if step_type == "sort":
        if isinstance(cfg.get("sort_cols"), list) and cfg.get("sort_cols"):
            cols = cfg.get("sort_cols") or []
            ascending = cfg.get("sort_ascending", True)
            header_row = int(cfg.get("header_row", 1) or 1)
        else:
            cols = [c.strip() for c in (cfg.get("columns", "") or "").split(",") if c.strip()]
            ascending = cfg.get("ascending", True)
            header_row = int(cfg.get("header_row", 1) or 1)
        return run_sort_step(file_path, cfg.get("sheet"), cols, ascending=ascending, header_row=header_row)
    if step_type == "insert":
        return run_insert_delete_step(file_path, cfg.get("sheet"), "insert", cfg.get("axis"), cfg.get("start"), cfg.get("end"))
    if step_type == "delete":
        if cfg.get("row_mode") == "advanced_condition":
            return run_advanced_delete_step(
                file_path=file_path,
                sheet_scope=cfg.get("sheet_scope", "single"),
                sheet_name=cfg.get("sheet", ""),
                selected_sheets=cfg.get("selected_sheets") or [],
                include_sheets=cfg.get("include_sheets") or [],
                exclude_sheets=cfg.get("exclude_sheets") or [],
                filters=cfg.get("filters") or [],
                filter_combine=cfg.get("filter_combine", "AND"),
                header_row=int(cfg.get("header_row", 1) or 1),
                save_log=bool(cfg.get("save_log", True)),
                log_path=cfg.get("log_path", ""),
            )
        if cfg.get("row_mode") == "condition":
            return run_delete_by_condition_step(
                file_path, cfg.get("sheet"),
                filters=cfg.get("filters") or [],
                filter_combine=cfg.get("filter_combine", "AND"),
                header_row=int(cfg.get("header_row", 1) or 1),
            )
        if str(cfg.get("axis", "")).lower() == "column"and cfg.get("columns"):
            if cfg.get("action") == "clear_data":
                return run_clear_columns_data_step(
                    file_path,
                    cfg.get("sheet"),
                    cfg.get("columns"),
                    header_row=int(cfg.get("header_row", 1) or 1),
                    end_row=cfg.get("end_row", "last"),
                )
            return run_delete_columns_step(
                file_path,
                cfg.get("sheet"),
                cfg.get("columns"),
                header_row=int(cfg.get("header_row", 1) or 1),
            )
        return run_insert_delete_step(file_path, cfg.get("sheet"), "delete", cfg.get("axis"), cfg.get("start"), cfg.get("end"))
    if step_type == "merge_cells":
        return run_merge_cells_step(file_path, cfg.get("sheet"), cfg.get("range"))
    if step_type == "delete_sheets":
        return run_delete_sheets_step(
            file_path,
            mode=cfg.get("mode", "named"),
            sheet_names=cfg.get("sheet_names") or [],
            pattern=cfg.get("pattern", ""),
            pattern_type=cfg.get("pattern_type", "contains"),
            case_sensitive=bool(cfg.get("case_sensitive", False)),
            protect_last=bool(cfg.get("protect_last", True)),
        )
    if step_type == "format":
        return run_format_step(
            file_path,
            cfg.get("sheet"),
            format_mode=cfg.get("format_mode", "static"),
            header_row=int(cfg.get("header_row", 1) or 1),
            col_mode=cfg.get("col_mode", "letter"),
            static_rules=cfg.get("static_rules") or [],
            cond_rules=cfg.get("cond_rules") or [],
        )
    if step_type == "sumifs":
        src_path_si = path_for(cfg.get("src_workbook")) or cfg.get("src_workbook") or ""
        return run_sumifs_step(
            target_file_path=file_path,
            target_sheet=cfg.get("sheet"),
            source_file_path=src_path_si,
            source_sheet=cfg.get("src_sheet"),
            src_lookup_col=cfg.get("src_lookup_col") or cfg.get("sum_range", "A"),
            tgt_lookup_col=cfg.get("tgt_lookup_col", "A"),
            sum_col=cfg.get("sum_col") or cfg.get("sum_range", "B"),
            output_col=cfg.get("output_col", "B"),
            src_header_row=int(cfg.get("src_header_row", 1) or 1),
            tgt_header_row=int(cfg.get("tgt_header_row", 1) or 1),
        )
    if step_type == "quota_update":
        src_path, err = resolve_required_path(cfg.get("source_file"), field="source_file")
        if err:
            return False, err
        tgt_path, err = resolve_required_path(cfg.get("target_file"), field="target_file")
        if err:
            return False, err
        return run_quota_update_step(
            source_path=src_path,
            source_sheet=cfg.get("source_sheet"),
            source_header_row=int(cfg.get("source_header_row", 1) or 1),
            source_filter_column=cfg.get("source_filter_column"),
            source_filter_operator=cfg.get("source_filter_operator", "equals"),
            source_filter_value=cfg.get("source_filter_value", ""),
            source_key_columns=list(cfg.get("source_key_columns") or []),
            source_key_norms=list(cfg.get("source_key_norms") or []),
            target_path=tgt_path,
            target_sheet=cfg.get("target_sheet"),
            target_header_row=int(cfg.get("target_header_row", 1) or 1),
            target_key_columns=list(cfg.get("target_key_columns") or []),
            target_key_norms=list(cfg.get("target_key_norms") or []),
            write_column=cfg.get("write_column"),
            write_value=cfg.get("write_value", ""),
            end_row=cfg.get("target_end_row", "last"),
        )
    if step_type == "group_rollup":
        return run_group_rollup_step(
            file_path=file_path,
            sheet_name=cfg.get("sheet"),
            header_row=int(cfg.get("header_row", 1) or 1),
            key_columns=list(cfg.get("key_columns") or []),
            type_column=cfg.get("type_column"),
            amount_column=cfg.get("amount_column"),
            from_type_value=cfg.get("from_type_value", ""),
            into_type_value=cfg.get("into_type_value", ""),
            include_from_in_sum=bool(cfg.get("include_from_in_sum", True)),
            delete_duplicate_into_rows=bool(cfg.get("delete_duplicate_into_rows", True)),
            delete_from_rows=bool(cfg.get("delete_from_rows", False)),
            end_row=cfg.get("end_row", "last"),
        )
    if step_type == "advance_vlookup":
        # Check if this is the new v2 format (has lookup_value_columns)
        if cfg.get('lookup_value_columns'):
            # New v2 format with full dropdown interface
            from operations import run_advance_vlookup_v2
            
            # Resolve file paths in config
            v2_config = dict(cfg)
            v2_config['condition_file'], err = resolve_required_path(cfg.get('condition_file'), field="condition_file")
            if err:
                return False, err
            v2_config['lookup_value_file'], err = resolve_required_path(cfg.get('lookup_value_file'), field="lookup_value_file")
            if err:
                return False, err
            v2_config['search_file'], err = resolve_required_path(cfg.get('search_file'), field="search_file")
            if err:
                return False, err
            v2_config['return_file'], err = resolve_required_path(cfg.get('return_file'), field="return_file")
            if err:
                return False, err
            v2_config['destination_file'], err = resolve_required_path(cfg.get('destination_file'), field="destination_file")
            if err:
                return False, err
            if cfg.get('else_source_file'):
                v2_config['else_source_file'], err = resolve_required_path(cfg.get('else_source_file'), field="else_source_file")
                if err:
                    return False, err
            
            return run_advance_vlookup_v2(v2_config)
        else:
            # Legacy format with prefix + source column
            lookup_path = path_for(cfg.get("lookup_file"))
            return run_advance_vlookup_step(
                file_path, cfg.get("sheet"), cfg.get("prefix"),
                cfg.get("src_col"), lookup_path, cfg.get("lookup_sheet"),
                cfg.get("key_col"), cfg.get("return_col"), cfg.get("output_col"),
                fallback_value=cfg.get("fallback_value"),
                fallback_col=cfg.get("fallback_col"),
                target_header_row=int(cfg.get("target_header_row", 1) or 1),
                lookup_header_row=int(cfg.get("lookup_header_row", 1) or 1),
                condition_enabled=bool(cfg.get("condition_enabled", False)),
                condition_column=cfg.get("condition_column"),
                condition_value=cfg.get("condition_value"),
                else_source_column=cfg.get("else_source_column"),
            )
    if step_type == "vlookup":
        # legacy
        lookup_path = path_for(cfg.get("lookup_file"))
        return run_vlookup_step(
            file_path, cfg.get("sheet"), cfg.get("main_col"),
            [{'file_path': lookup_path, 'sheet_name': cfg.get('lookup_sheet'),
              'key_col': cfg.get('key_col'), 'return_col': cfg.get('return_col')}],
            cfg.get('fallback'),
        )

    # data_input is effectively a no-op in a workflow (loads df), keep consistent
    if step_type == "data_input":
        success, msg, _df = run_data_input_step(file_path, cfg.get("sheet"))
        return success, msg

    if step_type == "pivot":
        src_p = path_for(cfg.get("src_file"))
        tgt_p = path_for(cfg.get("tgt_file"))
        if not src_p: return False, f"Source file not found: {cfg.get('src_file')}"
        if not tgt_p: return False, f"Target file not found: {cfg.get('tgt_file')}"
        return run_pivot_step(
            src_file_path=src_p,
            src_sheet=cfg.get("src_sheet"),
            src_header_row=int(cfg.get("src_header_row", 1) or 1),
            tgt_file_path=tgt_p,
            tgt_sheet=cfg.get("tgt_sheet", "Pivot"),
            tgt_start_row=int(cfg.get("tgt_start_row", 1) or 1),
            tgt_start_col=cfg.get("tgt_start_col", "A"),
            overwrite_sheet=bool(cfg.get("overwrite_sheet", True)),
            row_fields=cfg.get("row_fields") or [],
            col_fields=cfg.get("col_fields") or [],
            value_fields=cfg.get("value_fields") or [],
            filter_conditions=cfg.get("filter_conditions") or [],
            filter_combine=cfg.get("filter_combine", "AND"),
            fill_value=cfg.get("fill_value", 0),
            grand_total_rows=bool(cfg.get("grand_total_rows", False)),
            grand_total_cols=bool(cfg.get("grand_total_cols", False)),
            sort_field=cfg.get("sort_field", ""),
            sort_ascending=bool(cfg.get("sort_ascending", True)),
            apply_formatting=bool(cfg.get("apply_formatting", True)),
            header_bg_color=cfg.get("header_bg_color", "1F3864"),
            header_font_color=cfg.get("header_font_color", "FFFFFF"),
            total_bg_color=cfg.get("total_bg_color", "D9E1F2"),
        )

    if step_type == "folder_summary":
        src_label = cfg.get("src_file")
        src_p = path_for(src_label)
        tgt_p = path_for(cfg.get("tgt_file"))
        if not src_p:
            return False, f"Source file not found: {src_label}"
        if not tgt_p:
            return False, f"Target file not found: {cfg.get('tgt_file')}"
        return run_folder_summary_step(
            src_file_paths=[src_p],
            src_sheet=cfg.get("src_sheet", ""),
            header_row=int(cfg.get("header_row", 1) or 1),
            source_mode=cfg.get("source_mode", "folder"),
            sheet_selection=cfg.get("sheet_selection", "all"),
            sheet_pattern=cfg.get("sheet_pattern", ""),
            sheet_list=cfg.get("sheet_list") or [],
            sheet_exclude=cfg.get("sheet_exclude", ""),
            variables=cfg.get("variables") or [],
            formulas=cfg.get("formulas") or [],
            tgt_file=tgt_p,
            tgt_sheet=cfg.get("tgt_sheet", "Summary"),
            tgt_header_row=int(cfg.get("tgt_header_row", 1) or 1),
            filename_col=cfg.get("filename_col", "File Name"),
            sheet_name_col=cfg.get("sheet_name_col", ""),
            append_mode=bool(cfg.get("append_mode", True)),
            include_var_cols=bool(cfg.get("include_var_cols", True)),
            include_status_col=bool(cfg.get("include_status_col", True)),
            status_col_name=cfg.get("status_col_name", "Status"),
            on_missing_default=cfg.get("on_missing_default", 0),
        )

    if step_type == "unpivot":
        src_p = path_for(cfg.get("src_file"))
        if not src_p: return False, f"Source file not found: {cfg.get('src_file')}"
        tgt_p = path_for(cfg.get("tgt_file")) or ""
        return run_unpivot_step(
            src_file_path=src_p,
            src_sheet=cfg.get("src_sheet", ""),
            src_header_row=int(cfg.get("src_header_row", 1) or 1),
            id_cols=cfg.get("id_cols") or [],
            value_cols=cfg.get("value_cols") or [],
            var_name=cfg.get("var_name", "Category"),
            value_name=cfg.get("value_name", "Value"),
            tgt_file=tgt_p,
            tgt_sheet=cfg.get("tgt_sheet", "Unpivot"),
            append_mode=bool(cfg.get("append_mode", False)),
            drop_na=bool(cfg.get("drop_na", True)),
        )

    if step_type == "formula_broadcast":
        _fb_scope = cfg.get("scope", "all_sheets")
        _fb_fp = path_for(cfg.get("file")) or ""
        _fb_fps = [path_for(l) for l in (cfg.get("file_paths_labels") or []) if path_for(l)]
        _incl = [s.strip() for s in (cfg.get("include_sheets") or "").split(",") if s.strip()]
        _excl = [s.strip() for s in (cfg.get("exclude_sheets") or "").split(",") if s.strip()]
        return run_formula_broadcast_step(
            scope=_fb_scope,
            formula=cfg.get("formula", ""),
            cell=cfg.get("cell", ""),
            file_path=_fb_fp,
            file_paths=_fb_fps,
            sheet_name=cfg.get("sheet_name", ""),
            include_sheets=_incl,
            exclude_sheets=_excl,
        )

    if step_type == "sheet_updater":
        src_p = path_for(cfg.get("src_file"))
        if not src_p:
            return False, f"Source file not found: {cfg.get('src_file')}"
        tgt_p = path_for(cfg.get("tgt_file")) or ""
        return run_sheet_updater_step(
            src_file_path=src_p,
            src_sheet=cfg.get("src_sheet", ""),
            src_col_category=cfg.get("src_col_category", ""),
            src_col_particulars=cfg.get("src_col_particulars", ""),
            src_col_value=cfg.get("src_col_value", ""),
            tgt_file=tgt_p,
            sheet_match_mode=cfg.get("sheet_match_mode", "cat_in_sheet"),
            tgt_col_lookup=cfg.get("tgt_col_lookup", ""),
            tgt_col_write=cfg.get("tgt_col_write", ""),
            tgt_header_row=int(cfg.get("tgt_header_row", 1) or 1),
            case_sensitive=bool(cfg.get("case_sensitive", False)),
            particulars_match_mode=cfg.get("particulars_match_mode", "iexact"),
            particulars_rules=cfg.get("particulars_rules") or [],
        )

    if step_type == "sheet_row_inserter":
        src_p = path_for(cfg.get("src_file"))
        if not src_p:
            return False, f"Source file not found: {cfg.get('src_file')}"
        tgt_p = path_for(cfg.get("tgt_file")) or ""
        return run_sheet_row_inserter_step(
            src_file_path=src_p,
            src_sheet=cfg.get("src_sheet", ""),
            src_col_category=cfg.get("src_col_category", ""),
            src_col_particulars=cfg.get("src_col_particulars", ""),
            src_col_value=cfg.get("src_col_value", ""),
            tgt_file=tgt_p,
            sheet_match_mode=cfg.get("sheet_match_mode", "cat_in_sheet"),
            anchor_col=cfg.get("anchor_col", ""),
            anchor_mode=cfg.get("anchor_mode", "particulars"),
            anchor_value=cfg.get("anchor_value", ""),
            anchor_match=cfg.get("anchor_match", "exact"),
            insert_position=cfg.get("insert_position", "after_last"),
            tgt_col_particulars=cfg.get("tgt_col_particulars", ""),
            tgt_col_value=cfg.get("tgt_col_value", ""),
            extra_col_map=cfg.get("extra_col_map") or [],
            tgt_header_row=int(cfg.get("tgt_header_row", 1) or 1),
        )

    if step_type == "split_export":
        src_p = path_for(cfg.get("src_file"))
        if not src_p:
            return False, f"Source file not found: {cfg.get('src_file')}"
        _tgt_label = cfg.get("tgt_file")
        _tgt_p = path_for(_tgt_label) if _tgt_label else ""
        _tgt_resolved = _tgt_p or cfg.get("tgt_file_path", "")
        return run_split_export_step(
            src_file_path=src_p,
            src_sheet=cfg.get("src_sheet", ""),
            src_header_row=int(cfg.get("src_header_row", 1) or 1),
            split_mode=cfg.get("split_mode", "single"),
            split_col=cfg.get("split_col", ""),
            split_col2=cfg.get("split_col2", ""),
            output_folder=cfg.get("output_folder", ""),
            tgt_file=_tgt_resolved,
            tgt_sheet=cfg.get("tgt_sheet", "Export"),
            file_name_template=cfg.get("file_name_template", "{value}.xlsx"),
            sheet_name_template=cfg.get("sheet_name_template", "{value}"),
            sheet_name_template2=cfg.get("sheet_name_template2", "{value}"),
            filter_conditions=cfg.get("filter_conditions") or [],
            filter_combine=cfg.get("filter_combine", "AND"),
            include_columns=cfg.get("include_columns") or [],
            exclude_columns=cfg.get("exclude_columns") or [],
            sort_cols=cfg.get("sort_cols") or [],
            sort_ascending=cfg.get("sort_ascending", True),
            include_header=bool(cfg.get("include_header", True)),
            overwrite=bool(cfg.get("overwrite", True)),
            max_rows_per_sheet=int(cfg.get("max_rows_per_sheet", 0) or 0),
            add_summary_sheet=bool(cfg.get("add_summary_sheet", False)),
            summary_sheet_name=cfg.get("summary_sheet_name", "_Summary"),
            apply_formatting=bool(cfg.get("apply_formatting", True)),
            header_bg_color=cfg.get("header_bg_color", "1F3864"),
            header_font_color=cfg.get("header_font_color", "FFFFFF"),
        )

    return False, f"Unknown operation: {step_type}"


# =============================================================================
# EXPORT HELPERS (per-step "Export as File")
# =============================================================================

def _safe_slug(s: str, max_len: int = 60) -> str:
    s = (s or "").strip().lower()
    s = re.sub(r"[^a-z0-9]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    if not s:
        return "step"
    return s[:max_len]


def _determine_modified_file_labels(step_type: str, cfg: dict) -> list[str]:
    """
    Best-effort: which registered file label(s) are modified by this step.
    We export those files after the step runs.
    """
    if step_type in ("add_files", "data_input", "input_source", "split_export"):
        return []   # split_export writes to new files, never modifies the source
    if step_type in ("sheet_updater", "sheet_row_inserter"):
        return [cfg.get("tgt_file", "")]
    if step_type == "formula_broadcast":
        return [cfg.get("file", "")]
    if step_type == "unpivot":
        return [cfg.get("tgt_file", "")]
    if step_type in ("import", "append", "header_map"):
        return [cfg.get("tgt_file", "")]
    if step_type == "vlookup"and (cfg.get("mode") == "explicit"or cfg.get("lookup_value_file")):
        return [cfg.get("return_to_file", "")]
    if step_type == "quota_update":
        return [cfg.get("target_file", "")]
    # Default: single-file ops use cfg["file"]
    return [cfg.get("file", "")]


def _build_export_path(user_path: str, source_path: str, step_num: int, step_type: str) -> str:
    """
    user_path may be a folder or a file path.
    Always appends a suffix indicating the step.
    Exports as .xlsx (converts CSV/TXT -> XLSX).
    """
    up = (user_path or "").strip()
    if not up:
        up = (st.session_state.get("default_export_path") or "").strip()
    if not up:
        return ""

    src = Path(source_path)
    src_stem = src.stem
    suffix = f"__step{int(step_num):03d}_{_safe_slug(step_type)}"

    p = Path(up).expanduser()
    # If user_path looks like a file path, use it as base
    looks_like_file = p.suffix.lower() in (".xlsx", ".xls", ".csv", ".txt")

    if looks_like_file:
        base_dir = p.parent
        base_stem = p.stem
        out_stem = f"{base_stem}{suffix}"
    else:
        base_dir = p
        out_stem = f"{src_stem}{suffix}"

    # Always export as xlsx
    out_path = (base_dir / f"{out_stem}.xlsx").resolve()
    return str(out_path)


def _export_to_xlsx(source_path: str, export_path: str):
    """
    Export the current source file to an .xlsx at export_path.
    - If source is already xlsx/xls: copy.
    - If source is csv/txt: convert via pandas.
    """
    src = Path(source_path)
    dst = Path(export_path)
    dst.parent.mkdir(parents=True, exist_ok=True)

    lower = src.suffix.lower()
    if lower in (".xlsx", ".xls"):
        shutil.copy2(str(src), str(dst))
        return
    if lower in (".csv", ".txt"):
        df = pd.read_csv(str(src))
        with pd.ExcelWriter(str(dst), engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")
        return
    # Fallback: copy as-is but keep .xlsx extension
    shutil.copy2(str(src), str(dst))


def _prepare_non_destructive_export(step_num: int, step: dict) -> tuple[dict | None, list[tuple[str, str]]]:
    """
    NON-DESTRUCTIVE EXPORT: When export is enabled, copy original file(s) to export path
    BEFORE the operation runs, so the operation runs on the COPY, not the original.
    
    Returns:
        (swapped_paths, original_mappings)
        - swapped_paths: dict mapping label -> export_path (files to operate on)
        - original_mappings: list of (label, original_path) to restore after
        
        If export not enabled, returns (None, [])
    """
    cfg = step.get("config", {}) or {}
    export_cfg = cfg.get("export") or {}
    if not isinstance(export_cfg, dict) or not export_cfg.get("enabled"):
        return None, []

    step_type = (step.get("type") or "").strip().lower()
    user_path = export_cfg.get("path") or export_cfg.get("dir") or export_cfg.get("export_path") or ""

    labels = [x for x in _determine_modified_file_labels(step_type, cfg) if x]
    if not labels:
        return None, []

    swapped_paths = {}
    original_mappings = []

    for label in labels:
        src_path = get_file_path(label)
        if not src_path:
            continue

        out_path = _build_export_path(user_path, src_path, step_num, step_type)
        if not out_path:
            st.warning("Export enabled but no export path provided (set per-step or in sidebar).")
            return None, []

        try:
            # Copy original to export path BEFORE operation
            _export_to_xlsx(src_path, out_path)
            
            # Save original mapping for restoration
            original_mappings.append((label, src_path))
            
            # Temporarily swap registration to point to export copy
            st.session_state.uploaded_file_paths[label] = str(Path(out_path).resolve())
            swapped_paths[label] = out_path
            
        except Exception as e:
            st.warning(f"Export copy failed for '{label}': {e}")
            # Restore any swaps made so far
            for lbl, orig_path in original_mappings:
                st.session_state.uploaded_file_paths[lbl] = orig_path
            return None, []

    return swapped_paths, original_mappings


def _finalize_non_destructive_export(step_num: int, swapped_paths: dict, original_mappings: list[tuple[str, str]]):
    """
    After operation on export copy:
    1. Restore original file registrations (so original files are unchanged)
    2. Register export copies as NEW entries (available for subsequent steps)
    """
    # Restore original mappings (so original file labels point back to originals)
    for label, orig_path in original_mappings:
        st.session_state.uploaded_file_paths[label] = orig_path
    
    # Register export copies as new file entries
    for label, export_path in swapped_paths.items():
        export_name = os.path.basename(export_path)
        st.session_state.uploaded_file_paths[export_name] = str(Path(export_path).resolve())
        st.info(f"Step {step_num} exported (original unchanged): {export_path}")


def _maybe_export_after_step(step_num: int, step: dict):
    """
    LEGACY: If step.config.export.enabled is True BUT non-destructive export wasn't used,
    export the modified output file(s). (This is for backwards compatibility.)
    
    NOTE: With non-destructive export, this function is now a no-op when export is enabled
    because the copy+operation happens BEFORE via _prepare_non_destructive_export.
    """
    cfg = step.get("config", {}) or {}
    export_cfg = cfg.get("export") or {}
    if not isinstance(export_cfg, dict) or not export_cfg.get("enabled"):
        return
    
    # Non-destructive export already handled - skip legacy export
    # This function is kept for compatibility but does nothing when export is enabled
    # because _prepare_non_destructive_export + _finalize_non_destructive_export handles it
    pass


# =============================================================================
# STATUS TAB
# =============================================================================

def render_status_tab():
    """Show status and info"""
    st.markdown('<h2 style="color:#1a1a1a;font-family:IBM Plex Sans,sans-serif;text-transform:uppercase;letter-spacing:.1em;font-size:13px;font-weight:600;border-bottom:2px solid #c8392b;padding-bottom:4px;margin-bottom:12px;">System Status</h2>', unsafe_allow_html=True)
    
    st.success("**ALL 18 OPERATIONS COMPLETE!**")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Operations", "18", "100%")
        st.metric("Functions", "23", "Available")
    
    with col2:
        st.metric("Win32com", "0%", "Removed")
        st.metric("Tests", "12/12", "Passing")
    
    with col3:
        st.metric("Platform", "Cross", "Mac/Linux/Win")
        st.metric("Status", "Ready", "Production")
    
    st.divider()
    
    st.markdown("""
    ###  All 18 Operations Available:
    
    **Transform:** Formula, Copy/Paste, Convert Values, Forward Fill, Replace, Conditional  
    **Merge:** VLOOKUP, Advanced VLOOKUP, SUMIFS, Append, Import, Header Mapping  
    **Structure:** Filter, Sort, Insert R/C, Delete R/C, Merge Cells  
    **Input:** Data Input  
    
    ###  Features:
    -  NO win32com dependency
    -  Cross-platform (Mac/Linux/Windows)
    -  Save/Load workflows as JSON
    -  Visual workflow builder
    -  Step-by-step execution
    -  Error handling
    -  Progress tracking
    
    ###  Documentation:
    - `_START_HERE.md` - Quick start
    - `COMPLETION_SUMMARY.md` - Full details
    - `FINAL_STATUS.txt` - Visual summary
    """)


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def get_files():
    """Get list of uploaded files"""
    # Always include a blank option so selectboxes don't crash when no files exist yet.
    return [""] + list(st.session_state.uploaded_file_paths.keys())

def get_sheets_for_file(file_label: str):
    """
    Return available sheet names for a registered file label.
    For CSVs (or unknown), returns an empty list.
    """
    if not file_label:
        return []
    path = get_file_path(file_label)
    if not path:
        return []
    lower = path.lower()
    if lower.endswith(".csv") or lower.endswith(".txt"):
        return []
    return get_sheet_names_for_file(path) or []

def render_filter_builder(key_prefix: str, available_columns: list[str]) -> tuple[list[dict], str]:
    """
    Reusable multi-condition filter builder.
    Returns (filters_list, combine_logic) where combine_logic is "AND"or "OR".
    Stores state in session_state under key_prefix.
    """
    op_labels = list(OPERATORS.values())
    op_keys   = list(OPERATORS.keys())

    count_key   = f"{key_prefix}_fc"
    combine_key = f"{key_prefix}_combine"

    if count_key not in st.session_state:
        st.session_state[count_key] = 1

    col_add, col_clr, col_combine = st.columns([1, 1, 2])
    with col_add:
        if st.button("Add condition", key=f"{key_prefix}_add_cond"):
            st.session_state[count_key] += 1
    with col_clr:
        if st.button("Clear all", key=f"{key_prefix}_clr_cond"):
            st.session_state[count_key] = 1
            for k in list(st.session_state.keys()):
                if k.startswith(f"{key_prefix}_cond_"):
                    del st.session_state[k]
    with col_combine:
        combine = st.radio(
            "Combine logic",
            ["AND  (all must match)", "OR  (any can match)"],
            key=combine_key,
            horizontal=True,
            label_visibility="collapsed",
        )
    combine_logic = "AND"if combine.startswith("AND") else "OR"

    n = st.session_state[count_key]
    col_opts = [""] + (available_columns or [])

    filters = []
    h0, h1, h2, h3 = st.columns([3, 3, 3, 1])
    h0.markdown("**Column**"); h1.markdown("**Operator**")
    h2.markdown("**Value**");  h3.markdown("")

    for i in range(n):
        c0, c1, c2, c3 = st.columns([3, 3, 3, 1])
        col_key = f"{key_prefix}_cond_{i}_col"
        op_key  = f"{key_prefix}_cond_{i}_op"
        val_key = f"{key_prefix}_cond_{i}_val"

        if available_columns:
            prev_col = st.session_state.get(col_key, "")
            idx = col_opts.index(prev_col) if prev_col in col_opts else 0
            chosen_col = c0.selectbox(" ", col_opts, index=idx, key=col_key, label_visibility="collapsed")
        else:
            chosen_col = c0.text_input(" ", key=col_key, placeholder="Column name", label_visibility="collapsed")

        prev_op = st.session_state.get(op_key, op_keys[0])
        op_idx  = op_keys.index(prev_op) if prev_op in op_keys else 0
        chosen_op_label = c1.selectbox(" ", op_labels, index=op_idx, key=op_key, label_visibility="collapsed")
        chosen_op = op_keys[op_labels.index(chosen_op_label)]

        needs_val = chosen_op in NEEDS_VALUE
        if needs_val:
            chosen_val = c2.text_input(" ", key=val_key, placeholder="value", label_visibility="collapsed")
        else:
            c2.markdown("<span style='color:grey;font-size:0.85em'>no value needed</span>", unsafe_allow_html=True)
            chosen_val = ""

        if c3.button("", key=f"{key_prefix}_cond_{i}_del", use_container_width=True):
            st.session_state[count_key] = max(1, n - 1)
            for suffix in ("_col", "_op", "_val"):
                st.session_state.pop(f"{key_prefix}_cond_{i}{suffix}", None)
            st.rerun()

        if chosen_col:
            filters.append({
                "column": chosen_col,
                "operator": chosen_op,
                "value": chosen_val,
            })

    return filters, combine_logic


def sheet_selectbox(label: str, file_label: str, key: str) -> str:
    """
    Sheet selector that is ALWAYS a dropdown and ALWAYS driven by the selected file.

    Behavior:
    - If no file selected: shows a disabled dropdown and returns "".
    - If CSV selected: shows disabled dropdown (no sheets) and returns "".
    - If Excel selected: shows dropdown with [""] + real sheet names and returns selected sheet (or "").
    """
    if not file_label:
        st.selectbox(label, options=[""], index=0, key=f"{key}__no_file", disabled=True)
        return ""

    path = get_file_path(file_label)
    if not path:
        st.selectbox(label, options=[""], index=0, key=f"{key}__no_path", disabled=True)
        return ""

    lower = path.lower()
    if lower.endswith(".csv") or lower.endswith(".txt"):
        st.selectbox(label, options=["(CSV - no sheets)"], index=0, key=f"{key}__csv", disabled=True)
        return ""

    sheets = get_sheet_names_for_file(path) or []
    if not sheets:
        st.selectbox(label, options=["(No sheets detected)"], index=0, key=f"{key}__none", disabled=True)
        return ""

    # Keep it empty by default; user must choose.
    return st.selectbox(label, options=[""] + sheets, index=0, key=key)


def _is_excel_label(file_label: str) -> bool:
    path = get_file_path(file_label) if file_label else ""
    lower = (path or "").lower()
    return lower.endswith(".xlsx") or lower.endswith(".xls") or lower.endswith(".xlsm") or lower.endswith(".xlsb")

def _is_csv_label(file_label: str) -> bool:
    path = get_file_path(file_label) if file_label else ""
    lower = (path or "").lower()
    return lower.endswith(".csv") or lower.endswith(".txt")


def get_headers_for_file_sheet(file_label: str, sheet_name: str, header_row: int) -> list[str]:
    """
    Returns header names (strings) from the given file/sheet/header_row.
    - For Excel: reads the given `sheet_name` at `header_row`.
    - For CSV: reads the `header_row` line and returns non-empty values.
    """
    if not file_label:
        return []
    path = get_file_path(file_label)
    if not path:
        return []
    try:
        headers: list[str] = []

        if _is_excel_label(file_label):
            if not sheet_name:
                return []
            if path.lower().endswith(".xlsb"):
                # pyxlsb for binary Excel
                import pyxlsb
                hrow = int(header_row)
                with pyxlsb.open_workbook(path) as wb:
                    with wb.get_sheet(sheet_name) as ws:
                        for row_idx, row in enumerate(ws.rows(), start=1):
                            if row_idx == hrow:
                                for cell in row:
                                    v = cell.v
                                    if v is None:
                                        continue
                                    s = str(v).strip()
                                    if s:
                                        headers.append(s)
                                break
            else:
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
                for c in range(1, ws.max_column + 1):
                    v = ws.cell(row=int(header_row), column=c).value
                    if v is None:
                        continue
                    s = str(v).strip()
                    if not s:
                        continue
                    headers.append(s)
                wb.close()

        elif _is_csv_label(file_label):
            # Read up to the header_row and use that row as headers
            n = max(1, int(header_row))
            dfh = pd.read_csv(path, header=None, nrows=n, dtype=str, keep_default_na=False)
            if dfh.shape[0] < n:
                return []
            row = dfh.iloc[n - 1].tolist()
            for v in row:
                s = str(v).strip() if v is not None else ""
                if not s:
                    continue
                headers.append(s)
        else:
            return []

        # de-dupe preserving order
        seen = set()
        out = []
        for h in headers:
            if h in seen:
                continue
            seen.add(h)
            out.append(h)
        return out
    except Exception:
        return []


def get_column_letter_for_header(file_label: str, sheet_name: str, header_row: int, header_name: str) -> str:
    """
    Given a header name (from a dropdown), return the Excel column letter (e.g. 'K').
    Falls back to the header_name string if not found.
    """
    if not header_name:
        return ""
    path = get_file_path(file_label) if file_label else ""
    if not path:
        return header_name
    try:
        if path.lower().endswith(".xlsb"):
            import pyxlsb
            from openpyxl.utils import get_column_letter as _gcl
            hrow = int(header_row)
            with pyxlsb.open_workbook(path) as wb:
                with wb.get_sheet(sheet_name) as ws:
                    for row_idx, row in enumerate(ws.rows(), start=1):
                        if row_idx == hrow:
                            for c_idx, cell in enumerate(row, start=1):
                                v = cell.v
                                if v is not None and str(v).strip() == str(header_name).strip():
                                    return _gcl(c_idx)
                            break
        else:
            import openpyxl as _ox
            from openpyxl.utils import get_column_letter as _gcl
            _wb = _ox.load_workbook(path, read_only=True, data_only=True)
            _ws = _wb[sheet_name] if sheet_name and sheet_name in _wb.sheetnames else _wb.active
            hrow = int(header_row)
            for c in range(1, _ws.max_column + 1):
                v = _ws.cell(row=hrow, column=c).value
                if v is not None and str(v).strip() == str(header_name).strip():
                    _wb.close()
                    return _gcl(c)
            _wb.close()
    except Exception:
        pass
    return header_name  # fallback


def column_selectbox(label: str, file_label: str, sheet_name: str, header_row: int, key: str) -> str:
    """
    Dropdown-only column selector. Returns ""until a real header is selected.
    """
    if not file_label or not sheet_name:
        st.selectbox(label, options=[""], index=0, key=f"{key}__no_fs", disabled=True)
        return ""
    headers = get_headers_for_file_sheet(file_label, sheet_name, header_row)
    if not headers:
        st.selectbox(label, options=["(No headers detected)"], index=0, key=f"{key}__no_hdrs", disabled=True)
        return ""
    return st.selectbox(label, options=[""] + headers, index=0, key=key)


def export_options_ui(key_prefix: str) -> dict:
    """
    Per-step export option UI.
    Returns: {"enabled": bool, "path": str}
    """
    with st.expander("Export as File (optional)", expanded=False):
        enabled = st.checkbox("Export output after this step", value=False, key=f"{key_prefix}_export_enabled")
        default_val = st.session_state.get("default_export_path", "") or ""
        path = st.text_input(
            "Export folder (or file path)",
            value=default_val,
            key=f"{key_prefix}_export_path",
            disabled=not enabled,
            help="Folder example: `/Users/you/Exports/`  |  File example: `/Users/you/Exports/output.xlsx`",
        )
        if enabled and not path.strip():
            st.warning("Export is enabled but export path is empty (set here or in sidebar Export settings).")
        return {"enabled": bool(enabled), "path": str(path or "")}

def end_row_input(label: str, key_prefix: str, default_last: bool = True, default_number: int = 10) -> str:
    """
    End row input with explicit 'last'option.
    Returns either "last"or a stringified integer.
    """
    mode_default = "last"if default_last else "custom"
    mode = st.selectbox(f"{label}", options=["last", "custom"], key=f"{key_prefix}_mode", index=0 if mode_default == "last"else 1)
    if mode == "last":
        return "last"
    n = st.number_input("End row (number)", min_value=1, value=int(default_number), key=f"{key_prefix}_num")
    return str(int(n))

def end_index_input(label: str, key_prefix: str, default_last: bool = False, default_number: int = 1) -> str:
    """
    End index input for insert/delete with explicit 'last'option.
    Returns either "last"or a stringified integer.
    """
    mode_default = "last"if default_last else "custom"
    mode = st.selectbox(label, options=["custom", "last"], key=f"{key_prefix}_mode", index=1 if mode_default == "last"else 0)
    if mode == "last":
        return "last"
    n = st.number_input("End index (number)", min_value=1, value=int(default_number), key=f"{key_prefix}_num")
    return str(int(n))


def get_file_path(filename):
    """Get path for uploaded file"""
    return st.session_state.uploaded_file_paths.get(filename, "")


def add_step(step_type, name, config):
    """Add a step to the workflow"""
    # Basic validation for steps that require file selection
    if step_type not in ("add_files", "input_source", "normalize_import") and isinstance(config, dict):
        if "file"in config and not config.get("file"):
            st.warning("Please select a file (add/register files first).")
            return
        for k in ("src_file", "tgt_file", "lookup_file"):
            if k in config and not config.get(k):
                st.warning(f"Please select {k.replace('_', ' ')}.")
                return
        # Enforce sheet dropdown selection for Excel files (no default Sheet1)
        # Skip when the operation explicitly targets multiple sheets (no single sheet needed)
        _skip_sheet_check = (
            # Advanced Delete running on selected/all sheets
            (step_type == "delete" and config.get("row_mode") == "advanced_condition"
             and config.get("sheet_scope", "single") != "single")
            # Write Cell writing across all sheets
            or (step_type == "write_cell" and config.get("scope", "single") == "all_sheets")
            # Conditional Write across all sheets
            or (step_type == "conditional" and config.get("scope", "single") == "all_sheets")
        )
        if not _skip_sheet_check and "file"in config and "sheet"in config and _is_excel_label(config.get("file")) and not config.get("sheet"):
            st.warning("Please select a sheet (dropdown) for the chosen Excel file.")
            return
        # SUMIFS source sheet must also be chosen when source file is excel
        if step_type == "sumifs":
            if _is_excel_label(config.get("src_workbook")) and not config.get("src_sheet"):
                st.warning("Please select a source sheet (dropdown) for the SUMIFS source file.")
                return
        # Common multi-file sheet validations
        for file_key, sheet_key, msg in (
            ("lookup_file", "lookup_sheet", "Please select a lookup sheet (dropdown) for the lookup file."),
            ("src_file", "src_sheet", "Please select a source sheet (dropdown) for the source file."),
            ("tgt_file", "tgt_sheet", "Please select a target sheet (dropdown) for the target file."),
        ):
            # folder_summary only needs src_sheet in "folder" mode — skip check for sheets/both modes
            if step_type == "folder_summary" and sheet_key == "src_sheet":
                if config.get("source_mode", "folder") != "folder":
                    continue
            # split_export creates output sheets dynamically — no tgt_sheet required
            # (except single mode which writes to one specific sheet)
            if step_type == "split_export" and sheet_key == "tgt_sheet":
                if config.get("split_mode", "single") != "single":
                    continue
            # sheet_updater / sheet_row_inserter write to dynamically matched sheets — no tgt_sheet dropdown
            if step_type in ("sheet_updater", "sheet_row_inserter") and sheet_key == "tgt_sheet":
                continue
            # formula_broadcast targets sheets dynamically — no tgt_sheet dropdown
            if step_type == "formula_broadcast":
                continue
            # unpivot tgt_sheet is a free-text field, not a dropdown
            if step_type == "unpivot" and sheet_key == "tgt_sheet":
                continue
            # pivot writes to a user-named sheet — sheet name is stored in tgt_sheet as a text field,
            # not a dropdown, so the Excel-label guard is irrelevant; skip to avoid false positives
            if step_type == "pivot" and sheet_key == "tgt_sheet":
                continue
            if file_key in config and sheet_key in config and _is_excel_label(config.get(file_key)) and not config.get(sheet_key):
                st.warning(msg)
                return

        # Explicit VLOOKUP validations (File/Sheet/Column for all 4 parts)
        if step_type == "vlookup"and config.get("mode") == "explicit":
            required = [
                ("lookup_value_file", "Lookup Value file"),
                ("lookup_value_sheet", "Lookup Value sheet"),
                ("lookup_value_column", "Lookup Value column"),
                ("search_file", "Search Value file"),
                ("search_sheet", "Search Value sheet"),
                ("search_column", "Search Value column"),
                ("return_file", "Return Value file"),
                ("return_sheet", "Return Value sheet"),
                ("return_column", "Return Value column"),
                ("return_to_file", "Return To file"),
                ("return_to_sheet", "Return To sheet"),
                ("return_to_column", "Return To column"),
            ]
            for k, label in required:
                if not config.get(k):
                    st.warning(f"Please select {label}.")
                    return

        if step_type == "sort":
            # Require at least one sort column
            if isinstance(config.get("sort_cols"), list) and not config.get("sort_cols"):
                st.warning("Please add at least one sort column.")
                return

    # Check if we're in edit mode (updating an existing step)
    editing_idx = st.session_state.get("editing_step_idx")
    editing_data = st.session_state.get("editing_step_data")
    
    if editing_idx is not None and editing_data is not None:
        # Update existing step in place
        existing_step = st.session_state.workflow_steps[editing_idx]
        existing_step["type"] = step_type
        existing_step["name"] = name
        existing_step["config"] = config
        # Preserve id, depends_on, lane from existing step
        st.session_state.workflow_steps[editing_idx] = existing_step
        
        # Clear edit mode
        st.session_state.editing_step_idx = None
        st.session_state.editing_step_data = None
        st.session_state.pop("adding_step", None)
        
        st.session_state["_last_added_step_name"] = f"(Updated) {name}"
        return
    
    # Normal mode: create new step
    step = {
        "id": uuid.uuid4().hex[:12],
        "type": step_type,
        "name": name,
        "config": config,
        "depends_on": [],
        "lane": "",
    }
    insert_at = st.session_state.get("insert_at", None)
    if isinstance(insert_at, int):
        idx = max(0, min(int(insert_at), len(st.session_state.workflow_steps)))
        st.session_state.workflow_steps.insert(idx, step)
        # Clear insertion target after use
        st.session_state.insert_at = None
        st.session_state._workflow_notice = ""
    else:
        st.session_state.workflow_steps.append(step)
    # Save a message that survives the next rerun
    st.session_state["_last_added_step_name"] = name
    # Clear current step editor, but DON'T force rerun (button click already triggers a rerun)
    try:
        st.session_state.pop("adding_step", None)
    except Exception:
        pass


def register_file_bytes(display_name: str, data: bytes, overwrite: bool = True) -> str:
    """
    Disabled: user requested no temp/working copies.
    """
    raise RuntimeError("Byte uploads are disabled. Register local file paths instead (to overwrite originals).")


def register_file_path(path: str, display_name: str | None = None, overwrite: bool = True) -> str:
    """
    Register an existing local file path directly (no copy).
    This enables true in-place overwrite on the original file.
    Returns the resolved absolute path.
    
    Always overwrites existing registrations (no timestamp suffixes).
    """
    p = Path(path).expanduser()
    if not p.is_absolute():
        p = p.resolve()
    if not p.exists():
        raise FileNotFoundError(f"File does not exist: {p}")

    name = os.path.basename(str(p)) if not display_name else display_name
    # Always overwrite - no timestamp suffixes
    st.session_state.uploaded_file_paths[name] = str(p)
    return str(p)


# =============================================================================
# MAIN
# =============================================================================

if __name__ == "__main__":
    main()
