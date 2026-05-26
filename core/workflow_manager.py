"""
Workflow management for saving and loading workflows.
Handles persistence of workflow configurations to/from JSON files.
"""

import json
import os
import streamlit as st
from datetime import datetime
from typing import Dict, List, Any


def load_workflows_from_file(path: str = "workflows.json") -> dict:
    """
    Load workflows from JSON file and return a dict.
    
    Args:
        path: Path to workflows JSON file
        
    Returns:
        Dictionary mapping workflow names to workflow data
    """
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
    except Exception:
        return {}
    return {}


def save_workflow_to_file(name: str, data: dict, path: str = "workflows.json") -> tuple:
    """
    Save the workflow data under `name` into a JSON file at `path`.
    Merges with existing file contents if present.
    
    Args:
        name: Workflow name
        data: Workflow data dict
        path: Path to JSON file
        
    Returns:
        Tuple of (success: bool, error_message: str or None)
    """
    try:
        existing = {}
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                try:
                    existing = json.load(f)
                except Exception:
                    existing = {}

        # Merge with any previously-saved workflow
        prev = existing.get(name, {}) or {}
        prev_steps = prev.get('steps', []) or []
        prev_steps_by_id = {s.get('id'): s for s in prev_steps if s.get('id')}

        new_steps = []
        for s in data.get('steps', []) or []:
            sid = s.get('id')
            if sid and sid in prev_steps_by_id:
                merged = dict(prev_steps_by_id[sid])
                merged.update({k: v for k, v in s.items() if v is not None and v != ''})
                new_steps.append(merged)
            else:
                new_steps.append(s)

        payload = {
            "steps": new_steps,
            "target_file_label": data.get("target_file_label", prev.get('target_file_label', "")),
            "target_sheet": data.get("target_sheet", prev.get('target_sheet', "")),
            "uploaded_file_paths": data.get('uploaded_file_paths', prev.get('uploaded_file_paths', [])) or [],
            "input_paths": [],
            "output_paths": [],
            "saved_at": datetime.utcnow().isoformat() + "Z"
        }
        
        existing[name] = payload
        
        with open(path, "w", encoding="utf-8") as f:
            json.dump(existing, f, indent=2)
            
        return True, None
        
    except Exception as e:
        return False, str(e)


def collect_current_workflow_state() -> dict:
    """
    Collect the current workflow steps and their UI-configured parameters from session_state.
    
    Returns:
        Dictionary suitable for saving (includes steps list with parameters)
    """
    steps_out = []
    
    for step in st.session_state.get('workflow_steps', []):
        sid = step.get('id')
        raw_type = step.get('type')
        stype = (raw_type or '').lower().strip() if raw_type else ''
        s = {'id': sid, 'type': stype}

        try:
            if stype == 'formula':
                s['formula'] = st.session_state.get(f"formula_{sid}", '')
                s['column'] = st.session_state.get(f"column_{sid}", '')
                s['start'] = st.session_state.get(f"start_{sid}", 2)
                s['end'] = st.session_state.get(f"end_{sid}", 'last')
                s['convert_to_values'] = st.session_state.get(f"formula_convert_to_values_{sid}", False)

            # Add other step types as they are migrated
            # elif stype == 'vlookup':
            #     s['vlookup_maincol'] = st.session_state.get(f"vlookup_maincol_{sid}", '')
            #     ...
            
            # Preserve any extra keys from previous loads
            for k, v in step.items():
                if k not in ('id', 'type') and k not in s:
                    s[k] = v

        except Exception:
            pass

        # Include per-step target file label if present
        try:
            file_label = st.session_state.get(f"step_file_{sid}", '')
            if file_label:
                s['target_file_label'] = file_label
        except Exception:
            pass

        steps_out.append(s)

    return {
        'steps': steps_out,
        'target_file_label': st.session_state.get('main_file_select', ''),
        'target_sheet': st.session_state.get('main_sheet_name', ''),
        'uploaded_file_paths': st.session_state.get('uploaded_file_paths', []) or [],
        'input_paths': [],
        'output_paths': []
    }


def apply_loaded_workflow_to_session(data: dict):
    """
    Given a workflow dict, populate session_state with workflow steps and configurations.
    
    Args:
        data: Workflow data dict with 'steps', 'target_file_label', 'target_sheet', etc.
    """
    steps = data.get('steps', [])
    
    # Normalize steps
    st.session_state['workflow_steps'] = [
        {'id': s.get('id'), 'type': (s.get('type') or '').lower().strip()} 
        for s in steps
    ]

    # Restore uploaded file paths
    uploaded_paths = data.get('uploaded_file_paths', [])
    if uploaded_paths:
        st.session_state['uploaded_file_paths'] = uploaded_paths

    # Restore target file and sheet
    if 'target_file_label' in data:
        st.session_state['main_file_select'] = data['target_file_label']
    if 'target_sheet' in data:
        st.session_state['main_sheet_name'] = data['target_sheet']

    # Populate per-step widget-backed keys
    for s in steps:
        sid = s.get('id')
        stype = s.get('type')
        if not sid or not stype:
            continue
            
        if stype == 'formula':
            st.session_state[f"formula_{sid}"] = s.get('formula', '')
            st.session_state[f"column_{sid}"] = s.get('column', '')
            st.session_state[f"start_{sid}"] = s.get('start', 2)
            st.session_state[f"end_{sid}"] = s.get('end', 'last')
            st.session_state[f"formula_convert_to_values_{sid}"] = bool(s.get('convert_to_values', False))

        # Add other step types as they are migrated
        # elif stype == 'vlookup':
        #     st.session_state[f"vlookup_maincol_{sid}"] = s.get('vlookup_maincol', '')
        #     ...

        # Restore per-step target file label if present
        try:
            st.session_state[f"step_file_{sid}"] = s.get('target_file_label', '')
        except Exception:
            pass

        # Restore per-step sheet selection if present
        try:
            st.session_state[f"step_sheet_{sid}"] = s.get('step_sheet', '')
        except Exception:
            pass


