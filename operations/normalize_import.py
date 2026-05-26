"""
Normalize Import operation.

Reads a registered file class, applies a per-class header mapping
(source headers → standard target headers, with alternates + required validation),
then appends/replaces data into a target working file.
"""

import io
import os
import json
from pathlib import Path

import pandas as pd
import openpyxl
from core import wb_cache

# ---------------------------------------------------------------------------
# Mapping template persistence
# ---------------------------------------------------------------------------

_MAPPINGS_FILE = "mappings.json"


def load_all_templates() -> dict:
    """Return all saved mapping templates from mappings.json."""
    try:
        if os.path.exists(_MAPPINGS_FILE):
            with open(_MAPPINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
    except Exception:
        pass
    return {}


def save_template(name: str, source_class: str, mapping: list[dict]) -> tuple[bool, str]:
    """Save a mapping template to mappings.json."""
    try:
        existing = load_all_templates()
        existing[name] = {
            "source_class": source_class,
            "mapping": mapping,
        }
        with open(_MAPPINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(existing, f, indent=2)
        return True, None
    except Exception as e:
        return False, str(e)


def delete_template(name: str) -> bool:
    try:
        existing = load_all_templates()
        if name in existing:
            del existing[name]
            with open(_MAPPINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(existing, f, indent=2)
            return True
    except Exception:
        pass
    return False


# ---------------------------------------------------------------------------
# Header detection
# ---------------------------------------------------------------------------

def detect_headers(file_path: str, sheet: str = None, header_row: int = 1) -> list[str]:
    """Read the header row from a file and return column names."""
    if not file_path or not os.path.exists(file_path):
        return []
    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext in (".xlsx", ".xls", ".xlsm"):
            df = pd.read_excel(file_path, sheet_name=sheet, header=header_row - 1, nrows=0)
        else:
            df = pd.read_csv(file_path, header=header_row - 1, nrows=0)
        return [str(c).strip() for c in df.columns if str(c).strip()]
    except Exception:
        return []


# ---------------------------------------------------------------------------
# Mapping file parsing
# ---------------------------------------------------------------------------

def load_mapping_from_file(path: str) -> tuple[list[dict], str]:
    """
    Parse an uploaded mapping file (Excel or CSV).
    Expected columns (case-insensitive):
      Source Header, Target Header, Required, Data Type,
      Alternate Source Headers, Default Value
    Returns (list_of_mapping_dicts, error_string).
    """
    if not path or not os.path.exists(path):
        return [], "Mapping file not found"
    try:
        ext = os.path.splitext(path)[1].lower()
        if ext in (".xlsx", ".xls", ".xlsm"):
            df = pd.read_excel(path, nrows=500)
        else:
            df = pd.read_csv(path, nrows=500)

        df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]

        col_alias = {
            "source_header": ["source_header", "source", "src_header", "src", "from", "input"],
            "target_header": ["target_header", "target", "tgt_header", "tgt", "to", "output", "standard"],
            "required":       ["required", "mandatory"],
            "data_type":      ["data_type", "type", "dtype"],
            "alternate_source_headers": ["alternate_source_headers", "alternates", "alternate_headers",
                                         "aliases", "alt", "synonyms"],
            "default_value":  ["default_value", "default"],
        }

        def find_col(field):
            for alias in col_alias[field]:
                if alias in df.columns:
                    return alias
            return None

        src_col = find_col("source_header")
        tgt_col = find_col("target_header")
        if not src_col or not tgt_col:
            return [], f"Mapping file must have 'Source Header' and 'Target Header' columns. Found: {list(df.columns)}"

        req_col  = find_col("required")
        dt_col   = find_col("data_type")
        alt_col  = find_col("alternate_source_headers")
        def_col  = find_col("default_value")

        result = []
        for _, row in df.iterrows():
            src = str(row.get(src_col, "") or "").strip()
            tgt = str(row.get(tgt_col, "") or "").strip()
            if not src or not tgt:
                continue
            required = False
            if req_col:
                r = str(row.get(req_col, "") or "").strip().lower()
                required = r in ("yes", "true", "1", "y")
            dtype = ""
            if dt_col:
                dtype = str(row.get(dt_col, "") or "").strip().lower()
            alts = []
            if alt_col:
                raw = str(row.get(alt_col, "") or "").strip()
                alts = [a.strip() for a in raw.split(",") if a.strip()]
            default = ""
            if def_col:
                default = str(row.get(def_col, "") or "").strip()
            result.append({
                "source_header": src,
                "target_header": tgt,
                "required": required,
                "data_type": dtype,
                "alternate_headers": alts,
                "default_value": default,
            })

        return result, ""
    except Exception as e:
        return [], str(e)


def load_mapping_from_bytes(file_bytes: bytes, filename: str) -> tuple[list[dict], str]:
    """Same as load_mapping_from_file but from uploaded bytes."""
    tmp = Path("temp") / f"_mapping_{filename}"
    tmp.parent.mkdir(exist_ok=True)
    try:
        with open(str(tmp), "wb") as f:
            f.write(file_bytes)
        return load_mapping_from_file(str(tmp))
    finally:
        try:
            tmp.unlink()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Mapping application
# ---------------------------------------------------------------------------

def _find_source_col(detected: list[str], mapping_row: dict) -> str | None:
    """
    Find which detected column matches this mapping row.
    Checks: exact source_header, then each alternate_header (case-insensitive).
    Returns the actual column name found in detected, or None.
    """
    targets = [mapping_row["source_header"]] + (mapping_row.get("alternate_headers") or [])
    detected_lower = {c.lower(): c for c in detected}
    for candidate in targets:
        c = candidate.strip().lower()
        if c in detected_lower:
            return detected_lower[c]
    return None


def apply_header_mapping(
    df: pd.DataFrame,
    mapping: list[dict],
) -> tuple[pd.DataFrame, list[dict]]:
    """
    Apply column mapping to a DataFrame.
    Returns (mapped_df, report_rows).
    report_rows: list of dicts with keys: source_header, detected_as, target_header, status, note
    """
    detected = [str(c).strip() for c in df.columns]
    report = []
    rename_map = {}
    cols_to_keep = []

    for m in mapping:
        found = _find_source_col(detected, m)
        tgt = m["target_header"]
        if found:
            rename_map[found] = tgt
            cols_to_keep.append(tgt)
            status = "mapped"
            note = f"'{found}' → '{tgt}'" + (f" (via alternate)" if found.lower() != m["source_header"].lower() else "")
        elif m.get("required"):
            status = "missing_required"
            note = f"Required column '{m['source_header']}' not found"
        else:
            status = "missing_optional"
            note = f"Optional column '{m['source_header']}' not found"

        report.append({
            "Source Header": m["source_header"],
            "Detected As": found or "—",
            "Target Header": tgt,
            "Required": "Yes" if m.get("required") else "No",
            "Data Type": m.get("data_type", ""),
            "Status": "✅" if status == "mapped" else ("❌" if status == "missing_required" else "⚠️"),
            "Note": note,
        })

    # Rename columns in df
    df_renamed = df.rename(columns=rename_map)

    # Only keep mapped columns that exist
    keep = [c for c in cols_to_keep if c in df_renamed.columns]

    # Add defaults for missing optional columns
    for m in mapping:
        tgt = m["target_header"]
        if tgt not in df_renamed.columns and m.get("default_value"):
            df_renamed[tgt] = m["default_value"]
            if tgt not in keep:
                keep.append(tgt)

    df_out = df_renamed[keep] if keep else df_renamed

    return df_out, report


# ---------------------------------------------------------------------------
# Step runner (called by executor)
# ---------------------------------------------------------------------------

def _read_source_df(src_path: str, src_sheet: str, src_header_row: int) -> pd.DataFrame:
    ext = os.path.splitext(src_path)[1].lower()
    if ext in (".xlsx", ".xls", ".xlsm"):
        df = wb_cache.read_excel(src_path, sheet_name=src_sheet, header=src_header_row - 1)
    else:
        df = pd.read_csv(src_path, header=src_header_row - 1)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def run_normalize_import_step(
    src_path: str,
    src_sheet: str,
    src_header_row: int,
    mapping: list[dict],
    tgt_path: str,
    tgt_sheet: str,
    tgt_header_row: int,
    append_mode: bool = True,
    src_paths: list[str] | None = None,
) -> tuple[bool, str]:
    """
    Read one or more source files, apply header mapping to each,
    then write all rows to the target file.

    src_paths overrides src_path when provided (used for folder-based file classes
    where multiple files share the same format and mapping).
    """
    # Resolve the list of source files to process
    all_src = src_paths if src_paths else ([src_path] if src_path else [])
    all_src = [p for p in all_src if p and os.path.exists(p)]

    if not all_src:
        return False, f"No source files found (checked: {src_path})"
    if not tgt_path or not os.path.exists(tgt_path):
        return False, f"Target file not found: {tgt_path}"
    if not mapping:
        return False, "No mapping defined"

    try:
        # Map every source file and collect results
        frames = []
        total_src_rows = 0
        mapped_count = 0
        errors = []

        for i, sp in enumerate(all_src):
            try:
                df_src = _read_source_df(sp, src_sheet, src_header_row)
                df_mapped, report = apply_header_mapping(df_src, mapping)

                missing_req = [r["Source Header"] for r in report if r["Status"] == "❌"]
                if missing_req:
                    errors.append(f"{os.path.basename(sp)}: missing required columns {missing_req}")
                    continue

                if i == 0:
                    mapped_count = sum(1 for r in report if r["Status"] == "✅")

                total_src_rows += len(df_src)
                frames.append(df_mapped)
            except Exception as e:
                errors.append(f"{os.path.basename(sp)}: {e}")

        if not frames:
            err_detail = "; ".join(errors)
            return False, f"No files could be processed. Errors: {err_detail}"

        df_all_new = pd.concat(frames, ignore_index=True)

        # Read target
        tgt_ext = os.path.splitext(tgt_path)[1].lower()
        if tgt_ext in (".xlsx", ".xls", ".xlsm"):
            df_tgt = wb_cache.read_excel(tgt_path, sheet_name=tgt_sheet, header=tgt_header_row - 1)
            df_tgt.columns = [str(c).strip() for c in df_tgt.columns]
        else:
            df_tgt = pd.read_csv(tgt_path, header=tgt_header_row - 1)
            df_tgt.columns = [str(c).strip() for c in df_tgt.columns]

        if append_mode:
            df_result = pd.concat([df_tgt, df_all_new], ignore_index=True)
            mode_msg = (
                f"appended {len(df_all_new)} rows from {len(frames)}/{len(all_src)} file(s) "
                f"(target now has {len(df_result)} rows)"
            )
        else:
            df_result = df_all_new
            mode_msg = f"replaced with {len(df_all_new)} rows from {len(frames)}/{len(all_src)} file(s)"

        # Write back to target
        if tgt_ext in (".xlsx", ".xls", ".xlsm"):
            with wb_cache.excel_writer(tgt_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df_result.to_excel(writer, sheet_name=tgt_sheet, index=False, startrow=tgt_header_row - 1)
        else:
            df_result.to_csv(tgt_path, index=False)

        msg = f"Normalize Import: {mode_msg}, {mapped_count}/{len(mapping)} columns mapped"
        if errors:
            msg += f" | Skipped: {'; '.join(errors)}"
        return True, msg

    except Exception as e:
        return False, f"Error in normalize import: {str(e)}"
