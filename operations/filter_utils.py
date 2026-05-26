"""
Shared filtering utility for Import and Delete-by-condition steps.
Applies SQL-style WHERE conditions to a pandas DataFrame.
"""

import re
import pandas as pd
from openpyxl.utils import column_index_from_string

# ---------------------------------------------------------------------------
# Operator catalogue
# ---------------------------------------------------------------------------

OPERATORS = {
    "eq":           "= equals",
    "ne":           "≠ not equals",
    "contains":     "contains",
    "not_contains": "does not contain",
    "starts_with":  "starts with",
    "ends_with":    "ends with",
    "gt":           "> greater than",
    "lt":           "< less than",
    "gte":          "≥ greater or equal",
    "lte":          "≤ less or equal",
    "is_empty":     "is empty / blank",
    "is_not_empty": "is not empty / blank",
    "in_list":      "in list  (comma-sep values)",
    "not_in_list":  "not in list  (comma-sep values)",
}

# Operators that need a value input
NEEDS_VALUE = {k for k in OPERATORS if k not in ("is_empty", "is_not_empty")}

# Operators that work on numbers; will try to coerce column to numeric
NUMERIC_OPS = {"gt", "lt", "gte", "lte"}


def _is_numeric_op(op: str) -> bool:
    return op in NUMERIC_OPS


def _resolve_col(df: pd.DataFrame, col_spec: str) -> str:
    """
    Resolve a column spec to the actual column name in df.

    Accepts, in priority order:
      1. Exact header match             → "LedgerNo" matches column named "LedgerNo"
      2. Case-insensitive header match  → "ledgerno" matches "LedgerNo"
      3. Excel column letter            → "B" → second column (whatever its name is)
      4. 1-based integer string         → "2" → second column

    Returns the resolved column name, or the original spec if nothing matches
    (caller then does `col not in df.columns` as before).
    """
    if not col_spec:
        return col_spec

    # 1. Exact match
    if col_spec in df.columns:
        return col_spec

    # 2. Case-insensitive match
    lower_map = {str(c).lower(): c for c in df.columns}
    if col_spec.lower() in lower_map:
        return lower_map[col_spec.lower()]

    # 3. Excel column letter (A, B, AA …)
    stripped = col_spec.strip()
    if re.fullmatch(r"[A-Za-z]{1,3}", stripped):
        try:
            idx = column_index_from_string(stripped.upper()) - 1  # 0-based
            if 0 <= idx < len(df.columns):
                return df.columns[idx]
        except Exception:
            pass

    # 4. 1-based integer
    if stripped.lstrip("-").isdigit():
        try:
            idx = int(stripped) - 1  # 0-based
            if 0 <= idx < len(df.columns):
                return df.columns[idx]
        except Exception:
            pass

    return col_spec  # unresolved — caller handles


def _make_mask(df: pd.DataFrame, condition: dict) -> pd.Series:
    """Return a boolean Series for one filter condition."""
    raw_col = condition.get("column", "")
    col     = _resolve_col(df, str(raw_col).strip())
    op      = condition.get("operator", "eq")
    raw     = str(condition.get("value", "") or "")

    if col not in df.columns:
        # Unknown column → no rows match (return False series so AND logic survives)
        return pd.Series([False] * len(df), index=df.index)

    series = df[col]

    # ── Value-less operators ─────────────────────────────────────────────────
    if op == "is_empty":
        return series.isna() | (series.astype(str).str.strip() == "")
    if op == "is_not_empty":
        return ~(series.isna() | (series.astype(str).str.strip() == ""))

    # ── List operators ───────────────────────────────────────────────────────
    if op in ("in_list", "not_in_list"):
        items = [v.strip() for v in raw.split(",") if v.strip()]
        mask = series.astype(str).isin(items)
        return mask if op == "in_list" else ~mask

    # ── Numeric operators ────────────────────────────────────────────────────
    if op in NUMERIC_OPS:
        try:
            num_val = float(raw)
            num_series = pd.to_numeric(series, errors="coerce")
            if op == "gt":  return num_series > num_val
            if op == "lt":  return num_series < num_val
            if op == "gte": return num_series >= num_val
            if op == "lte": return num_series <= num_val
        except ValueError:
            return pd.Series([False] * len(df), index=df.index)

    # ── String operators ─────────────────────────────────────────────────────
    str_series = series.astype(str).fillna("")
    val_lower  = raw.lower()
    sl         = str_series.str.lower()

    if op == "eq":
        base = str_series == raw
        # Also match when value is stored as a number (e.g. 110800 vs "110800")
        try:
            num_val = float(raw)
            num_series = pd.to_numeric(df[col], errors="coerce")
            base = base | (num_series == num_val)
        except (ValueError, TypeError):
            pass
        return base
    if op == "ne":
        base = str_series != raw
        try:
            num_val = float(raw)
            num_series = pd.to_numeric(df[col], errors="coerce")
            base = base & (num_series != num_val)
        except (ValueError, TypeError):
            pass
        return base
    if op == "contains":     return sl.str.contains(val_lower, na=False, regex=False)
    if op == "not_contains": return ~sl.str.contains(val_lower, na=False, regex=False)
    if op == "starts_with":  return sl.str.startswith(val_lower, na=False)
    if op == "ends_with":    return sl.str.endswith(val_lower, na=False)

    # Fallback: equals
    return str_series == raw


def apply_filters(
    df: pd.DataFrame,
    filters: list[dict],
    combine: str = "AND",
) -> pd.DataFrame:
    """
    Filter a DataFrame by a list of conditions.

    Each condition dict:
        column   : str   – column name
        operator : str   – key from OPERATORS
        value    : str   – comparison value (ignored for is_empty / is_not_empty)

    combine : "AND" → all conditions must match
              "OR"  → any condition matches
    """
    if not filters:
        return df

    masks = [_make_mask(df, cond) for cond in filters]

    if combine.upper() == "OR":
        combined = masks[0]
        for m in masks[1:]:
            combined = combined | m
    else:  # AND (default)
        combined = masks[0]
        for m in masks[1:]:
            combined = combined & m

    return df[combined].reset_index(drop=True)
