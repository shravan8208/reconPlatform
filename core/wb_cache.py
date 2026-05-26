"""
In-memory workbook cache.

When active (during a workflow run), every load_workbook / wb.save / read_excel /
ExcelWriter call operates on BytesIO buffers instead of the real disk file.
The file is read from disk exactly once at first access and written back to disk
exactly once when flush_all() is called at the end of the run.

When inactive (direct calls, tests, one-off usage), every function falls through
to normal disk I/O so behaviour is identical to before.
"""

from __future__ import annotations

from contextlib import contextmanager
from io import BytesIO
from typing import Dict

import openpyxl
import pandas as pd

# ------------------------------------------------------------------
# Module-level state
# ------------------------------------------------------------------

_cache: Dict[str, bytes] = {}   # absolute path → raw xlsx bytes
_active: bool = False


# ------------------------------------------------------------------
# Lifecycle
# ------------------------------------------------------------------

def activate() -> None:
    """Call before a workflow run begins. Clears any leftover state."""
    global _active
    _cache.clear()
    _active = True


def deactivate() -> None:
    """Call after a workflow run ends (success or error). Flushes then clears."""
    global _active
    flush_all()
    _cache.clear()
    _active = False


def is_active() -> bool:
    return _active


# ------------------------------------------------------------------
# Internal helpers
# ------------------------------------------------------------------

def _read_disk(path: str) -> bytes:
    with open(path, "rb") as fh:
        return fh.read()


def _ensure(path: str) -> None:
    """Load file from disk into cache if not already there."""
    if path not in _cache:
        _cache[path] = _read_disk(path)


# ------------------------------------------------------------------
# openpyxl wrappers
# ------------------------------------------------------------------

def _xlsb_to_xlsx_bytes(path: str) -> bytes:
    """
    Convert a .xlsb file to xlsx bytes in memory using pyxlsb + pandas.
    Each sheet is read via pyxlsb and written into a fresh openpyxl workbook.
    Values only — formatting is not preserved (xlsb → openpyxl is one-way lossy).
    """
    import pyxlsb

    buf = BytesIO()
    new_wb = openpyxl.Workbook()
    new_wb.remove(new_wb.active)          # drop default empty "Sheet"

    with pyxlsb.open_workbook(path) as xlsb_wb:
        for sheet_name in xlsb_wb.sheets:
            ws = new_wb.create_sheet(sheet_name)
            with xlsb_wb.get_sheet(sheet_name) as xlsb_ws:
                for row in xlsb_ws.rows():
                    for cell in row:
                        if cell.v is not None:
                            # pyxlsb uses 0-based r/c
                            ws.cell(row=cell.r + 1, column=cell.c + 1, value=cell.v)

    new_wb.save(buf)
    new_wb.close()
    return buf.getvalue()


def load(path, data_only: bool = False, **kwargs):
    """
    Drop-in for openpyxl.load_workbook(path, data_only=...).

    .xlsb files are automatically converted to an in-memory xlsx workbook
    so all downstream openpyxl operations (read, format, write) work normally.
    The conversion is values-only (formatting from the original xlsb is dropped).
    """
    path = str(path)

    # .xlsb: openpyxl cannot read binary Excel — convert to xlsx first
    if path.lower().endswith(".xlsb"):
        xlsx_bytes = _xlsb_to_xlsx_bytes(path)
        if _active:
            # Cache the converted xlsx bytes under the original path key
            # so subsequent save() calls write back the converted version
            _cache[path] = xlsx_bytes
        return openpyxl.load_workbook(BytesIO(xlsx_bytes), data_only=data_only, **kwargs)

    if not _active:
        return openpyxl.load_workbook(path, data_only=data_only, **kwargs)
    _ensure(path)
    buf = BytesIO(_cache[path])
    return openpyxl.load_workbook(buf, data_only=data_only, **kwargs)


def save(wb, path) -> None:
    """Drop-in for wb.save(path). Writes to cache; does NOT touch disk."""
    path = str(path)
    if not _active:
        wb.save(path)
        return
    buf = BytesIO()
    wb.save(buf)
    _cache[path] = buf.getvalue()


# ------------------------------------------------------------------
# pandas wrappers
# ------------------------------------------------------------------

def read_excel(path, **kwargs) -> pd.DataFrame:
    """Drop-in for pd.read_excel(path, ...).
    Auto-selects pyxlsb engine for .xlsb files when reading from a BytesIO buffer
    (pandas cannot infer the format from a buffer without an explicit engine hint).
    """
    path = str(path)
    if not _active:
        # Disk path: pandas detects .xlsb automatically
        if path.lower().endswith(".xlsb") and "engine" not in kwargs:
            kwargs["engine"] = "pyxlsb"
        return pd.read_excel(path, **kwargs)
    _ensure(path)
    buf = BytesIO(_cache[path])
    # BytesIO has no extension — must tell pandas which engine to use for .xlsb
    if path.lower().endswith(".xlsb") and "engine" not in kwargs:
        kwargs["engine"] = "pyxlsb"
    return pd.read_excel(buf, **kwargs)


@contextmanager
def excel_writer(path, **kwargs):
    """
    Drop-in context manager for pd.ExcelWriter(path, ...).

    Usage (identical to the original):
        with wb_cache.excel_writer(target_path, engine='openpyxl', mode='a', ...) as writer:
            df.to_excel(writer, ...)
    """
    path = str(path)
    if not _active:
        with pd.ExcelWriter(path, **kwargs) as writer:
            yield writer
        return

    _ensure(path)
    buf = BytesIO(_cache[path])
    with pd.ExcelWriter(buf, **kwargs) as writer:
        yield writer
    _cache[path] = buf.getvalue()


# ------------------------------------------------------------------
# Flush
# ------------------------------------------------------------------

def flush(path: str) -> None:
    """Write a single cached file back to disk."""
    path = str(path)
    if path in _cache:
        with open(path, "wb") as fh:
            fh.write(_cache[path])


def flush_all() -> None:
    """Write every cached file back to disk."""
    for path in list(_cache.keys()):
        flush(path)
