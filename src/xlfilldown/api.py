# xlfilldown/api.py
from __future__ import annotations

from typing import Iterable, Optional, Sequence, Union, Any

# Re-export public helpers from core
from .core import (
    qident,
    normalize_headers,
    canon_list,
    sha256_hex,
)

# Import the existing core ingest functions (which still take pad_* params)
from .core import ingest_excel_to_sqlite as _core_ingest_excel_to_sqlite
from .core import ingest_excel_to_excel as _core_ingest_excel_to_excel


__all__ = [
    "ingest_excel_to_sqlite",
    "ingest_excel_to_excel",
    "qident",
    "normalize_headers",
    "canon_list",
    "sha256_hex",
]


def _resolve_fill_mode(
    *,
    fill_mode: Optional[str] = None,
    pad_hierarchical: Optional[bool] = None,
) -> bool:
    """
    Returns the boolean value expected by core.pad_hierarchical.

    Priority:
      1) explicit pad_hierarchical (for backward-compat)
      2) fill_mode string ("hierarchical" | "independent") — default is "hierarchical"
    """
    if pad_hierarchical is not None:
        return bool(pad_hierarchical)

    if fill_mode is None:
        return True  # default to hierarchical

    fm = str(fill_mode).strip().lower()
    if fm in ("hierarchical", "hier", "h"):
        return True
    if fm in ("independent", "ind", "i", "legacy"):
        return False
    raise ValueError("fill_mode must be 'hierarchical' or 'independent'.")


def _resolve_fill_cols(
    *,
    fill_cols: Optional[Sequence[str]] = None,
    pad_cols: Optional[Sequence[str]] = None,
) -> Sequence[str]:
    """
    Returns the list of column names to fill (aliasing pad_cols -> fill_cols).
    """
    if fill_cols is not None and pad_cols is not None:
        # If both provided, prefer fill_cols (new API).
        return list(fill_cols)
    if fill_cols is not None:
        return list(fill_cols)
    if pad_cols is not None:
        return list(pad_cols)
    # Let core handle empty/missing errors in its usual way.
    return []


def ingest_excel_to_sqlite(
    *,
    file: Union[str, "PathLike[str]"],
    sheet: str,
    header_row: int,
    # NEW API
    fill_cols: Optional[Sequence[str]] = None,
    fill_mode: Optional[str] = None,
    # OLD API (aliases)
    pad_cols: Optional[Sequence[str]] = None,
    pad_hierarchical: Optional[bool] = None,
    # passthrough options
    db: Union[str, "PathLike[str]"],
    table: Optional[str] = None,
    drop_blank_rows: bool = False,
    require_non_null: Optional[Sequence[str]] = None,
    row_hash: bool = False,
    excel_row_numbers: bool = False,
    if_exists: str = "fail",
    batch_size: int = 1000,
    **kwargs: Any,
) -> dict:
    """
    API wrapper that mirrors the CLI terms (fill_cols, fill_mode) while remaining
    backward-compatible with pad_cols / pad_hierarchical.
    """
    cols = _resolve_fill_cols(fill_cols=fill_cols, pad_cols=pad_cols)
    hierarchical = _resolve_fill_mode(fill_mode=fill_mode, pad_hierarchical=pad_hierarchical)

    # Forward to core using its expected parameter names
    return _core_ingest_excel_to_sqlite(
        file=file,
        sheet=sheet,
        header_row=header_row,
        pad_cols=cols,
        db=db,
        table=table,
        drop_blank_rows=drop_blank_rows,
        require_non_null=require_non_null or [],
        row_hash=row_hash,
        excel_row_numbers=excel_row_numbers,
        if_exists=if_exists,
        batch_size=batch_size,
        pad_hierarchical=hierarchical,
        **kwargs,
    )


def ingest_excel_to_excel(
    *,
    file: Union[str, "PathLike[str]"],
    sheet: str,
    header_row: int,
    # NEW API
    fill_cols: Optional[Sequence[str]] = None,
    fill_mode: Optional[str] = None,
    # OLD API (aliases)
    pad_cols: Optional[Sequence[str]] = None,
    pad_hierarchical: Optional[bool] = None,
    # passthrough options
    outfile: Union[str, "PathLike[str]"],
    outsheet: str,
    drop_blank_rows: bool = False,
    require_non_null: Optional[Sequence[str]] = None,
    row_hash: bool = False,
    excel_row_numbers: bool = False,
    if_exists: str = "fail",
    **kwargs: Any,
) -> dict:
    """
    API wrapper that mirrors the CLI terms (fill_cols, fill_mode) while remaining
    backward-compatible with pad_cols / pad_hierarchical.
    """
    cols = _resolve_fill_cols(fill_cols=fill_cols, pad_cols=pad_cols)
    hierarchical = _resolve_fill_mode(fill_mode=fill_mode, pad_hierarchical=pad_hierarchical)

    return _core_ingest_excel_to_excel(
        file=file,
        sheet=sheet,
        header_row=header_row,
        pad_cols=cols,
        outfile=outfile,
        outsheet=outsheet,
        drop_blank_rows=drop_blank_rows,
        require_non_null=require_non_null or [],
        row_hash=row_hash,
        excel_row_numbers=excel_row_numbers,
        if_exists=if_exists,
        pad_hierarchical=hierarchical,
        **kwargs,
    )


