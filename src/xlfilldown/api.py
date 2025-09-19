# Public API surface for xlfilldown
from .core import (
    ingest_excel_to_sqlite,
    ingest_excel_to_excel,
    qident,
    normalize_headers,
    canon_list,
    sha256_hex,
)

__all__ = [
    "ingest_excel_to_sqlite",
    "ingest_excel_to_excel",
    "qident",
    "normalize_headers",
    "canon_list",
    "sha256_hex",
]

