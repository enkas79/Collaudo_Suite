from __future__ import annotations

import re
from pathlib import Path


_COLUMN_RE = re.compile(r"^[A-Z]+$")
_CELL_RE = re.compile(r"^([A-Z]+)([1-9][0-9]*)$")
_ROW_RE = re.compile(r"^[1-9][0-9]*$")


def col_to_index(col_str: str) -> int:
    """Convert an Excel column label, e.g. A or AA, to a zero based index."""
    value = (col_str or "").strip().upper()
    if not _COLUMN_RE.fullmatch(value):
        raise ValueError(f"Colonna Excel non valida: {col_str!r}. Usa solo lettere, es. A, B, AA.")

    index = 0
    for char in value:
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def index_to_col(index_0_based: int) -> str:
    """Convert a zero based index to an Excel column label."""
    if index_0_based < 0:
        raise ValueError("Indice colonna negativo.")

    index = index_0_based
    result = ""
    while True:
        index, remainder = divmod(index, 26)
        result = chr(65 + remainder) + result
        if index == 0:
            break
        index -= 1
    return result


def cell_to_indices(cell_str: str) -> tuple[int, int]:
    """Convert an Excel start reference to indexes.

    Accepts the historical cell format (B14) and also a plain row number
    (14). If only the row is provided, the returned column index is -1.
    """
    value = (cell_str or "").strip().upper()
    if _ROW_RE.fullmatch(value):
        return int(value) - 1, -1

    match = _CELL_RE.fullmatch(value)
    if not match:
        raise ValueError(f"Cella/riga di partenza non valida: {cell_str!r}. Usa un formato tipo B14 oppure una riga tipo 14.")

    col_s, row_s = match.groups()
    row_idx = int(row_s) - 1
    col_idx = col_to_index(col_s)
    return row_idx, col_idx


def force_extension(path: str, extension: str) -> str:
    """Return path with the requested extension. Extension must include the dot."""
    if not extension.startswith("."):
        raise ValueError("L'estensione deve iniziare con un punto.")
    p = Path(path)
    if p.suffix.lower() != extension.lower():
        p = p.with_suffix(extension)
    return str(p)
