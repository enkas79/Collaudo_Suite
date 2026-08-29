from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(slots=True)
class AnalysisParams:
    files: list[str]
    keyword: str = ""
    threshold: int = 70
    min_occurrences: int = 3
    start_cell: str = "B14"
    target_col: str = "B"
    date_col: str = "E"
    sheet_name: str = "Punti Aperti"
    analyze_all_sheets: bool = False
    stemming: bool = True
    algorithm: str = "combinato"  # fuzzy | jaccard | combinato
    keyword_threshold: int = 90
    exhaustive_limit: int = 1800
    max_bucket_size: int = 900


@dataclass(slots=True)
class RowRecord:
    file_path: str
    file_name: str
    sheet: str
    row_index: int  # zero based pandas row index
    excel_row: int  # one based Excel row number
    target_col_index: int
    cell: str
    original_text: str
    normalized_text: str
    tokens: set[str] = field(default_factory=set)
    technical_codes: set[str] = field(default_factory=set)
    date_iso: str | None = None


@dataclass(slots=True)
class AnalysisReport:
    mode: str
    summary: list[dict[str, Any]] = field(default_factory=list)
    detail: list[dict[str, Any]] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    @property
    def has_results(self) -> bool:
        return bool(self.summary or self.detail)
