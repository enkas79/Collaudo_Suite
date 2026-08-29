from __future__ import annotations

import os
import re
from collections import Counter
from typing import Callable

import pandas as pd

from .excel_utils import index_to_col
from .models import AnalysisParams, RowRecord
from .text_utils import TextNormalizer


LogFn = Callable[[str], None]
ProgressFn = Callable[[int], None]
StatusFn = Callable[[str], None]
StopFn = Callable[[], bool]


class ExcelDataLoader:
    def __init__(
        self,
        params: AnalysisParams,
        normalizer: TextNormalizer,
        target_col_idx: int,
        date_col_idx: int,
        start_row_idx: int,
        log: LogFn,
        progress: ProgressFn,
        status: StatusFn,
        stop_requested: StopFn,
    ):
        self.params = params
        self.normalizer = normalizer
        self.target_col_idx = target_col_idx
        self.date_col_idx = date_col_idx
        self.start_row_idx = start_row_idx
        self.log = log
        self.progress = progress
        self.status = status
        self.stop_requested = stop_requested
        self.warnings: list[str] = []

    def _warn(self, message: str) -> None:
        self.warnings.append(message)
        self.log("AVVISO: " + message + "\n")

    def collect_records(self) -> tuple[list[RowRecord], list[str]]:
        records: list[RowRecord] = []
        estimated_units = max(len(self.params.files), 1)
        processed_files = 0

        for filepath in self.params.files:
            if self.stop_requested():
                break

            file_name = os.path.basename(filepath)
            if not os.path.exists(filepath):
                self._warn(f"File non trovato: {filepath}")
                processed_files += 1
                self.progress(int(processed_files / estimated_units * 35))
                continue

            self.status(f"Apertura {file_name}...")
            try:
                xls = pd.ExcelFile(filepath)
            except Exception as exc:
                self._warn(f"Impossibile aprire {file_name}: {exc}")
                processed_files += 1
                self.progress(int(processed_files / estimated_units * 35))
                continue

            if self.params.analyze_all_sheets:
                sheet_names = list(xls.sheet_names)
            else:
                sheet = self.params.sheet_name.strip() or "Punti Aperti"
                if sheet not in xls.sheet_names:
                    self._warn(f"Foglio '{sheet}' non trovato in {file_name}. Fogli disponibili: {', '.join(xls.sheet_names)}")
                    processed_files += 1
                    self.progress(int(processed_files / estimated_units * 35))
                    continue
                sheet_names = [sheet]

            for sheet_name in sheet_names:
                if self.stop_requested():
                    break
                self.status(f"Lettura {file_name} / {sheet_name}...")
                try:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                except Exception as exc:
                    self._warn(f"Errore lettura foglio '{sheet_name}' in {file_name}: {exc}")
                    continue

                if df.empty:
                    self._warn(f"Foglio '{sheet_name}' vuoto in {file_name}.")
                    continue

                if self.target_col_idx >= df.shape[1]:
                    self._warn(
                        f"Colonna analisi {index_to_col(self.target_col_idx)} assente in {file_name} / {sheet_name}; "
                        f"disponibili {df.shape[1]} colonne."
                    )
                    continue

                date_col_available = self.date_col_idx >= 0 and self.date_col_idx < df.shape[1]
                if self.date_col_idx >= 0 and not date_col_available:
                    self._warn(
                        f"Colonna data {index_to_col(self.date_col_idx)} assente in {file_name} / {sheet_name}; "
                        "analizzo comunque il testo senza data."
                    )

                invalid_dates = 0
                for row_idx in range(self.start_row_idx, df.shape[0]):
                    if self.stop_requested():
                        break

                    value = df.iat[row_idx, self.target_col_idx]
                    if pd.isna(value):
                        continue
                    original = str(value).strip()
                    if not original:
                        continue

                    normalized = self.normalizer.normalize(original)
                    if not normalized.normalized:
                        continue

                    date_iso = None
                    if date_col_available:
                        date_value = df.iat[row_idx, self.date_col_idx]
                        if pd.notna(date_value):
                            text_date = str(date_value).strip()
                            dayfirst = not bool(re.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}", text_date))
                            parsed = pd.to_datetime(date_value, errors="coerce", dayfirst=dayfirst)
                            if pd.notna(parsed):
                                date_iso = parsed.strftime("%Y-%m-%d")
                            else:
                                invalid_dates += 1

                    records.append(
                        RowRecord(
                            file_path=filepath,
                            file_name=file_name,
                            sheet=str(sheet_name),
                            row_index=row_idx,
                            excel_row=row_idx + 1,
                            target_col_index=self.target_col_idx,
                            cell=f"{index_to_col(self.target_col_idx)}{row_idx + 1}",
                            original_text=original,
                            normalized_text=normalized.normalized,
                            tokens=normalized.tokens,
                            technical_codes=normalized.technical_codes,
                            date_iso=date_iso,
                        )
                    )

                if invalid_dates:
                    self._warn(f"{invalid_dates} date non interpretabili in {file_name} / {sheet_name}.")

            processed_files += 1
            self.progress(int(processed_files / estimated_units * 35))

        by_file = Counter(r.file_name for r in records)
        if by_file:
            self.log("Righe utili lette:\n")
            for file_name, count in by_file.most_common():
                self.log(f"  - {file_name}: {count}\n")
        else:
            self.log("Nessuna riga utile letta.\n")

        return records, self.warnings
