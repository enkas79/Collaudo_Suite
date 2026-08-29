from __future__ import annotations

from collections import Counter
from datetime import datetime
from statistics import mean
from typing import Callable

from .data_loader import ExcelDataLoader
from .excel_utils import cell_to_indices, col_to_index, index_to_col
from .models import AnalysisParams, AnalysisReport, RowRecord
from .similarity import (
    UnionFind,
    calculate_similarity,
    candidate_pairs,
    compatible_for_comparison,
    count_candidate_pairs,
    partial_score,
)
from .text_utils import TextNormalizer, short_label


class AnalysisCancelled(Exception):
    pass


LogFn = Callable[[str], None]
ProgressFn = Callable[[int], None]
StatusFn = Callable[[str], None]
StopFn = Callable[[], bool]


def _noop_log(_: str) -> None:
    return None


def _noop_progress(_: int) -> None:
    return None


def _noop_status(_: str) -> None:
    return None


def _never_stop() -> bool:
    return False


class ExcelAnalyzer:
    def __init__(
        self,
        log: LogFn = _noop_log,
        progress: ProgressFn = _noop_progress,
        status: StatusFn = _noop_status,
        stop_requested: StopFn = _never_stop,
    ):
        self.log = log
        self.progress = progress
        self.status = status
        self.stop_requested = stop_requested
        self.warnings: list[str] = []
        self.errors: list[str] = []

    def _check_stop(self) -> None:
        if self.stop_requested():
            raise AnalysisCancelled("Analisi interrotta dall'utente.")

    def _warn(self, message: str) -> None:
        self.warnings.append(message)
        self.log("AVVISO: " + message + "\n")

    def run(self, params: AnalysisParams) -> AnalysisReport:
        self.warnings = []
        self.errors = []
        self.progress(0)
        self.status("Validazione parametri...")

        start_row_idx, start_col_idx = cell_to_indices(params.start_cell)
        target_col_idx = col_to_index(params.target_col)
        date_col_idx = col_to_index(params.date_col) if params.date_col.strip() else -1

        if start_col_idx >= 0 and start_col_idx != target_col_idx:
            self._warn(
                "La colonna della cella start e diversa dalla colonna analisi: uso solo la riga "
                f"{start_row_idx + 1}. Colonna analisi effettiva: {index_to_col(target_col_idx)}."
            )

        if not params.files:
            raise ValueError("Seleziona almeno un file Excel.")
        if not (0 <= params.threshold <= 100):
            raise ValueError("La soglia deve essere compresa tra 0 e 100.")
        if params.min_occurrences < 1:
            raise ValueError("Min occorrenze deve essere almeno 1.")

        normalizer = TextNormalizer(stemming=params.stemming)
        if normalizer.warning:
            self._warn(normalizer.warning)

        metadata = self._build_metadata(params, target_col_idx, date_col_idx, start_row_idx)
        self.log("Parametri analisi:\n")
        for key, value in metadata.items():
            self.log(f"  - {key}: {value}\n")
        self.log("\n")

        loader = ExcelDataLoader(
            params=params,
            normalizer=normalizer,
            target_col_idx=target_col_idx,
            date_col_idx=date_col_idx,
            start_row_idx=start_row_idx,
            log=self.log,
            progress=self.progress,
            status=self.status,
            stop_requested=self.stop_requested,
        )
        records, loader_warnings = loader.collect_records()
        self.warnings.extend(loader_warnings)
        self._check_stop()

        if params.keyword.strip():
            return self._run_keyword(params, records, metadata)
        return self._run_clustering(params, records, metadata)

    @staticmethod
    def _build_metadata(params: AnalysisParams, target_col_idx: int, date_col_idx: int, start_row_idx: int) -> dict[str, object]:
        return {
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "files_count": len(params.files),
            "sheet_mode": "tutti i fogli" if params.analyze_all_sheets else params.sheet_name,
            "start_row": start_row_idx + 1,
            "target_col": index_to_col(target_col_idx),
            "date_col": index_to_col(date_col_idx) if date_col_idx >= 0 else "N/D",
            "keyword": params.keyword.strip(),
            "keyword_threshold": params.keyword_threshold,
            "threshold": params.threshold,
            "min_occurrences": params.min_occurrences,
            "algorithm": params.algorithm,
            "stemming_requested": params.stemming,
            "exhaustive_limit": params.exhaustive_limit,
            "max_bucket_size": params.max_bucket_size,
        }

    def _run_keyword(self, params: AnalysisParams, records: list[RowRecord], metadata: dict[str, object]) -> AnalysisReport:
        keyword = params.keyword.strip().upper()
        self.status("Ricerca parola chiave...")
        self.log(f"Ricerca keyword: {keyword}\n")
        hits: list[dict[str, object]] = []
        total = max(len(records), 1)

        for idx, record in enumerate(records):
            self._check_stop()
            score = partial_score(record.original_text.upper(), keyword)
            if score >= params.keyword_threshold:
                row = {
                    "FILE": record.file_name,
                    "FOGLIO": record.sheet,
                    "RIGA": record.excel_row,
                    "CELLA": record.cell,
                    "TESTO TROVATO": record.original_text,
                    "DATA": record.date_iso or "N/D",
                    "SCORE_KEYWORD": round(score, 1),
                }
                hits.append(row)
                self.log(f"{record.file_name} | {record.sheet} | {record.cell}: {record.original_text}\n")

            if idx % 50 == 0 or idx == total - 1:
                self.progress(35 + int((idx + 1) / total * 60))

        self.progress(100)
        self.status("Completato.")
        if not hits:
            self.log("Nessun risultato trovato con la keyword indicata.\n")
        return AnalysisReport(
            mode="with_keywords",
            summary=hits,
            detail=hits.copy(),
            metadata={**metadata, "records_read": len(records), "hits": len(hits)},
            warnings=self.warnings.copy(),
            errors=self.errors.copy(),
        )

    def _run_clustering(self, params: AnalysisParams, records: list[RowRecord], metadata: dict[str, object]) -> AnalysisReport:
        total_records = len(records)
        self.log(f"Confronto righe utili: {total_records}\n")

        if total_records < params.min_occurrences:
            self.progress(100)
            self.status("Completato.")
            self.log("Righe insufficienti rispetto a Min Occorrenze.\n")
            return AnalysisReport(
                mode="no_keywords",
                metadata={**metadata, "records_read": total_records, "groups": 0},
                warnings=self.warnings.copy(),
                errors=self.errors.copy(),
            )

        pair_total = count_candidate_pairs(records, params.exhaustive_limit, params.max_bucket_size)
        if pair_total is None:
            self.log("Dataset grande: uso confronto per bucket, non confronto completo n^2.\n")
        else:
            self.log(f"Coppie candidate stimate: {pair_total}\n")

        uf = UnionFind(total_records)
        edge_scores: dict[tuple[int, int], float] = {}
        compared = 0
        accepted = 0

        self.status("Confronto similarita...")
        for i, j in candidate_pairs(records, params.exhaustive_limit, params.max_bucket_size):
            self._check_stop()
            r1 = records[i]
            r2 = records[j]
            if not compatible_for_comparison(r1.tokens, r2.tokens, r1.normalized_text, r2.normalized_text):
                continue

            score = calculate_similarity(
                r1.normalized_text,
                r2.normalized_text,
                r1.tokens,
                r2.tokens,
                params.algorithm,
            )
            compared += 1
            if score >= params.threshold:
                uf.union(i, j)
                edge_scores[(i, j)] = score
                accepted += 1

            if compared % 500 == 0:
                if pair_total:
                    self.progress(35 + min(55, int(compared / pair_total * 55)))
                else:
                    # Unknown total: move slowly but keep GUI alive.
                    self.progress(min(90, 35 + compared // 5000))
                self.status(f"Confrontate {compared} coppie, accettate {accepted}")

        self.status("Costruzione gruppi...")
        groups = [idxs for idxs in uf.groups().values() if len(idxs) >= params.min_occurrences]
        groups.sort(key=len, reverse=True)

        summary: list[dict[str, object]] = []
        detail: list[dict[str, object]] = []

        for group_number, indexes in enumerate(groups, start=1):
            self._check_stop()
            group_records = [records[idx] for idx in indexes]
            representative_index = self._choose_medoid(indexes, records, params.algorithm)
            representative = records[representative_index]
            latest_date = self._latest_date(group_records)
            file_count = len({r.file_name for r in group_records})
            sheet_count = len({r.sheet for r in group_records})

            scores_vs_rep: list[float] = []
            for idx in indexes:
                record = records[idx]
                if idx == representative_index:
                    score = 100.0
                else:
                    score = calculate_similarity(
                        representative.normalized_text,
                        record.normalized_text,
                        representative.tokens,
                        record.tokens,
                        params.algorithm,
                    )
                scores_vs_rep.append(score)
                detail.append(
                    {
                        "ID_GRUPPO": group_number,
                        "FILE": record.file_name,
                        "FOGLIO": record.sheet,
                        "RIGA": record.excel_row,
                        "CELLA": record.cell,
                        "TESTO_ORIGINALE": record.original_text,
                        "TESTO_NORMALIZZATO": record.normalized_text,
                        "DATA": record.date_iso or "N/D",
                        "SCORE_VS_RAPPRESENTANTE": round(score, 1),
                        "CODICI_TECNICI": ", ".join(sorted(record.technical_codes)),
                    }
                )

            row = {
                "ID_GRUPPO": group_number,
                "CONCETTO PRINCIPALE DEL GRUPPO": representative.original_text,
                "CONCETTO BREVE": short_label(representative.original_text),
                "NUMERO DI OCCORRENZE": len(indexes),
                "DATA PIU RECENTE": latest_date or "N/D",
                "FILE COINVOLTI": file_count,
                "FOGLI COINVOLTI": sheet_count,
                "SCORE MEDIO VS RAPPRESENTANTE": round(mean(scores_vs_rep), 1) if scores_vs_rep else 0,
            }
            summary.append(row)
            self.log(
                f"GRUPPO {group_number}: {row['CONCETTO PRINCIPALE DEL GRUPPO']}\n"
                f"  Occorrenze: {len(indexes)} | Data piu recente: {row['DATA PIU RECENTE']} | "
                f"Score medio: {row['SCORE MEDIO VS RAPPRESENTANTE']}\n\n"
            )

        if not summary:
            self.log("Nessun gruppo trovato con i parametri attuali.\n")

        self.progress(100)
        self.status("Completato.")
        return AnalysisReport(
            mode="no_keywords",
            summary=summary,
            detail=detail,
            metadata={
                **metadata,
                "records_read": total_records,
                "pairs_compared": compared,
                "pairs_accepted": accepted,
                "groups": len(summary),
            },
            warnings=self.warnings.copy(),
            errors=self.errors.copy(),
        )

    @staticmethod
    def _latest_date(records: list[RowRecord]) -> str | None:
        dates = [r.date_iso for r in records if r.date_iso]
        return max(dates) if dates else None

    @staticmethod
    def _choose_medoid(indexes: list[int], records: list[RowRecord], algorithm: str) -> int:
        if len(indexes) == 1:
            return indexes[0]

        # Full medoid on normal groups, capped for huge groups to keep the UI responsive.
        sample = indexes[:200]
        best_idx = sample[0]
        best_score = -1.0
        for candidate_idx in sample:
            candidate = records[candidate_idx]
            total = 0.0
            for other_idx in sample:
                if candidate_idx == other_idx:
                    continue
                other = records[other_idx]
                total += calculate_similarity(
                    candidate.normalized_text,
                    other.normalized_text,
                    candidate.tokens,
                    other.tokens,
                    algorithm,
                )
            if total > best_score:
                best_score = total
                best_idx = candidate_idx
        return best_idx
