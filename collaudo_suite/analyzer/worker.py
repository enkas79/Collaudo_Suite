from __future__ import annotations

import traceback

from PySide6.QtCore import QThread, Signal

from .analysis import AnalysisCancelled, ExcelAnalyzer
from .models import AnalysisParams, AnalysisReport


class AnalysisWorker(QThread):
    log_message = Signal(str)
    progress_update = Signal(int)
    status_update = Signal(str)
    analysis_finished = Signal(object)
    error_occurred = Signal(str)

    def __init__(self, params: AnalysisParams):
        super().__init__()
        self.params = params
        self._stop_requested = False

    def stop(self) -> None:
        self._stop_requested = True

    def _should_stop(self) -> bool:
        return self._stop_requested

    def run(self) -> None:
        analyzer = ExcelAnalyzer(
            log=self.log_message.emit,
            progress=self.progress_update.emit,
            status=self.status_update.emit,
            stop_requested=self._should_stop,
        )
        try:
            report = analyzer.run(self.params)
            self.analysis_finished.emit(report)
        except AnalysisCancelled as exc:
            report = AnalysisReport(mode="cancelled", warnings=[str(exc)])
            self.status_update.emit("Analisi interrotta.")
            self.log_message.emit(str(exc) + "\n")
            self.analysis_finished.emit(report)
        except Exception as exc:  # pragma: no cover - GUI error path
            details = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
            self.error_occurred.emit(f"{exc}\n\nDettagli tecnici:\n{details}")
