from __future__ import annotations

import json
import re
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any

from PySide6.QtCore import QDate, QObject, QSettings, QStandardPaths, Qt, QThread, QTimer, QUrl, Signal
from PySide6.QtGui import QColor, QDesktopServices, QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCompleter,
    QDateEdit,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QProgressBar,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QComboBox,
    QFrame,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .core import (
    ChecklistItem,
    ChecklistRecord,
    export_pdf,
    get_default_fixed_docx_path,
    get_default_map_xlsx_path,
    load_default_fixed_items,
    load_map_items_for_filter,
    available_map_filters,
    sample_items_from_pool,
    build_ticket_url,
    ticket_desc_sort_key,
)
from .jarvis_api import sync_map_xlsx_from_jarvis

from ..help_dialog import SuiteHelpDialog
from ..app_info import get_app_version
from ..control_exchange import ExternalControl, import_controls
from ..styles import CHECKLIST_TABLE_QSS

APP_NAME = "Collaudo Suite - Checklist"
ORG_NAME = "CollaudoTools"
FIXED_SOURCE_LABEL = "Check list interna"
PROJECT_VERSION = get_app_version()

COL_N = 0
COL_KIND = 1
COL_TICKET = 2
COL_TEXT = 3
COL_PASS = 4
COL_NOPASS = 5
COL_NOTE = 6


class JarvisSyncWorker(QObject):
    progress = Signal(int, int)
    finished = Signal(str, int)
    failed = Signal(str)

    def __init__(self, token: str, output_path: Path) -> None:
        super().__init__()
        self.token = token
        self.output_path = output_path

    def run(self) -> None:
        try:
            path = sync_map_xlsx_from_jarvis(
                self.token,
                self.output_path,
                progress_callback=lambda percent, processed: self.progress.emit(percent, processed),
            )
            self.finished.emit(str(path), 0)
        except Exception as exc:
            self.failed.emit(str(exc))


class MapLoadWorker(QObject):
    """Carica e filtra il file MAP senza bloccare il thread della GUI."""

    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, path: Path, filter_text: str, reference_date: date, months_back: int) -> None:
        super().__init__()
        self.path = path
        self.filter_text = filter_text
        self.reference_date = reference_date
        self.months_back = months_back

    def run(self) -> None:
        try:
            items = load_map_items_for_filter(
                self.path,
                self.filter_text,
                reference_date=self.reference_date,
                months_back=self.months_back,
            )
            self.finished.emit(items)
        except Exception as exc:
            self.failed.emit(str(exc))


class PdfExportWorker(QObject):
    """Genera il PDF della checklist fuori dal thread della GUI."""

    finished = Signal(str)
    failed = Signal(str)

    def __init__(self, records: list[ChecklistRecord], output_path: Path, header_info: dict[str, str]) -> None:
        super().__init__()
        self.records = records
        self.output_path = output_path
        self.header_info = header_info

    def run(self) -> None:
        try:
            path = export_pdf(
                self.records,
                self.output_path,
                source_random="",
                source_fixed=FIXED_SOURCE_LABEL,
                header_info=self.header_info,
                seed=None,
            )
            self.finished.emit(str(path))
        except Exception as exc:
            self.failed.emit(str(exc))


class ChecklistWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1280, 780)
        self.settings = QSettings(ORG_NAME, APP_NAME)
        self.fixed_items: list[ChecklistItem] = []
        self.map_pool: list[ChecklistItem] = []
        self.external_items: list[ChecklistItem] = []
        self.current_items: list[ChecklistItem] = []
        self.table_row_to_item_index: list[int | None] = []
        self.current_project_path: Path | None = None
        self.last_autosave_path: Path | None = None
        self._jarvis_thread: QThread | None = None
        self._jarvis_worker: JarvisSyncWorker | None = None
        self._jarvis_sync_silent = False
        self._map_load_thread: QThread | None = None
        self._map_load_worker: MapLoadWorker | None = None
        self._pdf_thread: QThread | None = None
        self._pdf_worker: PdfExportWorker | None = None
        self._pdf_for_print = False

        self._build_ui()
        self._load_settings()
        self._load_fixed_items()
        QTimer.singleShot(1200, self._auto_sync_jarvis_cache)

    def _build_ui(self) -> None:
        central = QWidget(self)
        self.setCentralWidget(central)

        root = QHBoxLayout(central)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(8)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        self.splitter = splitter
        root.addWidget(splitter)

        left_panel = QWidget()
        self.left_panel = left_panel
        left_panel.setMinimumWidth(310)
        left_panel.setMaximumWidth(410)
        left_panel.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.MinimumExpanding)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        header_box = QGroupBox("Intestazione")
        header_layout = QVBoxLayout(header_box)

        self.operator_edit = QLineEdit()
        self.operator_edit.setPlaceholderText("Nome e cognome")
        self.department_edit = QLineEdit()
        self.department_edit.setPlaceholderText("Reparto")
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("dd/MM/yyyy")
        self.date_edit.setDate(QDate.currentDate())
        self.order_number_edit = QLineEdit()
        self.order_number_edit.setPlaceholderText("Numero commessa")

        # I campi non devono comprimersi quando la scheda viene nascosta e mostrata
        # nuovamente dentro QTabWidget. Il pannello sinistro diventa scorrevole se
        # l'altezza disponibile non e sufficiente.
        for field in (
            self.operator_edit,
            self.department_edit,
            self.date_edit,
            self.order_number_edit,
        ):
            field.setMinimumHeight(28)
            field.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        header_layout.addWidget(QLabel("Collaudatore:"))
        header_layout.addWidget(self.operator_edit)
        header_layout.addWidget(QLabel("Reparto:"))
        header_layout.addWidget(self.department_edit)
        header_layout.addWidget(QLabel("Data:"))
        date_row = QHBoxLayout()
        date_row.addWidget(self.date_edit, 1)
        today_btn = QPushButton("Today")
        today_btn.setMaximumWidth(80)
        today_btn.clicked.connect(self.set_today)
        date_row.addWidget(today_btn)
        header_layout.addLayout(date_row)
        header_layout.addWidget(QLabel("Numero Commessa:"))
        header_layout.addWidget(self.order_number_edit)
        header_box.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        self.header_box = header_box
        left_layout.addWidget(header_box)

        map_box = QGroupBox("Estrazione")
        map_layout = QVBoxLayout(map_box)

        map_count_row = QHBoxLayout()
        self.map_count_spin = QSpinBox()
        self.map_count_spin.setMinimum(0)
        self.map_count_spin.setMaximum(999)
        self.map_count_spin.setValue(3)
        map_count_row.addWidget(QLabel("Segnalazioni MAP random:"))
        map_count_row.addWidget(self.map_count_spin)
        map_count_row.addStretch(1)
        map_layout.addLayout(map_count_row)

        self.map_filter_combo = QComboBox()
        self.map_filter_combo.setEditable(True)
        self.map_filter_combo.addItems(available_map_filters())
        self.map_filter_combo.setCurrentText("Tutti")
        map_layout.addWidget(QLabel("Prefisso Commercial code MAP:"))
        map_layout.addWidget(self.map_filter_combo)

        self.map_period_combo = QComboBox()
        self.map_period_combo.addItem("1 mese", 1)
        self.map_period_combo.addItem("3 mesi", 3)
        self.map_period_combo.addItem("6 mesi", 6)
        self.map_period_combo.addItem("1 anno", 12)
        self.map_period_combo.setCurrentIndex(self.map_period_combo.findData(12))
        self.map_period_combo.setToolTip(
            "Considera solo i MAP compresi tra la data di collaudo e il periodo precedente selezionato."
        )
        map_layout.addWidget(QLabel("Periodo MAP rispetto alla data di collaudo:"))
        map_layout.addWidget(self.map_period_combo)

        self.map_source_combo = QComboBox()
        self.map_source_combo.addItem("File Excel", "excel")
        self.map_source_combo.addItem("JARVIS API", "jarvis")
        self.map_source_combo.currentIndexChanged.connect(self._update_map_source_ui)
        map_layout.addWidget(QLabel("Sorgente MAP:"))
        map_layout.addWidget(self.map_source_combo)

        self.jarvis_token_edit = QLineEdit()
        self.jarvis_token_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.jarvis_token_edit.setPlaceholderText("Token con permesso ticket.read")
        self.jarvis_token_edit.setToolTip("Il token viene salvato solo nelle impostazioni locali dell'utente.")
        map_layout.addWidget(QLabel("Token JARVIS:"))
        map_layout.addWidget(self.jarvis_token_edit)

        self.jarvis_sync_btn = QPushButton("Aggiorna MAP da JARVIS")
        self.jarvis_sync_btn.clicked.connect(self.sync_jarvis_map)
        map_layout.addWidget(self.jarvis_sync_btn)

        self.jarvis_progress = QProgressBar()
        self.jarvis_progress.setRange(0, 100)
        self.jarvis_progress.setValue(0)
        self.jarvis_progress.setFormat("Sincronizzazione JARVIS: %p%")
        self.jarvis_progress.setVisible(False)
        map_layout.addWidget(self.jarvis_progress)

        self.map_file_edit = QLineEdit()
        self.map_file_edit.setReadOnly(True)
        self.map_file_edit.setPlaceholderText("File MAP interno")
        map_file_row = QHBoxLayout()
        map_file_row.addWidget(self.map_file_edit, 1)
        browse_map_btn = QPushButton("Sfoglia...")
        browse_map_btn.clicked.connect(self.browse_map_file)
        map_file_row.addWidget(browse_map_btn)
        map_layout.addWidget(QLabel("File MAP riepilogativo:"))
        map_layout.addLayout(map_file_row)

        self.preview_btn = QPushButton("Estrai / Aggiorna anteprima")
        self.preview_btn.clicked.connect(self.refresh_preview)
        map_layout.addWidget(self.preview_btn)
        map_box.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        self.map_box = map_box
        left_layout.addWidget(map_box)

        analyzer_box = QGroupBox("Controlli Analyzer")
        analyzer_layout = QVBoxLayout(analyzer_box)
        self.external_count_label = QLabel("Nessun controllo importato.")
        self.external_count_label.setWordWrap(True)
        analyzer_layout.addWidget(self.external_count_label)
        analyzer_buttons = QHBoxLayout()
        import_analyzer_btn = QPushButton("Importa...")
        import_analyzer_btn.setToolTip("Importa controlli creati con Analyzer da Excel, CSV o JSON.")
        import_analyzer_btn.clicked.connect(self.import_analyzer_controls)
        clear_analyzer_btn = QPushButton("Svuota")
        clear_analyzer_btn.clicked.connect(self.clear_external_controls)
        analyzer_buttons.addWidget(import_analyzer_btn, 1)
        analyzer_buttons.addWidget(clear_analyzer_btn)
        analyzer_layout.addLayout(analyzer_buttons)
        analyzer_box.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        self.analyzer_box = analyzer_box
        left_layout.addWidget(analyzer_box)


        left_layout.addStretch(1)

        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(6)

        self.summary_label = QLabel(
            "Compila l'intestazione, scegli quante segnalazioni MAP random aggiungere e genera l'anteprima. "
            "Il filtro MAP lavora sul Commercial code; usa Tutti per pescare da tutto il file. "
            "Il periodo MAP e calcolato a ritroso dalla data di collaudo selezionata. "
            "Le righe MAP includono anche il numero Ticket quando presente. "
            "I comandi di salvataggio, apertura, PDF e uscita sono disponibili come icone nella barra superiore."
        )
        self.summary_label.setWordWrap(True)
        right_layout.addWidget(self.summary_label)

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(["N.", "Tipo", "Ticket", "Controllo", "Pass", "No pass", "Note"])
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setWordWrap(True)
        self.table.setMouseTracking(True)
        self.table.viewport().setMouseTracking(True)
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked | QAbstractItemView.EditKeyPressed)
        self.table.setStyleSheet(CHECKLIST_TABLE_QSS)

        header = self.table.horizontalHeader()
        header.setMinimumSectionSize(36)
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setStretchLastSection(True)
        self.table.itemChanged.connect(self._on_table_item_changed)
        self.table.cellClicked.connect(self._on_table_cell_clicked)
        right_layout.addWidget(self.table, 1)

        left_scroll = QScrollArea()
        self.left_scroll = left_scroll
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.Shape.NoFrame)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        left_scroll.setMinimumWidth(320)
        left_scroll.setMaximumWidth(430)
        left_scroll.setWidget(left_panel)

        splitter.addWidget(left_scroll)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([350, 930])

        self._fit_table_columns()

        # La Checklist viene incorporata nella finestra principale della Suite.
        # Una seconda menu bar interna verrebbe visualizzata dentro la scheda
        # (in basso o in posizione anomala). I relativi comandi sono quindi
        # gestiti dalla barra superiore di SuiteMainWindow.
        self.menuBar().hide()

    def _update_map_source_ui(self) -> None:
        is_jarvis = self.map_source_combo.currentData() == "jarvis"
        self.jarvis_token_edit.setEnabled(is_jarvis)
        self.jarvis_sync_btn.setEnabled(is_jarvis)
        self.map_file_edit.setEnabled(not is_jarvis)

    def _refresh_map_filter_options(self, path: Path | None = None) -> None:
        current = self.map_filter_combo.currentText().strip() or "Tutti"
        options = available_map_filters(path if path and path.exists() else None)
        self.map_filter_combo.blockSignals(True)
        self.map_filter_combo.clear()
        self.map_filter_combo.addItems(options)
        self.map_filter_combo.setCurrentText(current if current else "Tutti")
        self.map_filter_combo.blockSignals(False)
        completer = QCompleter(options, self.map_filter_combo)
        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        completer.setFilterMode(Qt.MatchFlag.MatchContains)
        completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        self._map_filter_completer = completer
        self.map_filter_combo.setCompleter(completer)

    def _load_settings(self) -> None:
        # Mantiene solo le preferenze di estrazione.
        # L'intestazione deve essere sempre pulita a ogni nuovo avvio.
        self.map_count_spin.setValue(int(self.settings.value("map_count", 3)))
        self.map_filter_combo.setCurrentText(str(self.settings.value("map_filter", "Tutti")))
        saved_period = int(self.settings.value("map_period_months", 12))
        period_index = self.map_period_combo.findData(saved_period)
        self.map_period_combo.setCurrentIndex(period_index if period_index >= 0 else self.map_period_combo.findData(12))
        self.map_file_edit.setText(str(self.settings.value("map_file_path", "")))
        configured_map_path = Path(self.map_file_edit.text().strip()) if self.map_file_edit.text().strip() else get_default_map_xlsx_path()
        self._refresh_map_filter_options(configured_map_path)
        source = str(self.settings.value("map_source", "excel"))
        source_index = self.map_source_combo.findData(source)
        self.map_source_combo.setCurrentIndex(source_index if source_index >= 0 else 0)
        self.jarvis_token_edit.setText(str(self.settings.value("jarvis_token", "")))
        self._update_map_source_ui()
        self.operator_edit.clear()
        self.department_edit.clear()
        self.order_number_edit.clear()
        self.date_edit.setDate(QDate.currentDate())

    def _save_settings(self) -> None:
        # I dati di intestazione vengono salvati solo nei file lavoro, non nelle preferenze.
        self.settings.setValue("map_count", self.map_count_spin.value())
        self.settings.setValue("map_filter", self.map_filter_combo.currentText().strip() or "Tutti")
        self.settings.setValue("map_period_months", self._selected_map_period_months())
        self.settings.setValue("map_file_path", self.map_file_edit.text().strip())
        self.settings.setValue("map_source", self.map_source_combo.currentData() or "excel")
        self.settings.setValue("jarvis_token", self.jarvis_token_edit.text().strip())


    def set_today(self) -> None:
        self.date_edit.setDate(QDate.currentDate())

    def _header_info(self) -> dict[str, str]:
        return {
            "collaudatore": self.operator_edit.text().strip(),
            "reparto": self.department_edit.text().strip(),
            "data": self.date_edit.date().toString("dd/MM/yyyy"),
            "numero_commessa": self.order_number_edit.text().strip(),
        }

    def _load_fixed_items(self) -> None:
        try:
            self.fixed_items = load_default_fixed_items()
        except Exception as exc:
            QMessageBox.critical(self, "Errore checklist fissa", str(exc))

    def _selected_map_filter(self) -> str:
        text = self.map_filter_combo.currentText().strip()
        if not text or text.casefold() == "tutti":
            return ""
        return text

    def _selected_map_period_months(self) -> int:
        value = self.map_period_combo.currentData()
        try:
            months = int(value)
        except (TypeError, ValueError):
            months = 12
        return months if months in {1, 3, 6, 12} else 12

    def _collaudo_reference_date(self) -> date:
        selected = self.date_edit.date()
        return date(selected.year(), selected.month(), selected.day())

    def _map_period_label(self) -> str:
        return self.map_period_combo.currentText().strip() or "1 anno"

    def _selected_map_path(self) -> Path:
        path_text = self.map_file_edit.text().strip()
        if path_text:
            return Path(path_text)
        return get_default_map_xlsx_path()

    def _selected_map_source(self) -> str:
        return str(self.map_source_combo.currentData() or "excel")

    def _jarvis_cache_path(self) -> Path:
        base = QStandardPaths.writableLocation(QStandardPaths.AppDataLocation)
        root = Path(base) if base else Path.home() / ".collaudo_suite"
        return root / "map_cache_jarvis.xlsx"

    def _auto_sync_jarvis_cache(self) -> None:
        if self._selected_map_source() != "jarvis":
            return
        if not self._jarvis_cache_path().exists() or not self.jarvis_token_edit.text().strip():
            return
        self.sync_jarvis_map(silent=True)

    def sync_jarvis_map(self, *, silent: bool = False) -> None:
        """Refresh the local Excel cache used by the JARVIS MAP source."""
        if self._jarvis_thread is not None and self._jarvis_thread.isRunning():
            return
        token = self.jarvis_token_edit.text().strip()
        if not token:
            QMessageBox.warning(self, "Token JARVIS mancante", "Inserire un token con permesso ticket.read.")
            return
        self._save_settings()
        self._jarvis_sync_silent = silent
        self.jarvis_sync_btn.setEnabled(False)
        self.jarvis_progress.setVisible(True)
        self.jarvis_progress.setValue(0)
        self.jarvis_progress.setFormat("Sincronizzazione JARVIS: avvio...")

        thread = QThread(self)
        worker = JarvisSyncWorker(token, self._jarvis_cache_path())
        worker.moveToThread(thread)
        worker.progress.connect(self._on_jarvis_sync_progress)
        worker.finished.connect(self._on_jarvis_sync_finished)
        worker.failed.connect(self._on_jarvis_sync_failed)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        thread.started.connect(worker.run)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(self._on_jarvis_thread_finished)
        self._jarvis_thread = thread
        self._jarvis_worker = worker
        thread.start()

    def _on_jarvis_sync_progress(self, percent: int, processed: int) -> None:
        self.jarvis_progress.setValue(max(0, min(100, percent)))
        self.jarvis_progress.setFormat(f"Sincronizzazione JARVIS: {processed} ticket letti (%p%)")

    def _on_jarvis_sync_finished(self, path: str, _unused: int) -> None:
        self.jarvis_progress.setValue(100)
        self.jarvis_progress.setFormat("Cache MAP aggiornata (%p%)")
        self.map_file_edit.setText(path)
        self._save_settings()
        self._refresh_map_filter_options(Path(path))
        if not self._jarvis_sync_silent:
            QMessageBox.information(
                self,
                "MAP JARVIS aggiornato",
                f"Cache locale aggiornata con successo.\n\nFile: {path}",
            )
        if self._selected_map_source() == "jarvis":
            QTimer.singleShot(0, self.refresh_preview)

    def _on_jarvis_sync_failed(self, message: str) -> None:
        self.jarvis_progress.setVisible(False)
        self.jarvis_sync_btn.setEnabled(self._selected_map_source() == "jarvis")
        if self._jarvis_sync_silent:
            self.summary_label.setText(f"Aggiornamento automatico MAP non riuscito: {message}")
        else:
            QMessageBox.critical(self, "Errore aggiornamento MAP JARVIS", message)

    def _on_jarvis_thread_finished(self) -> None:
        self._jarvis_thread = None
        self._jarvis_worker = None
        self._jarvis_sync_silent = False
        self.jarvis_sync_btn.setEnabled(self._selected_map_source() == "jarvis")
        QTimer.singleShot(700, lambda: self.jarvis_progress.setVisible(False))

    def browse_map_file(self) -> None:
        start_dir = str(self._downloads_dir())
        current = self.map_file_edit.text().strip()
        if current and Path(current).exists():
            start_dir = str(Path(current).parent)
        path, _ = QFileDialog.getOpenFileName(self, "Scegli file MAP riepilogativo", start_dir, "Excel (*.xlsx)")
        if not path:
            return
        self.map_file_edit.setText(path)
        self._refresh_map_filter_options(Path(path))
        self._save_settings()

    @staticmethod
    def _control_key(text: str) -> str:
        return re.sub(r"\W+", "", str(text).casefold())

    def _update_external_label(self) -> None:
        count = len(self.external_items)
        self.external_count_label.setText(
            f"{count} controllo/i Analyzer pronti per la checklist." if count else "Nessun controllo importato."
        )

    def import_analyzer_controls(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Importa controlli Analyzer",
            str(self._downloads_dir()),
            "Controlli (*.xlsx *.xlsm *.csv *.json)",
        )
        if not path:
            return
        try:
            controls = import_controls(path)
            if not controls:
                raise ValueError("Il file non contiene controlli validi.")
            added = self.add_external_controls(controls, replace=False)
            QMessageBox.information(
                self,
                "Importazione completata",
                f"Controlli letti: {len(controls)}\nNuovi controlli aggiunti: {added}",
            )
        except Exception as exc:
            QMessageBox.critical(self, "Errore importazione", str(exc))

    def add_external_controls(self, controls: list[ExternalControl], *, replace: bool = False) -> int:
        existing = [] if replace else list(self.external_items)
        seen = {self._control_key(item.text) for item in existing}
        added = 0
        for control in controls:
            text = str(control.control).strip()
            key = self._control_key(text)
            if not text or not key or key in seen:
                continue
            existing.append(
                ChecklistItem(
                    text=text,
                    source=control.source_label(),
                    kind="ANALYZER",
                    ticket_number=str(control.ticket).strip(),
                    ticket_url="",
                )
            )
            seen.add(key)
            added += 1
        self.external_items = existing
        self._update_external_label()

        if self.current_items:
            previous = self._collect_records()
            records = [record for record in previous if record.kind != "ANALYZER"]
            analyzer_records = [
                ChecklistRecord(item.text, item.source, item.kind, ticket_number=item.ticket_number, ticket_url=item.ticket_url)
                for item in self.external_items
            ]
            analyzer_records = self._records_with_saved_state(analyzer_records, previous)
            self._replace_table_section("ANALYZER", analyzer_records, records)
            self.summary_label.setText(
                f"Checklist aggiornata con {len(self.external_items)} controlli importati da Analyzer. "
                "I risultati Pass/No pass già compilati sugli altri controlli sono stati mantenuti."
            )
        return added

    def clear_external_controls(self) -> None:
        if not self.external_items:
            return
        self.external_items = []
        self._update_external_label()
        if self.current_items:
            records = [record for record in self._collect_records() if record.kind != "ANALYZER"]
            self._replace_table_section("ANALYZER", [], records)
            self.summary_label.setText("Controlli Analyzer rimossi dalla checklist corrente.")

    def _start_map_load(self, path: Path, map_filter: str) -> None:
        if self._map_load_thread is not None and self._map_load_thread.isRunning():
            return
        self.preview_btn.setEnabled(False)
        self.summary_label.setText("Caricamento e filtro del file MAP in corso...")
        thread = QThread(self)
        worker = MapLoadWorker(
            path,
            map_filter,
            reference_date=self._collaudo_reference_date(),
            months_back=self._selected_map_period_months(),
        )
        worker.moveToThread(thread)
        worker.finished.connect(self._on_map_load_finished)
        worker.failed.connect(self._on_map_load_failed)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        thread.started.connect(worker.run)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(self._on_map_load_thread_finished)
        self._map_load_thread = thread
        self._map_load_worker = worker
        thread.start()

    def _on_map_load_finished(self, items: object) -> None:
        self.preview_btn.setEnabled(True)
        self.refresh_preview(_loaded_map_pool=list(items))

    def _on_map_load_failed(self, message: str) -> None:
        self.preview_btn.setEnabled(True)
        QMessageBox.critical(self, "Errore lettura MAP", message)

    def _on_map_load_thread_finished(self) -> None:
        self._map_load_thread = None
        self._map_load_worker = None

    def refresh_preview(self, *, _loaded_map_pool: list[ChecklistItem] | None = None) -> None:
        try:
            self._save_settings()
            if not self.fixed_items:
                self._load_fixed_items()
            if not self.fixed_items:
                raise RuntimeError("La checklist fissa interna non è disponibile.")

            base_items = [ChecklistItem(item.text, item.source, "Fisso", ticket_url=item.ticket_url) for item in self.fixed_items]
            warnings: list[str] = []

            map_filter = self._selected_map_filter()
            self.map_pool = []
            map_items: list[ChecklistItem] = []
            if self.map_count_spin.value() > 0:
                if _loaded_map_pool is None:
                    if self._selected_map_source() == "jarvis":
                        map_path = self._jarvis_cache_path()
                        if not map_path.exists():
                            self.sync_jarvis_map()
                            return
                    else:
                        map_path = self._selected_map_path()
                    if map_path.exists():
                        self._start_map_load(map_path, map_filter)
                        return

                if self._selected_map_source() == "jarvis":
                    cache_path = self._jarvis_cache_path()
                    self.map_file_edit.setText(str(cache_path))
                    self.map_pool = list(_loaded_map_pool or [])
                else:
                    map_path = self._selected_map_path()
                    if not map_path.exists():
                        warnings.append(
                            "Il file MAP interno non è presente. La lista è stata creata senza controlli MAP; "
                            "seleziona un file MAP esterno con Sfoglia per abilitarne l'estrazione."
                        )
                    else:
                        self.map_pool = list(_loaded_map_pool or [])
                if self.map_pool:
                    map_items, map_warnings = sample_items_from_pool(
                        self.map_pool,
                        self.map_count_spin.value(),
                        kind="MAP",
                        exclude_items=base_items,
                        seed=None,
                        label="segnalazioni MAP",
                    )
                    warnings.extend(map_warnings)

            used_keys = {self._control_key(item.text) for item in base_items + map_items}
            external_items = [item for item in self.external_items if self._control_key(item.text) not in used_keys]
            old_records = self._collect_records() if self.current_items else []
            map_records = self._records_with_saved_state(map_items, old_records)
            fixed_records = self._records_with_saved_state(base_items, old_records)
            analyzer_records = self._records_with_saved_state(external_items, old_records)
            all_records = fixed_records + map_records + analyzer_records
            self.current_project_path = None
            if self.current_items:
                self._replace_table_section("MAP", map_records, fixed_records + analyzer_records)
                self._replace_table_section("Fisso", fixed_records, map_records + analyzer_records)
                self._replace_table_section("ANALYZER", analyzer_records, fixed_records + map_records)
            else:
                self._apply_records_to_table(all_records)

            fixed_count = sum(1 for item in self.current_items if item.kind == "Fisso")
            map_count = sum(1 for item in self.current_items if item.kind == "MAP")
            analyzer_count = sum(1 for item in self.current_items if item.kind == "ANALYZER")
            msg = (
                f"Anteprima: {fixed_count} controlli fissi + {map_count} segnalazioni MAP random + "
                f"{analyzer_count} controlli Analyzer = {len(self.current_items)} controlli totali. "
                f"Pool MAP Commercial code '{map_filter or 'Tutti'}': {len(self.map_pool)}. "
                f"Periodo: {self._map_period_label()} fino al {self.date_edit.date().toString('dd/MM/yyyy')}. "
                f"Fonte MAP: {'JARVIS API' if self._selected_map_source() == 'jarvis' else 'file riepilogativo'}. "
                "Estrazione casuale a ogni aggiornamento."
            )
            if warnings:
                msg += "  Avvisi: " + " | ".join(warnings)

            autosave_error = ""
            try:
                autosave_path = self._autosave_work()
                if autosave_path:
                    msg += f"  Autosalvataggio: {autosave_path.name}."
            except Exception as exc:
                autosave_error = str(exc)
                msg += f"  Avviso autosalvataggio: {autosave_error}"
            self.summary_label.setText(msg)
        except Exception as exc:
            QMessageBox.critical(self, "Errore", str(exc))

    def _readonly_item(self, value: str, align: Qt.AlignmentFlag | Qt.Alignment = Qt.AlignLeft) -> QTableWidgetItem:
        item = QTableWidgetItem(value)
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        item.setTextAlignment(align)
        return item

    def _check_item(self) -> QTableWidgetItem:
        item = QTableWidgetItem("")
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        item.setTextAlignment(Qt.AlignCenter)
        font = QFont()
        font.setBold(True)
        font.setPointSize(12)
        item.setFont(font)
        item.setForeground(QColor("#111111"))
        item.setData(Qt.UserRole, False)
        return item

    def _section_item(self, title: str, count: int) -> QTableWidgetItem:
        item = QTableWidgetItem(f"{title} ({count})")
        item.setFlags(Qt.ItemIsEnabled)
        item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        font = QFont()
        font.setBold(True)
        item.setFont(font)
        item.setBackground(QColor("#e7e6e6"))
        item.setForeground(QColor("#222222"))
        return item

    def _row_background_for_kind(self, kind: str) -> QColor:
        if kind == "Fisso":
            return QColor("#f8f8f8")
        if kind == "MAP":
            return QColor("#fff8ed")
        if kind == "ANALYZER":
            return QColor("#eef7ff")
        return QColor("#ffffff")

    def _apply_row_background(self, row: int, kind: str) -> None:
        background = self._row_background_for_kind(kind)
        for col in range(self.table.columnCount()):
            table_item = self.table.item(row, col)
            if table_item is not None:
                table_item.setBackground(background)

    def _add_section_row(self, row: int, title: str, count: int) -> None:
        self.table.setSpan(row, 0, 1, self.table.columnCount())
        self.table.setItem(row, 0, self._section_item(title, count))
        self.table.setRowHeight(row, 28)
        self.table_row_to_item_index.append(None)

    @staticmethod
    def _record_key(record: ChecklistRecord | ChecklistItem) -> tuple[str, str, str]:
        return (
            str(record.kind).casefold(),
            ChecklistWindow._control_key(record.text),
            str(record.ticket_number).strip(),
        )

    def _records_with_saved_state(
        self, items: list[ChecklistItem], previous: list[ChecklistRecord]
    ) -> list[ChecklistRecord]:
        previous_by_key = {self._record_key(record): record for record in previous}
        records: list[ChecklistRecord] = []
        for item in items:
            saved = previous_by_key.get(self._record_key(item))
            records.append(
                ChecklistRecord(
                    item.text,
                    item.source,
                    item.kind,
                    result=saved.result if saved else "",
                    note=saved.note if saved else "",
                    ticket_number=item.ticket_number,
                    ticket_url=item.ticket_url,
                )
            )
        return records

    def _section_title(self, kind: str) -> str:
        return {
            "Fisso": "CONTROLLI FISSI",
            "MAP": "SEGNALAZIONI MAP",
            "ANALYZER": "CONTROLLI IMPORTATI DA ANALYZER",
        }.get(kind, kind)

    def _section_rows(self, kind: str) -> tuple[int, int] | None:
        title = self._section_title(kind)
        start = None
        for row in range(self.table.rowCount()):
            item = self.table.item(row, COL_N)
            if item is not None and item.text().split(" (")[0] == title:
                start = row
                break
        if start is None:
            return None
        end = start + 1
        while end < self.table.rowCount():
            item = self.table.item(end, COL_N)
            if item is not None and item.text().split(" (")[0] in {
                "CONTROLLI FISSI", "SEGNALAZIONI MAP", "CONTROLLI IMPORTATI DA ANALYZER"
            }:
                break
            end += 1
        return start, end

    def _write_data_row(self, row: int, item: ChecklistItem, display_number: int, record: ChecklistRecord | None = None) -> None:
        self.table.setItem(row, COL_N, self._readonly_item(str(display_number), Qt.AlignCenter))
        self.table.setItem(row, COL_KIND, self._readonly_item(item.kind, Qt.AlignCenter))
        ticket_cell = self._readonly_item(item.ticket_number, Qt.AlignCenter)
        ticket_url = item.ticket_url.strip() or build_ticket_url(item.ticket_number)
        if ticket_url:
            font = ticket_cell.font()
            font.setUnderline(True)
            ticket_cell.setFont(font)
            ticket_cell.setForeground(QColor("#0563c1"))
            ticket_cell.setToolTip("Clic per aprire il ticket")
            ticket_cell.setData(Qt.UserRole + 1, ticket_url)
        self.table.setItem(row, COL_TICKET, ticket_cell)
        self.table.setItem(row, COL_TEXT, self._readonly_item(item.text))
        self.table.setItem(row, COL_PASS, self._check_item())
        self.table.setItem(row, COL_NOPASS, self._check_item())
        note_item = QTableWidgetItem(record.note if record else "")
        note_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
        self.table.setItem(row, COL_NOTE, note_item)
        self._apply_row_background(row, item.kind)
        if record:
            if record.result == "Pass":
                self._set_result_cell(row, COL_PASS, True)
            elif record.result == "No pass":
                self._set_result_cell(row, COL_NOPASS, True)

    def _refresh_table_mapping(self) -> None:
        self.table_row_to_item_index = []
        display_number = 1
        for row in range(self.table.rowCount()):
            section = self.table.item(row, COL_N)
            if section is not None and section.text().split(" (")[0] in {
                "CONTROLLI FISSI", "SEGNALAZIONI MAP", "CONTROLLI IMPORTATI DA ANALYZER"
            }:
                self.table_row_to_item_index.append(None)
                continue
            kind_item = self.table.item(row, COL_KIND)
            if kind_item is None:
                self.table_row_to_item_index.append(None)
                continue
            kind = kind_item.text()
            candidates = [
                idx for idx, item in enumerate(self.current_items)
                if item.kind == kind and idx not in self.table_row_to_item_index
            ]
            candidates.sort(key=lambda idx: ticket_desc_sort_key(self.current_items[idx].ticket_number), reverse=True)
            idx = candidates[0] if candidates else None
            self.table_row_to_item_index.append(idx)
            if idx is not None:
                self.table.item(row, COL_N).setText(str(display_number))
                display_number += 1

    def _replace_table_section(
        self, kind: str, records: list[ChecklistRecord], other_records: list[ChecklistRecord]
    ) -> None:
        """Replace only one visual section, preserving widgets in all other sections."""
        ordered_records = records
        items = [ChecklistItem(r.text, r.source, r.kind, ticket_number=r.ticket_number, ticket_url=r.ticket_url) for r in ordered_records]
        self.current_items = [
            ChecklistItem(r.text, r.source, r.kind, ticket_number=r.ticket_number, ticket_url=r.ticket_url)
            for r in (other_records + records)
        ]
        self.table.blockSignals(True)
        section = self._section_rows(kind)
        if section is None:
            if not items:
                self.table.blockSignals(False)
                return
            insert_at = self.table.rowCount()
            group_order = ("Fisso", "MAP", "ANALYZER")
            kind_position = group_order.index(kind)
            for candidate_kind in group_order[kind_position + 1:]:
                found = self._section_rows(candidate_kind)
                if found:
                    insert_at = found[0]
                    break
            self.table.insertRow(insert_at)
            self.table.setSpan(insert_at, 0, 1, self.table.columnCount())
            self.table.setItem(insert_at, 0, self._section_item(self._section_title(kind), len(items)))
            section = (insert_at, insert_at + 1)
        start, end = section
        old_data_count = end - start - 1
        new_data_count = len(items)
        for row in range(start + old_data_count, start + new_data_count, -1):
            self.table.removeRow(row)
        for _ in range(new_data_count - old_data_count):
            self.table.insertRow(end)
        if not items:
            self.table.removeRow(start)
        else:
            self.table.item(start, 0).setText(f"{self._section_title(kind)} ({len(items)})")
            sorted_pairs = list(enumerate(ordered_records))
            sorted_pairs.sort(key=lambda pair: ticket_desc_sort_key(pair[1].ticket_number), reverse=True)
            for offset, (index, record) in enumerate(sorted_pairs, start=1):
                self._write_data_row(start + offset, items[index], 0, record)
        self._refresh_table_mapping()
        self.table.blockSignals(False)
        self.table.resizeRowsToContents()
        self._fit_table_columns()

    def _populate_table(self, items: list[ChecklistItem]) -> None:
        self.table.blockSignals(True)
        self.table.clearSpans()
        self.table_row_to_item_index = []

        groups = [
            ("Fisso", "CONTROLLI FISSI"),
            ("MAP", "SEGNALAZIONI MAP"),
            ("ANALYZER", "CONTROLLI IMPORTATI DA ANALYZER"),
        ]
        grouped_indices: list[tuple[str, str, list[int]]] = []
        for kind, title in groups:
            indices = [idx for idx, item in enumerate(items) if item.kind == kind]
            # Visualizza prima i ticket piu recenti. L'ordinamento viene applicato
            # alla sola vista: gli indici continuano a puntare agli elementi originali,
            # quindi esiti e note restano associati alla riga corretta.
            indices.sort(key=lambda idx: ticket_desc_sort_key(items[idx].ticket_number), reverse=True)
            if indices:
                grouped_indices.append((kind, title, indices))

        total_rows = sum(len(indices) + 1 for _, _, indices in grouped_indices)
        self.table.setRowCount(total_rows)

        row = 0
        display_number = 1
        for kind, title, indices in grouped_indices:
            self._add_section_row(row, title, len(indices))
            row += 1
            for item_index in indices:
                item = items[item_index]
                self.table_row_to_item_index.append(item_index)
                self.table.setItem(row, COL_N, self._readonly_item(str(display_number), Qt.AlignCenter))
                self.table.setItem(row, COL_KIND, self._readonly_item(item.kind, Qt.AlignCenter))
                ticket_cell = self._readonly_item(item.ticket_number, Qt.AlignCenter)
                ticket_url = item.ticket_url.strip() or build_ticket_url(item.ticket_number)
                if ticket_url:
                    font = ticket_cell.font()
                    font.setUnderline(True)
                    ticket_cell.setFont(font)
                    ticket_cell.setForeground(QColor("#0563c1"))
                    ticket_cell.setToolTip("Clic per aprire il ticket")
                    ticket_cell.setData(Qt.UserRole + 1, ticket_url)
                self.table.setItem(row, COL_TICKET, ticket_cell)
                self.table.setItem(row, COL_TEXT, self._readonly_item(item.text))
                self.table.setItem(row, COL_PASS, self._check_item())
                self.table.setItem(row, COL_NOPASS, self._check_item())
                note_item = QTableWidgetItem("")
                note_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
                self.table.setItem(row, COL_NOTE, note_item)
                self._apply_row_background(row, item.kind)
                row += 1
                display_number += 1

        self.table.blockSignals(False)
        self.table.resizeRowsToContents()
        self._fit_table_columns()

    def _apply_records_to_table(self, records: list[ChecklistRecord]) -> None:
        self.current_items = [ChecklistItem(record.text, record.source, record.kind, ticket_number=record.ticket_number, ticket_url=record.ticket_url) for record in records]
        self._populate_table(self.current_items)
        self.table.blockSignals(True)
        for row, item_index in enumerate(self.table_row_to_item_index):
            if item_index is None:
                continue
            record = records[item_index]
            if record.result == "Pass":
                self._set_result_cell(row, COL_PASS, True)
                self._set_result_cell(row, COL_NOPASS, False)
            elif record.result == "No pass":
                self._set_result_cell(row, COL_PASS, False)
                self._set_result_cell(row, COL_NOPASS, True)
            note_item = self.table.item(row, COL_NOTE)
            if note_item is not None:
                note_item.setText(record.note)
        self.table.blockSignals(False)
        self.table.resizeRowsToContents()
        self._fit_table_columns()

    def _fit_table_columns(self) -> None:
        if not hasattr(self, "table"):
            return
        viewport_width = max(self.table.viewport().width(), 820)
        self.table.resizeColumnsToContents()

        self.table.setColumnWidth(COL_N, max(44, min(62, self.table.columnWidth(COL_N) + 8)))
        self.table.setColumnWidth(COL_KIND, max(62, min(86, self.table.columnWidth(COL_KIND) + 8)))
        self.table.setColumnWidth(COL_TICKET, max(72, min(105, self.table.columnWidth(COL_TICKET) + 8)))
        self.table.setColumnWidth(COL_PASS, max(58, min(70, self.table.columnWidth(COL_PASS) + 8)))
        self.table.setColumnWidth(COL_NOPASS, max(76, min(92, self.table.columnWidth(COL_NOPASS) + 8)))

        fixed_width = (
            self.table.columnWidth(COL_N)
            + self.table.columnWidth(COL_KIND)
            + self.table.columnWidth(COL_TICKET)
            + self.table.columnWidth(COL_PASS)
            + self.table.columnWidth(COL_NOPASS)
            + 30
        )
        available = max(520, viewport_width - fixed_width)
        note_content_width = self.table.columnWidth(COL_NOTE)
        note_width = max(170, min(340, note_content_width + 28, int(available * 0.32)))
        control_width = max(390, available - note_width)
        self.table.setColumnWidth(COL_TEXT, control_width)
        self.table.setColumnWidth(COL_NOTE, note_width)

    def refresh_layout(self) -> None:
        """Ripristina le dimensioni minime dopo ogni riapertura della scheda."""
        for field in (
            self.operator_edit,
            self.department_edit,
            self.date_edit,
            self.order_number_edit,
            self.map_filter_combo,
            self.map_period_combo,
            self.map_file_edit,
            self.map_source_combo,
            self.jarvis_token_edit,
            self.jarvis_progress,
        ):
            field.setMinimumHeight(28)
            field.updateGeometry()

        for box in (self.header_box, self.map_box, self.analyzer_box):
            box.updateGeometry()

        self.left_panel.updateGeometry()
        self.left_scroll.updateGeometry()
        sizes = self.splitter.sizes()
        if not sizes or sizes[0] < 320:
            total = max(sum(sizes), self.width() - 24) if sizes else max(self.width() - 24, 1000)
            self.splitter.setSizes([350, max(650, total - 350)])
        self._fit_table_columns()

    def showEvent(self, event) -> None:  # noqa: N802 - Qt override
        super().showEvent(event)
        QTimer.singleShot(0, self.refresh_layout)

    def resizeEvent(self, event) -> None:  # noqa: N802 - Qt override
        super().resizeEvent(event)
        self._fit_table_columns()

    def _set_result_cell(self, row: int, column: int, checked: bool) -> None:
        cell = self.table.item(row, column)
        if cell is None:
            return
        cell.setData(Qt.UserRole, checked)
        cell.setText("X" if checked else "")

    def _is_result_cell_checked(self, row: int, column: int) -> bool:
        cell = self.table.item(row, column)
        if cell is None:
            return False
        return bool(cell.data(Qt.UserRole))

    def _on_table_cell_clicked(self, row: int, column: int) -> None:
        if column == COL_TICKET:
            # Il numero ticket e un collegamento operativo: un singolo clic apre il ticket.
            self._open_ticket_row(row, show_warning=False)
            return
        if column not in (COL_PASS, COL_NOPASS):
            return
        if row >= len(self.table_row_to_item_index) or self.table_row_to_item_index[row] is None:
            return

        currently_checked = self._is_result_cell_checked(row, column)
        other_col = COL_NOPASS if column == COL_PASS else COL_PASS

        self.table.blockSignals(True)
        self._set_result_cell(row, column, not currently_checked)
        if not currently_checked:
            self._set_result_cell(row, other_col, False)
        self.table.blockSignals(False)

    def _ticket_url_for_row(self, row: int) -> str:
        if row < 0 or row >= len(self.table_row_to_item_index):
            return ""
        item_index = self.table_row_to_item_index[row]
        if item_index is None or item_index >= len(self.current_items):
            return ""
        item = self.current_items[item_index]
        return item.ticket_url.strip() or build_ticket_url(item.ticket_number)

    def _open_ticket_row(self, row: int, *, show_warning: bool = True) -> bool:
        url = self._ticket_url_for_row(row)
        if not url:
            if show_warning:
                QMessageBox.warning(
                    self,
                    "Ticket MAP",
                    "La riga selezionata non contiene un numero ticket valido.",
                )
            return False
        if not QDesktopServices.openUrl(QUrl(url)):
            QMessageBox.critical(self, "Ticket MAP", f"Impossibile aprire il collegamento:\n{url}")
            return False
        return True

    def _on_table_item_changed(self, item: QTableWidgetItem) -> None:
        return

    def _ensure_preview(self) -> bool:
        if not self.current_items:
            self.refresh_preview()
        return bool(self.current_items)

    def _collect_records(self) -> list[ChecklistRecord]:
        records: list[ChecklistRecord] = []
        for row, item_index in enumerate(self.table_row_to_item_index):
            if item_index is None:
                continue
            item = self.current_items[item_index]
            pass_item = self.table.item(row, COL_PASS)
            nopass_item = self.table.item(row, COL_NOPASS)
            note_item = self.table.item(row, COL_NOTE)
            result = ""
            if pass_item and bool(pass_item.data(Qt.UserRole)):
                result = "Pass"
            elif nopass_item and bool(nopass_item.data(Qt.UserRole)):
                result = "No pass"
            note = note_item.text().strip() if note_item else ""
            records.append(ChecklistRecord(item.text, item.source, item.kind, result=result, note=note, ticket_number=item.ticket_number, ticket_url=item.ticket_url))
        return records

    def _downloads_dir(self) -> Path:
        location = QStandardPaths.writableLocation(QStandardPaths.DownloadLocation)
        if location:
            return Path(location)
        fallback = Path.home() / "Downloads"
        return fallback if fallback.exists() else Path.home()

    def _safe_stem(self, value: str, fallback: str = "collaudo") -> str:
        cleaned = re.sub(r"[^A-Za-z0-9_\-]+", "_", value.strip())
        cleaned = cleaned.strip("_")
        return cleaned or fallback

    def _default_export_name(self, suffix: str) -> str:
        stem = self._safe_stem(self.order_number_edit.text(), "collaudo")
        return f"checklist_{stem}.{suffix}"

    def _default_project_name(self) -> str:
        stem = self._safe_stem(self.order_number_edit.text(), "collaudo")
        return f"checklist_lavoro_{stem}.rcl.json"

    def _default_export_path(self, suffix: str) -> str:
        return str(self._downloads_dir() / self._default_export_name(suffix))

    def _default_project_path(self) -> str:
        return str(self._downloads_dir() / self._default_project_name())

    def _open_pdf(self, pdf_path: str | Path) -> None:
        path = Path(pdf_path).expanduser().resolve()
        if not path.is_file():
            raise FileNotFoundError(f"File PDF non trovato: {path}")
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(path))):
            raise RuntimeError("Impossibile aprire il PDF con il programma predefinito.")

    def _persistent_print_path(self) -> Path:
        target = Path(self._default_export_path("pdf"))
        if not target.exists():
            return target
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return target.with_name(f"{target.stem}_{timestamp}{target.suffix}")

    def save_pdf(self) -> None:
        if not self._ensure_preview():
            return
        path, _ = QFileDialog.getSaveFileName(self, "Salva check list PDF finale", self._default_export_path("pdf"), "PDF (*.pdf)")
        if not path:
            return
        if not path.lower().endswith(".pdf"):
            path += ".pdf"
        self._start_pdf_export(Path(path), for_print=False)

    def print_pdf(self) -> None:
        if not self._ensure_preview():
            return
        self._start_pdf_export(self._persistent_print_path(), for_print=True)

    def _start_pdf_export(self, target: Path, *, for_print: bool) -> None:
        if self._pdf_thread is not None:
            QMessageBox.information(self, "Esportazione PDF", "Una generazione PDF è già in corso, attendere il completamento.")
            return
        # I dati della tabella si leggono qui (thread GUI); il worker riceve solo copie.
        worker = PdfExportWorker(self._collect_records(), target, dict(self._header_info()))
        thread = QThread(self)
        worker.moveToThread(thread)
        # Slot legati a self (thread GUI): Qt usa una connessione accodata, mai lambda nel thread worker.
        worker.finished.connect(self._on_pdf_export_finished)
        worker.failed.connect(self._on_pdf_export_failed)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        thread.started.connect(worker.run)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(self._on_pdf_thread_finished)
        self._pdf_thread = thread
        self._pdf_worker = worker
        self._pdf_for_print = for_print
        self.summary_label.setText("Generazione PDF in corso...")
        QApplication.setOverrideCursor(Qt.CursorShape.BusyCursor)
        thread.start()

    def _on_pdf_thread_finished(self) -> None:
        QApplication.restoreOverrideCursor()
        thread = self._pdf_thread
        self._pdf_thread = None
        self._pdf_worker = None
        if thread is not None:
            thread.deleteLater()

    def _on_pdf_export_finished(self, pdf_path: str) -> None:
        self.summary_label.setText(f"PDF creato: {pdf_path}")
        try:
            self._open_pdf(pdf_path)
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Errore apertura PDF",
                f"Il PDF è stato creato ma non è stato possibile aprirlo.\n\nFile: {pdf_path}\nDettaglio: {exc}",
            )
            return
        if self._pdf_for_print:
            QMessageBox.information(
                self,
                "PDF pronto per la stampa",
                "Il PDF è stato salvato nella cartella Download e aperto normalmente.\n"
                "Il lettore resterà aperto: usa il comando Stampa del lettore PDF (ad esempio Ctrl+P).\n\n"
                f"File:\n{pdf_path}",
            )
        else:
            QMessageBox.information(
                self,
                "Esportazione completata",
                f"File creato e aperto con il lettore PDF predefinito:\n{pdf_path}",
            )

    def _on_pdf_export_failed(self, message: str) -> None:
        self.summary_label.setText("Generazione PDF non riuscita.")
        title = "Errore apertura PDF" if self._pdf_for_print else "Errore esportazione PDF"
        QMessageBox.critical(self, title, f"Non sono riuscito a creare il PDF.\n\nDettaglio: {message}")

    def wait_for_background_tasks(self, timeout_ms: int = 5000) -> None:
        """Attende i thread ancora attivi prima della chiusura dell'applicazione."""
        for thread in (self._pdf_thread, self._map_load_thread):
            if thread is not None and thread.isRunning():
                thread.quit()
                thread.wait(timeout_ms)

    def _project_payload(self) -> dict[str, Any]:
        return {
            "version": PROJECT_VERSION,
            "saved_at": datetime.now().isoformat(timespec="seconds"),
            "header": self._header_info(),
            "map_count": self.map_count_spin.value(),
            "map_filter": self.map_filter_combo.currentText().strip() or "Tutti",
            "map_period_months": self._selected_map_period_months(),
            "map_file_path": self.map_file_edit.text().strip(),
            "map_source": self._selected_map_source(),
            "items": [
                {
                    "text": record.text,
                    "source": record.source,
                    "kind": record.kind,
                    "ticket_number": record.ticket_number,
                    "ticket_url": record.ticket_url,
                    "result": record.result,
                    "pass": record.result == "Pass",
                    "nopass": record.result == "No pass",
                    "note": record.note,
                }
                for record in self._collect_records()
            ],
        }

    def _autosave_project_path(self) -> Path:
        stem = self._safe_stem(self.order_number_edit.text(), "collaudo")
        return self._downloads_dir() / f"checklist_autosave_{stem}.rcl.json"

    def _write_project_payload(self, target: Path) -> Path:
        """Write project data atomically to reduce corruption risk."""
        target = Path(target)
        target.parent.mkdir(parents=True, exist_ok=True)
        temporary = target.with_name(target.name + ".tmp")
        temporary.write_text(
            json.dumps(self._project_payload(), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        temporary.replace(target)
        return target

    def _autosave_work(self) -> Path | None:
        if not self.current_items:
            return None
        target = self._write_project_payload(self._autosave_project_path())
        self.last_autosave_path = target
        return target

    def autosave_on_exit(self) -> Path | None:
        """Persist the latest checklist state without changing manual-save path."""
        try:
            return self._autosave_work()
        except Exception as exc:
            print(f"Errore autosalvataggio in uscita: {exc}", file=sys.stderr)
            return None

    def closeEvent(self, event) -> None:  # noqa: N802 - Qt override
        self.autosave_on_exit()
        self._save_settings()
        self.wait_for_background_tasks()
        event.accept()

    @staticmethod
    def _ensure_project_extension(path: str) -> str:
        if path.lower().endswith((".rcl.json", ".json")):
            return path
        return path + ".rcl.json"

    def _save_work_to(self, target: Path, *, show_confirmation: bool) -> bool:
        try:
            saved_path = self._write_project_payload(target)
            self.current_project_path = saved_path
            self.summary_label.setText(f"Lavoro salvato: {saved_path}")
            if show_confirmation:
                QMessageBox.information(self, "Lavoro salvato", f"File creato:\n{saved_path}")
            return True
        except Exception as exc:
            QMessageBox.critical(self, "Errore salvataggio lavoro", str(exc))
            return False

    def save_work(self) -> None:
        """Salva sul file corrente; al primo salvataggio richiede il nome."""
        if not self._ensure_preview():
            return
        if self.current_project_path is None:
            self.save_work_as()
            return
        self._save_work_to(self.current_project_path, show_confirmation=False)

    def save_work_as(self) -> None:
        """Richiede sempre un nuovo nome e lo imposta come file corrente."""
        if not self._ensure_preview():
            return
        start_path = str(self.current_project_path) if self.current_project_path else self._default_project_path()
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Salva lavoro con nome",
            start_path,
            "Lavoro checklist (*.rcl.json);;JSON (*.json)",
        )
        if not path:
            return
        target = Path(self._ensure_project_extension(path))
        self._save_work_to(target, show_confirmation=True)

    def load_work(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Apri lavoro", str(self._downloads_dir()), "Lavoro checklist (*.rcl.json *.json)")
        if not path:
            return
        try:
            source = Path(path)
            data = json.loads(source.read_text(encoding="utf-8"))
            header = data.get("header", {}) if isinstance(data, dict) else {}
            self.operator_edit.setText(str(header.get("collaudatore", "")))
            self.department_edit.setText(str(header.get("reparto", "")))
            date_value = str(header.get("data", ""))
            parsed_date = QDate.fromString(date_value, "dd/MM/yyyy")
            if parsed_date.isValid():
                self.date_edit.setDate(parsed_date)
            self.order_number_edit.setText(str(header.get("numero_commessa", header.get("tipo_macchina", ""))))
            self.map_count_spin.setValue(int(data.get("map_count", self.map_count_spin.value())))
            self.map_filter_combo.setCurrentText(str(data.get("map_filter", self.map_filter_combo.currentText() or "Tutti")))
            saved_period = int(data.get("map_period_months", self._selected_map_period_months()))
            period_index = self.map_period_combo.findData(saved_period)
            if period_index >= 0:
                self.map_period_combo.setCurrentIndex(period_index)
            self.map_file_edit.setText(str(data.get("map_file_path", "")))
            source = str(data.get("map_source", "excel"))
            source_index = self.map_source_combo.findData(source)
            self.map_source_combo.setCurrentIndex(source_index if source_index >= 0 else 0)
            self._update_map_source_ui()

            records: list[ChecklistRecord] = []
            for raw in data.get("items", []):
                if not isinstance(raw, dict):
                    continue
                text = str(raw.get("text", "")).strip()
                if not text:
                    continue
                kind = str(raw.get("kind", "Fisso")).strip() or "Fisso"
                source_label = str(raw.get("source", ""))
                ticket_number = str(raw.get("ticket_number", raw.get("ticket", "")))
                ticket_url = str(raw.get("ticket_url", ""))
                result = str(raw.get("result", ""))
                if not result:
                    if bool(raw.get("pass", False)):
                        result = "Pass"
                    elif bool(raw.get("nopass", False)):
                        result = "No pass"
                if result not in {"", "Pass", "No pass"}:
                    result = ""
                note = str(raw.get("note", ""))
                records.append(ChecklistRecord(text, source_label, kind, result=result, note=note, ticket_number=ticket_number, ticket_url=ticket_url))
            if not records:
                raise ValueError("Il file lavoro non contiene controlli validi.")

            self.external_items = [
                ChecklistItem(record.text, record.source, "ANALYZER", ticket_number=record.ticket_number, ticket_url=record.ticket_url)
                for record in records if record.kind == "ANALYZER"
            ]
            self._update_external_label()
            self._save_settings()
            self.current_project_path = source
            self._apply_records_to_table(records)
            fixed_count = sum(1 for r in records if r.kind == "Fisso")
            map_count = sum(1 for r in records if r.kind == "MAP")
            analyzer_count = sum(1 for r in records if r.kind == "ANALYZER")
            self.summary_label.setText(
                f"Lavoro caricato: {fixed_count} controlli fissi + {map_count} segnalazioni MAP random + "
                f"{analyzer_count} controlli Analyzer = {len(records)} controlli totali. File: {source}"
            )
        except Exception as exc:
            QMessageBox.critical(self, "Errore apertura lavoro", str(exc))

    def show_fixed_info(self) -> None:
        try:
            fixed_path = get_default_fixed_docx_path()
            map_path = self._selected_map_path()
            text = (
                f"Checklist fissa interna caricata: {len(self.fixed_items)} controlli.\n\n"
                f"File checklist fissa:\n{fixed_path}\n\n"
                f"File segnalazioni MAP in uso:\n{map_path}\n\n"
                "Le schede collaudo random sono state rimosse: il programma usa la checklist fissa e, se richieste, "
                "le segnalazioni MAP filtrate sul Commercial code e sul periodo scelti in GUI, poi estratte casualmente. "
                "Con filtro 'Tutti' pesca da tutte le righe MAP con Title/Titolo valido. "
                "Il numero presente nella colonna Ticket/Numero Ticket viene importato e mostrato nella colonna Ticket.\n\n"
                "Nota Verniciatura: la riga generale G001 non viene usata come controllo autonomo; "
                "viene espansa sui sottocontrolli Basamento, Trave, Canotto, Trasporti e Siliconatura."
            )
        except Exception:
            text = "File interni non disponibili."
        QMessageBox.information(self, "Info file interni", text)

    def show_help(self) -> None:
        SuiteHelpDialog(self, initial_tab=3).exec()



def main() -> int:
    app = QApplication(sys.argv)
    window = ChecklistWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
