from __future__ import annotations

import os
import re
import textwrap
from pathlib import Path

from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg
from matplotlib.figure import Figure
from PySide6.QtCore import Qt, QStandardPaths, Signal
from PySide6.QtGui import QTextCursor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QAbstractSpinBox,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..control_exchange import ExternalControl, export_controls_xlsx
from .exporters import export_report
from .models import AnalysisParams, AnalysisReport
from .worker import AnalysisWorker


from ..help_dialog import SuiteHelpDialog
class AnalyzerWindow(QMainWindow):
    controls_ready = Signal(object)
    SIDEBAR_WIDTH = 410

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Analyzer anomalie")
        self.selected_files: list[str] = []
        self.analysis_report: AnalysisReport | None = None
        self.worker: AnalysisWorker | None = None
        self.canvas: FigureCanvasQTAgg | None = None

        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        self.sidebar = QFrame()
        self.sidebar.setObjectName("sidebar")
        self.sidebar.setFixedWidth(self.SIDEBAR_WIDTH)
        side_outer = QVBoxLayout(self.sidebar)
        side_outer.setContentsMargins(0, 0, 0, 0)
        side_outer.setSpacing(0)

        controls_frame = QFrame()
        controls_frame.setObjectName("sidebarControls")
        controls_layout = QVBoxLayout(controls_frame)
        controls_layout.setContentsMargins(14, 12, 14, 8)
        controls_layout.setSpacing(7)

        title_row = QHBoxLayout()
        title = QLabel("CONFIGURAZIONE ANALISI")
        title.setObjectName("sidebarTitle")
        title_row.addWidget(title, 1)
        help_btn = QPushButton("?")
        help_btn.setObjectName("btnMiniHelp")
        help_btn.clicked.connect(self.show_help)
        title_row.addWidget(help_btn)
        controls_layout.addLayout(title_row)

        controls_layout.addWidget(self._section_label("INPUT"))
        self.btn_select = QPushButton("Seleziona file Excel")
        self.btn_select.clicked.connect(self.select_files)
        controls_layout.addWidget(self.btn_select)
        self.lbl_files = QLabel("Nessun file selezionato.")
        self.lbl_files.setWordWrap(True)
        self.lbl_files.setObjectName("fileLabel")
        controls_layout.addWidget(self.lbl_files)

        controls_layout.addWidget(self._separator())
        controls_layout.addWidget(self._section_label("RICERCA"))
        self.txt_key = self._line_edit(placeholder="Parola chiave opzionale")
        controls_layout.addWidget(self.txt_key)

        search_grid = QGridLayout()
        search_grid.addWidget(QLabel("Soglia key"), 0, 0)
        search_grid.addWidget(QLabel("Soglia gruppi"), 0, 1)
        search_grid.addWidget(QLabel("Min. occ."), 0, 2)
        self.spin_keyword_threshold = self._spinbox(1, 100, 90)
        self.spin_thresh = self._spinbox(0, 100, 70)
        self.spin_occ = self._spinbox(1, 10000, 3)
        search_grid.addWidget(self.spin_keyword_threshold, 1, 0)
        search_grid.addWidget(self.spin_thresh, 1, 1)
        search_grid.addWidget(self.spin_occ, 1, 2)
        controls_layout.addLayout(search_grid)

        controls_layout.addWidget(self._separator())
        controls_layout.addWidget(self._section_label("ORIGINE DATI"))
        coord_grid = QGridLayout()
        coord_grid.addWidget(QLabel("Riga start"), 0, 0)
        coord_grid.addWidget(QLabel("Col. analisi"), 0, 1)
        coord_grid.addWidget(QLabel("Col. data"), 0, 2)
        self.txt_start = self._line_edit("14")
        self.txt_col = self._line_edit("B")
        self.txt_date = self._line_edit("E")
        coord_grid.addWidget(self.txt_start, 1, 0)
        coord_grid.addWidget(self.txt_col, 1, 1)
        coord_grid.addWidget(self.txt_date, 1, 2)
        controls_layout.addLayout(coord_grid)

        sheet_grid = QGridLayout()
        sheet_grid.addWidget(QLabel("Foglio Excel"), 0, 0)
        self.txt_sheet = self._line_edit("Punti Aperti")
        sheet_grid.addWidget(self.txt_sheet, 1, 0)
        self.chk_all_sheets = QCheckBox("Tutti i fogli")
        self.chk_all_sheets.stateChanged.connect(self._toggle_sheet_field)
        sheet_grid.addWidget(self.chk_all_sheets, 1, 1)
        controls_layout.addLayout(sheet_grid)

        self.chk_stem = QCheckBox("Stemming italiano")
        self.chk_stem.setChecked(True)
        controls_layout.addWidget(self.chk_stem)

        controls_layout.addWidget(self._separator())
        controls_layout.addWidget(self._section_label("ANALISI AVANZATA"))
        self.combo_algo = QComboBox()
        self.combo_algo.addItem("Combinato - consigliato", "combinato")
        self.combo_algo.addItem("Fuzzy", "fuzzy")
        self.combo_algo.addItem("Jaccard", "jaccard")
        controls_layout.addWidget(self.combo_algo)

        advanced = QGridLayout()
        advanced.addWidget(QLabel("Limite n2"), 0, 0)
        advanced.addWidget(QLabel("Max bucket"), 0, 1)
        self.spin_exhaustive = self._spinbox(100, 20000, 1800, 100)
        self.spin_bucket = self._spinbox(50, 10000, 900, 50)
        advanced.addWidget(self.spin_exhaustive, 1, 0)
        advanced.addWidget(self.spin_bucket, 1, 1)
        controls_layout.addLayout(advanced)

        scroll = QScrollArea()
        scroll.setObjectName("sidebarScroll")
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setWidget(controls_frame)
        side_outer.addWidget(scroll, 1)

        actions = QFrame()
        action_layout = QGridLayout(actions)
        action_layout.setContentsMargins(14, 8, 14, 12)
        self.btn_analyze = QPushButton("Avvia analisi")
        self.btn_analyze.setObjectName("btnStart")
        self.btn_analyze.clicked.connect(self.start_analysis)
        self.btn_stop = QPushButton("Stop")
        self.btn_stop.setObjectName("btnStop")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.stop_analysis)
        self.btn_export = QPushButton("Esporta report")
        self.btn_export.setObjectName("btnExport")
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self.export_analysis_report)
        self.btn_exit = QPushButton("Esci")
        self.btn_exit.setObjectName("btnExit")
        self.btn_exit.clicked.connect(QApplication.instance().quit)
        action_layout.addWidget(self.btn_analyze, 0, 0)
        action_layout.addWidget(self.btn_stop, 0, 1)
        action_layout.addWidget(self.btn_export, 1, 0)
        action_layout.addWidget(self.btn_exit, 1, 1)
        side_outer.addWidget(actions)

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(12, 12, 12, 10)
        header = QLabel("Analisi anomalie e preparazione controlli")
        header.setObjectName("contentTitle")
        content_layout.addWidget(header)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tabs.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        self.txt_out = QTextEdit()
        self.txt_out.setReadOnly(True)
        self.txt_out.setPlaceholderText("Avvia un'analisi per visualizzare il log.")
        self.tabs.addTab(self.txt_out, "Report testuale")

        chart_tab = QWidget()
        self.layout_chart = QVBoxLayout(chart_tab)
        self.tabs.addTab(chart_tab, "Grafico")
        self.tabs.setTabEnabled(1, False)

        controls_tab = QWidget()
        controls_tab_layout = QVBoxLayout(controls_tab)
        intro = QLabel(
            "Verifica e modifica il testo operativo. Seleziona i controlli da trasferire direttamente alla Checklist "
            "oppure esportali in un file Excel compatibile."
        )
        intro.setWordWrap(True)
        controls_tab_layout.addWidget(intro)
        self.controls_table = QTableWidget(0, 8)
        self.controls_table.setHorizontalHeaderLabels(
            ["Usa", "Anomalia originale", "Controllo da eseguire", "Occ.", "Data", "Macchina", "Categoria", "Ticket"]
        )
        self.controls_table.verticalHeader().setVisible(False)
        self.controls_table.setWordWrap(True)
        self.controls_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        header_view = self.controls_table.horizontalHeader()
        header_view.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header_view.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header_view.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        for col in (3, 4, 5, 6, 7):
            header_view.setSectionResizeMode(col, QHeaderView.ResizeMode.ResizeToContents)
        controls_tab_layout.addWidget(self.controls_table, 1)

        buttons = QHBoxLayout()
        select_all = QPushButton("Seleziona tutti")
        select_all.clicked.connect(lambda: self._set_all_controls_checked(True))
        deselect_all = QPushButton("Deseleziona tutti")
        deselect_all.clicked.connect(lambda: self._set_all_controls_checked(False))
        self.btn_export_controls = QPushButton("Esporta controlli Excel")
        self.btn_export_controls.setEnabled(False)
        self.btn_export_controls.clicked.connect(self.export_controls)
        self.btn_send_controls = QPushButton("Invia alla Checklist")
        self.btn_send_controls.setEnabled(False)
        self.btn_send_controls.clicked.connect(self.send_controls_to_checklist)
        buttons.addWidget(select_all)
        buttons.addWidget(deselect_all)
        buttons.addStretch(1)
        buttons.addWidget(self.btn_export_controls)
        buttons.addWidget(self.btn_send_controls)
        controls_tab_layout.addLayout(buttons)
        self.tabs.addTab(controls_tab, "Controlli per Checklist")
        content_layout.addWidget(self.tabs, 1)

        status = QFrame()
        status_layout = QHBoxLayout(status)
        status_layout.setContentsMargins(8, 5, 8, 5)
        self.lbl_status = QLabel("Pronto.")
        self.pbar = QProgressBar()
        self.pbar.setTextVisible(False)
        self.pbar.setFixedHeight(10)
        status_layout.addWidget(self.lbl_status, 1)
        status_layout.addWidget(self.pbar)
        content_layout.addWidget(status)

        root.addWidget(self.sidebar)
        root.addWidget(content, 1)
        self.apply_stylesheet()

    @staticmethod
    def _section_label(text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("sectionTitle")
        return label

    @staticmethod
    def _separator() -> QFrame:
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setObjectName("separator")
        return line

    @staticmethod
    def _line_edit(text: str = "", placeholder: str = "") -> QLineEdit:
        edit = QLineEdit(text)
        edit.setPlaceholderText(placeholder)
        edit.setMinimumHeight(29)
        return edit

    @staticmethod
    def _spinbox(minimum: int, maximum: int, value: int, step: int = 1) -> QSpinBox:
        spin = QSpinBox()
        spin.setRange(minimum, maximum)
        spin.setValue(value)
        spin.setSingleStep(step)
        spin.setAlignment(Qt.AlignmentFlag.AlignCenter)
        spin.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        return spin

    def _toggle_sheet_field(self) -> None:
        self.txt_sheet.setEnabled(not self.chk_all_sheets.isChecked())

    def show_help(self) -> None:
        SuiteHelpDialog(self, initial_tab=1).exec()

    def select_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "Apri Excel", "", "Excel (*.xlsx *.xlsm *.xls)")
        if not files:
            return
        self.selected_files = files
        names = [os.path.basename(path) for path in files[:3]]
        extra = "" if len(files) <= 3 else f" (+{len(files) - 3})"
        self.lbl_files.setText(f"{len(files)} file: " + ", ".join(names) + extra)

    def _collect_params(self) -> AnalysisParams:
        return AnalysisParams(
            files=self.selected_files,
            keyword=self.txt_key.text().strip(),
            keyword_threshold=self.spin_keyword_threshold.value(),
            threshold=self.spin_thresh.value(),
            min_occurrences=self.spin_occ.value(),
            start_cell=self.txt_start.text().strip().upper(),
            target_col=self.txt_col.text().strip().upper(),
            date_col=self.txt_date.text().strip().upper(),
            sheet_name=self.txt_sheet.text().strip() or "Punti Aperti",
            analyze_all_sheets=self.chk_all_sheets.isChecked(),
            stemming=self.chk_stem.isChecked(),
            algorithm=str(self.combo_algo.currentData() or "combinato"),
            exhaustive_limit=self.spin_exhaustive.value(),
            max_bucket_size=self.spin_bucket.value(),
        )

    def start_analysis(self) -> None:
        if not self.selected_files:
            QMessageBox.warning(self, "File mancanti", "Seleziona almeno un file Excel.")
            return
        self.btn_analyze.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.btn_export.setEnabled(False)
        self.btn_export_controls.setEnabled(False)
        self.btn_send_controls.setEnabled(False)
        self.analysis_report = None
        self.txt_out.clear()
        self.controls_table.setRowCount(0)
        self._clear_chart()
        self.pbar.setValue(0)
        self.tabs.setCurrentIndex(0)

        try:
            params = self._collect_params()
        except Exception as exc:
            self._analysis_ended()
            QMessageBox.critical(self, "Parametri non validi", str(exc))
            return

        self.worker = AnalysisWorker(params)
        self.worker.log_message.connect(self._append_log)
        self.worker.progress_update.connect(self.pbar.setValue)
        self.worker.status_update.connect(self.lbl_status.setText)
        self.worker.analysis_finished.connect(self.on_finished)
        self.worker.error_occurred.connect(self.on_error)
        self.worker.start()

    def _append_log(self, message: str) -> None:
        self.txt_out.insertPlainText(message)
        self.txt_out.moveCursor(QTextCursor.MoveOperation.End)

    def stop_analysis(self) -> None:
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.btn_stop.setEnabled(False)
            self.lbl_status.setText("Arresto in corso...")

    def on_error(self, message: str) -> None:
        self._analysis_ended()
        self.lbl_status.setText("Errore.")
        QMessageBox.critical(self, "Errore analisi", message)

    def on_finished(self, report: AnalysisReport) -> None:
        self.analysis_report = report
        self._analysis_ended()
        if report.mode == "cancelled":
            self.lbl_status.setText("Analisi interrotta.")
            return
        self.lbl_status.setText("Completato.")
        self.btn_export.setEnabled(bool(report.has_results or report.metadata or report.warnings or report.errors))
        if report.mode == "no_keywords" and report.summary:
            self.generate_embedded_chart(report.summary)
            self.tabs.setTabEnabled(1, True)
        self._populate_controls_table(report)
        if self.controls_table.rowCount():
            self.btn_export_controls.setEnabled(True)
            self.btn_send_controls.setEnabled(True)
            self.tabs.setCurrentIndex(2)

    def _analysis_ended(self) -> None:
        self.btn_analyze.setEnabled(True)
        self.btn_stop.setEnabled(False)

    def _clear_chart(self) -> None:
        if self.canvas:
            self.layout_chart.removeWidget(self.canvas)
            self.canvas.deleteLater()
            self.canvas = None
        self.tabs.setTabEnabled(1, False)

    def generate_embedded_chart(self, summary: list[dict[str, object]]) -> None:
        self._clear_chart()
        top = sorted(summary, key=lambda row: int(row.get("NUMERO DI OCCORRENZE", 0)), reverse=True)[:20]
        if not top:
            return
        labels = [str(row.get("CONCETTO BREVE", "N/D")) for row in top]
        values = [int(row.get("NUMERO DI OCCORRENZE", 0)) for row in top]
        figure = Figure(figsize=(6.4, 3.8), dpi=100)
        axis = figure.add_subplot(111)
        positions = list(range(len(labels)))
        axis.bar(positions, values)
        axis.set_title("Primi 20 gruppi per occorrenze")
        axis.set_ylabel("Occorrenze")
        axis.set_xticks(positions)
        axis.set_xticklabels(["\n".join(textwrap.wrap(label, 16)) for label in labels], rotation=45, ha="right", fontsize=9)
        figure.tight_layout()
        self.canvas = FigureCanvasQTAgg(figure)
        self.layout_chart.addWidget(self.canvas)

    @staticmethod
    def _suggest_control(original: str) -> str:
        text = re.sub(r"\s+", " ", str(original or "")).strip()
        if not text:
            return ""
        lowered = text.casefold()
        action_starts = ("verific", "controll", "accert", "esegu", "misur", "prov", "impost", "registr")
        if lowered.startswith(action_starts):
            return text[0].upper() + text[1:]
        return f"Verificare: {text}"

    def _populate_controls_table(self, report: AnalysisReport) -> None:
        rows: list[tuple[str, int, str]] = []
        if report.mode == "no_keywords":
            for row in report.summary:
                rows.append((
                    str(row.get("CONCETTO PRINCIPALE DEL GRUPPO", "")),
                    int(row.get("NUMERO DI OCCORRENZE", 1) or 1),
                    str(row.get("DATA PIU RECENTE", "")),
                ))
        elif report.mode == "with_keywords":
            for row in report.summary:
                rows.append((str(row.get("TESTO TROVATO", "")), 1, str(row.get("DATA", ""))))

        self.controls_table.setRowCount(len(rows))
        for index, (original, occurrences, latest_date) in enumerate(rows):
            use_item = QTableWidgetItem("")
            use_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsSelectable)
            use_item.setCheckState(Qt.CheckState.Checked)
            self.controls_table.setItem(index, 0, use_item)

            original_item = QTableWidgetItem(original)
            original_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.controls_table.setItem(index, 1, original_item)
            self.controls_table.setItem(index, 2, QTableWidgetItem(self._suggest_control(original)))

            occ_item = QTableWidgetItem(str(occurrences))
            occ_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.controls_table.setItem(index, 3, occ_item)
            self.controls_table.setItem(index, 4, QTableWidgetItem(latest_date))
            self.controls_table.setItem(index, 5, QTableWidgetItem(""))
            self.controls_table.setItem(index, 6, QTableWidgetItem(""))
            self.controls_table.setItem(index, 7, QTableWidgetItem(""))
        self.controls_table.resizeRowsToContents()

    def _set_all_controls_checked(self, checked: bool) -> None:
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        for row in range(self.controls_table.rowCount()):
            item = self.controls_table.item(row, 0)
            if item:
                item.setCheckState(state)

    def selected_controls(self) -> list[ExternalControl]:
        controls: list[ExternalControl] = []
        for row in range(self.controls_table.rowCount()):
            use_item = self.controls_table.item(row, 0)
            if not use_item or use_item.checkState() != Qt.CheckState.Checked:
                continue
            def value(col: int) -> str:
                item = self.controls_table.item(row, col)
                return item.text().strip() if item else ""
            control_text = value(2)
            if not control_text:
                continue
            try:
                occurrences = max(1, int(value(3) or "1"))
            except ValueError:
                occurrences = 1
            controls.append(ExternalControl(
                control=control_text,
                original=value(1),
                occurrences=occurrences,
                latest_date=value(4),
                machine=value(5),
                category=value(6),
                ticket=value(7),
                source="Analyzer",
            ))
        return controls

    def export_analysis_report(self) -> None:
        if not self.analysis_report:
            return
        path, selected_filter = QFileDialog.getSaveFileName(
            self, "Salva report", str(self._downloads_dir() / "Report_Analisi"),
            "Excel (*.xlsx);;Word (*.docx);;Text (*.txt);;CSV (*.csv)",
        )
        if not path:
            return
        try:
            saved_paths = export_report(self.analysis_report, path, selected_filter)
            QMessageBox.information(self, "Export completato", "File salvati:\n" + "\n".join(saved_paths))
        except Exception as exc:
            QMessageBox.critical(self, "Errore export", str(exc))

    def export_controls(self) -> None:
        controls = self.selected_controls()
        if not controls:
            QMessageBox.warning(self, "Nessun controllo", "Seleziona almeno un controllo valido.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Esporta controlli per Checklist", str(self._downloads_dir() / "controlli_analyzer.xlsx"), "Excel (*.xlsx)"
        )
        if not path:
            return
        try:
            target = export_controls_xlsx(controls, path)
            QMessageBox.information(self, "Controlli esportati", f"File creato:\n{target}")
        except Exception as exc:
            QMessageBox.critical(self, "Errore esportazione", str(exc))

    def send_controls_to_checklist(self) -> None:
        controls = self.selected_controls()
        if not controls:
            QMessageBox.warning(self, "Nessun controllo", "Seleziona almeno un controllo valido.")
            return
        self.controls_ready.emit(controls)

    @staticmethod
    def _downloads_dir() -> Path:
        location = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DownloadLocation)
        return Path(location) if location else Path.home() / "Downloads"

    def closeEvent(self, event) -> None:  # noqa: N802
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.worker.wait(1500)
        event.accept()

    def apply_stylesheet(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow { background: #f5f6fa; }
            #sidebar, #sidebarControls, #sidebarScroll, #sidebarScroll > QWidget > QWidget { background: #2c3e50; border: none; }
            #sidebar QLabel { color: #ecf0f1; }
            #sidebarTitle { color: #67b7f7; font-size: 16px; font-weight: 800; }
            #sectionTitle { color: #9fd1ff; font-weight: 800; margin-top: 4px; }
            #fileLabel { color: #d0d7de; font-style: italic; }
            #separator { background: #4b6175; max-height: 1px; }
            #sidebar QLineEdit, #sidebar QSpinBox, #sidebar QComboBox {
                min-height: 28px; padding: 3px 6px; border: 1px solid #60758a; border-radius: 4px;
                background: #34495e; color: white;
            }
            #sidebar QCheckBox { color: white; }
            #sidebar QPushButton { min-height: 31px; background: #405a73; color: white; border: 0; border-radius: 5px; }
            #sidebar QPushButton:hover { background: #4d6d8b; }
            #btnStart { background: #268c4f; font-weight: 700; }
            #btnStop { background: #a63a32; font-weight: 700; }
            #btnExport { background: #2878a8; font-weight: 700; }
            #btnExit { background: #6d7478; font-weight: 700; }
            #btnMiniHelp { min-width: 28px; max-width: 28px; border-radius: 14px; background: #2878a8; font-weight: 800; }
            #contentTitle { background: white; border: 1px solid #e0e4e8; border-radius: 7px; padding: 9px; font-size: 16px; font-weight: 800; }
            QTextEdit, QTableWidget { background: white; border: 1px solid #d9dee3; border-radius: 6px; }
            QTabWidget::pane { border: 1px solid #d9dee3; background: white; }
            QTabBar::tab { padding: 7px 13px; background: #e8ebef; }
            QTabBar::tab:selected { background: white; font-weight: 700; border-bottom: 2px solid #2878a8; }
            """
        )
