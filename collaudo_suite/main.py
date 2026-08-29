from __future__ import annotations

import sys
import traceback

from PySide6.QtCore import QSize, Qt, QTimer
from PySide6.QtGui import QAction, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QToolBar,
    QVBoxLayout,
    QStyle,
    QWidget,
)

from .analyzer.app import AnalyzerWindow
from .checklist.app import ChecklistWindow
from .help_dialog import SuiteHelpDialog

APP_TITLE = "Collaudo Suite"
APP_VERSION = "1.1.7"


class HomePage(QWidget):
    def __init__(self, open_analyzer, open_checklist) -> None:
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 36, 40, 36)
        layout.setSpacing(18)

        title = QLabel("Collaudo Suite")
        title.setObjectName("homeTitle")
        subtitle = QLabel(
            "Un unico programma per analizzare anomalie ricorrenti, trasformarle in controlli operativi "
            "e inserirle nella checklist di collaudo."
        )
        subtitle.setWordWrap(True)
        subtitle.setObjectName("homeSubtitle")
        layout.addWidget(title)
        layout.addWidget(subtitle)

        workflow = QFrame()
        workflow.setObjectName("workflowCard")
        workflow_layout = QVBoxLayout(workflow)
        workflow_layout.setContentsMargins(22, 20, 22, 20)
        workflow_layout.setSpacing(12)
        workflow_layout.addWidget(QLabel("1. ANALYZER — carica i file anomalie e raggruppa i problemi ricorrenti."))
        workflow_layout.addWidget(QLabel("2. CONTROLLI — verifica e modifica il testo operativo proposto."))
        workflow_layout.addWidget(QLabel("3. CHECKLIST — trasferisci direttamente i controlli oppure importali da Excel."))
        workflow_layout.addWidget(QLabel("4. OUTPUT — compila Pass/No pass, salva il lavoro ed esporta il PDF."))
        layout.addWidget(workflow)

        buttons = QHBoxLayout()
        analyzer_btn = QPushButton("Apri Analyzer")
        analyzer_btn.setObjectName("primaryButton")
        analyzer_btn.clicked.connect(open_analyzer)
        checklist_btn = QPushButton("Apri Checklist")
        checklist_btn.setObjectName("secondaryButton")
        checklist_btn.clicked.connect(open_checklist)
        buttons.addWidget(analyzer_btn)
        buttons.addWidget(checklist_btn)
        buttons.addStretch(1)
        layout.addLayout(buttons)
        layout.addStretch(1)

        version = QLabel(f"Versione {APP_VERSION}")
        version.setAlignment(Qt.AlignmentFlag.AlignRight)
        version.setObjectName("versionLabel")
        layout.addWidget(version)

        self.setStyleSheet(
            """
            #homeTitle { font-size: 30px; font-weight: 800; color: #22313f; }
            #homeSubtitle { font-size: 15px; color: #52616b; }
            #workflowCard { background: white; border: 1px solid #dfe5ea; border-radius: 10px; }
            #workflowCard QLabel { font-size: 13px; color: #2d3e4f; }
            #primaryButton, #secondaryButton { min-height: 42px; padding: 0 20px; border-radius: 7px; font-weight: 700; }
            #primaryButton { background: #2878a8; color: white; border: none; }
            #primaryButton:hover { background: #3490c4; }
            #secondaryButton { background: #2f855a; color: white; border: none; }
            #secondaryButton:hover { background: #38a169; }
            #versionLabel { color: #7a8793; }
            """
        )


class SuiteMainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"{APP_TITLE} {APP_VERSION}")
        self.resize(1500, 900)
        self.setMinimumSize(1100, 700)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tabs.setMovable(False)
        self.setCentralWidget(self.tabs)

        self.analyzer = AnalyzerWindow()
        self.checklist = ChecklistWindow()
        self.home = HomePage(
            open_analyzer=lambda: self.tabs.setCurrentIndex(1),
            open_checklist=lambda: self.tabs.setCurrentIndex(2),
        )

        self.tabs.addTab(self.home, "Home")
        self.tabs.addTab(self.analyzer, "Analyzer anomalie")
        self.tabs.addTab(self.checklist, "Checklist")
        self.tabs.currentChanged.connect(self._on_tab_changed)

        self.analyzer.controls_ready.connect(self._receive_controls)

        # Comandi principali sempre visibili nella barra superiore.
        # Non vengono inseriti in un menu a tendina, così restano accessibili
        # indipendentemente dalla scheda attiva.
        guide_action = QAction("Guida", self)
        guide_action.triggered.connect(lambda: SuiteHelpDialog(self).exec())
        self.menuBar().addAction(guide_action)

        fixed_info_action = QAction("Info file interni", self)
        fixed_info_action.triggered.connect(self.checklist.show_fixed_info)
        self.menuBar().addAction(fixed_info_action)

        self._build_command_toolbar()

        self.statusBar().showMessage("Pronto")
        self._apply_style()
        # Reapply module-specific styles after the suite stylesheet, which otherwise propagates to child windows.
        self.analyzer.apply_stylesheet()

    def _toolbar_icon(self, theme_name: str, fallback: QStyle.StandardPixmap) -> QIcon:
        icon = QIcon.fromTheme(theme_name)
        if icon.isNull():
            icon = self.style().standardIcon(fallback)
        return icon

    def _add_toolbar_action(
        self,
        toolbar: QToolBar,
        *,
        text: str,
        theme_icon: str,
        fallback_icon: QStyle.StandardPixmap,
        callback,
        shortcut: str = "",
    ) -> QAction:
        action = QAction(self._toolbar_icon(theme_icon, fallback_icon), text, self)
        action.setToolTip(text)
        action.setStatusTip(text)
        if shortcut:
            action.setShortcut(shortcut)
        action.triggered.connect(callback)
        toolbar.addAction(action)
        return action

    def _build_command_toolbar(self) -> None:
        toolbar = QToolBar("Comandi Checklist", self)
        toolbar.setObjectName("checklistCommandToolbar")
        toolbar.setMovable(False)
        toolbar.setFloatable(False)
        toolbar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonIconOnly)
        toolbar.setIconSize(QSize(24, 24))
        self.addToolBar(Qt.ToolBarArea.TopToolBarArea, toolbar)

        self._add_toolbar_action(
            toolbar,
            text="Esporta PDF",
            theme_icon="document-export",
            fallback_icon=QStyle.StandardPixmap.SP_DialogSaveButton,
            callback=self.checklist.save_pdf,
        )
        self._add_toolbar_action(
            toolbar,
            text="Stampa PDF",
            theme_icon="document-print",
            fallback_icon=QStyle.StandardPixmap.SP_FileIcon,
            callback=self.checklist.print_pdf,
            shortcut="Ctrl+P",
        )
        toolbar.addSeparator()
        self._add_toolbar_action(
            toolbar,
            text="Salva lavoro",
            theme_icon="document-save",
            fallback_icon=QStyle.StandardPixmap.SP_DialogSaveButton,
            callback=self.checklist.save_work,
            shortcut="Ctrl+S",
        )
        self._add_toolbar_action(
            toolbar,
            text="Salva lavoro con nome",
            theme_icon="document-save-as",
            fallback_icon=QStyle.StandardPixmap.SP_DialogSaveButton,
            callback=self.checklist.save_work_as,
            shortcut="Ctrl+Shift+S",
        )
        self._add_toolbar_action(
            toolbar,
            text="Apri lavoro",
            theme_icon="document-open",
            fallback_icon=QStyle.StandardPixmap.SP_DialogOpenButton,
            callback=self.checklist.load_work,
            shortcut="Ctrl+O",
        )
        toolbar.addSeparator()
        self._add_toolbar_action(
            toolbar,
            text="Esci",
            theme_icon="application-exit",
            fallback_icon=QStyle.StandardPixmap.SP_DialogCloseButton,
            callback=QApplication.instance().quit,
            shortcut="Ctrl+Q",
        )
        self.command_toolbar = toolbar

    def _on_tab_changed(self, index: int) -> None:
        if self.tabs.widget(index) is self.checklist:
            QTimer.singleShot(0, self.checklist.refresh_layout)

    def _receive_controls(self, controls) -> None:
        added = self.checklist.add_external_controls(list(controls), replace=False)
        self.tabs.setCurrentIndex(2)
        self.statusBar().showMessage(f"Trasferiti {added} nuovi controlli alla Checklist", 8000)
        QMessageBox.information(
            self,
            "Trasferimento completato",
            f"Controlli selezionati: {len(controls)}\n"
            f"Nuovi controlli aggiunti alla Checklist: {added}\n\n"
            "I duplicati sono stati ignorati. I controlli sono già visibili se la checklist era stata generata; "
            "in caso contrario premi 'Estrai / Aggiorna anteprima'.",
        )

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow { background: #f4f6f8; }
            QTabWidget::pane { border: 0; }
            QTabBar::tab {
                min-width: 150px; padding: 9px 16px; background: #e7ebef; color: #334;
                border: 1px solid #d0d6dc; border-bottom: none;
            }
            QTabBar::tab:selected { background: white; font-weight: 800; border-top: 3px solid #2878a8; }
            QStatusBar { background: white; border-top: 1px solid #dfe5ea; }
            """
        )

    def closeEvent(self, event) -> None:  # noqa: N802
        # ChecklistWindow is embedded as a tab, so its own closeEvent is not
        # guaranteed to run when the suite closes. Trigger autosave explicitly.
        self.checklist.autosave_on_exit()
        self.checklist._save_settings()
        if self.analyzer.worker and self.analyzer.worker.isRunning():
            self.analyzer.worker.stop()
            self.analyzer.worker.wait(1500)
        event.accept()


def exception_hook(exctype, value, tb) -> None:
    details = "".join(traceback.format_exception(exctype, value, tb))
    print(details, file=sys.stderr)
    app = QApplication.instance()
    if app:
        QMessageBox.critical(None, "Errore non gestito", f"{value}\n\nDettagli tecnici:\n{details}")


def main() -> int:
    sys.excepthook = exception_hook
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setOrganizationName("CollaudoTools")
    app.setStyle("Fusion")
    window = SuiteMainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
