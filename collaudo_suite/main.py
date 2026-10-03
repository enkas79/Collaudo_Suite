from __future__ import annotations

import sys
import traceback
from pathlib import Path

from PySide6.QtCore import QSize, Qt, QTimer, QUrl
from PySide6.QtGui import QAction, QDesktopServices, QIcon, QKeySequence
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QProgressDialog,
    QPushButton,
    QTabWidget,
    QToolBar,
    QVBoxLayout,
    QStyle,
    QWidget,
)

from .analyzer.app import AnalyzerWindow
from .app_info import APP_AUTHOR, APP_TITLE, RELEASES_PAGE_URL, get_app_version
from .checklist.app import ChecklistWindow
from .help_dialog import SuiteHelpDialog
from .styles import HOME_QSS, SUITE_QSS
from .updater import (
    UpdateCheckWorker,
    UpdateDownloadWorker,
    UpdateInfo,
    can_self_install,
    installer_download_path,
    launch_installer,
)

APP_VERSION = get_app_version()

# Attesa prima del controllo automatico all'avvio, per non competere con il
# caricamento iniziale della GUI e dei file interni.
STARTUP_UPDATE_CHECK_DELAY_MS = 2500


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

        self.setStyleSheet(HOME_QSS)


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

        # La toolbar resta sempre visibile; il menu File riusa le stesse azioni
        # (stesse scorciatoie) così i comandi sono raggiungibili da entrambi.
        self._build_command_toolbar()
        self._build_menu_bar()

        self.statusBar().showMessage("Pronto")
        self._apply_style()
        # Reapply module-specific styles after the suite stylesheet, which otherwise propagates to child windows.
        self.analyzer.apply_stylesheet()

        self._update_worker: UpdateCheckWorker | None = None
        self._update_check_running = False
        self._download_worker: UpdateDownloadWorker | None = None
        self._download_progress: QProgressDialog | None = None
        self._download_version = ""
        self._download_cancelled = False
        QTimer.singleShot(STARTUP_UPDATE_CHECK_DELAY_MS, lambda: self._check_for_updates(silent=True))

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

    def _build_menu_bar(self) -> None:
        menu_bar = self.menuBar()

        file_menu = menu_bar.addMenu("&File")
        file_menu.addActions(self.command_toolbar.actions())

        tools_menu = menu_bar.addMenu("&Strumenti")
        fixed_info_action = QAction("Info file interni", self)
        fixed_info_action.triggered.connect(self.checklist.show_fixed_info)
        tools_menu.addAction(fixed_info_action)

        help_menu = menu_bar.addMenu("&Aiuto")
        guide_action = QAction("Guida", self)
        guide_action.setShortcut(QKeySequence(QKeySequence.StandardKey.HelpContents))
        guide_action.triggered.connect(self.show_guide)
        help_menu.addAction(guide_action)

        self.update_action = QAction("Controlla aggiornamenti", self)
        self.update_action.triggered.connect(lambda: self._check_for_updates(silent=False))
        help_menu.addAction(self.update_action)

        help_menu.addSeparator()
        about_action = QAction(f"Informazioni su {APP_TITLE}", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def show_guide(self) -> None:
        SuiteHelpDialog(self).exec()

    def show_about(self) -> None:
        # Versione riletta a ogni apertura: riflette sempre il contenuto di version.txt.
        QMessageBox.about(
            self,
            f"Informazioni su {APP_TITLE}",
            f"<h3>{APP_TITLE}</h3>"
            f"<p>Versione <b>{get_app_version()}</b></p>"
            f"<p>Autore: {APP_AUTHOR}</p>"
            "<p>Analisi delle anomalie ricorrenti, preparazione dei controlli "
            "e checklist di collaudo.</p>"
            f'<p><a href="{RELEASES_PAGE_URL}">Pagina delle release</a></p>',
        )

    def _check_for_updates(self, *, silent: bool) -> None:
        if self._update_check_running:
            if not silent:
                self.statusBar().showMessage("Verifica aggiornamenti già in corso...", 4000)
            return

        self._update_silent = silent
        self._update_check_running = True
        worker = UpdateCheckWorker(APP_VERSION, self)
        worker.update_available.connect(self._on_update_available)
        worker.no_update.connect(self._on_no_update)
        worker.check_failed.connect(self._on_update_check_failed)
        worker.finished.connect(self._on_update_check_finished)
        self._update_worker = worker

        if not silent:
            self.statusBar().showMessage("Verifica aggiornamenti in corso...")
        worker.start()

    def _on_update_check_finished(self) -> None:
        # Azzera il riferimento prima di programmare la distruzione dell'oggetto Qt:
        # isRunning() su un QThread già cancellato da Shiboken solleva RuntimeError.
        self._update_check_running = False
        worker = self._update_worker
        self._update_worker = None
        if worker is not None:
            worker.deleteLater()

    def _on_update_available(self, info: UpdateInfo) -> None:
        self.statusBar().showMessage(f"Nuova versione disponibile: {info.version}", 10000)
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Information)
        box.setWindowTitle("Aggiornamento disponibile")
        box.setText(f"È disponibile la versione {info.version} (versione attuale: {APP_VERSION}).")
        if info.notes:
            box.setDetailedText(info.notes)
        install_button = None
        if can_self_install(info):
            install_button = box.addButton("Scarica e installa", QMessageBox.ButtonRole.AcceptRole)
        open_button = box.addButton("Apri pagina download", QMessageBox.ButtonRole.ActionRole)
        box.addButton("Più tardi", QMessageBox.ButtonRole.RejectRole)
        box.exec()
        clicked = box.clickedButton()
        if install_button is not None and clicked is install_button:
            self._start_update_download(info)
        elif clicked is open_button:
            target = info.release_url or info.download_url
            if target:
                QDesktopServices.openUrl(QUrl(target))

    def _start_update_download(self, info: UpdateInfo) -> None:
        if self._download_worker is not None:
            return
        destination = installer_download_path(info)
        worker = UpdateDownloadWorker(info.download_url, destination, self)
        progress = QProgressDialog(f"Download della versione {info.version} in corso...", "Annulla", 0, 0, self)
        progress.setWindowTitle("Aggiornamento")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)
        progress.setAutoClose(False)
        progress.setAutoReset(False)
        progress.canceled.connect(self._cancel_update_download)
        worker.progress.connect(self._on_update_download_progress)
        worker.download_finished.connect(self._on_update_downloaded)
        worker.download_failed.connect(self._on_update_download_failed)
        worker.finished.connect(self._on_update_download_thread_finished)
        self._download_worker = worker
        self._download_progress = progress
        self._download_version = info.version
        self._download_cancelled = False
        progress.show()
        worker.start()

    def _cancel_update_download(self) -> None:
        # Nota: QProgressDialog emette canceled anche quando viene chiuso da codice.
        if self._download_worker is not None and self._download_worker.isRunning():
            self._download_cancelled = True
            self._download_worker.requestInterruption()

    def _on_update_download_progress(self, received: int, total: int) -> None:
        progress = self._download_progress
        if progress is None:
            return
        if total > 0:
            # Valori in KB per restare nel range int di QProgressDialog anche con file grandi.
            progress.setMaximum(max(1, total // 1024))
            progress.setValue(min(received, total) // 1024)
            progress.setLabelText(
                f"Download della versione {self._download_version}: "
                f"{received / 1_048_576:.1f} / {total / 1_048_576:.1f} MB"
            )
        else:
            progress.setLabelText(f"Download della versione {self._download_version}: {received / 1_048_576:.1f} MB")

    def _close_download_progress(self) -> None:
        if self._download_progress is not None:
            self._download_progress.close()
            self._download_progress.deleteLater()
            self._download_progress = None

    def _on_update_downloaded(self, path: str) -> None:
        self._close_download_progress()
        answer = QMessageBox.question(
            self,
            "Installa aggiornamento",
            f"Download completato. Installare ora la versione {self._download_version}?\n\n"
            "Il lavoro corrente verrà salvato automaticamente e l'applicazione verrà chiusa "
            "per permettere l'installazione.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if answer != QMessageBox.StandardButton.Yes:
            self.statusBar().showMessage(f"Installer salvato in: {path}", 10000)
            return
        try:
            launch_installer(Path(path))
        except OSError as exc:
            QMessageBox.critical(
                self,
                "Aggiornamento",
                f"Impossibile avviare l'installer.\n\nFile: {path}\nDettaglio: {exc}",
            )
            return
        # close() passa da closeEvent: autosalvataggio e chiusura ordinata dei thread.
        self.close()

    def _on_update_download_failed(self, message: str) -> None:
        cancelled = self._download_cancelled
        self._close_download_progress()
        if cancelled:
            self.statusBar().showMessage("Download dell'aggiornamento annullato.", 6000)
            return
        QMessageBox.warning(
            self,
            "Aggiornamento",
            f"Download dell'aggiornamento non riuscito.\n\nDettaglio: {message}",
        )

    def _on_update_download_thread_finished(self) -> None:
        worker = self._download_worker
        self._download_worker = None
        if worker is not None:
            worker.deleteLater()

    def _on_no_update(self) -> None:
        if not self._update_silent:
            self.statusBar().showMessage("Nessun aggiornamento disponibile.", 6000)
            QMessageBox.information(self, "Aggiornamenti", "Stai già usando l'ultima versione disponibile.")

    def _on_update_check_failed(self, message: str) -> None:
        if not self._update_silent:
            self.statusBar().showMessage("Verifica aggiornamenti non riuscita.", 6000)
            QMessageBox.warning(
                self,
                "Verifica aggiornamenti",
                "Non è stato possibile verificare la presenza di aggiornamenti.\n"
                "Controlla la connessione a Internet e riprova.\n\n"
                f"Dettagli: {message}",
            )

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
        self.setStyleSheet(SUITE_QSS)

    def closeEvent(self, event) -> None:  # noqa: N802
        # ChecklistWindow is embedded as a tab, so its own closeEvent is not
        # guaranteed to run when the suite closes. Trigger autosave explicitly.
        self.checklist.autosave_on_exit()
        self.checklist._save_settings()
        self.checklist.wait_for_background_tasks()
        if self.analyzer.worker and self.analyzer.worker.isRunning():
            self.analyzer.worker.stop()
            self.analyzer.worker.wait(1500)
        if self._update_worker is not None and self._update_worker.isRunning():
            self._update_worker.wait(4000)
        if self._download_worker is not None and self._download_worker.isRunning():
            self._download_worker.requestInterruption()
            self._download_worker.wait(4000)
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
