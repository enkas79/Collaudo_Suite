"""GUI Windows per scaricare i documenti Giornale Macchina da JARVIS."""

from __future__ import annotations

import contextlib
import base64
import ctypes
import io
import os
import sys
import threading
from datetime import datetime
from pathlib import Path

# Permette l'avvio anche quando PyCharm usa una working directory diversa
# dalla cartella che contiene questo file.
SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from PySide6.QtCore import QObject, QSettings, QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from scarica_giornali_macchina import download_documents


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", ctypes.c_uint32), ("pbData", ctypes.POINTER(ctypes.c_ubyte))]


def _protect_text(value: str) -> str:
    """Protect a local secret with the current Windows user DPAPI key."""
    if not value or os.name != "nt":
        return ""
    raw = value.encode("utf-8")
    source = ctypes.create_string_buffer(raw)
    source_blob = _DataBlob(len(raw), ctypes.cast(source, ctypes.POINTER(ctypes.c_ubyte)))
    result_blob = _DataBlob()
    crypt32 = ctypes.windll.crypt32
    if not crypt32.CryptProtectData(ctypes.byref(source_blob), None, None, None, None, 0, ctypes.byref(result_blob)):
        return ""
    try:
        protected = ctypes.string_at(result_blob.pbData, result_blob.cbData)
        return base64.b64encode(protected).decode("ascii")
    finally:
        ctypes.windll.kernel32.LocalFree(result_blob.pbData)


def _unprotect_text(value: str) -> str:
    if not value or os.name != "nt":
        return ""
    try:
        raw = base64.b64decode(value.encode("ascii"), validate=True)
        source = ctypes.create_string_buffer(raw)
        source_blob = _DataBlob(len(raw), ctypes.cast(source, ctypes.POINTER(ctypes.c_ubyte)))
        result_blob = _DataBlob()
        crypt32 = ctypes.windll.crypt32
        if not crypt32.CryptUnprotectData(ctypes.byref(source_blob), None, None, None, None, 0, ctypes.byref(result_blob)):
            return ""
        try:
            return ctypes.string_at(result_blob.pbData, result_blob.cbData).decode("utf-8")
        finally:
            ctypes.windll.kernel32.LocalFree(result_blob.pbData)
    except (ValueError, OSError, UnicodeDecodeError):
        return ""


class _LogStream(io.TextIOBase):
    def __init__(self, callback, file_handle, *secrets: str) -> None:
        super().__init__()
        self.callback = callback
        self.file_handle = file_handle
        self.secrets = tuple(secret for secret in secrets if secret)
        self.pending = ""

    def _emit_line(self, line: str) -> None:
        for secret in self.secrets:
            line = line.replace(secret, "[SEGRETO OSCURATO]")
        if line.strip():
            self.callback.emit(line)
            timestamp = datetime.now().astimezone().isoformat(timespec="seconds")
            self.file_handle.write(f"{timestamp} {line}\n")
            self.file_handle.flush()

    def write(self, text: str) -> int:
        self.pending += str(text)
        while "\n" in self.pending:
            line, self.pending = self.pending.split("\n", 1)
            self._emit_line(line)
        return len(text)

    def flush(self) -> None:
        if self.pending.strip():
            self._emit_line(self.pending)
            self.pending = ""
        self.file_handle.flush()


class DownloadWorker(QObject):
    log = Signal(str)
    finished = Signal(int, int)
    stopped = Signal(int, int)
    failed = Signal(str)

    def __init__(
        self,
        token: str,
        output: Path,
        page_size: int,
        max_pages: int,
        dry_run: bool,
        organize_by_machine: bool,
        wbs: str,
        auth_cookie: str,
    ) -> None:
        super().__init__()
        self.token = token
        self.output = output
        self.page_size = page_size
        self.max_pages = max_pages
        self.dry_run = dry_run
        self.organize_by_machine = organize_by_machine
        self.wbs = wbs
        self.auth_cookie = auth_cookie
        self.stop_event = threading.Event()

    def request_stop(self) -> None:
        self.stop_event.set()

    def run(self) -> None:
        stream = None
        file_handle = None
        try:
            self.output.mkdir(parents=True, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_path = self.output / f"log_giornali_macchina_{timestamp}.txt"
            file_handle = log_path.open("w", encoding="utf-8", newline="\n")
            stream = _LogStream(self.log, file_handle, self.token, self.auth_cookie)
            stream.write(f"[LOG] File log: {log_path}\n")
            with contextlib.redirect_stdout(stream), contextlib.redirect_stderr(stream):
                found, downloaded = download_documents(
                    self.token,
                    self.output,
                    page_size=self.page_size,
                    max_pages=self.max_pages,
                    dry_run=self.dry_run,
                    organize_by_machine=self.organize_by_machine,
                    wbs=self.wbs,
                    auth_cookie=self.auth_cookie,
                    stop_event=self.stop_event,
                )
            stream.flush()
            stream.write(f"[FINE] {found} corrispondenze, {downloaded} file scaricati.\n")
            if self.stop_event.is_set():
                self.stopped.emit(found, downloaded)
            else:
                self.finished.emit(found, downloaded)
        except Exception as exc:
            if stream is not None:
                stream.write(f"[ERRORE] {exc}\n")
                stream.flush()
            self.failed.emit(str(exc))
        finally:
            if stream is not None:
                stream.flush()
            if file_handle is not None:
                file_handle.close()


class GiornaliMacchinaWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Downloader Giornali Macchina - JARVIS")
        self.resize(760, 560)
        self.settings = QSettings("Breton", "CollaudoSuiteGiornaliMacchina")
        self.thread: QThread | None = None
        self.worker: DownloadWorker | None = None

        central = QWidget(self)
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        form = QFormLayout()

        self.token_edit = QLineEdit()
        self.token_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.token_edit.setPlaceholderText("Token JARVIS per la ricerca Assets e il download")
        self.token_edit.setText(
            _unprotect_text(str(self.settings.value("jarvis_token_dpapi", "")))
            or os.environ.get("JARVIS_AUTH_TOKEN", "")
        )
        form.addRow("Token JARVIS:", self.token_edit)

        self.auth_cookie_edit = QLineEdit()
        self.auth_cookie_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.auth_cookie_edit.setPlaceholderText("Facoltativo: solo il valore del cookie AuthCookie")
        self.auth_cookie_edit.setText(
            _unprotect_text(str(self.settings.value("jarvis_cookie_dpapi", "")))
            or os.environ.get("JARVIS_AUTH_COOKIE", "")
        )
        form.addRow("Cookie sessione:", self.auth_cookie_edit)

        self.wbs_edit = QLineEdit()
        self.wbs_edit.setPlaceholderText("es. 86249 (lasciare vuoto per tutti gli Assets)")
        form.addRow("WBS / macchina:", self.wbs_edit)

        output_row = QHBoxLayout()
        self.output_edit = QLineEdit(str(Path.cwd() / "Giornali_Macchina"))
        self.output_edit.setReadOnly(True)
        browse = QPushButton("Scegli...")
        browse.clicked.connect(self.choose_output)
        output_row.addWidget(self.output_edit, 1)
        output_row.addWidget(browse)
        form.addRow("Cartella destinazione:", output_row)

        self.page_size_spin = QSpinBox()
        self.page_size_spin.setRange(1, 500)
        self.page_size_spin.setValue(100)
        form.addRow("Documenti per pagina:", self.page_size_spin)

        self.max_pages_spin = QSpinBox()
        self.max_pages_spin.setRange(1, 100000)
        self.max_pages_spin.setValue(100000)
        form.addRow("Numero massimo cartelle:", self.max_pages_spin)
        layout.addLayout(form)

        self.dry_run = QCheckBox("Simula la ricerca senza scaricare i file")
        layout.addWidget(self.dry_run)
        self.organize_by_machine = QCheckBox("Organizza i file per tipologia macchina (NC300, Trinity, Genya...)")
        self.organize_by_machine.setChecked(True)
        layout.addWidget(self.organize_by_machine)

        self.start_button = QPushButton("Cerca e scarica dagli Assets JARVIS")
        self.start_button.clicked.connect(self.start_download)
        layout.addWidget(self.start_button)

        self.stop_button = QPushButton("Stop")
        self.stop_button.setEnabled(False)
        self.stop_button.clicked.connect(self.stop_download)
        layout.addWidget(self.stop_button)

        layout.addWidget(QLabel("Log operazioni:"))
        self.log_edit = QPlainTextEdit()
        self.log_edit.setReadOnly(True)
        layout.addWidget(self.log_edit, 1)

        self.status_label = QLabel("Pronto. La ricerca viene eseguita esclusivamente nella sezione Assets JARVIS.")
        layout.addWidget(self.status_label)

    def choose_output(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "Scegli cartella di destinazione", self.output_edit.text())
        if selected:
            self.output_edit.setText(selected)

    def start_download(self) -> None:
        token = self.token_edit.text().strip()
        if not token:
            QMessageBox.warning(self, "Token mancante", "Inserire il token JARVIS.")
            return
        if self.thread and self.thread.isRunning():
            return

        token_secret = _protect_text(token)
        cookie_secret = _protect_text(self.auth_cookie_edit.text().strip())
        if token_secret:
            self.settings.setValue("jarvis_token_dpapi", token_secret)
        else:
            self.settings.remove("jarvis_token_dpapi")
        if cookie_secret:
            self.settings.setValue("jarvis_cookie_dpapi", cookie_secret)
        else:
            self.settings.remove("jarvis_cookie_dpapi")
        self.settings.sync()

        self.log_edit.clear()
        self.start_button.setEnabled(False)
        self.status_label.setText("Ricerca documenti in corso...")
        self.thread = QThread(self)
        self.worker = DownloadWorker(
            token,
            Path(self.output_edit.text()),
            self.page_size_spin.value(),
            self.max_pages_spin.value(),
            self.dry_run.isChecked(),
            self.organize_by_machine.isChecked(),
            self.wbs_edit.text().strip(),
            self.auth_cookie_edit.text().strip(),
        )
        self.worker.moveToThread(self.thread)
        self.worker.log.connect(self.log_edit.appendPlainText)
        self.worker.finished.connect(self.download_finished)
        self.worker.stopped.connect(self.download_stopped)
        self.worker.failed.connect(self.download_failed)
        self.worker.finished.connect(self.thread.quit)
        self.worker.stopped.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.started.connect(self.worker.run)
        self.thread.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread_finished)
        self.stop_button.setEnabled(True)
        self.thread.start()

    def stop_download(self) -> None:
        if self.worker is not None and self.thread is not None and self.thread.isRunning():
            self.stop_button.setEnabled(False)
            self.status_label.setText("Arresto richiesto: attendo la fine della richiesta corrente...")
            self.worker.request_stop()

    def download_finished(self, found: int, downloaded: int) -> None:
        self.status_label.setText(f"Completato: {found} corrispondenze, {downloaded} file scaricati.")
        if not self.dry_run.isChecked():
            QMessageBox.information(self, "Download completato", self.status_label.text())

    def download_stopped(self, found: int, downloaded: int) -> None:
        self.status_label.setText(f"Interrotto: {found} corrispondenze, {downloaded} file scaricati.")

    def download_failed(self, message: str) -> None:
        self.status_label.setText("Errore durante il download.")
        QMessageBox.critical(self, "Errore JARVIS", message)

    def thread_finished(self) -> None:
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.thread = None
        self.worker = None

    def closeEvent(self, event) -> None:  # noqa: N802 - Qt override
        if self.thread and self.thread.isRunning():
            QMessageBox.information(self, "Download in corso", "Attendere la fine del download prima di chiudere la finestra.")
            event.ignore()
            return
        event.accept()


def main() -> int:
    app = QApplication(sys.argv)
    window = GiornaliMacchinaWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
