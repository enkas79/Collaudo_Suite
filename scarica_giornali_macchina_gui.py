"""GUI Windows per scaricare i documenti Giornale Macchina da JARVIS."""

from __future__ import annotations

import contextlib
import io
import sys
from datetime import datetime
from pathlib import Path

# Permette l'avvio anche quando PyCharm usa una working directory diversa
# dalla cartella che contiene questo file.
SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from PySide6.QtCore import QObject, QThread, Signal
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


class _LogStream(io.TextIOBase):
    def __init__(self, callback, file_handle, secret: str) -> None:
        super().__init__()
        self.callback = callback
        self.file_handle = file_handle
        self.secret = secret
        self.pending = ""

    def _emit_line(self, line: str) -> None:
        if self.secret:
            line = line.replace(self.secret, "[TOKEN OSCURATO]")
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
    failed = Signal(str)

    def __init__(
        self,
        token: str,
        output: Path,
        page_size: int,
        max_pages: int,
        dry_run: bool,
        organize_by_machine: bool,
    ) -> None:
        super().__init__()
        self.token = token
        self.output = output
        self.page_size = page_size
        self.max_pages = max_pages
        self.dry_run = dry_run
        self.organize_by_machine = organize_by_machine

    def run(self) -> None:
        stream = None
        file_handle = None
        try:
            self.output.mkdir(parents=True, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_path = self.output / f"log_giornali_macchina_{timestamp}.txt"
            file_handle = log_path.open("w", encoding="utf-8", newline="\n")
            stream = _LogStream(self.log, file_handle, self.token)
            stream.write(f"[LOG] File log: {log_path}\n")
            with contextlib.redirect_stdout(stream), contextlib.redirect_stderr(stream):
                found, downloaded = download_documents(
                    self.token,
                    self.output,
                    page_size=self.page_size,
                    max_pages=self.max_pages,
                    dry_run=self.dry_run,
                    organize_by_machine=self.organize_by_machine,
                )
            stream.flush()
            stream.write(f"[FINE] {found} corrispondenze, {downloaded} file scaricati.\n")
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
        self.thread: QThread | None = None
        self.worker: DownloadWorker | None = None

        central = QWidget(self)
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        form = QFormLayout()

        self.token_edit = QLineEdit()
        self.token_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.token_edit.setPlaceholderText("Token con permesso documents-read / download")
        form.addRow("Token JARVIS:", self.token_edit)

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

        self.start_button = QPushButton("Cerca e scarica dalla directory JARVIS")
        self.start_button.clicked.connect(self.start_download)
        layout.addWidget(self.start_button)

        layout.addWidget(QLabel("Log operazioni:"))
        self.log_edit = QPlainTextEdit()
        self.log_edit.setReadOnly(True)
        layout.addWidget(self.log_edit, 1)

        self.status_label = QLabel("Pronto. La ricerca viene eseguita nella directory Documentale JARVIS.")
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
        )
        self.worker.moveToThread(self.thread)
        self.worker.log.connect(self.log_edit.appendPlainText)
        self.worker.finished.connect(self.download_finished)
        self.worker.failed.connect(self.download_failed)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.started.connect(self.worker.run)
        self.thread.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread_finished)
        self.thread.start()

    def download_finished(self, found: int, downloaded: int) -> None:
        self.status_label.setText(f"Completato: {found} corrispondenze, {downloaded} file scaricati.")
        if not self.dry_run.isChecked():
            QMessageBox.information(self, "Download completato", self.status_label.text())

    def download_failed(self, message: str) -> None:
        self.status_label.setText("Errore durante il download.")
        QMessageBox.critical(self, "Errore JARVIS", message)

    def thread_finished(self) -> None:
        self.start_button.setEnabled(True)
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
