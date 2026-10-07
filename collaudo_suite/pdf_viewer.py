from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QHBoxLayout, QLabel, QMessageBox, QPushButton, QVBoxLayout
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView


GUIDE_FILENAME = "Guida_operativa_Collaudo_Suite.pdf"


def guide_pdf_path() -> Path:
    """Return the guide location in source runs and PyInstaller builds."""
    return Path(__file__).resolve().parent / "assets" / GUIDE_FILENAME


class GuidePdfDialog(QDialog):
    """In-app PDF guide with page scrolling and zoom controls."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Guida operativa - Collaudo Suite")
        self.resize(1120, 820)
        self.setMinimumSize(760, 560)

        layout = QVBoxLayout(self)
        toolbar = QHBoxLayout()
        self.page_label = QLabel("Caricamento guida...")
        toolbar.addWidget(self.page_label)
        toolbar.addStretch(1)

        self.zoom_out_button = QPushButton("-")
        self.zoom_out_button.setToolTip("Riduci zoom")
        self.zoom_out_button.clicked.connect(lambda: self._zoom(0.85))
        toolbar.addWidget(self.zoom_out_button)

        self.zoom_label = QLabel("Adatta larghezza")
        self.zoom_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        toolbar.addWidget(self.zoom_label)

        self.zoom_in_button = QPushButton("+")
        self.zoom_in_button.setToolTip("Aumenta zoom")
        self.zoom_in_button.clicked.connect(lambda: self._zoom(1.18))
        toolbar.addWidget(self.zoom_in_button)

        self.fit_button = QPushButton("Adatta larghezza")
        self.fit_button.clicked.connect(self._fit_width)
        toolbar.addWidget(self.fit_button)
        layout.addLayout(toolbar)

        self.document = QPdfDocument(self)
        self.document.statusChanged.connect(self._update_document_status)
        self.view = QPdfView(self)
        self.view.setDocument(self.document)
        self.view.setPageMode(QPdfView.PageMode.MultiPage)
        self.view.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        layout.addWidget(self.view, 1)

        path = guide_pdf_path()
        if not path.is_file():
            self.page_label.setText("Guida PDF non trovata")
            self.view.setEnabled(False)
            QMessageBox.warning(
                self,
                "Guida non disponibile",
                f"Non trovo il PDF della guida:\n{path}\n\n"
                "Rigenera o reinstalla Collaudo Suite includendo l'asset PDF della guida.",
            )
            return

        self.document.load(str(path))
        self._update_document_status()

    def _update_document_status(self) -> None:
        status = str(self.document.status()).rsplit(".", 1)[-1]
        if status == "Ready":
            self.page_label.setText(f"{self.document.pageCount()} pagine - scorri per navigare")
        elif status == "Error":
            self.page_label.setText(f"Errore nel caricamento del PDF ({self.document.error()})")

    def _zoom(self, factor: float) -> None:
        self.view.setZoomMode(QPdfView.ZoomMode.Custom)
        self.view.setZoomFactor(max(0.5, min(3.0, self.view.zoomFactor() * factor)))
        self.zoom_label.setText(f"{round(self.view.zoomFactor() * 100)}%")

    def _fit_width(self) -> None:
        self.view.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        self.zoom_label.setText("Adatta larghezza")


def show_guide_pdf(parent=None) -> None:
    dialog = GuidePdfDialog(parent)
    dialog.exec()
