import os
import tempfile
import unittest
from pathlib import Path
from unittest import mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

try:
    from PySide6.QtCore import QEventLoop, QTimer
    from PySide6.QtWidgets import QApplication, QStyledItemDelegate
except ImportError:  # PySide6 (o le librerie grafiche di sistema) non disponibili
    QApplication = None

if QApplication is not None:
    from collaudo_suite.main import SuiteMainWindow
    from collaudo_suite.pdf_viewer import GuidePdfDialog, guide_pdf_path


@unittest.skipIf(QApplication is None, "PySide6 non disponibile")
class SuiteGuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        # Nessuna chiamata di rete durante i test.
        patcher = mock.patch.object(SuiteMainWindow, "_check_for_updates")
        patcher.start()
        self.addCleanup(patcher.stop)
        self.window = SuiteMainWindow()
        self.addCleanup(self.window.deleteLater)

    def _menu_titles(self, title: str) -> list[str]:
        for action in self.window.menuBar().actions():
            if action.text().replace("&", "") == title:
                return [a.text() for a in action.menu().actions() if not a.isSeparator()]
        self.fail(f"Menu {title!r} assente")

    def test_help_menu_has_guide_updates_and_about(self):
        entries = self._menu_titles("Aiuto")
        self.assertIn("Guida", entries)
        self.assertIn("Controlla aggiornamenti", entries)
        self.assertTrue(any(entry.startswith("Informazioni su") for entry in entries))

    def test_file_menu_mirrors_toolbar_commands(self):
        entries = self._menu_titles("File")
        for command in ("Esporta PDF", "Stampa PDF", "Salva lavoro", "Apri lavoro", "Esci"):
            self.assertIn(command, entries)

    def test_about_dialog_shows_version_and_author(self):
        with mock.patch("collaudo_suite.main.QMessageBox.about") as about:
            self.window.show_about()
        text = about.call_args.args[2]
        self.assertIn("Versione", text)
        self.assertIn("Autore", text)

    def test_about_dialog_shows_author_name(self):
        with mock.patch("collaudo_suite.main.QMessageBox.about") as about:
            self.window.show_about()
        self.assertIn("Enrico Martini", about.call_args.args[2])

    def test_app_icon_is_bundled_and_valid(self):
        from PySide6.QtGui import QIcon

        from collaudo_suite.app_info import app_icon_path

        self.assertTrue(app_icon_path().is_file())
        self.assertTrue(app_icon_path().with_suffix(".ico").is_file())
        self.assertFalse(QIcon(str(app_icon_path())).isNull())

    def test_guide_pdf_is_bundled_and_loads(self):
        self.assertTrue(guide_pdf_path().is_file())
        dialog = GuidePdfDialog(self.window)
        self.addCleanup(dialog.deleteLater)
        self.assertGreater(dialog.document.pageCount(), 0)

    def test_sidebar_combo_popups_use_styled_delegate(self):
        # Regressione: con il delegate "menu" di Fusion le opzioni erano bianche su fondo chiaro.
        for combo in (self.window.analyzer.combo_period, self.window.analyzer.combo_algo):
            self.assertIsInstance(combo.itemDelegate(), QStyledItemDelegate)

    def test_pdf_export_runs_in_background_thread(self):
        checklist = self.window.checklist
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "checklist.pdf"
            loop = QEventLoop()
            opened: list[str] = []
            with mock.patch.object(checklist, "_open_pdf", side_effect=lambda path: opened.append(str(path))), \
                    mock.patch("collaudo_suite.checklist.app.QMessageBox") as box:
                checklist._start_pdf_export(target, for_print=False)
                self.assertIsNotNone(checklist._pdf_thread)
                checklist._pdf_thread.finished.connect(loop.quit)
                QTimer.singleShot(20000, loop.quit)
                loop.exec()
                self.app.processEvents()

            box.critical.assert_not_called()
            self.assertTrue(target.exists())
            self.assertEqual(opened, [str(target)])
            self.assertIsNone(checklist._pdf_thread)


if __name__ == "__main__":
    unittest.main()
