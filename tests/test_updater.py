import os
import tempfile
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

try:
    from PySide6.QtCore import QCoreApplication, QEventLoop, QTimer
except ImportError:  # PySide6 non installato nell'ambiente di test
    QCoreApplication = None

if QCoreApplication is not None:
    from collaudo_suite.updater import (
        UpdateDownloadWorker,
        UpdateInfo,
        can_self_install,
        installer_download_path,
        is_newer_version,
    )


def _info(download_url: str) -> "UpdateInfo":
    return UpdateInfo(version="1.2.0", notes="", download_url=download_url, release_url="https://example.invalid/r")


@unittest.skipIf(QCoreApplication is None, "PySide6 non disponibile")
class UpdaterLogicTests(unittest.TestCase):
    def test_version_comparison(self):
        self.assertTrue(is_newer_version("v1.2.0", "1.1.13"))
        self.assertTrue(is_newer_version("1.10.0", "1.9.9"))
        self.assertFalse(is_newer_version("1.1.13", "1.1.13"))
        self.assertFalse(is_newer_version("1.1.9", "1.1.10"))

    def test_self_install_only_for_windows_exe(self):
        self.assertTrue(can_self_install(_info("https://x/CollaudoSuite-Setup-1.2.0.exe"), platform="win32"))
        self.assertFalse(can_self_install(_info("https://x/CollaudoSuite-Setup-1.2.0.exe"), platform="linux"))
        self.assertFalse(can_self_install(_info("https://x/CollaudoSuite-1.2.0.zip"), platform="win32"))
        self.assertFalse(can_self_install(_info(""), platform="win32"))

    def test_download_path_stays_in_target_directory(self):
        base = Path(tempfile.gettempdir()) / "cs-test"
        path = installer_download_path(_info("https://x/dl/CollaudoSuite-Setup-1.2.0.exe"), base)
        self.assertEqual(path, base / "CollaudoSuite-Setup-1.2.0.exe")
        fallback = installer_download_path(_info("https://x/"), base)
        self.assertEqual(fallback.parent, base)


@unittest.skipIf(QCoreApplication is None, "PySide6 non disponibile")
class UpdateDownloadWorkerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QCoreApplication.instance() or QCoreApplication([])

    def _run(self, worker: "UpdateDownloadWorker") -> dict:
        result: dict = {}
        loop = QEventLoop()
        worker.download_finished.connect(lambda path: result.update(path=path))
        worker.download_failed.connect(lambda message: result.update(error=message))
        worker.finished.connect(loop.quit)
        QTimer.singleShot(10000, loop.quit)
        worker.start()
        loop.exec()
        worker.wait(2000)
        return result

    def test_downloads_file_with_progress(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.exe"
            payload = os.urandom(700_000)
            source.write_bytes(payload)
            destination = Path(tmp) / "out" / "CollaudoSuite-Setup-1.2.0.exe"
            worker = UpdateDownloadWorker(source.as_uri(), destination)
            progress: list[tuple[int, int]] = []
            worker.progress.connect(lambda received, total: progress.append((received, total)))
            result = self._run(worker)

            self.assertNotIn("error", result)
            self.assertEqual(Path(result["path"]), destination)
            self.assertEqual(destination.read_bytes(), payload)
            self.assertFalse(destination.with_name(destination.name + ".part").exists())
            self.assertTrue(progress)
            self.assertEqual(progress[-1], (len(payload), len(payload)))

    def test_failed_download_leaves_no_partial_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            destination = Path(tmp) / "setup.exe"
            worker = UpdateDownloadWorker((Path(tmp) / "missing.exe").as_uri(), destination)
            result = self._run(worker)

            self.assertIn("error", result)
            self.assertFalse(destination.exists())
            self.assertFalse(destination.with_name("setup.exe.part").exists())


if __name__ == "__main__":
    unittest.main()
