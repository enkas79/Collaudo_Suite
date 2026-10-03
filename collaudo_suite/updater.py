from __future__ import annotations

import json
import os
import sys
import tempfile
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import QThread, Signal

from .app_info import GITHUB_OWNER, GITHUB_REPO

RELEASES_API_URL = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
REQUEST_TIMEOUT = 6
DOWNLOAD_TIMEOUT = 30
DOWNLOAD_CHUNK_SIZE = 256 * 1024
USER_AGENT = "CollaudoSuite-Updater"


@dataclass(frozen=True)
class UpdateInfo:
    version: str
    notes: str
    download_url: str
    release_url: str


def _parse_version(text: str) -> tuple[int, ...]:
    text = text.strip().lstrip("vV")
    parts: list[int] = []
    for chunk in text.split("."):
        digits = "".join(ch for ch in chunk if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts) or (0,)


def is_newer_version(remote: str, local: str) -> bool:
    return _parse_version(remote) > _parse_version(local)


def _pick_asset_url(assets: list[dict]) -> str:
    for asset in assets:
        name = str(asset.get("name", "")).lower()
        if name.endswith((".exe", ".msi", ".zip")):
            return str(asset.get("browser_download_url", ""))
    return ""


def fetch_latest_release() -> UpdateInfo | None:
    """Interroga le GitHub Releases per l'ultima versione pubblicata.

    Nessuna eccezione di rete viene soppressa qui: il chiamante (il worker in
    background) decide come trattare l'assenza di connessione, tipica sui PC
    di reparto senza accesso a Internet.
    """
    request = urllib.request.Request(
        RELEASES_API_URL,
        headers={"Accept": "application/vnd.github+json", "User-Agent": USER_AGENT},
    )
    with urllib.request.urlopen(request, timeout=REQUEST_TIMEOUT) as response:
        payload = json.loads(response.read().decode("utf-8"))

    tag = str(payload.get("tag_name", "")).strip()
    if not tag:
        return None

    assets = payload.get("assets", []) or []
    return UpdateInfo(
        version=tag.lstrip("vV"),
        notes=str(payload.get("body", "")).strip(),
        download_url=_pick_asset_url(assets),
        release_url=str(payload.get("html_url", "")),
    )


def can_self_install(info: UpdateInfo, platform: str | None = None) -> bool:
    """True se l'aggiornamento può essere scaricato e avviato direttamente dall'app.

    Serve un installer Windows (.exe): negli altri casi resta solo l'apertura
    della pagina di download.
    """
    platform = sys.platform if platform is None else platform
    return platform == "win32" and info.download_url.lower().endswith(".exe")


def installer_download_path(info: UpdateInfo, directory: Path | None = None) -> Path:
    """Percorso locale (cartella temporanea) in cui salvare l'installer scaricato."""
    name = info.download_url.rstrip("/").rsplit("/", 1)[-1] or f"CollaudoSuite-Setup-{info.version}.exe"
    # Evita che un nome remoto anomalo esca dalla cartella di destinazione.
    name = Path(name).name
    base = directory if directory is not None else Path(tempfile.gettempdir()) / "CollaudoSuite-update"
    return base / name


def launch_installer(path: Path) -> None:
    """Avvia l'installer scaricato tramite la shell di Windows (gestisce il prompt UAC)."""
    if sys.platform != "win32":
        raise OSError("L'installazione automatica è supportata solo su Windows.")
    os.startfile(str(path))  # type: ignore[attr-defined]  # disponibile solo su Windows


class UpdateDownloadWorker(QThread):
    """Scarica l'installer della nuova versione in background, con avanzamento."""

    progress = Signal(int, int)  # byte ricevuti, byte totali (0 se sconosciuti)
    download_finished = Signal(str)  # percorso del file scaricato
    download_failed = Signal(str)

    def __init__(self, url: str, destination: Path, parent=None) -> None:
        super().__init__(parent)
        self._url = url
        self._destination = destination

    def run(self) -> None:
        partial = self._destination.with_name(self._destination.name + ".part")
        try:
            self._destination.parent.mkdir(parents=True, exist_ok=True)
            request = urllib.request.Request(self._url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(request, timeout=DOWNLOAD_TIMEOUT) as response:
                total = int(response.headers.get("Content-Length") or 0)
                received = 0
                with open(partial, "wb") as handle:
                    while True:
                        if self.isInterruptionRequested():
                            raise InterruptedError("Download annullato dall'utente.")
                        chunk = response.read(DOWNLOAD_CHUNK_SIZE)
                        if not chunk:
                            break
                        handle.write(chunk)
                        received += len(chunk)
                        self.progress.emit(received, total)
            if total and received != total:
                raise OSError(f"Download incompleto: ricevuti {received} byte su {total}.")
            os.replace(partial, self._destination)
        except Exception as exc:  # rete, disco o annullamento: mai crash della GUI
            try:
                partial.unlink(missing_ok=True)
            except OSError:
                pass
            self.download_failed.emit(str(exc))
            return
        self.download_finished.emit(str(self._destination))


class UpdateCheckWorker(QThread):
    """Verifica in background la presenza di una nuova versione su GitHub Releases.

    Eseguito in QThread per non bloccare mai la GUI con una chiamata di rete,
    come richiesto per qualunque I/O nella suite.
    """

    update_available = Signal(object)  # UpdateInfo
    no_update = Signal()
    check_failed = Signal(str)

    def __init__(self, current_version: str, parent=None) -> None:
        super().__init__(parent)
        self._current_version = current_version

    def run(self) -> None:
        try:
            info = fetch_latest_release()
        except (urllib.error.URLError, TimeoutError, ValueError, OSError) as exc:
            self.check_failed.emit(str(exc))
            return
        except Exception as exc:  # difesa da risposte GitHub inattese
            self.check_failed.emit(str(exc))
            return

        if info is None:
            self.check_failed.emit("Risposta GitHub non valida: 'tag_name' mancante.")
            return

        if is_newer_version(info.version, self._current_version):
            self.update_available.emit(info)
        else:
            self.no_update.emit()
