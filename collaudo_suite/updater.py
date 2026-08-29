from __future__ import annotations

import json
import urllib.error
import urllib.request
from dataclasses import dataclass

from PySide6.QtCore import QThread, Signal

from .app_info import GITHUB_OWNER, GITHUB_REPO

RELEASES_API_URL = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
REQUEST_TIMEOUT = 6


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
        headers={"Accept": "application/vnd.github+json", "User-Agent": "CollaudoSuite-Updater"},
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
