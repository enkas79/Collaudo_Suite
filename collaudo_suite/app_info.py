from __future__ import annotations

import sys
from pathlib import Path

APP_TITLE = "Collaudo Suite"
GITHUB_OWNER = "enkas79"
GITHUB_REPO = "Collaudo_Suite"

# Usato solo se version.txt non è raggiungibile da nessuna posizione nota.
_FALLBACK_VERSION = "1.1.8"


def _version_file_candidates() -> list[Path]:
    candidates = [Path(__file__).resolve().parent.parent / "version.txt"]

    bundle_root = getattr(sys, "_MEIPASS", None)
    if bundle_root:
        candidates.append(Path(bundle_root) / "version.txt")

    executable_dir = Path(sys.executable).resolve().parent
    candidates.append(executable_dir / "version.txt")
    candidates.append(executable_dir / "_internal" / "version.txt")
    return candidates


def get_app_version() -> str:
    """Legge la versione corrente da version.txt (sorgente o build PyInstaller)."""
    for candidate in _version_file_candidates():
        try:
            if candidate.exists():
                text = candidate.read_text(encoding="utf-8").strip()
                if text:
                    return text
        except OSError:
            continue
    return _FALLBACK_VERSION
