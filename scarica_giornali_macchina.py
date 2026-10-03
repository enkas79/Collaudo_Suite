"""Downloader dei documenti ``Giornale Macchina`` dalla directory DMS JARVIS."""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import unicodedata
from pathlib import Path
from typing import Any, Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode, unquote
from urllib.request import Request, urlopen

BASE_URL = "https://jarvis.breton.it"
API_VERSION = "2"
SEARCH_PATH = "/dms/api/public/document/search"
CONTAINER_SEARCH_PATH = "/dms/api/public/container/search"
DOCUMENT_DETAIL_PATH = "/dms/api/public/document/{document_id}"
BROWSE_CONTAINER_PATH = "/dms/api/public/container/{container_id}/browse"
DOWNLOAD_PATH = "/dms/api/public/document/{document_id}/download/{file_id}"
DOWNLOAD_DOCUMENT_PATH = "/dms/api/public/document/{document_id}/download"
MACHINE_PREFIXES = (
    "MASTERGEVNC", "NC1200", "NC300", "NC400", "NC600", "TRINITY", "GENYA",
    "EVONIX", "VIPER", "EAGLE", "WMEE", "WMFE", "WMGE", "WMHE", "WMME",
    "WMRE", "WME", "WMF", "WMG", "WMH", "WMR", "WM", "WP", "NC",
)


def _normalise_name(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"[^a-z0-9]+", "", text.casefold())


def is_giornale_macchina_name(value: Any) -> bool:
    return "giornalemacchina" in _normalise_name(value)


def machine_type_from_text(*values: Any) -> str:
    """Detect a stable machine family from order path/title/file name."""
    text = _normalise_name(" ".join(str(value or "") for value in values))
    for prefix in MACHINE_PREFIXES:
        if prefix.casefold() in text:
            return prefix
    return "NON_IDENTIFICATA"


def _safe_filename(value: str, fallback: str) -> str:
    name = Path(str(value or "")).name.strip()
    name = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", name).strip(" .")
    return name or fallback


def _request_json(token: str, path: str, payload: dict[str, Any]) -> dict[str, Any]:
    url = f"{BASE_URL}{path}?{urlencode({'api-version': API_VERSION})}"
    print(f"[API >] POST {path}")
    request = Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Accept": "application/json",
            "Content-Type": "application/json",
            "jarvis-auth-token": token,
        },
        method="POST",
    )
    try:
        with urlopen(request, timeout=60) as response:
            body = response.read().decode("utf-8")
            print(f"[API <] POST {path} -> HTTP {response.status}")
            result = json.loads(body)
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")[:500]
        raise RuntimeError(f"HTTP {exc.code} su {path}: {detail}") from exc
    except URLError as exc:
        raise RuntimeError(f"JARVIS non raggiungibile: {exc.reason}") from exc
    if not isinstance(result, dict):
        raise RuntimeError(f"Risposta non valida da {path}.")
    if result.get("success") is False:
        raise RuntimeError(str(result.get("message") or result.get("error") or result))
    return result


def _request_binary(token: str, path: str, *, query: dict[str, Any] | None = None) -> tuple[bytes, str]:
    query = {"api-version": API_VERSION, **(query or {})}
    url = f"{BASE_URL}{path}?{urlencode(query)}"
    print(f"[API >] GET {path} (download binario)")
    request = Request(url, headers={"Accept": "application/octet-stream", "jarvis-auth-token": token})
    try:
        with urlopen(request, timeout=120) as response:
            content = response.read()
            print(f"[API <] GET {path} -> HTTP {response.status}, {len(content)} byte")
            return content, response.headers.get("Content-Disposition", "")
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")[:500]
        raise RuntimeError(f"HTTP {exc.code} su {path}: {detail}") from exc
    except URLError as exc:
        raise RuntimeError(f"Download non riuscito: {exc.reason}") from exc


def _request_json_get(token: str, path: str) -> dict[str, Any]:
    url = f"{BASE_URL}{path}?{urlencode({'api-version': API_VERSION})}"
    request = Request(url, headers={"Accept": "application/json", "jarvis-auth-token": token})
    print(f"[API >] GET {path}")
    try:
        with urlopen(request, timeout=60) as response:
            body = response.read().decode("utf-8")
            print(f"[API <] GET {path} -> HTTP {response.status}")
            result = json.loads(body)
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")[:500]
        raise RuntimeError(f"HTTP {exc.code} su {path}: {detail}") from exc
    except URLError as exc:
        raise RuntimeError(f"JARVIS non raggiungibile: {exc.reason}") from exc
    if not isinstance(result, dict):
        raise RuntimeError(f"Risposta non valida da {path}.")
    if result.get("success") is False:
        raise RuntimeError(str(result.get("message") or result.get("error") or result))
    return result


def _browse_container(token: str, container_id: str) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    response = _request_json_get(token, BROWSE_CONTAINER_PATH.format(container_id=quote(container_id, safe="")))
    data = response.get("data")
    data = data if isinstance(data, dict) else {}
    containers = [item for item in data.get("containers") or [] if isinstance(item, dict)]
    documents = [item for item in data.get("documents") or [] if isinstance(item, dict)]
    return containers, documents


def _directory_documents(token: str, *, max_containers: int = 100000) -> Iterable[tuple[dict[str, Any], str]]:
    """Walk only the ``Doc WBS`` trees of the order folders."""
    pending = [("Container_0", "", False)]
    visited: set[str] = set()
    while pending and len(visited) < max_containers:
        container_id, container_path, inside_wbs = pending.pop()
        if not container_id or container_id in visited:
            continue
        visited.add(container_id)
        print(f"[DIR] Apro {container_path or 'root'} ({container_id})")
        try:
            containers, documents = _browse_container(token, container_id)
        except RuntimeError as exc:
            # Some tokens can search containers/documents but cannot browse the root.
            # Resolve the Doc WBS containers first, then browse only those trees.
            if container_id == "Container_0" and "HTTP 403" in str(exc):
                print("[INFO] Browse root non autorizzato: cerco i contenitori Doc WBS.")
                wbs_containers = list(_search_wbs_containers(token))
                for wbs_id, wbs_path in wbs_containers:
                    pending.append((wbs_id, wbs_path, True))
                continue
            raise
        print(f"[DIR] {container_id}: {len(containers)} sottocartelle, {len(documents)} documenti")
        if inside_wbs:
            for document in documents:
                yield document, container_path
        for child in containers:
            child_id = str(child.get("id") or "").strip()
            child_name = str(child.get("name") or child_id).strip()
            if child_id:
                child_path = f"{container_path} / {child_name}" if container_path else child_name
                child_is_wbs = inside_wbs or _normalise_name(child_name) in {"docwbs", "documentowbs"}
                if child_is_wbs and not inside_wbs:
                    print(f"[WBS] {child_path}")
                pending.append((child_id, child_path, child_is_wbs))


def _search_documents_under_doc_wbs(token: str) -> Iterable[tuple[dict[str, Any], str]]:
    """Fallback for tokens without permission to browse the DMS root."""
    for document in _pages(token, page_size=100, max_pages=100000, search_text="Giornale"):
        path = str(
            document.get("documentFullPath")
            or document.get("fullPath")
            or document.get("path")
            or ""
        ).strip()
        if "docwbs" not in _normalise_name(path):
            continue
        yield document, path


def _search_wbs_containers(token: str) -> Iterable[tuple[str, str]]:
    """Find Doc WBS containers without traversing the forbidden root container."""
    seen: set[str] = set()
    for container in _pages(
        token,
        page_size=100,
        max_pages=100000,
        search_text="Doc WBS",
        search_path=CONTAINER_SEARCH_PATH,
    ):
        container_id = str(container.get("id") or container.get("containerId") or "").strip()
        name = str(container.get("name") or container.get("title") or "").strip()
        path = str(container.get("fullPath") or container.get("path") or name).strip()
        if not container_id or container_id in seen:
            continue
        if "docwbs" not in _normalise_name(f"{name} {path}"):
            continue
        seen.add(container_id)
        print(f"[WBS] {path or name}")
        yield container_id, path or name


def _pages(
    token: str,
    *,
    page_size: int,
    max_pages: int,
    search_text: str = "Giornale",
    search_path: str = SEARCH_PATH,
) -> Iterable[dict[str, Any]]:
    for page in range(max_pages):
        if page * page_size >= 10000:
            print("[WARN] JARVIS limita una ricerca a 10.000 risultati; ricerca interrotta al limite API.", file=sys.stderr)
            break
        response = _request_json(
            token,
            search_path,
            {
                "pageSize": page_size,
                "startIndex": page * page_size,
                "showArchived": False,
                # Riduce il result set prima della paginazione: JARVIS impone
                # un limite globale di 10.000 risultati per una ricerca.
                "text": search_text,
                "sortField": None,
                "sortDescending": True,
            },
        )
        rows = response.get("data")
        if not isinstance(rows, list) or not rows:
            print(f"[RICERCA] Pagina {page + 1}: nessun risultato")
            break
        print(f"[RICERCA] Pagina {page + 1}: {len(rows)} risultati")
        for row in rows:
            if isinstance(row, dict):
                yield row
        total = response.get("total")
        if len(rows) < page_size or (isinstance(total, int) and (page + 1) * page_size >= total):
            break


def _filename_from_headers(disposition: str) -> str:
    match = re.search(r"filename\*=UTF-8''([^;]+)|filename=\"?([^;\"]+)", disposition, re.I)
    return unquote(match.group(1) or match.group(2) or "").strip() if match else ""


def _write_file(folder: Path, suggested_name: str, content: bytes, disposition: str, fallback: str) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / _safe_filename(_filename_from_headers(disposition) or suggested_name, fallback)
    if target.exists():
        stem, suffix = target.stem, target.suffix
        index = 2
        while target.exists():
            target = folder / f"{stem}_{index}{suffix}"
            index += 1
    target.write_bytes(content)
    return target


def download_documents(
    token: str,
    output: Path,
    *,
    page_size: int = 100,
    max_pages: int = 100000,
    dry_run: bool = False,
    organize_by_machine: bool = False,
) -> tuple[int, int]:
    """Search the JARVIS Documentale directory and download matching documents."""
    token = str(token or "").strip()
    if not token:
        raise ValueError("Inserire un token JARVIS.")
    found = downloaded = 0
    del page_size
    print(f"[START] Ricerca giornali macchina | simulazione={'sì' if dry_run else 'no'} | output={output}")
    print(f"[START] Organizza per macchina={'sì' if organize_by_machine else 'no'} | limite cartelle={max_pages}")
    directory_documents = list(_directory_documents(token, max_containers=max(1, max_pages)))
    if not directory_documents:
        print("[INFO] Nessun documento trovato attraversando Doc WBS: verifico solo i percorsi Doc WBS indicizzati.")
        directory_documents = list(_search_documents_under_doc_wbs(token))
    for document, container_path in directory_documents:
        name = str(document.get("title") or document.get("fileName") or "").strip()
        if not name or not is_giornale_macchina_name(name):
            continue
        found += 1
        document_id = str(document.get("id") or "").strip()
        print(f"[MATCH] {name} ({document_id}) | cartella: {container_path or 'root'}")
        if dry_run:
            continue
        if not document_id:
            print(f"[WARN] documento senza id: {name}", file=sys.stderr)
            continue
        try:
            detail = _request_json_get(
                token,
                DOCUMENT_DETAIL_PATH.format(document_id=quote(document_id, safe="")),
            ).get("data")
            detail = detail if isinstance(detail, dict) else {}
            file_id = str(detail.get("defaultFileId") or detail.get("defaultBlobId") or "").strip()
            file_info_name = ""
            if not file_id:
                # The current JARVIS UI document model exposes file IDs in
                # filesInfo rather than defaultFileId/defaultBlobId.
                file_infos = detail.get("filesInfo") or detail.get("files") or []
                if isinstance(file_infos, dict):
                    file_infos = [file_infos]
                if isinstance(file_infos, list):
                    main_file = next(
                        (item for item in file_infos if isinstance(item, dict) and item.get("isMain")),
                        next((item for item in file_infos if isinstance(item, dict)), {}),
                    )
                    file_id = str(main_file.get("fileId") or main_file.get("blobId") or "").strip()
                    file_info_name = str(main_file.get("name") or "").strip()
            download_name = str(detail.get("fileName") or file_info_name or name).strip()
            if file_id:
                content, disposition = _request_binary(
                    token,
                    DOWNLOAD_PATH.format(document_id=quote(document_id, safe=""), file_id=quote(file_id, safe="")),
                    query={"fileName": download_name, "format": "original"},
                )
            else:
                content, disposition = _request_binary(
                    token,
                    DOWNLOAD_DOCUMENT_PATH.format(document_id=quote(document_id, safe="")),
                    query={"fileName": download_name},
                )
            machine_type = machine_type_from_text(container_path, name)
            document_folder = output / "documenti"
            if organize_by_machine:
                document_folder /= machine_type
            if container_path:
                for segment in container_path.split(" / "):
                    document_folder /= _safe_filename(segment, "cartella")
            target = _write_file(document_folder, download_name, content, disposition, f"{document_id}.bin")
            downloaded += 1
            print(f"[OK] {target}")
        except RuntimeError as exc:
            print(f"[WARN] {name}: {exc}", file=sys.stderr)
    return found, downloaded


def main() -> int:
    parser = argparse.ArgumentParser(description="Scarica i Giornali Macchina dalla directory DMS JARVIS.")
    parser.add_argument("--token", default=os.environ.get("JARVIS_AUTH_TOKEN", ""))
    parser.add_argument("--output", type=Path, default=Path("Giornali_Macchina"))
    parser.add_argument("--page-size", type=int, default=100)
    parser.add_argument("--max-pages", type=int, default=100000, help="Numero massimo di cartelle DMS da attraversare.")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--organize-by-machine", action="store_true", help="Crea una sottocartella per tipologia macchina.")
    args = parser.parse_args()
    if not args.token.strip():
        parser.error("specificare --token oppure impostare JARVIS_AUTH_TOKEN")
    try:
        found, downloaded = download_documents(
            args.token,
            args.output,
            page_size=args.page_size,
            max_pages=args.max_pages,
            dry_run=args.dry_run,
            organize_by_machine=args.organize_by_machine,
        )
    except Exception as exc:
        print(f"Errore: {exc}", file=sys.stderr)
        return 2
    print(f"Completato: {found} corrispondenze, {downloaded} file scaricati.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
