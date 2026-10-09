"""Downloader dei documenti ``Giornale Macchina`` dalla sezione Assets JARVIS."""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import threading
import unicodedata
from pathlib import Path
from typing import Any, Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode, unquote
from urllib.request import Request, urlopen

BASE_URL = "https://jarvis.breton.it"
API_VERSION = "2"
ASSET_SEARCH_PATH = "/api/v1/omnisearch/search"
ASSET_DOCUMENTS_PATH = "/api/v2/SystemDocuments/Search"
ASSET_DOWNLOAD_PATH = "/api/v2/download/openFile/{document_id}/{blob_id}/{file_name}"
SEARCH_PATH = "/dms/api/public/document/search"
CONTAINER_SEARCH_PATH = "/dms/api/public/container/search"
DOCUMENT_DETAIL_PATH = "/dms/api/public/document/{document_id}"
BROWSE_CONTAINER_PATH = "/dms/api/public/container/{id}/browse"
DOWNLOAD_PATH = "/dms/api/public/document/{document_id}/download/{file_id}"
DOWNLOAD_DOCUMENT_PATH = "/dms/api/public/document/{document_id}/download"
MACHINE_PREFIXES = (
    "MASTERGEVNC", "NC1200", "NC300", "NC400", "NC600", "TRINITY", "GENYA",
    "EVONIX", "VIPER", "EAGLE", "WMEE", "WMFE", "WMGE", "WMHE", "WMME",
    "WMRE", "WME", "WMF", "WMG", "WMH", "WMR", "WM", "WP", "NC",
)
ASSET_RESULT_WINDOW = 10000  # index.max_result_window di Elasticsearch
# Caratteri usati per suddividere la ricerca WBS ("contiene") oltre il limite.
WBS_SPLIT_ALPHABET = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-._/"
WBS_SPLIT_MAX_DEPTH = 6
EXCEL_EXTENSIONS = {".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm"}


def _normalise_name(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"[^a-z0-9]+", "", text.casefold())


def is_giornale_macchina_name(value: Any) -> bool:
    return "giornalemacchina" in _normalise_name(value)


def is_excel_filename(value: Any) -> bool:
    return Path(str(value or "").strip()).suffix.casefold() in EXCEL_EXTENSIONS


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


def _request_asset_json(token: str, path: str, payload: dict[str, Any], *, auth_cookie: str = "") -> dict[str, Any]:
    """Call the internal API used by the JARVIS Assets page."""
    url = f"{BASE_URL}{path}"
    print(f"[API >] POST {path}")
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "X-Jarvis-Language": "it",
    }
    if auth_cookie.strip():
        headers["Cookie"] = f"AuthCookie={auth_cookie.strip()}"
    else:
        headers["jarvis-auth-token"] = token
        headers["Authorization"] = f"Bearer {token}"
    request = Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers=headers,
        method="POST",
    )
    try:
        with urlopen(request, timeout=60) as response:
            raw = response.read().decode("utf-8", errors="replace")
            content_type = response.headers.get("Content-Type", "")
            print(f"[API <] POST {path} -> HTTP {response.status} {content_type}, {len(raw)} caratteri")
            try:
                result = json.loads(raw)
            except json.JSONDecodeError as exc:
                raise RuntimeError(
                    f"Risposta non JSON da {path} (HTTP {response.status}, {content_type or 'tipo ignoto'}): "
                    f"{raw[:200].strip() or 'corpo vuoto'}. Token o cookie probabilmente scaduti: rieffettuare il login in JARVIS."
                ) from exc
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")[:500]
        if exc.code == 403:
            raise RuntimeError(
                "JARVIS ha negato l'accesso alla sezione Assets (HTTP 403). "
                "Usare un token/sessione con permesso di lettura Assets."
            ) from exc
        raise RuntimeError(f"HTTP {exc.code} su {path}: {detail}") from exc
    except URLError as exc:
        raise RuntimeError(f"JARVIS non raggiungibile: {exc.reason}") from exc
    if not isinstance(result, dict):
        raise RuntimeError(f"Risposta non valida da {path}.")
    if result.get("success") is False:
        raise RuntimeError(str(result.get("errorMessage") or result.get("message") or result))
    return result


def _request_asset_binary(token: str, path: str, *, auth_cookie: str = "") -> tuple[bytes, str]:
    url = f"{BASE_URL}{path}"
    print(f"[API >] GET {path} (download binario)")
    headers = {
        "Accept": "application/octet-stream",
        "X-Jarvis-Language": "it",
    }
    if auth_cookie.strip():
        headers["Cookie"] = f"AuthCookie={auth_cookie.strip()}"
    else:
        headers["jarvis-auth-token"] = token
        headers["Authorization"] = f"Bearer {token}"
    request = Request(
        url,
        headers=headers,
    )
    try:
        with urlopen(request, timeout=120) as response:
            content = response.read()
            print(f"[API <] GET download -> HTTP {response.status}, {len(content)} byte")
            return content, response.headers.get("Content-Disposition", "")
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")[:500]
        raise RuntimeError(f"HTTP {exc.code} su download: {detail}") from exc
    except URLError as exc:
        raise RuntimeError(f"Download non riuscito: {exc.reason}") from exc


def _asset_string_property(asset: dict[str, Any], key: str) -> str:
    for prop in asset.get("stringProperties") or []:
        if isinstance(prop, dict) and str(prop.get("key") or "") == key:
            values = prop.get("value") or []
            return str(values[0] or "").strip() if values else ""
    return ""


def _search_assets(
    token: str,
    *,
    start_index: int,
    chunk_size: int,
    wbs: str = "",
    auth_cookie: str = "",
    exact: bool = True,
) -> dict[str, Any]:
    filters = []
    if wbs.strip():
        filters.append({"field": "wbs", "dataType": 0, "values": [wbs.strip()], "exactSearch": exact})
    payload = {
        "domainContext": "Assets",
        "query": {
            "filters": filters,
            "startIndex": start_index,
            "chunkSize": chunk_size,
            "showArchived": False,
        },
    }
    return _request_asset_json(token, ASSET_SEARCH_PATH, payload, auth_cookie=auth_cookie)


def _search_asset_documents(token: str, asset_id: str, *, auth_cookie: str = "") -> list[dict[str, Any]]:
    response = _request_asset_json(
        token,
        ASSET_DOCUMENTS_PATH,
        {
            "aggregateId": asset_id,
            "filter": "",
            "showVirtualPath": False,
            "withoutFiles": False,
            "getOnlyDocumentsWithoutSecondaryContext": False,
            "metadataFilterComposition": 0,
        },
        auth_cookie=auth_cookie,
    )
    root = response.get("root") or {}
    documents: list[dict[str, Any]] = []
    visited: set[str] = set()

    def walk(node: Any) -> None:
        if isinstance(node, list):
            for item in node:
                walk(item)
            return
        if not isinstance(node, dict):
            return
        document_id = str(node.get("documentId") or "").strip()
        if document_id and document_id not in visited and ("files" in node or "title" in node):
            visited.add(document_id)
            documents.append(node)
        for key in ("documents", "folders", "containers", "children", "root"):
            walk(node.get(key))

    walk(root)
    return documents


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
    response = _request_json_get(token, BROWSE_CONTAINER_PATH.format(id=quote(container_id, safe="")))
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


def _iter_assets(
    token: str,
    *,
    page_size: int,
    max_pages: int,
    wbs: str,
    exact: bool,
    auth_cookie: str,
    stop_event: threading.Event | None,
    seen: set[str],
    depth: int = 0,
) -> Iterable[dict[str, Any]]:
    """Itera gli asset a blocchi; oltre i 10.000 risultati suddivide per WBS "contiene"."""
    limit_hit = False
    for page in range(max(1, int(max_pages))):
        if stop_event is not None and stop_event.is_set():
            return
        start_index = page * page_size
        if start_index >= ASSET_RESULT_WINDOW:
            limit_hit = True
            break
        # startIndex + chunkSize non deve superare index.max_result_window.
        chunk_size = min(page_size, ASSET_RESULT_WINDOW - start_index)
        response = _search_assets(
            token, start_index=start_index, chunk_size=chunk_size,
            wbs=wbs, auth_cookie=auth_cookie, exact=exact,
        )
        rows = [item.get("item") for item in response.get("items") or [] if isinstance(item, dict)]
        rows = [item for item in rows if isinstance(item, dict)]
        if not rows:
            return
        print(f"[ASSETS] WBS~'{wbs or '*'}' pagina {page + 1}: {len(rows)} asset")
        for asset in rows:
            asset_id = str(asset.get("id") or "").strip()
            if asset_id and asset_id in seen:
                continue
            if asset_id:
                seen.add(asset_id)
            yield asset
        total = response.get("total")
        if len(rows) < chunk_size or (isinstance(total, int) and start_index + len(rows) >= total):
            return
    if not limit_hit:
        return
    if exact or depth >= WBS_SPLIT_MAX_DEPTH:
        print(f"[WARN] Limite di {ASSET_RESULT_WINDOW} risultati raggiunto per WBS '{wbs}': alcuni asset potrebbero mancare.", file=sys.stderr)
        return
    print(f"[INFO] Oltre {ASSET_RESULT_WINDOW} risultati per WBS '{wbs or '*'}': suddivido la ricerca.")
    for char in WBS_SPLIT_ALPHABET:
        if stop_event is not None and stop_event.is_set():
            return
        yield from _iter_assets(
            token, page_size=page_size, max_pages=max_pages, wbs=wbs + char, exact=False,
            auth_cookie=auth_cookie, stop_event=stop_event, seen=seen, depth=depth + 1,
        )


def download_documents(
    token: str,
    output: Path,
    *,
    page_size: int = 100,
    max_pages: int = 100000,
    dry_run: bool = False,
    organize_by_machine: bool = False,
    wbs: str = "",
    auth_cookie: str = "",
    stop_event: threading.Event | None = None,
) -> tuple[int, int]:
    """Search only JARVIS Assets and download their Giornale Macchina files."""
    token = str(token or "").strip()
    if not token:
        raise ValueError("Inserire un token JARVIS.")
    if "\r" in token or "\n" in token:
        raise ValueError("Il token JARVIS non può contenere righe multiple: incollare solo il token, non il log.")
    found = downloaded = 0
    print(f"[START] Ricerca giornali macchina | simulazione={'sì' if dry_run else 'no'} | output={output}")
    print(f"[START] Sorgente: JARVIS Assets | WBS={wbs.strip() or 'tutti'} | organizzazione per macchina={'sì' if organize_by_machine else 'no'}")
    page_size = max(1, min(int(page_size), 100))
    assets = _iter_assets(
        token, page_size=page_size, max_pages=max_pages, wbs=wbs.strip(), exact=bool(wbs.strip()),
        auth_cookie=auth_cookie, stop_event=stop_event, seen=set(),
    )
    for asset in assets:
        if stop_event is not None and stop_event.is_set():
            print("[STOP] Arresto richiesto dall'utente.")
            return found, downloaded
        asset_id = str(asset.get("id") or "").strip()
        asset_title = str(asset.get("title") or asset_id).strip()
        if not asset_id:
            continue
        try:
            documents = _search_asset_documents(token, asset_id, auth_cookie=auth_cookie)
        except RuntimeError as exc:
            print(f"[WARN] {asset_title}: {exc}", file=sys.stderr)
            continue
        for document in documents:
            if stop_event is not None and stop_event.is_set():
                print("[STOP] Arresto richiesto dall'utente.")
                return found, downloaded
            name = str(document.get("title") or "").strip()
            files = [item for item in document.get("files") or [] if isinstance(item, dict)]
            matching_files = [
                item for item in files
                if is_excel_filename(item.get("fileName"))
                and is_giornale_macchina_name(item.get("fileName"))
            ]
            if is_giornale_macchina_name(name):
                matching_files = [item for item in files if is_excel_filename(item.get("fileName"))]
            if not matching_files:
                continue
            for file_info in matching_files:
                document_id = str(document.get("documentId") or "").strip()
                blob_id = str(file_info.get("blobId") or "").strip()
                download_name = str(file_info.get("fileName") or name or "giornale_macchina.bin").strip()
                found += 1
                print(f"[MATCH] {name} | {asset_title} ({asset_id}) | {download_name}")
                if dry_run:
                    continue
                if not document_id or not blob_id:
                    print(f"[WARN] file senza documentId/blobId: {download_name}", file=sys.stderr)
                    continue
                try:
                    content, disposition = _request_asset_binary(
                        token,
                        ASSET_DOWNLOAD_PATH.format(
                            document_id=quote(document_id, safe=""),
                            blob_id=quote(blob_id, safe=""),
                            file_name=quote(download_name, safe=""),
                        ),
                        auth_cookie=auth_cookie,
                    )
                    machine_type = machine_type_from_text(asset_title, _asset_string_property(asset, "matnr"))
                    document_folder = output / "documenti"
                    if organize_by_machine:
                        document_folder /= machine_type
                    document_folder /= _safe_filename(asset_title, asset_id)
                    target = _write_file(document_folder, download_name, content, disposition, f"{document_id}.bin")
                    downloaded += 1
                    print(f"[OK] {target}")
                except RuntimeError as exc:
                    print(f"[WARN] {download_name}: {exc}", file=sys.stderr)
    if stop_event is not None and stop_event.is_set():
        print("[STOP] Arresto richiesto dall'utente.")
    return found, downloaded


def main() -> int:
    parser = argparse.ArgumentParser(description="Scarica i Giornali Macchina dalla sezione Assets JARVIS.")
    parser.add_argument("--token", default=os.environ.get("JARVIS_AUTH_TOKEN", ""))
    parser.add_argument("--output", type=Path, default=Path("Giornali_Macchina"))
    parser.add_argument("--page-size", type=int, default=100)
    parser.add_argument("--max-pages", type=int, default=100000, help="Numero massimo di pagine Assets da leggere.")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--organize-by-machine", action="store_true", help="Crea una sottocartella per tipologia macchina.")
    parser.add_argument("--wbs", default="", help="Limita la ricerca all'Asset con questo codice WBS.")
    parser.add_argument("--auth-cookie", default=os.environ.get("JARVIS_AUTH_COOKIE", ""), help="Valore del cookie AuthCookie; preferire JARVIS_AUTH_COOKIE.")
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
            wbs=args.wbs,
            auth_cookie=args.auth_cookie,
        )
    except Exception as exc:
        print(f"Errore: {exc}", file=sys.stderr)
        return 2
    print(f"Completato: {found} corrispondenze, {downloaded} file scaricati.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())