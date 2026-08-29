from __future__ import annotations

import calendar
import csv
from functools import lru_cache
import random
import re
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import sys
from typing import Iterable, Mapping, Sequence
from xml.sax.saxutils import escape as xml_escape
import xml.etree.ElementTree as ET

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

CHECKBOX_CHARS = ("☐", "□", "☑", "☒", "✓", "✔")
BULLET_PREFIXES = ("·", "•", "-", "–", "—", "*", "")
WHITESPACE_RE = re.compile(r"\s+")
NR_RE = re.compile(r"^\d+(?:\.\d+)+$")  # es. 1.12, 20.100, 39.24
TICKET_BASE_URL = "https://jarvis.breton.it/UI/#/ticketing/ticket/Ticket_"

ACTION_WORDS = (
    "controll", "verific", "prova", "provare", "esegu", "misura", "misurare",
    "registr", "compil", "azion", "premere", "selezion", "salvare", "aprire",
    "inserire", "impost", "attiv", "disattiv", "ridurre", "scolleg", "colleg",
    "ripristin", "lanciare", "caricare", "creare", "montare", "rimuovere",
    "taratura", "movimento", "tensione", "messa in bolla", "rettilineità",
    "planarità", "perpendicolarità", "quadratura", "allineamento",
)

META_PATTERNS = [
    re.compile(r"^modulo\s+istruzioni$", re.I),
    re.compile(r"^check\s*list$", re.I),
    re.compile(r"^manuale\s+generale", re.I),
    re.compile(r"^documento\s+qualit", re.I),
    re.compile(r"^sezione\s+di\s+collaudo$", re.I),
    re.compile(r"^modulo\s+collaudo", re.I),
    re.compile(r"^protocollo$", re.I),
    re.compile(r"^titolo$", re.I),
    re.compile(r"^tipo$", re.I),
    re.compile(r"^wbs$", re.I),
    re.compile(r"^stato$", re.I),
    re.compile(r"^rev\.?\s*n", re.I),
    re.compile(r"^pag\.?\s*\d+", re.I),
    re.compile(r"^pagina\s+\d+", re.I),
    re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4}$"),
    re.compile(r"^\[digitare qui\]$", re.I),
    re.compile(r"^ogni diritto", re.I),
    re.compile(r"^errore\.\s*nome\s+della\s+propriet", re.I),
]

COLUMN_HEADERS = {
    "nr", "n", "controllo", "strumento", "richiesta", "misurata", "operatore", "data",
    "collaudo", "note del collaudo", "schema", "osservazioni e riferimenti",
    "fornitore", "anno", "nr serie", "gruppo preassemblato",
}

LOW_VALUE_FIELDS = {
    "costruttore", "matricola", "fornitore", "anno", "nr serie", "operatore", "data",
    "collaudo", "note del collaudo", "wbs", "stato", "tipo", "titolo", "protocollo", "contratto",
}

VERNICIATURA_CHILDREN = {
    "basamento", "trave", "canotto", "trasporti", "siliconatura",
}


@dataclass(frozen=True)
class ChecklistItem:
    text: str
    source: str
    kind: str  # "Fisso", "MAP" or "ANALYZER"
    ticket_number: str = ""
    ticket_url: str = ""


@dataclass(frozen=True)
class ChecklistRecord:
    text: str
    source: str
    kind: str  # "Fisso", "MAP" or "ANALYZER"
    result: str = ""  # "Pass", "No pass" or ""
    note: str = ""
    ticket_number: str = ""
    ticket_url: str = ""


def clean_text(text: str) -> str:
    """Normalize Word paragraph/table-cell text into a readable checklist row."""
    text = str(text).replace("\xa0", " ").replace("\u200b", " ")
    text = text.replace("\r", "\n")
    for char in CHECKBOX_CHARS:
        text = text.replace(char, " ")
    text = text.strip()
    while text.startswith(BULLET_PREFIXES):
        text = text[1:].strip()
    text = WHITESPACE_RE.sub(" ", text)
    return text.strip(" \t-–—")


def _norm_key(text: str) -> str:
    return re.sub(r"\W+", "", text.casefold())


def is_meta_line(text: str) -> bool:
    if not text:
        return True
    if len(text) <= 1:
        return True
    key = text.casefold().strip()
    if key in COLUMN_HEADERS or key in LOW_VALUE_FIELDS:
        return True
    for pattern in META_PATTERNS:
        if pattern.search(text):
            return True
    return False


def is_probable_section_title(text: str) -> bool:
    """Reject generic sector titles such as CONTROLLO ASSE X or RETTILINEITA MOVIMENTO ASSE Y.

    Numbered table rows are handled separately and are not rejected only because they contain
    uppercase words.
    """
    cleaned = clean_text(text)
    if is_meta_line(cleaned):
        return True
    if NR_RE.fullmatch(cleaned):
        return True

    words = cleaned.split()
    # Sector titles are often uppercase, sometimes followed by a lowercase note in parentheses
    # such as "(Se presente)". Remove parenthetical notes before testing uppercase shape.
    title_core = re.sub(r"\([^)]*\)", "", cleaned).strip(" -")
    title_words = title_core.split()
    if len(words) <= 10:
        letters = "".join(ch for ch in title_core if ch.isalpha())
        if letters and title_core.upper() == title_core:
            return True

    lower = cleaned.casefold()
    if lower in {"lay-out", "layout"}:
        return True
    if re.fullmatch(r"\d+(?:\.\d+)?\s*-\s*lay-?out", lower):
        return True
    if len(words) <= 5 and lower.startswith(("controllo ", "controlli ", "verifica ", "verifiche ")):
        return True
    if len(words) <= 4 and lower in COLUMN_HEADERS:
        return True
    return False


def _paragraph_texts(document: Document) -> Iterable[str]:
    for paragraph in document.paragraphs:
        yield paragraph.text
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    yield paragraph.text


def _has_checkbox(raw_text: str) -> bool:
    return any(char in raw_text for char in CHECKBOX_CHARS)


def _looks_actionable(cleaned: str, raw_text: str) -> bool:
    if is_probable_section_title(cleaned):
        return False
    lower = cleaned.casefold()
    if _has_checkbox(raw_text):
        return True
    if raw_text.strip().startswith(BULLET_PREFIXES) and any(word in lower for word in ACTION_WORDS):
        return True
    if any(lower.startswith(word) for word in ACTION_WORDS):
        return True
    if any(word in lower for word in ACTION_WORDS) and 12 <= len(cleaned) <= 260:
        return True
    return False


def unique_texts(values: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        normalized = clean_text(value)
        key = _norm_key(normalized)
        if not normalized or key in seen:
            continue
        seen.add(key)
        result.append(normalized)
    return result


def _cell_text(cell) -> str:
    parts = [p.text for p in cell.paragraphs if clean_text(p.text)]
    return clean_text(" ".join(parts))


def _find_control_column(cells: list[str]) -> int | None:
    for idx, text in enumerate(cells):
        if text.casefold().strip() == "controllo":
            return idx
    return None


def _extract_structured_table_controls(document: Document, source_label: str) -> list[ChecklistItem]:
    """Extract real checklist rows from structured collaudo tables.

    The Word sheets often contain sector titles in merged rows (for example CONTROLLO ASSE X)
    and then a header row with a CONTROLLO column. We only take rows below a CONTROLLO header
    and reject generic titles such as SCHEMA, OSSERVAZIONI, CONTROLLO ASSE X, etc.
    """
    values: list[str] = []
    for table in document.tables:
        control_idx: int | None = None
        for row in table.rows:
            cells = [_cell_text(cell) for cell in row.cells]
            if not any(cells):
                continue

            nonempty_cells = [c for c in cells if c]
            repeated_title = False
            if nonempty_cells:
                most_common_count = max(nonempty_cells.count(c) for c in set(nonempty_cells))
                repeated_title = most_common_count >= max(2, len(nonempty_cells) // 2) and is_probable_section_title(nonempty_cells[0])
            if repeated_title or any(c.casefold() in {"schema", "osservazioni e riferimenti"} for c in nonempty_cells):
                control_idx = None
                continue

            header_idx = _find_control_column(cells)
            if header_idx is not None:
                control_idx = header_idx
                continue
            if control_idx is None or control_idx >= len(cells):
                continue

            control = cells[control_idx]
            if not control or is_meta_line(control):
                continue
            if is_probable_section_title(control):
                control_idx = None
                continue
            if _norm_key(control) in {_norm_key(x) for x in LOW_VALUE_FIELDS}:
                continue
            if len(control) < 6:
                continue

            number = ""
            for prefix_cell in cells[:control_idx]:
                if NR_RE.fullmatch(prefix_cell.strip()):
                    number = prefix_cell.strip()
                    break
            values.append(f"{number} - {control}" if number else control)
    return [ChecklistItem(text=t, source=source_label, kind="") for t in unique_texts(values)]


def extract_items_from_docx(path: str | Path, *, mode: str = "auto", source_label: str | None = None) -> list[ChecklistItem]:
    """Extract checklist items from a DOCX.

    Modes:
    - fixed: keeps almost every non-metadata line. Good for a dedicated fixed-checklist DOCX.
    - random: prefers numbered rows from collaudo tables; this avoids generic sector titles.
    - auto: same as random for structured tables, otherwise falls back to checkbox/actionable lines.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(path)
    if path.suffix.lower() != ".docx":
        raise ValueError(f"Formato non supportato: {path.suffix}. Usa file .docx")

    doc = Document(str(path))
    label = source_label or path.name

    if mode not in {"auto", "fixed", "random"}:
        raise ValueError("mode deve essere: auto, fixed o random")

    if mode in {"auto", "random"}:
        numbered = _extract_structured_table_controls(doc, label)
        if numbered:
            return numbered

    raw_lines = [line for line in _paragraph_texts(doc)]
    cleaned_pairs = [(raw, clean_text(raw)) for raw in raw_lines]
    cleaned_pairs = [(raw, clean) for raw, clean in cleaned_pairs if not is_meta_line(clean)]

    if mode == "fixed":
        # Dedicated fixed-checklist files are intentionally short and may contain valid controls
        # such as "Controllo ordine di vendita". Do not apply the sector-title filter here.
        chosen = [clean for raw, clean in cleaned_pairs]
    elif mode == "random":
        chosen = [clean for raw, clean in cleaned_pairs if _looks_actionable(clean, raw)]
    else:
        checkbox_items = [clean for raw, clean in cleaned_pairs if _has_checkbox(raw) and not is_probable_section_title(clean)]
        chosen = checkbox_items or [clean for raw, clean in cleaned_pairs if _looks_actionable(clean, raw)]

    return [ChecklistItem(text=t, source=label, kind="") for t in unique_texts(chosen)]


def _resource_path(package: str, *parts: str) -> Path:
    """Return a bundled resource path in source and PyInstaller builds.

    ``importlib.resources.files()`` with a package name supplied as a string is
    difficult for PyInstaller to detect statically. In previous builds the
    physical ``data`` directory was copied, but the Python subpackage
    ``collaudo_suite.checklist.data`` was omitted, causing:
    ``No module named 'collaudo_suite.checklist.data'``.

    Resolve the resource from the filesystem instead. PyInstaller exposes its
    onedir data directory through ``sys._MEIPASS``; ``__file__`` remains the
    most reliable location when running directly from source.
    """
    relative_package = Path(*package.split("."))
    relative_resource = relative_package.joinpath(*parts)

    candidates: list[Path] = []

    # Source tree and frozen package location. For this module the requested
    # package is normally ``collaudo_suite.checklist.data``.
    if package == "collaudo_suite.checklist.data":
        candidates.append(Path(__file__).resolve().parent / "data" / Path(*parts))

    bundle_root = getattr(sys, "_MEIPASS", None)
    if bundle_root:
        candidates.append(Path(bundle_root) / relative_resource)

    executable_dir = Path(sys.executable).resolve().parent
    candidates.extend(
        [
            executable_dir / relative_resource,
            executable_dir / "_internal" / relative_resource,
        ]
    )

    for candidate in candidates:
        if candidate.exists():
            return candidate

    # Keep the historical contract: callers may ask for an optional resource
    # and then test ``Path.exists()`` themselves (for example MAP and schede).
    # The first candidate is the canonical expected location.
    if candidates:
        return candidates[0]
    return Path(*parts)


def get_default_fixed_docx_path() -> Path:
    """Return the bundled Word file used for the mandatory fixed checklist."""
    return _resource_path("collaudo_suite.checklist.data", "Check list.docx")


def get_default_map_xlsx_path() -> Path:
    """Return the bundled MAP summary workbook.

    Older project packages used different file names for the internal MAP workbook.
    Prefer the compact bundled name, but keep backward compatibility.
    """
    for filename in ("map.xlsx",):
        candidate = _resource_path("collaudo_suite.checklist.data", filename)
        if candidate.exists():
            return candidate
    return _resource_path("collaudo_suite.checklist.data", "map.xlsx")


def get_builtin_schede_dir() -> Path:
    return _resource_path("collaudo_suite.checklist.data", "schede")


def get_builtin_schede() -> dict[str, Path]:
    """Return bundled collaudo sheets that can be used without browsing files."""
    schede_dir = get_builtin_schede_dir()
    if not schede_dir.exists():
        return {}
    return {p.stem: p for p in sorted(schede_dir.glob("*.docx"))}



XLSX_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _xlsx_col_to_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha()).upper()
    index = 0
    for char in letters:
        index = index * 26 + (ord(char) - ord("A") + 1)
    return max(0, index - 1)


def _xlsx_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(data)
    values: list[str] = []
    for si in root.findall("main:si", XLSX_NS):
        texts = [node.text or "" for node in si.findall(".//main:t", XLSX_NS)]
        values.append("".join(texts))
    return values


def _xlsx_sheet_path(zf: zipfile.ZipFile, preferred_name: str = "Data") -> str:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_by_id: dict[str, str] = {}
    for rel in rels.findall("pkgrel:Relationship", XLSX_NS):
        rel_id = rel.attrib.get("Id", "")
        target = rel.attrib.get("Target", "")
        if rel_id and target:
            clean_target = target.lstrip("/")
            rel_by_id[rel_id] = clean_target if clean_target.startswith("xl/") else "xl/" + clean_target

    first_path = ""
    for sheet in workbook.findall("main:sheets/main:sheet", XLSX_NS):
        name = sheet.attrib.get("name", "")
        rel_id = sheet.attrib.get(f"{{{XLSX_NS['rel']}}}id", "")
        target = rel_by_id.get(rel_id, "")
        if target and not first_path:
            first_path = target
        if target and name.casefold() == preferred_name.casefold():
            return target
    if first_path:
        return first_path
    raise ValueError("Nessun foglio leggibile trovato nel file XLSX MAP.")


def _xlsx_cell_text(cell: ET.Element, shared_strings: Sequence[str]) -> str:
    cell_type = cell.attrib.get("t", "")
    if cell_type == "inlineStr":
        texts = [node.text or "" for node in cell.findall(".//main:t", XLSX_NS)]
        return clean_text("".join(texts))

    value_node = cell.find("main:v", XLSX_NS)
    if value_node is None or value_node.text is None:
        return ""
    raw = value_node.text
    if cell_type == "s":
        try:
            return clean_text(shared_strings[int(raw)])
        except (ValueError, IndexError):
            return ""
    return clean_text(raw)


def _read_xlsx_rows(path: str | Path, *, sheet_name: str = "Data") -> list[list[str]]:
    """Read XLSX values with stdlib only. Enough for simple tabular MAP files."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(path)
    if path.suffix.lower() != ".xlsx":
        raise ValueError(f"Formato non supportato: {path.suffix}. Usa file .xlsx")

    with zipfile.ZipFile(path) as zf:
        shared_strings = _xlsx_shared_strings(zf)
        sheet_path = _xlsx_sheet_path(zf, sheet_name)
        root = ET.fromstring(zf.read(sheet_path))
        rows: list[list[str]] = []
        for row in root.findall("main:sheetData/main:row", XLSX_NS):
            values: list[str] = []
            for cell in row.findall("main:c", XLSX_NS):
                ref = cell.attrib.get("r", "")
                col_idx = _xlsx_col_to_index(ref) if ref else len(values)
                while len(values) <= col_idx:
                    values.append("")
                values[col_idx] = _xlsx_cell_text(cell, shared_strings)
            while values and values[-1] == "":
                values.pop()
            rows.append(values)
        return rows


def _find_header_index(headers: Sequence[str], *needles: str) -> int | None:
    normalized_needles = [_norm_key(n) for n in needles]
    for idx, header in enumerate(headers):
        key = _norm_key(header)
        if all(needle in key for needle in normalized_needles):
            return idx
    return None


def _find_header_index_any(headers: Sequence[str], candidates: Sequence[str]) -> int | None:
    """Find a column index using robust English/Italian header aliases."""
    normalized_candidates = [_norm_key(candidate) for candidate in candidates if _norm_key(candidate)]
    for idx, header in enumerate(headers):
        key = _norm_key(header)
        if not key:
            continue
        for candidate in normalized_candidates:
            if key == candidate or candidate in key or key in candidate:
                return idx
    return None


MAP_DATE_HEADER_GROUPS: tuple[tuple[str, ...], ...] = (
    (
        "Last modify",
        "Last modified",
        "Last modification",
        "LastModify",
        "Data ultima modifica",
        "Ultima modifica",
        "Data modifica",
        "Modified date",
        "Modification date",
        "Update date",
        "Updated at",
    ),
    (
        "Ticket date",
        "Data ticket",
        "Creation date",
        "Created date",
        "Created at",
        "Data creazione",
        "Data apertura",
        "Open date",
    ),
    ("Date", "Data"),
)


def _find_map_date_index(headers: Sequence[str]) -> int | None:
    """Locate the most meaningful MAP date column using ordered aliases.

    Last-modification fields are preferred because the MAP export and TicketV2
    terminology normally expose the most recent ticket activity through those
    columns. Generic Date/Data is used only as a final fallback.
    """
    normalized_headers = [_norm_key(header) for header in headers]
    for group_index, aliases in enumerate(MAP_DATE_HEADER_GROUPS):
        normalized_aliases = [_norm_key(alias) for alias in aliases]
        for idx, key in enumerate(normalized_headers):
            if not key:
                continue
            if any(key == alias for alias in normalized_aliases):
                return idx

        # Do not use substring matching for the generic Date/Data fallback:
        # headers such as "Data Module" are not ticket dates.
        if group_index == len(MAP_DATE_HEADER_GROUPS) - 1:
            continue
        for idx, key in enumerate(normalized_headers):
            if not key:
                continue
            if any(alias and (alias in key or key in alias) for alias in normalized_aliases):
                return idx
    return None


def _subtract_calendar_months(reference: date, months: int) -> date:
    """Subtract whole calendar months, clamping the day to the target month."""
    if months < 0:
        raise ValueError("Il numero di mesi non puo essere negativo.")
    absolute_month = reference.year * 12 + (reference.month - 1) - months
    year, zero_based_month = divmod(absolute_month, 12)
    month = zero_based_month + 1
    day = min(reference.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


def _coerce_reference_date(value: date | datetime | str | None) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date()
    except ValueError as exc:
        raise ValueError(f"Data di collaudo non valida: {text}") from exc


def _coerce_map_date(value: object, *, workbook_epoch: datetime | None = None) -> date | None:
    """Convert typical Excel/string date values to ``date``.

    Numeric Excel serials are supported when the workbook epoch is available.
    Invalid or empty values return ``None`` and are excluded by an active period
    filter rather than being silently treated as recent.
    """
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            from openpyxl.utils.datetime import from_excel
            converted = from_excel(value, epoch=workbook_epoch) if workbook_epoch is not None else from_excel(value)
            return converted.date() if isinstance(converted, datetime) else converted
        except Exception:
            return None

    text = str(value).strip()
    if not text:
        return None
    text = re.sub(r"\s+", " ", text)
    for fmt in (
        "%d/%m/%Y",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%d-%m-%Y",
        "%d.%m.%Y",
        "%m/%d/%Y",
    ):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date()
    except ValueError:
        return None


def map_period_bounds(reference_date: date | datetime | str, months_back: int) -> tuple[date, date]:
    """Return inclusive lower/upper dates for the selected MAP period."""
    reference = _coerce_reference_date(reference_date)
    if reference is None:
        raise ValueError("La data di collaudo e obbligatoria per il filtro temporale MAP.")
    months = int(months_back)
    if months not in {1, 3, 6, 12}:
        raise ValueError("Periodo MAP non valido. Sono ammessi 1, 3, 6 o 12 mesi.")
    return _subtract_calendar_months(reference, months), reference


def _find_map_header_row(rows: Sequence[Sequence[str]]) -> tuple[int, int, int, int | None]:
    """Return row/title/code/ticket column indices for a MAP workbook.

    Accepts common English headers and Italian exports such as
    'Numero Ticket', 'Titolo' and 'Commessa S Codice commerciale'.
    """
    for row_idx, headers in enumerate(rows[:25]):
        title_idx = _find_header_index_any(headers, ("Title", "Titolo"))
        code_idx = _find_header_index_any(
            headers,
            (
                "Commercial code",
                "Codice commerciale",
                "Commessa S Codice commerciale",
                "Commessa Codice commerciale",
            ),
        )
        ticket_idx = _find_header_index_any(headers, ("Ticket Number", "Numero Ticket", "N Ticket", "Ticket"))
        if title_idx is not None and code_idx is not None:
            return row_idx, title_idx, code_idx, ticket_idx
    raise ValueError(
        "Nel file MAP non trovo le colonne richieste. Servono almeno 'Title/Titolo' e "
        "'Commercial code/Codice commerciale'. La colonna ticket e' opzionale ma consigliata."
    )


def _row_value(row: Sequence[str], index: int | None) -> str:
    if index is None or index >= len(row):
        return ""
    return str(row[index]).strip()


@lru_cache(maxsize=8)
def _cached_xlsx_rows(path_text: str, mtime_ns: int, size: int) -> tuple[tuple[str, ...], ...]:
    """Cache MAP workbook rows while invalidating when the XLSX file changes."""
    del mtime_ns, size
    return tuple(tuple(row) for row in _read_xlsx_rows(path_text, sheet_name="Data"))


def _xlsx_rows_for_path(path: Path) -> list[list[str]]:
    """Read MAP workbook rows through a cache keyed by path and file signature."""
    stat = path.stat()
    return [list(row) for row in _cached_xlsx_rows(str(path.resolve()), stat.st_mtime_ns, stat.st_size)]


def _map_filter_matches(commercial_code: str, title: str, ticket_number: str, filter_text: str) -> bool:
    """Return True when a MAP row matches the selected commercial-code filter.

    Primary rule: Commercial code contains the filter text.
    Practical fallback: some MAP rows expose the useful machine code inside the Title, or the user
    may paste a ticket/order reference. In those cases the Title/Ticket fields are also checked.
    An empty filter means all MAP rows with a Title are eligible.
    """
    normalized_filter = _norm_key(filter_text)
    if not normalized_filter:
        return True

    normalized_code = _norm_key(commercial_code)
    normalized_title = _norm_key(title)
    normalized_ticket = _norm_key(ticket_number)

    # Main requested behavior: Commercial code contains the selected key.
    if normalized_filter in normalized_code:
        return True

    # Useful aliases seen in the MAP export. Keep this conservative.
    aliases = {normalized_filter}
    if normalized_filter == "nc300":
        aliases.update({"wp0300", "wp0300k"})

    if any(alias and alias in normalized_code for alias in aliases):
        return True

    # Fallback for real-world MAP rows where the commercial-code cell is generic but the title
    # contains the machine/order reference. This prevents an empty pool when the operator enters
    # a commessa or an old code that appears in the Title.
    return normalized_filter in normalized_title or normalized_filter in normalized_ticket



def build_ticket_url(ticket_number: str) -> str:
    """Build the direct ticket URL from a ticket number.

    The MAP may contain a numeric value, ``Ticket_12345`` or text with spaces.
    Only the final ticket identifier is appended to the known ticket route.
    """
    value = str(ticket_number or "").strip()
    if not value:
        return ""
    if value.lower().startswith("ticket_"):
        value = value[7:].strip()
    value = re.sub(r"\s+", "", value)
    if not value:
        return ""
    return f"{TICKET_BASE_URL}{value}"

def ticket_desc_sort_key(ticket_number: str) -> tuple[int, int, str]:
    """Return a stable key for displaying ticket identifiers in descending order.

    Numeric identifiers are ordered by their final number, so ``Ticket_100`` is
    correctly placed before ``Ticket_99``. Empty values remain at the end.
    """
    value = str(ticket_number or "").strip()
    if not value:
        return (0, -1, "")
    matches = re.findall(r"\d+", value)
    if matches:
        return (2, int(matches[-1]), value.casefold())
    return (1, -1, value.casefold())

def _map_row_metadata(
    path: Path,
    header_row_idx: int,
    ticket_idx: int | None,
    date_idx: int | None,
    *,
    require_dates: bool,
) -> tuple[dict[int, str], dict[int, date]]:
    """Read ticket hyperlinks and MAP dates in a single workbook pass."""
    if ticket_idx is None and date_idx is None:
        return {}, {}

    workbook = None
    try:
        from openpyxl import load_workbook

        workbook = load_workbook(path, read_only=False, data_only=False)
        sheet = workbook["Data"] if "Data" in workbook.sheetnames else workbook.active
        workbook_epoch = getattr(workbook, "epoch", None)
        links: dict[int, str] = {}
        dates: dict[int, date] = {}

        for excel_row in range(header_row_idx + 2, sheet.max_row + 1):
            if ticket_idx is not None:
                ticket_cell = sheet.cell(row=excel_row, column=ticket_idx + 1)
                target = ""
                if ticket_cell.hyperlink and ticket_cell.hyperlink.target:
                    target = str(ticket_cell.hyperlink.target).strip()
                elif isinstance(ticket_cell.value, str):
                    match = re.match(r'^=HYPERLINK\(\s*["\']([^"\']+)["\']', ticket_cell.value, re.I)
                    if match:
                        target = match.group(1).strip()
                if target:
                    links[excel_row] = target

            if date_idx is not None:
                parsed = _coerce_map_date(
                    sheet.cell(row=excel_row, column=date_idx + 1).value,
                    workbook_epoch=workbook_epoch,
                )
                if parsed is not None:
                    dates[excel_row] = parsed

        return links, dates
    except Exception as exc:
        if require_dates:
            raise ValueError(f"Non riesco a leggere le date dal file MAP: {path.name}") from exc
        return {}, {}
    finally:
        if workbook is not None:
            workbook.close()


def extract_map_items_from_xlsx(
    path: str | Path,
    filter_text: str = "",
    *,
    source_label: str | None = None,
    reference_date: date | datetime | str | None = None,
    months_back: int | None = None,
) -> list[ChecklistItem]:
    """Extract MAP rows from the summary workbook.

    - Checklist text: Title/Titolo column.
    - Ticket number: Ticket Number/Numero Ticket column, when present.
    - Main filter: Commercial code/Codice commerciale contains ``filter_text``.
    - Empty filter: use all rows with a valid Title.
    - Practical fallback: Title and Ticket Number are also searched for the typed filter.
    - Optional period filter: include only rows dated between ``reference_date - months_back``
      and the reference date, both inclusive.
    """
    path = Path(path)
    rows = _xlsx_rows_for_path(path)
    if not rows:
        return []

    header_row_idx, title_idx, code_idx, ticket_idx = _find_map_header_row(rows)
    headers = rows[header_row_idx]
    date_idx = _find_map_date_index(headers)

    period_start: date | None = None
    period_end: date | None = None
    if months_back is not None:
        period_start, period_end = map_period_bounds(reference_date, int(months_back))
        if date_idx is None:
            raise ValueError(
                "Nel file MAP non trovo una colonna data utilizzabile per il filtro temporale. "
                "Sono riconosciuti, ad esempio: Last Modify, Last Modified, Data ultima modifica, "
                "Created Date, Data ticket o Data."
            )

    ticket_links, row_dates = _map_row_metadata(
        path,
        header_row_idx,
        ticket_idx,
        date_idx if months_back is not None else None,
        require_dates=months_back is not None,
    )
    if months_back is not None and not row_dates:
        raise ValueError(
            "La colonna data del file MAP e stata riconosciuta, ma non contiene date valide. "
            "Verifica il formato delle date nel foglio Data."
        )

    label = source_label or Path(path).name
    result: list[ChecklistItem] = []
    seen: set[str] = set()
    for excel_row, row in enumerate(rows[header_row_idx + 1:], start=header_row_idx + 2):
        title = _row_value(row, title_idx)
        commercial_code = _row_value(row, code_idx)
        ticket_number = _row_value(row, ticket_idx)
        if not title:
            continue
        if period_start is not None and period_end is not None:
            row_date = row_dates.get(excel_row)
            if row_date is None or row_date < period_start or row_date > period_end:
                continue
        if not _map_filter_matches(commercial_code, title, ticket_number, filter_text):
            continue
        key = _norm_key(f"{ticket_number} {title}")
        if key in seen:
            continue
        seen.add(key)
        ticket_url = ticket_links.get(excel_row, "") or build_ticket_url(ticket_number)
        result.append(ChecklistItem(text=title, source=label, kind="", ticket_number=ticket_number, ticket_url=ticket_url))

    return result

def load_map_items_for_filter(
    path: str | Path,
    filter_text: str = "",
    *,
    reference_date: date | datetime | str | None = None,
    months_back: int | None = None,
) -> list[ChecklistItem]:
    """Load MAP titles using Commercial-code and optional temporal filters."""
    path = Path(path)
    return extract_map_items_from_xlsx(
        path,
        filter_text,
        source_label=path.name,
        reference_date=reference_date,
        months_back=months_back,
    )


def load_default_map_items_for_filter(filter_text: str = "") -> list[ChecklistItem]:
    """Load bundled MAP titles using a Commercial-code filter.

    Use an empty filter to make all MAP titles eligible.
    """
    return load_map_items_for_filter(get_default_map_xlsx_path(), filter_text)


# Backward-compatible alias for older saved code imports.
def load_default_map_items_for_machine(machine_name: str) -> list[ChecklistItem]:
    return load_default_map_items_for_filter(machine_name)


def available_map_filters() -> list[str]:
    """Return common Commercial-code prefixes found in the bundled MAP file.

    This is intentionally compact: the GUI shows the three filters most useful for this project,
    but the operator can still type any other filter.
    """
    return ["Tutti", "NC300", "GENYA", "TRINITY"]


def map_pool_size(filter_text: str = "") -> int:
    return len(load_default_map_items_for_filter("" if filter_text.casefold() == "tutti" else filter_text))

def sample_items_from_pool(
    pool: Sequence[ChecklistItem],
    count: int,
    *,
    kind: str,
    exclude_items: Sequence[ChecklistItem] = (),
    seed: int | None = None,
    label: str = "elementi",
) -> tuple[list[ChecklistItem], list[str]]:
    """Pick random items from a pool and assign a final kind label."""
    warnings: list[str] = []
    pick_count = max(0, int(count))
    if pick_count == 0:
        return [], warnings

    excluded = {_norm_key(item.text) for item in exclude_items}
    available = [item for item in pool if _norm_key(item.text) not in excluded]
    if pick_count > len(available):
        warnings.append(
            f"Richiesti {pick_count} {label}, ma disponibili solo {len(available)} elementi utili. Uso tutti quelli disponibili."
        )
        pick_count = len(available)

    rng = random.Random(seed)
    picked = rng.sample(list(available), pick_count) if pick_count else []
    return [ChecklistItem(i.text, i.source, kind, ticket_number=i.ticket_number, ticket_url=i.ticket_url) for i in picked], warnings


def _normalize_fixed_groups(items: Sequence[ChecklistItem]) -> list[ChecklistItem]:
    """Normalize the fixed checklist.

    The source Word has a varnishing parent line followed by bare sub-items (Basamento, Trave,
    Canotto, Trasporti, Siliconatura). Bare child rows are ambiguous in the final checklist, so the
    parent context is merged into each sub-item and the non-checkable parent row is removed.
    """
    result: list[ChecklistItem] = []
    i = 0
    while i < len(items):
        item = items[i]
        text_key = item.text.casefold()
        if "verniciatura" in text_key and "g001" in text_key:
            parent = item.text
            children: list[ChecklistItem] = []
            j = i + 1
            while j < len(items) and items[j].text.casefold().strip() in VERNICIATURA_CHILDREN:
                children.append(items[j])
                j += 1
            if children:
                for child in children:
                    result.append(ChecklistItem(f"{parent} - {child.text}", item.source, item.kind))
                i = j
                continue
        result.append(item)
        i += 1
    return result


def load_default_fixed_items() -> list[ChecklistItem]:
    """Load the mandatory checklist bundled with the application.

    The user does not need to select this file in the GUI. To change the fixed list,
    replace collaudo_suite/checklist/data/Check list.docx in the project/package.
    """
    path = get_default_fixed_docx_path()
    items = extract_items_from_docx(path, mode="fixed", source_label="Check list interna")
    normalized = _normalize_fixed_groups(items)
    return [ChecklistItem(item.text, item.source, "Fisso") for item in normalized]


def build_checklist(
    fixed_items: Sequence[ChecklistItem],
    random_pool: Sequence[ChecklistItem],
    random_count: int,
    *,
    seed: int | None = None,
    include_fixed_in_total: bool = False,
) -> tuple[list[ChecklistItem], list[str]]:
    """Return final checklist and warnings.

    Fixed items are always included. If include_fixed_in_total=True, random_count is interpreted as
    desired total count. When fixed items exceed it, fixed items still win and a warning is returned.
    """
    warnings: list[str] = []
    fixed = [ChecklistItem(i.text, i.source, "Fisso") for i in fixed_items]

    fixed_keys = {_norm_key(i.text) for i in fixed}
    pool = [i for i in random_pool if _norm_key(i.text) not in fixed_keys]

    if include_fixed_in_total:
        desired_total = max(1, random_count)
        if len(fixed) >= desired_total:
            pick_count = 0
            warnings.append(
                f"I controlli fissi sono {len(fixed)}, quindi superano o eguagliano il totale richiesto ({desired_total}). Ho mantenuto tutti i fissi."
            )
        else:
            pick_count = desired_total - len(fixed)
    else:
        pick_count = max(1, random_count)

    if pick_count > len(pool):
        warnings.append(
            f"Richiesti {pick_count} controlli casuali, ma disponibili solo {len(pool)} controlli utili. Uso tutti quelli disponibili."
        )
        pick_count = len(pool)

    rng = random.Random(seed)
    picked = rng.sample(list(pool), pick_count) if pick_count else []
    random_items = [ChecklistItem(i.text, i.source, "Random", ticket_number=i.ticket_number, ticket_url=i.ticket_url) for i in picked]
    return fixed + random_items, warnings


def records_from_items(items: Sequence[ChecklistItem]) -> list[ChecklistRecord]:
    return [ChecklistRecord(i.text, i.source, i.kind, ticket_number=i.ticket_number, ticket_url=i.ticket_url) for i in items]


def _record_fields(record: ChecklistRecord | ChecklistItem) -> ChecklistRecord:
    if isinstance(record, ChecklistRecord):
        return record
    return ChecklistRecord(record.text, record.source, record.kind, ticket_number=record.ticket_number, ticket_url=record.ticket_url)


def _pass_text(result: str, target: str) -> str:
    return "X" if result.casefold().strip() == target.casefold().strip() else ""


def _header_rows(header_info: Mapping[str, str] | None) -> list[tuple[str, str]]:
    info = dict(header_info or {})
    rows = [
        ("Collaudatore", info.get("collaudatore", "")),
        ("Reparto", info.get("reparto", "")),
        ("Data", info.get("data", "")),
        ("Numero Commessa", info.get("numero_commessa", info.get("tipo_macchina", ""))),
    ]
    return rows


def _header_value(header_info: Mapping[str, str] | None, key: str) -> str:
    return str((header_info or {}).get(key, "")).strip()




def _kind_section_label(kind: str) -> str:
    labels = {
        "Fisso": "CONTROLLI FISSI",
        "MAP": "SEGNALAZIONI MAP",
        "ANALYZER": "CONTROLLI IMPORTATI DA ANALYZER",
    }
    return labels.get(kind, kind.upper() if kind else "ALTRI CONTROLLI")

def _set_docx_cell_width(cell, width_inches: float) -> None:
    cell.width = Inches(width_inches)
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.first_child_found_in("w:tcW")
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(int(width_inches * 1440)))
    tc_w.set(qn("w:type"), "dxa")




def _set_docx_cell_shading(cell, fill: str = "E7E6E6") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.first_child_found_in("w:shd")
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)

def _set_docx_table_layout(table, widths_inches: Sequence[float]) -> None:
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    layout = tbl_pr.first_child_found_in("w:tblLayout")
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")

    tbl_w = tbl_pr.first_child_found_in("w:tblW")
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(int(sum(widths_inches) * 1440)))
    tbl_w.set(qn("w:type"), "dxa")

    old_grid = tbl.find(qn("w:tblGrid"))
    if old_grid is not None:
        tbl.remove(old_grid)
    grid = OxmlElement("w:tblGrid")
    for width in widths_inches:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(int(width * 1440)))
        grid.append(col)
    tbl.insert(0, grid)

def export_docx(
    items: Sequence[ChecklistRecord | ChecklistItem],
    output_path: str | Path,
    *,
    title: str = "Check list di collaudo",
    source_random: str = "",
    source_fixed: str = "Check list interna",
    header_info: Mapping[str, str] | None = None,
    seed: int | None = None,
) -> Path:
    output_path = Path(output_path)
    doc = Document()

    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.top_margin = Inches(0.45)
    section.bottom_margin = Inches(0.45)
    section.left_margin = Inches(0.45)
    section.right_margin = Inches(0.45)

    heading = doc.add_heading(title, level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    header_table = doc.add_table(rows=2, cols=4)
    header_table.style = "Table Grid"
    header_table.autofit = False
    header_widths = [2.25, 2.25, 1.4, 2.35]
    _set_docx_table_layout(header_table, header_widths)
    labels = ["Collaudatore", "Reparto", "Data", "Numero Commessa"]
    values = [
        _header_value(header_info, "collaudatore"),
        _header_value(header_info, "reparto"),
        _header_value(header_info, "data"),
        _header_value(header_info, "numero_commessa") or _header_value(header_info, "tipo_macchina"),
    ]
    for idx, (label, value, width) in enumerate(zip(labels, values, header_widths)):
        header_table.rows[0].cells[idx].text = label
        header_table.rows[1].cells[idx].text = value
        _set_docx_cell_width(header_table.rows[0].cells[idx], width)
        _set_docx_cell_width(header_table.rows[1].cells[idx], width)
        for paragraph in header_table.rows[0].cells[idx].paragraphs:
            for run in paragraph.runs:
                run.bold = True

    meta = doc.add_paragraph()
    meta.add_run("Generata il: ").bold = True
    meta.add_run(datetime.now().strftime("%d/%m/%Y %H:%M"))
    if source_random:
        meta.add_run(" | Sorgente random: ").bold = True
        meta.add_run(source_random)
    if source_fixed:
        meta.add_run(" | Checklist fissa: ").bold = True
        meta.add_run(source_fixed)
    if seed is not None:
        meta.add_run(" | Seed casuale: ").bold = True
        meta.add_run(str(seed))

    table = doc.add_table(rows=1, cols=7)
    table.style = "Table Grid"
    table.autofit = False
    hdr = table.rows[0].cells
    headers = ["N.", "Tipo", "Ticket", "Controllo", "Pass", "No pass", "Note"]
    widths = [0.35, 0.55, 0.75, 5.85, 0.5, 0.65, 1.95]
    _set_docx_table_layout(table, widths)
    for cell, header, width in zip(hdr, headers, widths):
        cell.text = header
        _set_docx_cell_width(cell, width)
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True

    last_kind = None
    display_idx = 1
    for item_obj in items:
        item = _record_fields(item_obj)
        if item.kind != last_kind:
            section_row = table.add_row().cells
            section_cell = section_row[0].merge(section_row[-1])
            section_cell.text = _kind_section_label(item.kind)
            _set_docx_cell_shading(section_cell, "E7E6E6")
            for paragraph in section_cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)
                for run in paragraph.runs:
                    run.bold = True
                    run.font.size = Pt(8)
            last_kind = item.kind

        row = table.add_row().cells
        row[0].text = str(display_idx)
        row[1].text = item.kind
        row[2].text = item.ticket_number
        row[3].text = item.text
        row[4].text = _pass_text(item.result, "Pass")
        row[5].text = _pass_text(item.result, "No pass")
        row[6].text = item.note
        for cell, width in zip(row, widths):
            _set_docx_cell_width(cell, width)
        display_idx += 1

    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)
                for run in paragraph.runs:
                    run.font.size = Pt(8)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


def export_csv(items: Sequence[ChecklistRecord | ChecklistItem], output_path: str | Path, *, header_info: Mapping[str, str] | None = None) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        for label, value in _header_rows(header_info):
            writer.writerow([label, value])
        writer.writerow([])
        writer.writerow(["N", "Tipo", "Ticket", "Controllo", "Pass", "No pass", "Note"])
        for idx, item_obj in enumerate(items, start=1):
            item = _record_fields(item_obj)
            writer.writerow([
                idx,
                item.kind,
                item.ticket_number,
                item.text,
                _pass_text(item.result, "Pass"),
                _pass_text(item.result, "No pass"),
                item.note,
            ])
    return output_path


def _pdf_paragraph(text: str, style):
    from reportlab.platypus import Paragraph

    parts = str(text).split("<br/>")
    safe = "<br/>".join(xml_escape(part).replace("\n", "<br/>") for part in parts)
    return Paragraph(safe, style)


def export_pdf(
    items: Sequence[ChecklistRecord | ChecklistItem],
    output_path: str | Path,
    *,
    title: str = "Check list di collaudo",
    source_random: str = "",
    source_fixed: str = "Check list interna",
    header_info: Mapping[str, str] | None = None,
    seed: int | None = None,
) -> Path:
    """Export a printable PDF. Pass / No pass cells are filled from the GUI state."""
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Spacer, Table, TableStyle

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=landscape(A4),
        leftMargin=10 * mm,
        rightMargin=10 * mm,
        topMargin=12 * mm,
        bottomMargin=11 * mm,
        title=title,
        author="Random Checklist Collaudo",
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "ChecklistTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=15,
        leading=18,
        alignment=TA_CENTER,
        spaceAfter=6,
    )
    meta_style = ParagraphStyle(
        "ChecklistMeta",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        alignment=TA_LEFT,
        spaceAfter=6,
    )
    header_style = ParagraphStyle(
        "ChecklistHeader",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.5,
        leading=9,
        alignment=TA_CENTER,
    )
    cell_style = ParagraphStyle(
        "ChecklistCell",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.3,
        leading=8.6,
        alignment=TA_LEFT,
    )
    center_style = ParagraphStyle(
        "ChecklistCenter",
        parent=cell_style,
        alignment=TA_CENTER,
        fontName="Helvetica-Bold",
    )

    header_data = [
        [_pdf_paragraph("Collaudatore", header_style), _pdf_paragraph("Reparto", header_style), _pdf_paragraph("Data", header_style), _pdf_paragraph("Numero Commessa", header_style)],
        [
            _pdf_paragraph(_header_value(header_info, "collaudatore"), cell_style),
            _pdf_paragraph(_header_value(header_info, "reparto"), cell_style),
            _pdf_paragraph(_header_value(header_info, "data"), cell_style),
            _pdf_paragraph(_header_value(header_info, "numero_commessa") or _header_value(header_info, "tipo_macchina"), cell_style),
        ],
    ]
    header_table = Table(header_data, colWidths=[70 * mm, 55 * mm, 35 * mm, 70 * mm])
    header_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EDEDED")),
        ("GRID", (0, 0), (-1, -1), 0.45, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))

    meta_lines = [f"Generata il: {datetime.now().strftime('%d/%m/%Y %H:%M')}"]
    if source_random:
        meta_lines.append(f"Sorgente random: {source_random}")
    if source_fixed:
        meta_lines.append(f"Checklist fissa: {source_fixed}")
    if seed is not None:
        meta_lines.append(f"Seed casuale: {seed}")
    meta_lines.append("I controlli Pass / No pass sono quelli selezionati nel programma.")

    data = [[
        _pdf_paragraph("N.", header_style),
        _pdf_paragraph("Tipo", header_style),
        _pdf_paragraph("Ticket", header_style),
        _pdf_paragraph("Controllo", header_style),
        _pdf_paragraph("Pass", header_style),
        _pdf_paragraph("No pass", header_style),
        _pdf_paragraph("Note", header_style),
    ]]
    section_style_commands = []
    last_kind = None
    display_idx = 1
    for item_obj in items:
        item = _record_fields(item_obj)
        if item.kind != last_kind:
            section_row_index = len(data)
            data.append([
                _pdf_paragraph(_kind_section_label(item.kind), header_style),
                "", "", "", "", "", "",
            ])
            section_style_commands.extend([
                ("SPAN", (0, section_row_index), (-1, section_row_index)),
                ("BACKGROUND", (0, section_row_index), (-1, section_row_index), colors.HexColor("#E7E6E6")),
                ("TEXTCOLOR", (0, section_row_index), (-1, section_row_index), colors.black),
                ("FONTNAME", (0, section_row_index), (-1, section_row_index), "Helvetica-Bold"),
                ("ALIGN", (0, section_row_index), (-1, section_row_index), "LEFT"),
            ])
            last_kind = item.kind
        data.append([
            _pdf_paragraph(str(display_idx), center_style),
            _pdf_paragraph(item.kind, center_style),
            _pdf_paragraph(item.ticket_number, center_style),
            _pdf_paragraph(item.text, cell_style),
            _pdf_paragraph(_pass_text(item.result, "Pass"), center_style),
            _pdf_paragraph(_pass_text(item.result, "No pass"), center_style),
            _pdf_paragraph(item.note, cell_style),
        ])
        display_idx += 1

    table = Table(
        data,
        colWidths=[9 * mm, 17 * mm, 20 * mm, 139 * mm, 18 * mm, 22 * mm, 50 * mm],
        repeatRows=1,
        splitByRow=True,
    )
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EDEDED")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("GRID", (0, 0), (-1, -1), 0.45, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (2, -1), "CENTER"),
        ("ALIGN", (4, 0), (5, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.2),
        ("LEFTPADDING", (0, 0), (-1, -1), 2.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2.5),
        ("TOPPADDING", (0, 0), (-1, -1), 3.0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.0),
        *section_style_commands,
    ]))

    story = [
        _pdf_paragraph(title, title_style),
        header_table,
        Spacer(1, 2 * mm),
        _pdf_paragraph("<br/>".join(meta_lines), meta_style),
        Spacer(1, 2 * mm),
        table,
    ]

    def _footer(canvas, document):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.drawRightString(document.pagesize[0] - 10 * mm, 6 * mm, f"Pag. {document.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    return output_path
