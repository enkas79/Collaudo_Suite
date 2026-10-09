# Download giornali macchina da SharePoint - Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** L'Analyzer scarica da SharePoint Breton i giornali macchina del modello scelto dalla cartella `{anno}_Giornale_Macchina` e li passa all'analisi esistente.

**Architecture:** Modulo `sharepoint.py` senza Qt (login MSAL + Microsoft Graph via `requests`), un `QThread` worker che lo esegue fuori dal main thread, e un blocco "Sorgente SharePoint" nella sidebar dell'Analyzer che riempie `selected_files`.

**Tech Stack:** Python 3.11, PySide6, msal, requests, pytest (unittest.mock).

**Spec:** `docs/superpowers/specs/2026-10-09-sharepoint-giornali-macchina-design.md`

## Global Constraints

- Commenti e testi UI in italiano; type hints obbligatori; GUI separata dalla logica.
- Nessuna operazione di rete nel thread GUI.
- Struttura: `{anno}_Giornale_Macchina/<sottocartella modello>/*.xlsx`; match sottocartella case-insensitive per inclusione; anno da `date.today().year`.
- Sito di prova: `https://breton.sharepoint.com/sites/CorsoFesto`, libreria `BRETONFS1`, cartella base `Test_Martini`.
- Token mai nei log; cache token in `%LOCALAPPDATA%\CollaudoSuite`.
- `version.txt` 1.4.1 -> 1.5.0; aggiungere `msal` e `requests` a `pyproject.toml` (dependencies) e il modulo `collaudo_suite.analyzer` e' gia' in `packages`.
- Test da `tests/` con `unittest`/pytest, `QT_QPA_PLATFORM=offscreen`.

## Review Focus

- Cartella dell'anno assente: errore chiaro `SharePointError`, nessun fallback automatico.
- Sottocartelle modello multiple che corrispondono (es. `NC300` e `NC300_old`): preferire il nome esatto, altrimenti la prima in ordine alfabetico.
- Risultati Graph paginati (`@odata.nextLink`): elencare tutte le pagine.
- File non xlsx o temporanei (`~$...xlsx`) nella sottocartella: ignorati.
- Nomi file duplicati o con caratteri non validi: il download non deve sovrascrivere ne' fallire.

---

### Task 1: Configurazione e logica di ricerca (pura)

**Files:**
- Create: `collaudo_suite/sharepoint_config.py`
- Create: `collaudo_suite/sharepoint.py` (parte pura)
- Test: `tests/test_sharepoint.py`

**Interfaces:**
- Produces:
  - `@dataclass(frozen=True) SharePointConfig(site_url: str, library: str, base_folder: str, client_id: str, tenant: str)` e `load_config() -> SharePointConfig` (legge `%LOCALAPPDATA%\CollaudoSuite\sharepoint.json` se esiste, altrimenti default: site `https://breton.sharepoint.com/sites/CorsoFesto`, library `BRETONFS1`, base_folder `Test_Martini`, client_id `14d82eec-204b-4c2f-b7e8-296a70dab67e`, tenant `organizations`).
  - `class SharePointError(Exception)`.
  - `year_folder_name(year: int) -> str` -> `"2026_Giornale_Macchina"`.
  - `pick_model_folder(names: list[str], model: str) -> str` (esatto case-insensitive, altrimenti per inclusione, a parita' ordine alfabetico; `SharePointError` se nessuno).
  - `filter_xlsx(names: list[str]) -> list[str]` (solo `.xlsx`/`.xlsm`, esclude `~$*`).
  - `unique_local_name(name: str, taken: set[str]) -> str` (sanifica caratteri `<>:"/\|?*`, aggiunge `_2`, `_3` se duplicato).

- [ ] **Step 1: Test** in `tests/test_sharepoint.py`: `year_folder_name(2026) == "2026_Giornale_Macchina"`; `pick_model_folder(["NC300_old","nc300","Genya"], "NC300") == "nc300"`; `pick_model_folder(["Genya_2","Genya_1"], "genya") == "Genya_1"`; `pick_model_folder(["Trinity"], "NC300")` solleva `SharePointError`; `filter_xlsx(["a.xlsx","~$a.xlsx","b.pdf","c.XLSM"]) == ["a.xlsx","c.XLSM"]`; `unique_local_name("a:b.xlsx", {"a_b.xlsx"}) == "a_b_2.xlsx"`; `load_config()` con `LOCALAPPDATA` puntato a una tmp dir vuota restituisce i default.
- [ ] **Step 2:** `pytest tests/test_sharepoint.py -v` -> FAIL (modulo mancante).
- [ ] **Step 3:** Implementare le funzioni sopra nei due file.
- [ ] **Step 4:** `pytest tests/test_sharepoint.py -v` -> PASS.
- [ ] **Step 5:** `git add collaudo_suite/sharepoint_config.py collaudo_suite/sharepoint.py tests/test_sharepoint.py && git commit -m "feat: logica ricerca giornali macchina SharePoint"`

### Task 2: SharePointClient (login + Graph)

**Files:**
- Modify: `collaudo_suite/sharepoint.py`
- Modify: `pyproject.toml` (dependencies: `msal>=1.28`, `requests>=2.31`)
- Test: `tests/test_sharepoint.py`

**Interfaces:**
- Consumes: Task 1.
- Produces:
  - `class SharePointClient(config: SharePointConfig, session: requests.Session | None = None, token_provider: Callable[[], str] | None = None)`. `token_provider` default = login MSAL `PublicClientApplication` (authority `https://login.microsoftonline.com/{tenant}`, scope `["Files.Read.All","Sites.Read.All"]`, prima `acquire_token_silent` dalla `SerializableTokenCache` salvata in `%LOCALAPPDATA%\CollaudoSuite\msal_cache.bin`, poi `acquire_token_interactive`); errori -> `SharePointError`.
  - `list_children(folder_path: str) -> list[dict]` (path relativo alla libreria; GET `/drives/{drive_id}/root:/{path}:/children` seguendo `@odata.nextLink`; risolve e memorizza site id e drive id: GET `/sites/{host}:{site_path}`, poi `/sites/{id}/drives` scegliendo il drive il cui `webUrl` termina con `/{library}` o `name == library`).
  - `download(item: dict, dest_dir: Path, taken: set[str]) -> Path` (usa `@microsoft.graph.downloadUrl` senza header Authorization).
  - `download_machine_files(client, model: str, year: int, dest_dir: Path, progress: Callable[[int], None] = lambda _: None) -> list[Path]`: lista `base_folder`, cerca `year_folder_name(year)` (esatto case-insensitive, altrimenti `SharePointError` "Cartella {nome} non trovata"), lista le sottocartelle, `pick_model_folder`, lista i file, `filter_xlsx`, scarica; lista vuota -> `SharePointError`.

- [ ] **Step 1: Test** con `session` mockata (`unittest.mock.MagicMock`, `token_provider=lambda: "tok"`): (a) `download_machine_files` su struttura finta scarica solo gli xlsx della sottocartella `Genya` e ritorna 2 Path esistenti; (b) cartella anno assente -> `SharePointError` con testo contenente `2026_Giornale_Macchina`; (c) `list_children` con due pagine (`@odata.nextLink`) ritorna tutti gli elementi; (d) la richiesta di download del file non include header `Authorization`; (e) nessun xlsx -> `SharePointError`; (f) il token non compare in nessun messaggio di eccezione.
- [ ] **Step 2:** `pytest tests/test_sharepoint.py -v` -> i nuovi test FAIL.
- [ ] **Step 3:** Implementare `SharePointClient` e `download_machine_files`; timeout `requests` 30 s; `raise_for_status` convertito in `SharePointError` con codice HTTP (senza token). Import di `msal` dentro il token provider di default, per non rompere i test se manca.
- [ ] **Step 4:** `pytest tests/test_sharepoint.py -v` -> PASS.
- [ ] **Step 5:** `git add collaudo_suite/sharepoint.py pyproject.toml tests/test_sharepoint.py && git commit -m "feat: client SharePoint con login MSAL e Graph"`

### Task 3: Worker Qt

**Files:**
- Create: `collaudo_suite/analyzer/sharepoint_worker.py`
- Test: `tests/test_sharepoint.py`

**Interfaces:**
- Consumes: `download_machine_files`, `SharePointClient`, `load_config`, `SharePointError`.
- Produces: `class SharePointWorker(QThread)` con `__init__(self, model: str, year: int, dest_dir: Path, client: SharePointClient | None = None)`; segnali `log_message = Signal(str)`, `progress_update = Signal(int)`, `files_ready = Signal(list)` (lista di `str`), `error_occurred = Signal(str)`. Mai eccezioni non gestite: ogni `Exception` -> `error_occurred`.

- [ ] **Step 1: Test** (offscreen, `worker.run()` chiamato direttamente senza thread) con `client` finto: successo emette `files_ready` con i percorsi; `SharePointError("x")` emette `error_occurred` contenente `x`; un `RuntimeError` generico emette comunque `error_occurred`.
- [ ] **Step 2:** `pytest tests/test_sharepoint.py -k worker -v` -> FAIL.
- [ ] **Step 3:** Implementare il worker sul modello di `collaudo_suite/analyzer/worker.py`.
- [ ] **Step 4:** `pytest tests/test_sharepoint.py -v` -> PASS.
- [ ] **Step 5:** `git add collaudo_suite/analyzer/sharepoint_worker.py tests/test_sharepoint.py && git commit -m "feat: worker download SharePoint"`

### Task 4: GUI Analyzer, versione e verifica

**Files:**
- Modify: `collaudo_suite/analyzer/app.py` (sezione INPUT dopo `lbl_files`, riga ~96; nuovi metodi accanto a `select_files`, riga ~310; `closeEvent` per fermare il worker)
- Modify: `collaudo_suite/styles.py` solo se serve uno stile (usare gli object name esistenti `btnMiniHelp`/pulsanti standard se possibile)
- Modify: `version.txt` -> `1.5.0`
- Modify: `collaudo_suite/help_dialog.py` (breve paragrafo sulla sorgente SharePoint)
- Test: `tests/test_gui.py`

**Interfaces:**
- Consumes: `SharePointWorker`.
- Produces: in `AnalyzerWindow`: `combo_sp_model: QComboBox` (editabile; voci `GENYA`, `NC300`, `TRINITY`, `EVONIX`), `btn_sp_download: QPushButton("Scarica da SharePoint")`, metodi `download_from_sharepoint() -> None` (disabilita il pulsante, avvia il worker con `dest_dir = tempfile.mkdtemp(prefix="giornali_")`, `year=date.today().year`), `_on_sharepoint_files(files: list[str]) -> None` (imposta `selected_files` e `lbl_files` come `select_files`, riabilita il pulsante), `_on_sharepoint_error(message: str) -> None` (`QMessageBox.warning`, riabilita il pulsante). Il log del worker va su `_append_log`.

- [ ] **Step 1: Test** in `tests/test_gui.py` (stile esistente): dopo `window._on_sharepoint_files(["a.xlsx","b.xlsx"])` `window.selected_files == ["a.xlsx","b.xlsx"]` e `btn_sp_download.isEnabled()`; `combo_sp_model` contiene `GENYA` e `NC300`.
- [ ] **Step 2:** `pytest tests/test_gui.py -k sharepoint -v` -> FAIL.
- [ ] **Step 3:** Implementare il blocco UI (spaziature multipli di 4/8, etichetta di sezione "SORGENTE SHAREPOINT") e i tre metodi; fermare/attendere il worker in `closeEvent`.
- [ ] **Step 4:** Aggiornare `version.txt` a `1.5.0` e `help_dialog.py`; eseguire `pytest -q` (tutta la suite) e `ruff check . --fix` -> nessun errore nuovo.
- [ ] **Step 5:** Verifica manuale: `python run_suite.py`, Analyzer, modello `GENYA`, "Scarica da SharePoint": login Microsoft, poi `lbl_files` mostra i file scaricati e l'analisi parte. Riportare l'esito reale, inclusi eventuali blocchi del tenant sul Client ID.
- [ ] **Step 6:** `git add -A collaudo_suite tests version.txt && git commit -m "feat: Analyzer scarica i giornali macchina da SharePoint (1.5.0)"`; poi push e merge su `master` come da CLAUDE.md.
