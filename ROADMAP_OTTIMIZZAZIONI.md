# Roadmap ottimizzazioni — Collaudo Suite

Analisi del codice (`collaudo_suite/`) alla ricerca di colli di bottiglia prestazionali e violazioni delle linee guida di progetto (threading GUI). Nessuna modifica funzionale è stata applicata: questo documento è la base per pianificare gli interventi.

## 1. Priorità alta

### 1.1 Caricamento MAP/PDF/DOCX sul thread principale della GUI
**File:** `collaudo_suite/checklist/app.py` (`load_map_items_for_filter`, `export_pdf`, `export_docx`, `export_csv`, `extract_items_from_docx`)

Queste operazioni I/O (lettura file XLSX/DOCX potenzialmente grandi, generazione PDF/DOCX) vengono eseguite in modo sincrono nel thread GUI, senza `QThread`/`Signal` come invece già avviene in `analyzer/worker.py`. Con file MAP o checklist voluminosi l'interfaccia si blocca (violazione della regola CLAUDE.md sul threading).

**Intervento proposto:** introdurre un worker `QThread` dedicato per `load_map_items_for_filter` e per l'export (PDF/DOCX/CSV), replicando il pattern già usato in `analyzer/worker.py`, con segnali di progresso/errore.

### 1.2 Doppia lettura del workbook MAP in modalità non ottimizzata
**File:** `collaudo_suite/checklist/core.py` (`_map_row_metadata`, righe 761-812)

`extract_map_items_from_xlsx` legge già le righe tramite `_xlsx_rows_for_path`, che è **cachata** con `lru_cache` e invalidata su mtime/size (ottimo pattern). Tuttavia, subito dopo, `_map_row_metadata` riapre lo stesso file con `openpyxl.load_workbook(..., read_only=False)` per estrarre hyperlink/date — una seconda parsificazione completa del workbook, **non cachata**, ripetuta a ogni cambio di filtro o periodo.

**Intervento proposto:**
- Cachare anche l'output di `_map_row_metadata` con la stessa chiave (path, mtime, size) usata da `_cached_xlsx_rows`.
- Valutare `read_only=True` in `load_workbook` per ridurre memoria/tempo — da verificare con test che gli hyperlink (`cell.hyperlink`) restino leggibili in quella modalità, perché in read-only openpyxl alcuni metadati per-cella hanno un percorso diverso.

### 1.3 Lettura riga-per-riga degli Excel in ingresso all'Analyzer
**File:** `collaudo_suite/analyzer/data_loader.py` (`collect_records`, righe 113-159)

Per ogni riga del foglio si accede a `df.iat[row_idx, col]` e si chiama `pd.to_datetime(...)` singolarmente. Su fogli con decine di migliaia di righe questo è molto più lento delle operazioni vettoriali di pandas.

**Intervento proposto:** estrarre le colonne target/data come `Series` (`df.iloc[start:, col]`), applicare `.astype(str).str.strip()` e `pd.to_datetime(..., errors="coerce")` in blocco, poi iterare solo sui valori già puliti per costruire i `RowRecord`. Da fare con attenzione: il calcolo di `dayfirst` oggi è per singola cella in base al formato testuale della data — andrebbe preservato o generalizzato prima di vettorizzare.

## 2. Priorità media

### 2.1 `_extract_structured_table_controls` — controllo titoli ripetuti O(n²) per riga
**File:** `collaudo_suite/checklist/core.py`, riga 233

```python
most_common_count = max(nonempty_cells.count(c) for c in set(nonempty_cells))
```
Per ogni riga della tabella Word si ricalcola il conteggio con un doppio scan. Le tabelle di collaudo hanno poche colonne quindi l'impatto reale è trascurabile, ma può essere reso lineare con `collections.Counter(nonempty_cells).most_common(1)`.

### 2.2 Analyzer: limite di confronto esaustivo
**File:** `collaudo_suite/analyzer/similarity.py` (`candidate_pairs`, `exhaustive_limit=1800`)

Il blocking per bucket è già implementato correttamente per dataset grandi. Sotto la soglia di 1800 record il confronto resta O(n²) (~1,6M coppie al limite), che può risultare pesante su hardware datato. Non è un bug, ma se in pratica si osservano rallentamenti con dataset di 1000-1800 righe, abbassare la soglia di `exhaustive_limit` o estendere il blocking anche a quell'intervallo è un intervento a basso rischio.

## 3. Priorità bassa / osservazioni

- `collaudo_suite/checklist/core.py`: buon uso di `lru_cache` per il MAP e di euristiche di blocking nell'Analyzer — pattern da mantenere come riferimento per i punti 1.2 e 1.3.
- Nessun problema di sicurezza rilevato nei percorsi analizzati (niente `eval`, comandi shell, o concatenazioni SQL/HTML non sanificate: l'export PDF usa correttamente `xml_escape`).

## Prossimi passi consigliati

1. Implementare 1.1 (worker per MAP/export) — impatto UX più alto, rischio di regressione basso se si replica il pattern esistente in `analyzer/worker.py`.
2. Implementare 1.2 (cache metadati MAP) — impatto prestazionale immediato su ogni cambio filtro/periodo, rischio basso.
3. Valutare 1.3 dopo aver raccolto un caso reale con Excel di grandi dimensioni, per non introdurre regressioni sul parsing delle date.
