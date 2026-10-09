# Analyzer: download giornali macchina da SharePoint

## Obiettivo
L'Analyzer scarica in automatico dal SharePoint Breton i giornali macchina (xlsx) del modello selezionato (GENYA, NC300, TRINITY, EVONIX...) dalla cartella `{anno}_Giornale_Macchina` (anno = data odierna) e li passa all'analisi esistente. Sito di prova: `https://breton.sharepoint.com/sites/CorsoFesto`, libreria `BRETONFS1`, cartella `Test_Martini`.

## Decisioni
- Autenticazione: login Microsoft interattivo (MSAL, public client) con account aziendale; nessuna password gestita dall'app; cache token in `%LOCALAPPDATA%\CollaudoSuite`.
- Accesso ai dati: Microsoft Graph via `requests`.
- Struttura: `{anno}_Giornale_Macchina/<sottocartella modello>/*.xlsx`. Match sottocartella case-insensitive per inclusione.
- Selezione manuale dei file Excel resta disponibile.

## Moduli
- `collaudo_suite/sharepoint_config.py`: URL sito, libreria, cartella base, Client ID, tenant (JSON in `%LOCALAPPDATA%\CollaudoSuite\sharepoint.json`, con default di prova).
- `collaudo_suite/sharepoint.py`: `SharePointClient` (login, risoluzione site/drive, list, download), `find_year_folder`, `pick_model_folder`, `download_machine_files`. Nessuna dipendenza Qt.
- `collaudo_suite/analyzer/sharepoint_worker.py`: `QThread` con segnali `log_message`, `progress_update`, `files_ready(list)`, `error_occurred(str)`.
- `collaudo_suite/analyzer/app.py`: blocco "Sorgente SharePoint" in INPUT (combo modello + pulsante); a download finito riempie `selected_files`.

## Errori
Login annullato, rete assente, cartella anno o sottocartella modello assente, nessun xlsx: `QMessageBox.warning`, nessun crash. Token mai nei log.

## Test
pytest con `requests` mockato: ricerca cartella anno, scelta sottocartella, filtro xlsx. Login reale verificato a mano.

## Dipendenze / versione
Aggiunte `msal`, `requests`; `version.txt` 1.4.1 -> 1.5.0.
