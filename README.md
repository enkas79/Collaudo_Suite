# Collaudo Suite

Applicazione PySide6 che integra Analyzer anomalie, preparazione dei controlli e Checklist di collaudo.

La versione corrente è indicata in `version.txt`, unica fonte usata da applicazione, `pyproject.toml` e workflow di build.

## Modifiche della versione 1.3.1

- nuova icona dell'applicazione (finestra, barra delle applicazioni, eseguibile e installer), generata da `tools/build_icon.py` in `collaudo_suite/assets/app_icon.png` / `app_icon.ico`;
- autore indicato in *Aiuto > Informazioni*, nei metadati del pacchetto e nell'installer: **Enrico Martini**.

## Modifiche della versione 1.3.0

- rimossa la checklist fissa interna (`Check list.docx`): la checklist è ora composta solo dalle segnalazioni MAP estratte e dai controlli importati dall'Analyzer; nei lavoro `.rcl.json` salvati con versioni precedenti le righe "Fisso" vengono ignorate all'apertura;
- la **Guida** apre un manuale PDF integrato (`collaudo_suite/assets/Guida_operativa_Collaudo_Suite.pdf`) con scorrimento e zoom; il PDF si rigenera da `collaudo_suite/help_dialog.py` con `python tools/build_help_pdf.py`;
- JARVIS: il codice commerciale viene letto dal campo nativo OmniSearch `propertydefinition_2674/matnr` e la ricerca usa il filtro Commercial code dell'API; scartati i riferimenti opachi (`DatasetElement_…`) come codici; se la cache Excel è aperta in un altro programma, viene salvata con un nuovo nome invece di fallire;
- Checklist: sezioni JARVIS ed Excel mostrate solo per la sorgente selezionata, riepilogo con conteggi Pass / No pass / da verificare, celle esito colorate, link ticket ricavato dalla cella visibile;
- menu *Strumenti*: la voce diventa **Info file MAP**.

## Modifiche della versione 1.2.0

- corretto il popup delle tendine nella barra laterale dell'Analyzer (*Anomalie negli ultimi*, algoritmo di analisi): le opzioni erano bianche su fondo chiaro e quindi illeggibili; ora il popup usa la palette scura della barra laterale;
- barra dei menu riorganizzata: **File** (gli stessi comandi della toolbar), **Strumenti** (*Info file interni*) e **Aiuto** con *Guida* (F1), *Controlla aggiornamenti* e *Informazioni su Collaudo Suite* (autore e versione letta da `version.txt`);
- aggiornamento dall'app: su Windows il pulsante **Scarica e installa** scarica in background l'installer della nuova release (con avanzamento e possibilità di annullare), salva il lavoro, chiude la suite e avvia l'installazione;
- *Esporta PDF* e *Stampa PDF* generano il file in un thread separato: l'interfaccia non si blocca più durante la creazione del PDF;
- stili dell'interfaccia (QSS) centralizzati nel modulo `collaudo_suite/styles.py`;
- `pyproject.toml` legge la versione da `version.txt`.

## Modifiche della versione 1.1.9

- corretto un crash ("Internal C++ object (UpdateCheckWorker) already deleted") che si verificava premendo **Verifica aggiornamenti** dopo il primo controllo (automatico o manuale): lo stato del controllo ora è tracciato con un flag dedicato invece di interrogare un `QThread` potenzialmente già distrutto da `deleteLater`.

## Modifiche della versione 1.1.8

- aggiunta la verifica automatica degli aggiornamenti: all'avvio la suite controlla in background (senza bloccare la GUI) l'ultima release pubblicata su GitHub e, se disponibile una versione più recente, propone di aprire la pagina di download;
- aggiunta la voce di menu **Verifica aggiornamenti** per lanciare il controllo manualmente in qualsiasi momento;
- la versione dell'applicazione viene ora letta dinamicamente da `version.txt` nella root del progetto (o nella cartella dell'eseguibile per le build Windows), anziché essere fissata nel codice;
- aggiunto il workflow GitHub Actions `.github/workflows/build-installers.yml`: a ogni push su `master` che tocca `version.txt` (o manualmente), compila l'eseguibile Windows con PyInstaller, genera l'installer NSIS `CollaudoSuite-Setup-<versione>.exe` (script in `packaging\windows\installer.nsi`) e pubblica automaticamente una GitHub Release (tag `v<versione>`) con l'installer allegato — è la release che la verifica aggiornamenti nell'app va a controllare.

## Modifiche della versione 1.1.7

- aggiunta configurazione `pyproject.toml` con dipendenze, entry point `collaudo-suite` e dati package inclusi;
- corretto il caching del file MAP: se il riepilogativo Excel viene sostituito o aggiornato mantenendo lo stesso percorso, la checklist rilegge i dati aggiornati;
- aggiunto test di regressione per verificare che la cache MAP si invalidi quando il file cambia;
- mantenuta compatibilità con `python run_suite.py` e con il flusso di build Windows esistente.

## Modifiche della versione 1.1.6

- aggiunta la scelta del periodo temporale MAP: **1 mese, 3 mesi, 6 mesi oppure 1 anno**;
- il periodo viene calcolato a ritroso usando come riferimento la **Data di collaudo** dell'intestazione;
- sono inclusi i MAP con data compresa tra il limite iniziale e il giorno del collaudo, estremi inclusi;
- i MAP successivi alla data del collaudo e quelli senza una data valida vengono esclusi;
- riconoscimento automatico delle principali colonne data, tra cui `Last Modify`, `Last Modified`, `Data ultima modifica`, `Created Date`, `Data ticket` e `Data`;
- periodo selezionato memorizzato nelle preferenze e nei file lavoro `.rcl.json`;
- mantenute le funzioni della 1.1.5: collegamento ticket con un clic, ordinamento decrescente, toolbar superiore, Salva e Salva con nome.

Il valore predefinito e **1 anno**, per ridurre il rischio di restringere eccessivamente il pool MAP al primo utilizzo.

## Scorciatoie

- `Ctrl+S`: Salva lavoro
- `Ctrl+Shift+S`: Salva lavoro con nome
- `Ctrl+O`: Apri lavoro
- `Ctrl+P`: Stampa PDF
- `Ctrl+Q`: Esci

## Flusso consigliato

1. Selezionare la data effettiva del collaudo.
2. Scegliere Commercial code, periodo MAP e file riepilogativo.
3. Premere **Estrai / Aggiorna anteprima**.
4. Compilare Pass/No pass e note.
5. Usare le icone nella barra superiore per salvare, aprire o produrre il PDF.

## Regola temporale MAP

Esempio: data collaudo `24/07/2026` e periodo `1 mese`.
Sono ammessi i MAP dal `24/06/2026` al `24/07/2026`, inclusi. Un MAP del `25/07/2026` viene escluso perche successivo al collaudo.

## Sorgente MAP JARVIS

Nella scheda Checklist è possibile selezionare `JARVIS API` al posto del file Excel. La modalità API usa gli endpoint pubblici di ricerca ticket con un token locale avente il permesso `ticket.read`. Il codice viene letto dalla proprietà equivalente alla colonna `Commessa S Codice commerciale` del riepilogativo Excel. Il filtro `Commercial code` cerca il testo anche all'interno del codice: ad esempio `300` trova `NC30000068`.

Il pulsante `Aggiorna MAP da JARVIS` crea una cache Excel locale con i ticket e le relative proprietà. Le anteprime successive leggono la cache, senza ripetere le chiamate API. Dopo il primo caricamento, agli avvii successivi la cache viene aggiornata automaticamente in background e mantiene anche le righe già presenti. Il token non viene inserito nei file lavoro `.rcl.json`; viene mantenuto nelle impostazioni locali dell'applicazione. Se il token non è configurato o JARVIS non è raggiungibile, la modalità Excel resta disponibile come alternativa.

## Installazione da sorgente

```bat
py -3 -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python run_suite.py
```

Oppure, con un ambiente Python moderno:

```bat
python -m pip install -e .
collaudo-suite
```

## Creazione eseguibile Windows

Eseguire `build_windows.bat`. Il risultato viene creato in `dist\CollaudoSuite\CollaudoSuite.exe`.

## Build e release automatiche (CI)

Il workflow `.github/workflows/build-installers.yml` si attiva a ogni push su `master` che modifica `version.txt` (oppure manualmente da GitHub Actions). Compila l'eseguibile Windows con PyInstaller, genera un vero installer con NSIS (script in `packaging\windows\installer.nsi`) e pubblica una GitHub Release con tag `v<versione presa da version.txt>` e `CollaudoSuite-Setup-<versione>.exe` come asset. L'installer crea le voci nel menu Start, il collegamento sul desktop e una disinstallazione da Pannello di controllo. La funzione "Verifica aggiornamenti" dell'app legge proprio questa release per proporre il download della nuova versione.

## Downloader Giornali Macchina

Lo script autonomo `scarica_giornali_macchina.py` cerca esclusivamente nella sezione Assets di JARVIS, apre i documenti associati a ciascun Asset e scarica quelli denominati `Giornale Macchina`. Sono accettate le varianti di scrittura (`giornale_macchina`, spazi, trattini e maiuscole/minuscole), ma vengono scaricati soltanto file Excel.

È disponibile anche la versione grafica:

```bat
python scarica_giornali_macchina_gui.py
```

La GUI permette di inserire il token, scegliere la cartella di destinazione, simulare la ricerca e visualizzare il log senza bloccare la finestra. È possibile attivare l'organizzazione automatica per tipologia macchina (`NC300`, `TRINITY`, `GENYA`, `EVONIX`, `VIPER`, `EAGLE`, `WM`, `WP`); i casi non riconosciuti vengono salvati in `NON_IDENTIFICATA`.

```bat
python scarica_giornali_macchina.py --token IL_TOKEN --output Giornali_Macchina
```

Il token può essere passato anche con la variabile `JARVIS_AUTH_TOKEN`. Per verificare prima le corrispondenze senza scaricare file usare `--dry-run`.

Per le API interne della pagina Assets può essere necessario il cookie di sessione del browser. Usare la variabile temporanea `JARVIS_AUTH_COOKIE` oppure il campo "Cookie sessione" della GUI; inserire solo il valore del cookie, senza `AuthCookie=`. Non salvare o condividere il cookie.

La guida completa per token, cookie e troubleshooting è in [GUIDA_TOKEN_COOKIE_JARVIS.md](GUIDA_TOKEN_COOKIE_JARVIS.md).
