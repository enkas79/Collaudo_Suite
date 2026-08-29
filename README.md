# Collaudo Suite 1.1.8

Applicazione PySide6 che integra Analyzer anomalie, preparazione dei controlli e Checklist di collaudo.

## Modifiche della versione 1.1.8

- aggiunta la verifica automatica degli aggiornamenti: all'avvio la suite controlla in background (senza bloccare la GUI) l'ultima release pubblicata su GitHub e, se disponibile una versione più recente, propone di aprire la pagina di download;
- aggiunta la voce di menu **Verifica aggiornamenti** per lanciare il controllo manualmente in qualsiasi momento;
- la versione dell'applicazione viene ora letta dinamicamente da `version.txt` nella root del progetto (o nella cartella dell'eseguibile per le build Windows), anziché essere fissata nel codice;
- aggiunto il workflow GitHub Actions `.github/workflows/build-installers.yml`: a ogni push su `main` che tocca `version.txt` o il codice, compila l'eseguibile Windows con PyInstaller, lo comprime in `CollaudoSuite-<versione>-win64.zip` e pubblica automaticamente una GitHub Release (tag `v<versione>`) con l'archivio allegato — è la release che la verifica aggiornamenti nell'app va a controllare.

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

Il workflow `.github/workflows/build-installers.yml` si attiva a ogni push su `main` che modifica `version.txt` (oppure manualmente da GitHub Actions). Compila l'eseguibile Windows con PyInstaller, genera un vero installer con NSIS (script in `packaging\windows\installer.nsi`) e pubblica una GitHub Release con tag `v<versione presa da version.txt>` e `CollaudoSuite-Setup-<versione>.exe` come asset. L'installer crea le voci nel menu Start, il collegamento sul desktop e una disinstallazione da Pannello di controllo. La funzione "Verifica aggiornamenti" dell'app legge proprio questa release per proporre il download della nuova versione.
