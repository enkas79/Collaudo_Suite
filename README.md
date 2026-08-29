# Collaudo Suite 1.1.7

Applicazione PySide6 che integra Analyzer anomalie, preparazione dei controlli e Checklist di collaudo.

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
