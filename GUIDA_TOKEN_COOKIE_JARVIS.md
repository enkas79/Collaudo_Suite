# Guida JARVIS: token, cookie e download dei Giornali Macchina

## 1. Cosa fa il downloader

Il programma cerca esclusivamente nella sezione **Assets** di JARVIS. Per ogni Asset:

1. esegue la ricerca Assets;
2. legge i documenti collegati all'Asset;
3. riconosce `giornale_macchina`, `Giornale Macchina`, `GIORNALE-MACCHINA` e varianti equivalenti;
4. scarica soltanto file Excel: `.xls`, `.xlsx`, `.xlsm`, `.xlsb`, `.xlt`, `.xltx`, `.xltm`.

PDF, DOC, TXT e altri formati vengono ignorati.

## 2. Permessi del token

Quando si crea il token JARVIS, abilitare almeno:

- `wbs-read`: ricerca degli Assets/WBS;
- `documents-read`: lettura dei documenti collegati;
- `external-documents-download`: download dei file.

Il token API può essere inserito nel campo **Token JARVIS** della GUI. Non incollare il comando cURL completo: serve solo il valore del token.

## 3. Cookie di sessione

Gli endpoint interni usati dalla pagina Assets possono richiedere la sessione browser. In questo caso usare il cookie `AuthCookie`.

### Recupero da Edge/Chrome

1. Aprire `https://jarvis.breton.it/UI/#/assets/search` e accedere.
2. Premere `F12`.
3. Aprire **Application** → **Storage** → **Cookies** → `https://jarvis.breton.it`.
4. Cercare `AuthCookie`.
5. Copiare esclusivamente il valore del cookie.

Nel campo **Cookie sessione** della GUI incollare il valore senza il prefisso `AuthCookie=`.

Il cookie è una credenziale temporanea: non inserirlo in ticket, chat, screenshot, repository o file condivisi. Dopo averlo condiviso per errore, disconnettersi da JARVIS o revocare la sessione e copiarne uno nuovo.

La GUI salva token e cookie cifrati con DPAPI nel profilo dell'utente Windows. Non vengono salvati nel progetto, nel README o nei log. Se si svuota un campo prima di avviare una ricerca, il relativo segreto salvato viene rimosso.

## 4. Uso dalla GUI

1. Avviare `avvia_giornali_macchina.bat`.
2. Inserire il token nel campo **Token JARVIS**.
3. Se necessario, inserire il valore di `AuthCookie` nel campo **Cookie sessione**.
4. Inserire una **WBS / macchina** per limitare la ricerca; lasciando vuoto il campo vengono analizzati tutti gli Assets.
5. Selezionare la cartella di destinazione.
6. Usare prima **Simula la ricerca senza scaricare i file**.
7. Quando i risultati sono corretti, disattivare la simulazione e avviare il download.

Durante la ricerca è possibile premere **Stop**. L'arresto è cooperativo: la richiesta HTTP o il file eventualmente in corso viene completato, poi il programma interrompe la scansione senza creare file parziali.

## 5. Uso dalla riga di comando

Con token e cookie passati direttamente:

```bat
python scarica_giornali_macchina.py --token IL_TOKEN --auth-cookie IL_COOKIE --wbs 86249 --output Giornali_Macchina
```

È preferibile usare variabili d'ambiente, così token e cookie non finiscono nella cronologia del terminale:

```bat
set "JARVIS_AUTH_TOKEN=IL_TOKEN"
set "JARVIS_AUTH_COOKIE=IL_COOKIE"
python scarica_giornali_macchina.py --wbs 86249 --output Giornali_Macchina
```

Per una prova senza scrivere i file:

```bat
python scarica_giornali_macchina.py --wbs 86249 --dry-run
```

## 6. Interpretazione del log

- `POST /api/v1/omnisearch/search`: ricerca degli Assets;
- `POST /api/v2/SystemDocuments/Search`: lettura dei documenti dell'Asset;
- `[MATCH]`: trovato un documento/file Excel compatibile;
- `[OK]`: file scaricato;
- `HTTP 403`: token o cookie non autorizzato/scaduto;
- `HTTP 401`: autenticazione mancante o non valida.

I file vengono salvati nella cartella scelta, eventualmente organizzati per tipologia macchina e per Asset.
