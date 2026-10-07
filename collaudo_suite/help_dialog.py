from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QTabWidget,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)


class SuiteHelpDialog(QDialog):
    """Guida operativa unificata di Collaudo Suite."""

    def __init__(self, parent: QWidget | None = None, initial_tab: int = 0) -> None:
        super().__init__(parent)
        self.setWindowTitle("Guida operativa - Collaudo Suite")
        self.resize(940, 720)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._page(self._workflow_html()), "Processo completo")
        self.tabs.addTab(self._page(self._analyzer_html()), "Analyzer")
        self.tabs.addTab(self._page(self._controls_html()), "Controlli")
        self.tabs.addTab(self._page(self._checklist_html()), "Checklist")
        self.tabs.addTab(self._page(self._save_output_html()), "Salvataggio e PDF")
        self.tabs.addTab(self._page(self._method_html()), "Metodo e limiti")
        self.tabs.setCurrentIndex(max(0, min(initial_tab, self.tabs.count() - 1)))
        layout.addWidget(self.tabs, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    @staticmethod
    def _page(html: str) -> QTextBrowser:
        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setHtml(html)
        return browser

    @staticmethod
    def _base(body: str) -> str:
        return f"""
        <html><head><style>
        body {{ font-family: Arial, sans-serif; font-size: 10.5pt; color: #243442; line-height: 1.45; }}
        h1 {{ color: #22313f; font-size: 19pt; margin: 0 0 10px 0; }}
        h2 {{ color: #2878a8; font-size: 13.5pt; margin-top: 18px; }}
        h3 {{ color: #334e5c; font-size: 11.5pt; margin-top: 14px; }}
        table {{ border-collapse: collapse; width: 100%; margin: 8px 0 14px 0; }}
        th, td {{ border: 1px solid #d7dfe5; padding: 7px; vertical-align: top; }}
        th {{ background: #edf4f8; color: #22313f; }}
        code {{ background: #edf1f4; padding: 1px 4px; border-radius: 3px; }}
        .flow {{ background: #eef6fb; border: 1px solid #b9d9ec; padding: 11px; border-radius: 7px; font-weight: bold; }}
        .warn {{ background: #fff5d9; border: 1px solid #efd183; padding: 10px; border-radius: 7px; }}
        .ok {{ background: #eaf7ef; border: 1px solid #b7ddc4; padding: 10px; border-radius: 7px; }}
        .step {{ color: #2878a8; font-weight: bold; }}
        </style></head><body>{body}</body></html>
        """

    @classmethod
    def _workflow_html(cls) -> str:
        return cls._base("""
        <h1>Processo completo di Collaudo Suite</h1>
        <p>Collaudo Suite collega l'analisi delle anomalie alla compilazione della checklist di collaudo. Questa guida descrive il flusso completo e le due modalita di lavoro: analisi storica delle anomalie e preparazione del verbale di collaudo.</p>
        <div class="flow">File anomalie Excel &rarr; Analyzer &rarr; Controlli operativi &rarr; Checklist &rarr; Pass/No pass &rarr; Salvataggio lavoro &rarr; PDF finale</div>

        <h2>Flusso consigliato</h2>
        <ol>
          <li><span class="step">Apri Analyzer</span> e carica uno o più file Excel contenenti anomalie, segnalazioni o punti aperti.</li>
          <li>Configura il foglio, le colonne e il periodo storico, quindi avvia l'analisi.</li>
          <li>Esamina i gruppi trovati e trasforma le anomalie rilevanti in controlli tecnici verificabili.</li>
          <li>Seleziona i controlli validi e usa <b>Invia alla Checklist</b>, oppure esportali in Excel per un uso successivo.</li>
          <li><span class="step">Apri Checklist</span>, compila l'intestazione e configura l'eventuale estrazione MAP.</li>
          <li>Genera l'anteprima: il programma unisce i controlli MAP e quelli importati da Analyzer. I MAP sono filtrati prima per codice commerciale e periodo, poi campionati casualmente.</li>
          <li>Registra per ogni riga l'esito <b>Pass</b> o <b>No pass</b> e le eventuali note.</li>
          <li>Salva il lavoro per poterlo riprendere oppure esporta/stampa il PDF finale.</li>
        </ol>

        <h2>Origine dei controlli</h2>
        <table>
          <tr><th>Gruppo</th><th>Origine</th><th>Funzione</th></tr>
          <tr><td>MAP</td><td>File MAP scelto dall'utente o file interno</td><td>Segnalazioni estratte casualmente e filtrate per Commercial code.</td></tr>
          <tr><td>Analyzer</td><td>Anomalie ricorrenti analizzate</td><td>Controlli aggiuntivi derivati dall'esperienza e dai problemi reali.</td></tr>
        </table>
        <div class="ok"><b>Obiettivo:</b> trasformare le anomalie ricorrenti in verifiche preventive tracciabili durante i collaudi successivi.</div>

        <h2>Orientarsi nell'applicazione</h2>
        <p>Dalla scheda Home si apre Analyzer o Checklist. Il menu <b>File</b> ripropone i comandi principali; <b>Strumenti &gt; Info file interni</b> mostra i dati delle risorse locali; <b>Aiuto &gt; Guida</b> apre questo manuale. F1 apre la guida. Le icone della barra superiore mostrano il nome del comando al passaggio del mouse.</p>
        """)

    @classmethod
    def _analyzer_html(cls) -> str:
        return cls._base("""
        <h1>Uso dell'Analyzer anomalie</h1>
        <h2>Procedura rapida</h2>
        <ol>
          <li>Premi <b>Seleziona file Excel</b> e carica uno o più file <code>.xlsx</code>, <code>.xls</code> o <code>.xlsm</code>.</li>
          <li>Indica la prima riga dati, la colonna del testo da analizzare, la colonna data e il nome del foglio. Sono accettati riferimenti come <code>14</code> o <code>B14</code> per la riga, e lettere di colonna come <code>B</code> ed <code>E</code>.</li>
          <li>Lascia vuota la parola chiave per individuare gruppi ricorrenti; compilala per cercare un tema preciso.</li>
          <li>Imposta soglie e numero minimo di occorrenze.</li>
          <li>Premi <b>Avvia analisi</b> e controlla sintesi, dettaglio e grafico.</li>
        </ol>

        <h2>Significato delle opzioni</h2>
        <table>
          <tr><th>Opzione</th><th>Cosa fa</th><th>Attenzione</th></tr>
          <tr><td>Parola chiave</td><td>Ricerca mirata di un termine o codice, ad esempio <code>E32.0</code>, <code>DM30</code> o <code>pressostato</code>. Se vuota, si cercano gruppi ricorrenti senza tema predefinito.</td><td>Una parola generica produce molti falsi positivi; una parola troppo specifica puo nascondere varianti.</td></tr>
          <tr><td>Soglia keyword</td><td>Somiglianza minima tra testo e parola chiave.</td><td>Alta = più precisione; bassa = più risultati e più rumore.</td></tr>
          <tr><td>Soglia gruppi</td><td>Somiglianza minima per collegare due descrizioni nello stesso gruppo.</td><td>Alta separa troppo; bassa può unire problemi differenti.</td></tr>
          <tr><td>Min. occ.</td><td>Numero minimo di righe necessario per mostrare un gruppo.</td><td>Valori alti nascondono anomalie rare ma importanti.</td></tr>
          <tr><td>Riga start</td><td>Prima riga Excel da leggere. Accetta <code>14</code> oppure <code>B14</code>.</td><td>Una riga errata include intestazioni o esclude dati.</td></tr>
          <tr><td>Col. analisi</td><td>Colonna contenente la descrizione dell'anomalia.</td><td>È il parametro più importante dell'importazione.</td></tr>
          <tr><td>Col. data</td><td>Colonna usata per datare le occorrenze e applicare il periodo selezionato.</td><td>Verificare che la colonna contenga date leggibili.</td></tr>
          <tr><td>Anomalie negli ultimi</td><td>Limita l'analisi agli ultimi 1, 3, 6 o 12 mesi di calendario rispetto a oggi. La finestra inizia dal primo giorno del mese iniziale e termina oggi: per esempio, il 2 ottobre, 6 mesi significa 1 maggio-2 ottobre.</td><td>Le righe senza una data valida o fuori finestra sono escluse. Verificare che la colonna data sia corretta.</td></tr>
          <tr><td>Foglio / Tutti i fogli</td><td>Limita l'analisi a un foglio oppure legge l'intero file.</td><td>Tutti i fogli può mescolare dati non omogenei.</td></tr>
          <tr><td>Algoritmo</td><td>Combinato è consigliato; Fuzzy è tollerante; Jaccard è più lessicale.</td><td>Confrontare periodi diversi usando sempre lo stesso algoritmo.</td></tr>
          <tr><td>Limite n2 / Max bucket</td><td>Regolano il numero di confronti candidati e la dimensione dei gruppi candidati sui dataset grandi.</td><td>Aumentarli puo migliorare la copertura, ma richiede piu tempo e memoria.</td></tr>
        </table>

        <h2>Lettura dei risultati</h2>
        <p>Il <b>Report testuale</b> contiene i messaggi di avanzamento e la sintesi dei gruppi. Il <b>Grafico</b> e disponibile dopo l'analisi. Nella scheda <b>Controlli per Checklist</b> ogni riga mostra anomalia originale, testo operativo proposto, occorrenze, data, macchina, categoria e ticket quando disponibili. Esporta il report per conservare i risultati. <b>Stop</b> chiede l'interruzione: attendere lo stato di completamento/interruzione prima di riavviare con parametri diversi.</p>
        """)

    @classmethod
    def _controls_html(cls) -> str:
        return cls._base("""
        <h1>Creazione dei controlli per la Checklist</h1>
        <p>Dopo l'analisi, il programma propone un testo di controllo per i gruppi selezionabili. Il testo e una bozza da validare, non una diagnosi automatica.</p>
        <h2>Come procedere</h2>
        <ol>
          <li>Apri la scheda <b>Controlli per Checklist</b>.</li>
          <li>Leggi il gruppo origine e verifica le righe dettagliate che lo compongono.</li>
          <li>Correggi il testo proposto rendendolo un'azione osservabile e verificabile.</li>
          <li>Mantieni, quando disponibile, il riferimento al ticket per la tracciabilità.</li>
          <li>Seleziona solo i controlli tecnicamente validi.</li>
          <li>Usa <b>Seleziona tutti</b> o <b>Deseleziona tutti</b> come punto di partenza, quindi premi <b>Invia alla Checklist</b> per il trasferimento diretto oppure esporta i controlli in Excel.</li>
        </ol>

        <h2>Regole di scrittura consigliate</h2>
        <table>
          <tr><th>Debole</th><th>Migliore</th></tr>
          <tr><td>Problema pressostato.</td><td>Verificare il corretto intervento del pressostato aria generale e la comparsa dell'allarme associato.</td></tr>
          <tr><td>Manca rondella.</td><td>Verificare la presenza e il corretto montaggio della rondella prevista dal disegno.</td></tr>
          <tr><td>Controllare sensore.</td><td>Verificare attivazione, disattivazione e timeout del sensore nelle due posizioni di lavoro.</td></tr>
        </table>
        <div class="warn"><b>Controllo scientifico:</b> una ricorrenza statistica non dimostra da sola una causa comune. Prima di creare il controllo, verificare che le anomalie raggruppate descrivano realmente lo stesso fenomeno tecnico.</div>
        <h2>Duplicati</h2>
        <p>Durante il trasferimento alla Checklist, i controlli equivalenti già presenti vengono ignorati mediante confronto del testo normalizzato.</p>
        <p>Il trasferimento diretto porta alla scheda Checklist e aggiorna la tabella gia generata senza cancellare gli esiti degli altri controlli. L'importazione da file e disponibile in Checklist per Excel, CSV e JSON. Il pulsante <b>Svuota</b> rimuove i controlli Analyzer importati dalla sessione corrente; non modifica i file sorgente.</p>
        """)

    @classmethod
    def _checklist_html(cls) -> str:
        return cls._base("""
        <h1>Compilazione della Checklist</h1>
        <h2>1. Intestazione</h2>
        <p>Compila <b>Collaudatore</b>, <b>Reparto</b>, <b>Data</b> e <b>Numero Commessa</b>. Il pulsante <b>Today</b> imposta la data odierna. L'intestazione di un nuovo lavoro parte vuota; usa <b>Apri lavoro</b> per riprendere una sessione salvata.</p>


        <h2>3. Controlli MAP</h2>
        <ol>
          <li>Imposta il numero di segnalazioni da estrarre (predefinito 3); usa <b>0</b> per non aggiungere MAP.</li>
          <li>Scegli il filtro Commercial code: per esempio <b>NC300</b>, <b>GENYA</b>, <b>TRINITY</b>, <b>Tutti</b>, oppure digita una stringa personalizzata. La corrispondenza e per inclusione nel valore del codice commerciale: <code>300</code> trova, ad esempio, <code>NC30000068</code>. Non filtra per titolo o numero ticket.</li>
          <li>Scegli il periodo di 1, 3, 6 o 12 mesi. Il riferimento e la <b>Data</b> del collaudo (non necessariamente oggi). Gli estremi sono inclusi: il limite iniziale si calcola a mesi di calendario e il giorno del collaudo e incluso. Le righe successive alla data di collaudo e le righe senza data valida sono escluse.</li>
          <li>Imposta la sorgente MAP. Con <b>File Excel</b>, premi <b>Sfoglia...</b> e seleziona il riepilogativo; se il percorso e vuoto viene usato il file MAP interno, se presente. Sono riconosciute intestazioni equivalenti a <code>Titolo</code>, <code>Commessa S Codice commerciale</code> e <code>Numero Ticket</code>, oltre a vari nomi per la data (ad esempio <code>Last Modify</code>, <code>Data ultima modifica</code>, <code>Created Date</code> o <code>Data ticket</code>).</li>
          <li>Le righe vengono filtrate per stato non escluso, titolo, data e codice commerciale; i duplicati vengono eliminati e il numero richiesto e campionato casualmente dal pool residuo. Ogni aggiornamento puo quindi restituire MAP diversi. Il riepilogo sopra la tabella riporta dimensione del pool, filtro, periodo, sorgente ed eventuali avvisi.</li>
        </ol>

        <h3>Usare JARVIS API</h3>
        <ol>
          <li>Seleziona <b>JARVIS API</b> come sorgente.</li>
          <li>Inserisci un token JARVIS autorizzato alla lettura ticket (<code>ticket.read</code>). Il campo e mascherato; il token resta nelle impostazioni locali dell'utente e non viene incluso nel file lavoro.</li>
          <li>Premi <b>Aggiorna MAP da JARVIS</b>. La sincronizzazione cerca i ticket MAP e legge il valore della proprieta corrispondente a <b>Commessa S Codice commerciale</b> (non identificativi interni opachi come <code>DatasetElement_...</code>).</li>
          <li>La suite crea/aggiorna una cache Excel locale e poi usa la stessa procedura di filtro e campionamento della sorgente Excel. Se la cache e gia presente, agli avvii successivi puo essere aggiornata automaticamente in background.</li>
        </ol>
        <div class="warn"><b>Se la sincronizzazione fallisce:</b> controlla token/permesso, connessione e messaggio di errore. Se Windows segnala accesso negato durante la sostituzione della cache, chiudi eventuali Excel che tengono aperto il file e riprova; se il file resta bloccato, la suite usa un percorso cache alternativo versionato. La sorgente File Excel rimane utilizzabile.</div>

        <h3>Scaricare manualmente il file MAP da JARVIS</h3>
        <p>Se non usi la sorgente JARVIS API, puoi scaricare manualmente il riepilogativo Excel direttamente da JARVIS.</p>
        <ol>
          <li>Apri <b>Ticketing &gt; Ricevuti</b>.</li>
          <li>Apri il pannello <b>Filtra</b> e collega le condizioni con <b>AND</b>.</li>
          <li>Imposta <b>Linea di business</b> su <b>Uno tra</b> e scegli la linea corretta per il codice commerciale richiesto. La linea di business non e fissa e puo variare.</li>
          <li>Imposta <b>Commessa S - Codice commerciale</b> su <b>Contiene</b> e inserisci il codice, ad esempio <code>NC300</code>.</li>
          <li>Imposta <b>Ticket - Modello ticket</b> su <b>Uno tra</b> e scegli <b>Segnalazione MAP</b>.</li>
          <li>Applica i filtri e verifica il codice commerciale e il numero dei risultati.</li>
          <li>Apri il menu con i tre puntini in alto a destra, scegli <b>Esporta</b> e salva il risultato in formato Excel.</li>
          <li>Usa un nome descrittivo, ad esempio <code>MAP_NC300_2026-10-07.xlsx</code>, quindi seleziona il file nella sorgente <b>File Excel</b> della Checklist.</li>
        </ol>
        <p><b>Sequenza visuale JARVIS:</b></p>
        <img src="jarvis_map_01.png" alt="JARVIS - ricerca documentale" />
        <img src="jarvis_map_02.png" alt="JARVIS - ticket ricevuti" />
        <img src="jarvis_map_03.png" alt="JARVIS - filtri MAP" />
        <img src="jarvis_map_04.png" alt="JARVIS - esportazione Excel" />
        <div class="warn"><b>Controllo del file:</b> il riepilogativo deve riportare almeno il codice commerciale, il titolo della segnalazione, il numero ticket, la data di modifica e lo stato. In Collaudo Suite il codice viene cercato nella colonna equivalente a <b>Commessa S - Codice commerciale</b>.</div>

        <h2>4. Controlli Analyzer</h2>
        <p>Possono arrivare direttamente dal modulo Analyzer o essere importati dal file Excel esportato in precedenza.</p>

        <h2>5. Generazione e compilazione</h2>
        <ol>
          <li>Premi <b>Estrai / Aggiorna anteprima</b>.</li>
          <li>Controlla l'elenco completo prima di iniziare il collaudo.</li>
          <li>Per ogni riga seleziona <b>Pass</b> oppure <b>No pass</b>. Le due opzioni sono alternative.</li>
          <li>Inserisci nelle note valori misurati, anomalie, riferimenti o azioni eseguite.</li>
        </ol>
        <div class="warn"><b>Attenzione:</b> non usare Pass per indicare semplicemente che il controllo è stato letto. Pass significa che la condizione verificata è conforme ai criteri applicabili.</div>

        <h2>Ticket e aggiornamento dell'elenco</h2>
        <p>Il numero nella colonna <b>Ticket</b> e un collegamento: un clic apre il ticket JARVIS nel browser predefinito. Se dopo <b>Estrai / Aggiorna anteprima</b> cambia l'elenco o l'ordine, clicca il numero nella riga desiderata: il collegamento e associato alla cella del ticket visualizzata, non alla posizione precedente nella lista. Se non si apre, controlla che il browser sia configurato e che JARVIS sia raggiungibile; verifica il ticket dalla sua pagina e riprova.</p>
        <p>La tabella e divisa in sezioni <b>SEGNALAZIONI MAP</b> e <b>CONTROLLI IMPORTATI DA ANALYZER</b>. L'aggiornamento MAP rigenera il campione MAP; cerca di preservare gli esiti e le note associandoli ai controlli che restano nel nuovo elenco. Prima di cambiare filtri o sorgente, salva il lavoro se i risultati inseriti devono essere conservati.</p>

        <h2>Importare e gestire controlli Analyzer</h2>
        <p>Il trasferimento diretto da Analyzer e preferibile per la sessione in corso. In alternativa usa <b>Importa...</b> per selezionare un file Excel, CSV o JSON con i controlli. <b>Svuota</b> rimuove il gruppo Analyzer dalla checklist corrente. Controlla sempre numero e testo dei controlli importati prima di generare il documento finale.</p>
        """)

    @classmethod
    def _save_output_html(cls) -> str:
        return cls._base("""
        <h1>Salvataggio, ripresa ed esportazione</h1>
        <h2>Comandi nella barra superiore</h2>
        <p>Le icone della barra superiore permettono di esportare o stampare il PDF, salvare, salvare con nome, aprire un lavoro ed uscire dal programma. Posizionando il mouse sull'icona viene mostrata la funzione.</p>

        <h2>Salva lavoro</h2>
        <p>Al primo salvataggio viene richiesto il nome del file. I salvataggi successivi aggiornano direttamente lo stesso file senza richiedere nuovamente il nome. Usa <b>Salva con nome</b> per creare una copia o cambiare destinazione.</p>
        <p>Memorizza la sessione corrente in un file lavoro JSON con estensione <code>.rcl.json</code>, inclusi:</p>
        <ul>
          <li>dati di intestazione;</li>
          <li>controlli MAP e Analyzer;</li>
          <li>ticket;</li>
          <li>codice/filtro e periodo MAP, percorso della sorgente e dati dei controlli importati;</li>
          <li>esiti Pass/No pass;</li>
          <li>note.</li>
        </ul>
        <p>I percorsi proposti sono nella cartella <b>Download</b>. La suite effettua anche un autosalvataggio del lavoro corrente. I file di lavoro possono contenere informazioni di collaudo: conservarli in una posizione autorizzata.</p>

        <h2>Apri lavoro</h2>
        <p>Ricarica una sessione salvata e permette di proseguire il collaudo senza perdere gli esiti già inseriti.</p>

        <h2>Esporta PDF</h2>
        <p>Crea il report finale senza richiedere la stampa immediata. Usalo per archiviazione o invio.</p>

        <h2>Stampa PDF</h2>
        <p>Genera il PDF e lo apre nel visualizzatore predefinito, lasciando all'utente il controllo della stampa.</p>

        <h2>Aggiornamenti</h2>
        <p>All'avvio la suite verifica in background se è disponibile una nuova versione; il controllo si può avviare anche da <b>Aiuto &gt; Controlla aggiornamenti</b>. Su Windows il pulsante <b>Scarica e installa</b> scarica l'installer, salva il lavoro corrente, chiude l'applicazione e avvia l'installazione.</p>

        <h2>Verifiche prima della chiusura</h2>
        <ol>
          <li>Controllare che l'intestazione sia completa.</li>
          <li>Verificare che ogni controllo applicabile abbia un esito.</li>
          <li>Motivare i No pass con note sufficienti.</li>
          <li>Controllare la presenza dei ticket importati.</li>
          <li>Aprire il PDF e verificare impaginazione e completezza.</li>
        </ol>

        <h2>Scorciatoie da tastiera</h2>
        <table><tr><th>Combinazione</th><th>Azione</th></tr>
          <tr><td><code>Ctrl+S</code></td><td>Salva lavoro</td></tr>
          <tr><td><code>Ctrl+Shift+S</code></td><td>Salva lavoro con nome</td></tr>
          <tr><td><code>Ctrl+O</code></td><td>Apri lavoro</td></tr>
          <tr><td><code>Ctrl+P</code></td><td>Stampa PDF</td></tr>
          <tr><td><code>Ctrl+Q</code></td><td>Esci</td></tr>
          <tr><td><code>F1</code></td><td>Apri Guida</td></tr>
        </table>

        <h2>Risoluzione dei problemi</h2>
        <table><tr><th>Sintomo</th><th>Controllo / soluzione</th></tr>
          <tr><td>Nessun MAP estratto</td><td>Controlla numero richiesto, codice commerciale, data collaudo, periodo e sorgente. Il riepilogo indica quanti record sono entrati nel pool. Con Excel verifica intestazioni e colonna data; date future e non valide sono escluse.</td></tr>
          <tr><td>Filtro MAP senza risultati</td><td>Usa temporaneamente <b>Tutti</b>, poi confronta il valore effettivo della colonna <b>Commessa S Codice commerciale</b>. Il filtro e una sottostringa: deve comparire nel codice.</td></tr>
          <tr><td>Ticket JARVIS non si apre</td><td>Riprova dopo l'aggiornamento della lista cliccando il numero ticket nella riga corretta. Verifica connessione, browser predefinito e accesso JARVIS.</td></tr>
          <tr><td>Token non accettato / sync fallita</td><td>Verifica che sia un token valido con <code>ticket.read</code>; controlla rete e messaggio di errore. Puoi passare temporaneamente alla sorgente File Excel.</td></tr>
          <tr><td>Errore accesso negato sulla cache</td><td>Chiudi il file cache aperto in Excel e ripeti la sincronizzazione. Non eliminare la cache mentre la suite e in esecuzione; il percorso alternativo puo comparire dopo una sostituzione bloccata.</td></tr>
          <tr><td>Modifiche o esiti non ritrovati</td><td>Apri il file lavoro salvato piu recente. Controlla il nome/percorso indicato dopo il salvataggio e verifica che non sia stato aperto un altro lavoro.</td></tr>
          <tr><td>PDF incompleto o non aggiornato</td><td>Completa la checklist, salva il lavoro, esporta un nuovo PDF e apri il file appena creato. Prima di distribuire il PDF, verifica intestazione, tutte le pagine, esiti e note.</td></tr>
        </table>
        """)

    @classmethod
    def _method_html(cls) -> str:
        return cls._base("""
        <h1>Metodo, interpretazione e limiti</h1>
        <h2>Come lavora l'Analyzer</h2>
        <p>Il testo viene normalizzato conservando, per quanto possibile, codici tecnici come <code>E32.0</code>, <code>DM30</code>, <code>S41A</code> o <code>BH4.0</code>. Le righe candidate vengono confrontate e collegate in gruppi di similarità. Il rappresentante del gruppo è scelto come testo centrale, non semplicemente come prima riga.</p>

        <h2>Ipotesi da verificare</h2>
        <ul>
          <li>Descrizioni simili possono riferirsi a cause differenti.</li>
          <li>Descrizioni differenti possono rappresentare lo stesso problema.</li>
          <li>Una frequenza alta può dipendere da un difetto ricorrente, ma anche da una modalità di registrazione ripetitiva.</li>
          <li>Un'anomalia rara può avere criticità elevata e non deve essere esclusa solo perché non raggiunge il minimo di occorrenze.</li>
        </ul>

        <h2>Buone pratiche</h2>
        <ul>
          <li>Leggere sempre le righe originali nel dettaglio.</li>
          <li>Mantenere costanti algoritmo e soglie quando si confrontano periodi diversi.</li>
          <li>Usare il ticket per risalire alla segnalazione originale.</li>
          <li>Separare frequenza, gravità e probabilità di rilevazione: non sono la stessa cosa.</li>
        </ul>

        <div class="ok"><b>Conclusione:</b> Collaudo Suite è uno strumento di supporto decisionale. L'analisi automatica individua candidati; la validazione tecnica resta responsabilità dell'operatore e del processo qualità.</div>
        """)
