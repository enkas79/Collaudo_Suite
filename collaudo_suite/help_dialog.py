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
        <p>Collaudo Suite collega l'analisi delle anomalie alla compilazione della checklist di collaudo.</p>
        <div class="flow">File anomalie Excel &rarr; Analyzer &rarr; Controlli operativi &rarr; Checklist &rarr; Pass/No pass &rarr; Salvataggio lavoro &rarr; PDF finale</div>

        <h2>Flusso consigliato</h2>
        <ol>
          <li><span class="step">Apri Analyzer</span> e carica uno o più file Excel contenenti anomalie, segnalazioni o punti aperti.</li>
          <li>Configura il foglio e le colonne corrette, quindi avvia l'analisi.</li>
          <li>Esamina i gruppi trovati e trasforma le anomalie rilevanti in controlli tecnici verificabili.</li>
          <li>Seleziona i controlli validi e usa <b>Invia alla Checklist</b>, oppure esportali in Excel per un uso successivo.</li>
          <li><span class="step">Apri Checklist</span>, compila l'intestazione e configura l'eventuale estrazione MAP.</li>
          <li>Genera l'anteprima: il programma unisce checklist fissa, controlli MAP e controlli Analyzer.</li>
          <li>Registra per ogni riga l'esito <b>Pass</b> o <b>No pass</b> e le eventuali note.</li>
          <li>Salva il lavoro per poterlo riprendere oppure esporta/stampa il PDF finale.</li>
        </ol>

        <h2>Origine dei controlli</h2>
        <table>
          <tr><th>Gruppo</th><th>Origine</th><th>Funzione</th></tr>
          <tr><td>Checklist fissa</td><td>Documento interno al programma</td><td>Controlli standard sempre presenti.</td></tr>
          <tr><td>MAP</td><td>File MAP scelto dall'utente o file interno</td><td>Segnalazioni estratte casualmente e filtrate per Commercial code.</td></tr>
          <tr><td>Analyzer</td><td>Anomalie ricorrenti analizzate</td><td>Controlli aggiuntivi derivati dall'esperienza e dai problemi reali.</td></tr>
        </table>
        <div class="ok"><b>Obiettivo:</b> trasformare le anomalie ricorrenti in verifiche preventive tracciabili durante i collaudi successivi.</div>
        """)

    @classmethod
    def _analyzer_html(cls) -> str:
        return cls._base("""
        <h1>Uso dell'Analyzer anomalie</h1>
        <h2>Procedura rapida</h2>
        <ol>
          <li>Premi <b>Seleziona file Excel</b> e carica uno o più file <code>.xlsx</code>, <code>.xls</code> o <code>.xlsm</code>.</li>
          <li>Indica la prima riga dati, la colonna del testo da analizzare, la colonna data e il nome del foglio.</li>
          <li>Lascia vuota la parola chiave per individuare gruppi ricorrenti; compilala per cercare un tema preciso.</li>
          <li>Imposta soglie e numero minimo di occorrenze.</li>
          <li>Premi <b>Avvia analisi</b> e controlla sintesi, dettaglio e grafico.</li>
        </ol>

        <h2>Significato delle opzioni</h2>
        <table>
          <tr><th>Opzione</th><th>Cosa fa</th><th>Attenzione</th></tr>
          <tr><td>Parola chiave</td><td>Ricerca mirata di un termine o codice, ad esempio <code>E32.0</code>, <code>DM30</code> o <code>pressostato</code>.</td><td>Una parola generica produce molti falsi positivi.</td></tr>
          <tr><td>Soglia keyword</td><td>Somiglianza minima tra testo e parola chiave.</td><td>Alta = più precisione; bassa = più risultati e più rumore.</td></tr>
          <tr><td>Soglia gruppi</td><td>Somiglianza minima per collegare due descrizioni nello stesso gruppo.</td><td>Alta separa troppo; bassa può unire problemi differenti.</td></tr>
          <tr><td>Min. occ.</td><td>Numero minimo di righe necessario per mostrare un gruppo.</td><td>Valori alti nascondono anomalie rare ma importanti.</td></tr>
          <tr><td>Riga start</td><td>Prima riga Excel da leggere. Accetta <code>14</code> oppure <code>B14</code>.</td><td>Una riga errata include intestazioni o esclude dati.</td></tr>
          <tr><td>Col. analisi</td><td>Colonna contenente la descrizione dell'anomalia.</td><td>È il parametro più importante dell'importazione.</td></tr>
          <tr><td>Col. data</td><td>Colonna usata per determinare la data più recente del gruppo.</td><td>Non modifica il raggruppamento, ma la tracciabilità.</td></tr>
          <tr><td>Foglio / Tutti i fogli</td><td>Limita l'analisi a un foglio oppure legge l'intero file.</td><td>Tutti i fogli può mescolare dati non omogenei.</td></tr>
          <tr><td>Algoritmo</td><td>Combinato è consigliato; Fuzzy è tollerante; Jaccard è più lessicale.</td><td>Confrontare periodi diversi usando sempre lo stesso algoritmo.</td></tr>
          <tr><td>Limite n2 / Max bucket</td><td>Regolano copertura e velocità sui dataset grandi.</td><td>Aumentare i valori migliora la copertura ma rallenta l'analisi.</td></tr>
        </table>
        """)

    @classmethod
    def _controls_html(cls) -> str:
        return cls._base("""
        <h1>Creazione dei controlli per la Checklist</h1>
        <p>Dopo l'analisi, il programma propone un testo di controllo per ogni gruppo di anomalie.</p>
        <h2>Come procedere</h2>
        <ol>
          <li>Apri la scheda <b>Controlli per Checklist</b>.</li>
          <li>Leggi il gruppo origine e verifica le righe dettagliate che lo compongono.</li>
          <li>Correggi il testo proposto rendendolo un'azione osservabile e verificabile.</li>
          <li>Mantieni, quando disponibile, il riferimento al ticket per la tracciabilità.</li>
          <li>Seleziona solo i controlli tecnicamente validi.</li>
          <li>Premi <b>Invia alla Checklist</b> per il trasferimento diretto oppure esporta in Excel.</li>
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
        """)

    @classmethod
    def _checklist_html(cls) -> str:
        return cls._base("""
        <h1>Compilazione della Checklist</h1>
        <h2>1. Intestazione</h2>
        <p>Compila <b>Collaudatore</b>, <b>Reparto</b>, <b>Data</b> e <b>Numero Commessa</b>. Il pulsante <b>Today</b> imposta la data odierna. L'intestazione parte vuota a ogni nuovo avvio.</p>

        <h2>2. Checklist fissa</h2>
        <p>La checklist standard è interna al programma e viene aggiunta automaticamente. Non deve essere selezionata manualmente.</p>

        <h2>3. Controlli MAP</h2>
        <ol>
          <li>Imposta il numero di segnalazioni da estrarre; usa <b>0</b> per non aggiungerne.</li>
          <li>Scegli il filtro Commercial code: <b>NC300</b>, <b>GENYA</b>, <b>TRINITY</b>, <b>Tutti</b> oppure digita un valore personalizzato.</li>
          <li>Con <b>Sfoglia</b> seleziona il file MAP riepilogativo. In assenza di selezione viene usato il file interno.</li>
          <li>Il file può contenere colonne equivalenti a Title/Titolo, Commercial code/Codice commerciale e Ticket Number/Numero Ticket.</li>
          <li>Le righe vengono prima filtrate e poi estratte casualmente. Il ticket viene riportato nella colonna dedicata.</li>
        </ol>

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
        """)

    @classmethod
    def _save_output_html(cls) -> str:
        return cls._base("""
        <h1>Salvataggio, ripresa ed esportazione</h1>
        <h2>Comandi nella barra superiore</h2>
        <p>Le icone della barra superiore permettono di esportare o stampare il PDF, salvare, salvare con nome, aprire un lavoro ed uscire dal programma. Posizionando il mouse sull'icona viene mostrata la funzione.</p>

        <h2>Salva lavoro</h2>
        <p>Al primo salvataggio viene richiesto il nome del file. I salvataggi successivi aggiornano direttamente lo stesso file senza richiedere nuovamente il nome. Usa <b>Salva con nome</b> per creare una copia o cambiare destinazione.</p>
        <p>Memorizza la sessione corrente, inclusi:</p>
        <ul>
          <li>dati di intestazione;</li>
          <li>controlli fissi, MAP e Analyzer;</li>
          <li>ticket;</li>
          <li>esiti Pass/No pass;</li>
          <li>note.</li>
        </ul>
        <p>I percorsi proposti sono nella cartella <b>Download</b>.</p>

        <h2>Apri lavoro</h2>
        <p>Ricarica una sessione salvata e permette di proseguire il collaudo senza perdere gli esiti già inseriti.</p>

        <h2>Esporta PDF</h2>
        <p>Crea il report finale senza richiedere la stampa immediata. Usalo per archiviazione o invio.</p>

        <h2>Stampa PDF</h2>
        <p>Genera il PDF e lo apre nel visualizzatore predefinito, lasciando all'utente il controllo della stampa.</p>

        <h2>Verifiche prima della chiusura</h2>
        <ol>
          <li>Controllare che l'intestazione sia completa.</li>
          <li>Verificare che ogni controllo applicabile abbia un esito.</li>
          <li>Motivare i No pass con note sufficienti.</li>
          <li>Controllare la presenza dei ticket importati.</li>
          <li>Aprire il PDF e verificare impaginazione e completezza.</li>
        </ol>
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
          <li>Rivedere periodicamente i controlli Analyzer ed eliminare quelli ormai assorbiti nella checklist standard.</li>
        </ul>

        <div class="ok"><b>Conclusione:</b> Collaudo Suite è uno strumento di supporto decisionale. L'analisi automatica individua candidati; la validazione tecnica resta responsabilità dell'operatore e del processo qualità.</div>
        """)
