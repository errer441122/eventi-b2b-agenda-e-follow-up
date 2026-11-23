# Come gestire un cambio turno al desk

Il foglio aiuta chi aggiorna l’agenda delle dimostrazioni di uno stand, anche su più giorni e più aree. Quando una delegazione chiede di spostarsi, l’operatore controlla capienza, stato dell’invito, presenza di un referente e prima alternativa nella stessa area e giornata. Quando il gruppo accetta, l’agenda indica chi deve ancora ricevere il nuovo orario.

**[Scarica l’esempio Excel](agenda-demo.xlsx?raw=true)**. Il file funziona in locale e i nomi sono fittizi.

## Prova il cambio in 90 secondi

| Passo | Cosa fare | Cosa osservare |
| --- | --- | --- |
| 1. Ricevi la richiesta | Apri **Cambio turno**. Sono già selezionati visita **V101** e turno **S02**. | Il gruppo è di tre persone. Alle 10:30 restano due posti. La Sala demo B alle 10:45 ha posto, ma è un’altra area: il foglio propone **S03**, alle **11:00** nella stessa sala, con tre posti e Luca come referente. |
| 2. Proponi l’alternativa | In **Cambio turno**, cambia **E6** da S02 a **S03**. | Compare *Cambio proponibile*. In **Turni**, i tre posti della delegazione rimangono occupati alle 10:00. |
| 3. Registra il sì | Solo dopo la conferma del visitatore, in **Agenda** cambia **D6** da S01 a **S03**. | Alle 10:00 si liberano tre posti. Alle 11:00 i residui passano da tre a zero. Sulla riga del gruppo compare *Avvisare visitatore e referente*. |
| 4. Passa le consegne | Dopo aver comunicato l’orario al visitatore, in **H6** scegli dal menu il testo di G6. Dopo aver avvisato Luca, in **J6** scegli dal menu il testo di I6. | Prima resta *Avvisare il referente*. Solo dopo il secondo aggiornamento compare *Entrambi aggiornati*. |
| 5. Il turno slitta | In **Turni** cambia l’ora di S03, cella **C8**, da 11:00 a 11:20. | Sulle righe dei gruppi delle 11:00 torna *Avvisare visitatore e referente*: chi aveva ricevuto le 11:00 va richiamato. |

Se il visitatore rifiuta la proposta al passo 2, puoi chiudere la richiesta lasciando invariata l’agenda. Per ripetere la demo, riapri una copia del file originale.

## Come usarlo al banco

- **Prima dell’apertura.** In *Turni* prepara le sessioni: giorno, ora, area, capienza in persone e referente. In *Agenda* inserisci una riga per gruppo prenotato, con un codice univoco e lo stato dell’invito verificato dall’organizzatore. Le celle azzurre sono modificabili.
- **Quando arriva la richiesta.** Aggiorna giorno e ora al desk in **B4** ed **E4**, scegli visita e turno e leggi la decisione. Le alternative restano nella stessa area e giornata e iniziano dopo l’ora indicata e dopo il turno richiesto. Un gruppo resta insieme. Un turno sospeso o senza referente viene saltato. Prima di registrare il sì, ricontrolla i posti.
- **Quando cambia l’operatore.** Il collega apre *Agenda* e legge *Azione ancora aperta*. Dopo ogni avviso, nelle colonne *Detto al visitatore* e *Detto al referente* sceglie dal menu il testo che ha comunicato. Se dopo cambiano orario, area o referente del turno, il testo non coincide più e la riga torna da avvisare. Telefonata, messaggio e comunicazione al banco spettano all’operatore.

Il caso copre **più giorni e più aree, fino a 150 turni e 300 gruppi**, con un solo operatore che conferma modifiche alla volta. Se un gruppo arriva più numeroso, il turno risulta *oltre capienza* ma il desk continua a funzionare: propone un turno con posto per tutti, oppure si alza la capienza e il referente va avvisato del nuovo numero. La colonna *Note* raccoglie richieste come un interprete. Il foglio segnala i dati incoerenti inseriti a mano, come un codice doppio o due turni alla stessa ora nella stessa area. Le formule forniscono un controllo operativo, mentre l’ammissione all’area rimane responsabilità del personale incaricato.

## Quando usarlo

Il foglio è pensato per un desk con una sola agenda e un operatore alla volta. Il beneficio è concreto quando due colleghi si danno il cambio, o quando un orario cambia a giornata iniziata: chi arriva vede chi ha già ricevuto l’orario giusto e chi aspetta ancora una telefonata.

Per valutarlo su un evento, si possono ripercorrere tre situazioni con il personale del desk: un gruppo troppo numeroso, un turno sospeso e un turno che slitta. Il confronto con l’agenda già in uso riguarda due cose osservabili: il tempo necessario per trovare un’alternativa e la possibilità di riconoscere, al cambio turno, chi aspetta ancora una comunicazione.

## Verifica

Il file è stato aperto in Microsoft Excel e ricalcolato dopo ogni modifica. Sessantadue controlli verificano il caso iniziale, l’area e il giorno delle alternative, il trasferimento dei posti, i due avvisi indipendenti, un turno che slitta, un cambio di referente, un gruppo più numeroso, il secondo giorno, l’orario già trascorso, i turni sospesi o senza referente, l’invito da confermare e i dati incoerenti. [Script per ripetere la verifica con Excel su Windows](verifica-excel.ps1)
