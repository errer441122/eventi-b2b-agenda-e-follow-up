# Come gestire un cambio turno al desk

Il foglio aiuta una persona che aggiorna l’agenda di una sala demo. Quando una delegazione chiede di spostarsi, l’operatore può controllare capienza, stato dell’invito, presenza di un referente e prima alternativa successiva. Quando il gruppo accetta, l’agenda indica chi deve ancora ricevere il nuovo orario.

**[Scarica l’esempio Excel](agenda-demo.xlsx?raw=true)**. Il file funziona in locale e i nomi sono fittizi.

## Prova il cambio in 90 secondi

| Passo | Cosa fare | Cosa osservare |
| --- | --- | --- |
| 1. Ricevi la richiesta | Apri **Cambio turno**. Sono già selezionati visita **V101** e turno **S02**. | Il gruppo è di tre persone. Alle 10:30 restano due posti. Il primo turno successivo utile è **S03**, alle **11:00**, con tre posti e Luca come referente. |
| 2. Proponi l’alternativa | In **Cambio turno**, cambia **E6** da S02 a **S03**. | Compare *Cambio proponibile*. In **Turni**, i tre posti della delegazione rimangono occupati alle 10:00. |
| 3. Registra il sì | Solo dopo la conferma del visitatore, in **Agenda** cambia **D6** da S01 a **S03**. | Alle 10:00 si liberano tre posti. Alle 11:00 i residui passano da tre a zero. Sulla riga del gruppo compare *Avvisare visitatore e referente*. |
| 4. Passa le consegne | Dopo aver comunicato l’orario al visitatore, scrivi **S03** in **G6**. Dopo aver avvisato Luca, scrivi **S03** in **H6**. | Prima resta *Avvisare il referente*. Solo dopo il secondo aggiornamento compare *Entrambi aggiornati*. |

Se il visitatore rifiuta la proposta al passo 2, puoi chiudere la richiesta lasciando invariata l’agenda. Per ripetere la demo, riapri una copia del file originale.

## Come usarlo al banco

- **Prima dell’apertura.** In *Turni* prepara le sessioni della giornata, gli orari, la capienza e il referente. In *Agenda* inserisci una riga per delegazione prenotata, con un identificativo univoco e lo stato dell’invito verificato dall’organizzatore. Le celle azzurre sono modificabili. La capienza è espressa in persone.
- **Quando arriva la richiesta.** Aggiorna l’ora al desk in **B4**, seleziona visita e turno e leggi la decisione. Le alternative devono iniziare dopo l’ora indicata e dopo il turno richiesto. Un gruppo resta insieme. Un turno sospeso o senza referente viene saltato. Prima di registrare il sì, ricontrolla i posti.
- **Quando cambia l’operatore.** Il collega apre *Agenda* e legge *Azione ancora aperta*. Registra il codice del turno che ciascuna persona ha effettivamente ricevuto. Telefonata, messaggio e comunicazione al banco spettano all’operatore.

Il caso copre **una sala, una giornata, fino a 12 turni e 40 delegazioni**, con un solo operatore che conferma modifiche alla volta. Conserva il codice, l’orario e il referente dei turni già comunicati. Per sostituire una sessione crea un nuovo turno e riprogramma le delegazioni. Il foglio segnala eventuali incoerenze inserite a mano, anche una sovracapienza. Le formule forniscono un controllo operativo, mentre l’ammissione all’area rimane responsabilità del personale incaricato.

## Quando usarlo

Il foglio è pensato per un desk con una sola agenda e un operatore alla volta. Il beneficio è concreto quando due colleghi si danno il cambio: chi arriva può vedere che il visitatore ha già ricevuto il nuovo orario e che resta da avvisare il referente della sala.

Per valutarlo su un evento, si possono ripercorrere tre situazioni con il personale del desk: un gruppo troppo numeroso, un turno sospeso e un cambio con un avviso ancora da dare. Il confronto con l’agenda già in uso riguarda due cose osservabili: il tempo necessario per trovare un’alternativa e la possibilità di riconoscere, al cambio turno, chi aspetta ancora una comunicazione.

## Verifica

Il file è stato aperto in Microsoft Excel e ricalcolato dopo ogni modifica. Trenta controlli verificano il caso iniziale, il trasferimento dei posti, i due avvisi indipendenti, una cancellazione, una seconda richiesta per lo stesso turno, l’orario già trascorso, i turni sospesi o senza referente, l’invito da confermare e i dati incoerenti. [Script per ripetere la verifica con Excel su Windows](verifica-excel.ps1)
