# Chi ricontattare dopo un webinar B2B?

Il giorno dopo un webinar, Sales chiede a chi proporre una demo e il team eventi deve spedire i materiali. Il file delle iscrizioni registra anche chi ha annullato. Il registro degli ingressi conta due volte chi si è riconnesso. Il CRM conserva le preferenze email e il nome del commerciale che segue il contatto. Quale lista consegneresti ai due team?

Questo **caso di studio simulato di Marketing Operations** mette a confronto i tre export in Excel. Per ogni ticket restituisce una decisione, un destinatario interno e il motivo. Tutti i nomi, gli indirizzi e gli eventi sono di esempio.

**[Scarica il file Excel](riconciliazione-evento.xlsx?raw=true)** · [Lista per Sales e Marketing](liste/follow-up.csv) · [Coda di revisione](liste/revisione.csv) · [Esclusioni](liste/esclusi.csv)

![Foglio Follow-up con 19 pratiche, i destinatari interni e le motivazioni](assets/anteprima.png)

## Risultato del caso

Gli export contengono **18 righe di iscrizione, 17 ingressi e 16 record CRM**. La riconciliazione produce **19 pratiche**. Un ticket presente in più fonti conta una volta. Un ingresso senza iscrizione apre comunque una pratica. Una riga priva di ticket riceve un riferimento provvisorio, così resta visibile a chi deve verificarla. Il conteggio misura pratiche, anche quando l'identità della persona è ancora incerta.

| Destinazione | Pratiche | Azione |
| --- | ---: | --- |
| Sales | 3 | Due richieste di demo passano al commerciale già assegnato. Una va nella coda Sales. |
| Marketing | 2 | Il team eventi prepara i materiali richiesti. |
| Revisione | 11 | CRM Ops verifica identità, iscrizione o permesso prima di qualsiasi contatto. |
| Escluso | 3 | Il team lascia fuori un assente, un contatto che ha cancellato senza partecipare e un contatto che ha revocato il permesso. |

## Tre ticket da seguire

| Ticket | Cosa mostrano le fonti | Decisione |
| --- | --- | --- |
| **T-0015 · Pia** | Si registra una volta, accede alle 10:24 e rientra alle 10:26 con lo stesso indirizzo. Ha chiesto i materiali e il CRM conferma il permesso. | Una sola riga al team eventi. Un invio per accesso avrebbe duplicato il messaggio. |
| **T-0006 · Fabio** | Il registro riporta un ingresso alle 10:06. L'iscrizione risulta annullata alle 10:54. | Revisione. Gli orari spiegano come possano essere vere entrambe le fonti. Il team eventi verifica se rimane valida la richiesta di demo. |
| **T-0016 · Rico** | Due righe d'iscrizione portano lo stesso ticket e due email diverse. Al check-in ne compare una sola. | Revisione. Prima di passare il caso a Sales, CRM Ops deve verificare l'indirizzo corretto con l'organizzatore. |

## Come funziona

1. **Riunire le tracce dell'evento.** Il ticket collega iscrizioni e ingressi. Spazi esterni e maiuscole delle email vengono normalizzati. Il nome serve a leggere il caso. Un nome uguale, da solo, non autorizza a unire due persone.
2. **Controllare l'identità.** Ticket mancante o malformato, più iscrizioni sullo stesso ticket, email discordanti, un indirizzo condiviso fra ticket diversi o più record CRM con la stessa email richiedono revisione. Il controllo degli indirizzi intercetta errori elementari e non verifica l'esistenza di una casella.
3. **Controllare presenza e permesso.** La politica scelta per il webinar limita questa campagna ai presenti con permesso email confermato nel CRM. Una revoca esclude il contatto identificato con certezza. Un'iscrizione annullata con un accesso va verificata. Un assente o un annullato senza accessi resta fuori dalla campagna. Per un presente, un record CRM mancante o un permesso sconosciuto richiedono revisione.
4. **Consegnare la lista.** Le richieste di demo passano al commerciale esistente, oppure alla coda Sales. Chi ha chiesto solo i materiali passa al team eventi. La cartella Excel calcola le liste localmente. Invio email, importazione nel CRM e aggiornamento dei permessi restano attività esterne.

HubSpot consente già di importare i partecipanti a un evento usando l'email come identificativo. Qui la decisione viene prima dell'importazione, per esempio quando una persona compare con due indirizzi o un annullamento arriva dopo il check-in. I CSV di uscita sono liste di lavoro per i team. [Documentazione HubSpot sull'importazione dei partecipanti](https://knowledge.hubspot.com/integrations/import-marketing-event-participants)

## Demo in due minuti

1. Apri il foglio **Follow-up** e filtra la colonna **Esito** su **Sales**. Il ticket **T-0001** passa a **Sara Neri**, il commerciale assegnato nel CRM. Annota il totale Sales, poi rimuovi il filtro.
2. Nel foglio **CRM**, cambia il permesso di Ada in **C6** da **Confermato** a **Revocato**. Torna a **Follow-up**. Il ticket passa a **Escluso**, il motivo diventa **Permesso revocato** e i casi Sales scendono da **3 a 2**. Ripristina il valore per tornare al caso originale.
3. Apri **Riscontri** e confronta **T-0007** con **T-0015**. Il primo mostra un accesso con un'altra email. Il secondo mostra due accessi con la stessa email. Spiega perché solo il primo richiede una verifica.

**Per provare altri dati**, modifica una copia del file. Nei tre fogli sorgente scrivi nelle colonne blu delle righe 6–35, fino a 30 righe per fonte. Per sostituire un export, svuota prima i vecchi valori delle colonne blu, conservando le formule a destra. Il formato dei ticket è `T-` seguito da quattro cifre. I riferimenti `VER-I` e `VER-P` identificano le righe sorgente da controllare nell'export corrente. I [CSV originali](dati/) permettono di ripristinare il caso. I CSV in [liste/](liste/) fotografano l'esempio pubblicato. Dopo una modifica, filtra il foglio **Follow-up** e copia le sole righe visibili per ottenere una nuova lista.

## Cosa puoi dimostrare in colloquio

- **Perché Excel?** Il team ha tre file piccoli e deve approvare una lista. Chi esamina una decisione può risalire dal motivo ai tre export. Per decidere se automatizzare servirebbero frequenza degli eventi, volumi e tempo effettivamente speso in revisione.
- **Perché 11 pratiche da rivedere?** Il campione contiene intenzionalmente molti casi difficili per verificare le regole. Quel conteggio misura la copertura del caso di studio. Una stima del carico del team richiederebbe dati di eventi reali.
- **Come sapresti se il processo aiuta un'azienda?** Su un evento pilota misurerei i minuti necessari a preparare la lista e confronterei le destinazioni suggerite con quelle approvate dal responsabile CRM. In questo esercizio sono verificabili gli esiti delle regole. Tempo risparmiato, conversioni e ricavi richiedono osservazioni sul campo.

**Collaudo.** I 19 esiti sono stati controllati rispetto a una lista di risultati attesi. Il file è stato riaperto e ricalcolato in Microsoft Excel. Sono stati verificati i tre CSV di uscita, il filtro, una revoca, la correzione di un'email discordante e l'aggiunta di una pratica nelle tre fonti. Il caso copre un singolo webinar con export finali e un ticket condiviso fra piattaforme.
