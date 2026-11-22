# Un cambio di appuntamento al desk di una fiera

Sono le **09:45**. Un gruppo di tre persone, già prenotato per le 10:00, chiede di entrare alle 10:30. In quel turno restano **due posti**. Cosa può promettere chi è al desk, e chi deve essere avvisato se il gruppo accetta un altro orario?

**[Apri il foglio Excel per il desk](fiera/agenda-demo.xlsx?raw=true)** · [Prova il cambio in 90 secondi](fiera/README.md#prova-il-cambio-in-90-secondi) · [Vai al caso post-evento](#dopo-il-webinar-del-partner-chi-richiama-sales)

![Esempio simulato. Il gruppo mantiene la visita delle 10:00 e il desk può proporre le 11:00, quando sono disponibili tre posti.](fiera/cambio-turno.png)

Il foglio propone le **11:00**. Finché il gruppo valuta la proposta, mantiene i tre posti alle 10:00. Se accetta, l’operatore aggiorna l’agenda una volta e vede chi deve ancora ricevere il nuovo orario. Chi subentra al desk può riprendere da quell’avviso.

**Da dove nasce l’esempio.** Per Interpack 2026 Marchesini Group ha presentato ORBITA-L in un’area descritta come accessibile su invito e appuntamento. La prenotazione è un requisito documentato. Orari, delegazioni, capienza, referenti e richieste di cambio qui sono inventati per provare la decisione. L’applicabilità dipende da come il team gestisce già questi passaggi. [Pagina ufficiale di Interpack di Marchesini Group](https://www.marchesini.com/en/trade-show) · [Comunicato italiano](https://www.marchesini.com/it/news-media/news/marchesini-group-interpack-2026)

## Dopo il webinar del partner: chi richiama Sales?

Un partner organizza un webinar B2B. La mattina dopo manda al team marketing le iscrizioni e il registro degli ingressi. Il commerciale chiede «Chi ha chiesto una demo?». La collega che ha gestito l’evento chiede «A chi mando le slide?». Il partner conosce le presenze. Il CRM dell’azienda contiene il permesso per il follow-up e il commerciale che segue ciascun contatto.

**Il lavoro di Marketing Ops è consegnare due liste diverse.** Nel caso simulato, dai **17 ingressi registrati** emergono **3 richieste di demo da affidare a Sales** e **2 persone a cui inviare i materiali**. Una persona può entrare due volte, chiedere soltanto le slide oppure aver revocato il permesso. Trattare ogni ingresso come una richiesta commerciale produrrebbe errori già nel primo invio.

| Chi apre il file | Cosa riceve | Esempio dal caso |
| --- | --- | --- |
| **Responsabile Sales** | [3 richieste di demo](liste/demo-per-sales.csv), ciascuna con ID CRM, commerciale o coda e prossima azione | Ada ha richiesto una demo ed è già seguita da Sara Neri. Sara apre il contatto C-001 e le propone due orari. |
| **Team eventi** | [2 destinatari dei materiali](liste/materiali-da-inviare.csv) | Pia si è collegata due volte. Riceve le slide una volta sola. Clara ha già un commerciale, ma ha chiesto solo i materiali. |
| **CRM Ops** | [11 pratiche da verificare](liste/revisione.csv), ognuna con il controllo da fare | Giada si è registrata con un indirizzo e collegata con un altro. Il team chiede conferma al partner prima di usare l’uno o l’altro. |

**[Apri la cartella Excel e prova a cambiare un permesso](riconciliazione-evento.xlsx?raw=true)** · [Tre export di partenza](dati/) · [3 esclusioni motivate](liste/esclusi.csv)

![Lista operativa con richieste, responsabili, record CRM e prossima azione](assets/anteprima.png)

## Dove servirebbe davvero

Il caso riguarda un **evento occasionale gestito da un partner**. L’azienda riceve file, ha un proprio CRM e deve approvare il follow-up prima di contattare i partecipanti. Il ticket condiviso fra iscrizioni e accessi è il requisito per l’abbinamento. L’email mette poi in relazione il partecipante con il CRM.

Per un webinar organizzato direttamente su Zoom con HubSpot già collegato, partirei dalla sincronizzazione nativa di iscrizioni e presenze. HubSpot documenta anche due limiti concreti: alcuni ospiti arrivano senza email e l’iscrizione con un indirizzo secondario non crea l’evento nella cronologia del contatto. Il foglio rende visibili i casi da chiedere all’organizzatore. In Italia, il controllo delle liste di invitati e partecipanti compare anche tra le mansioni di Junior Marketing Automation Analyst di KPMG. [Integrazione ufficiale Zoom–HubSpot](https://knowledge.hubspot.com/integrations/use-hubspot-and-zoom-webinars) · [Ruolo KPMG](https://careers.kpmg.it/job/Milano-Marketing-Junior-Automation-Analyst-Milano/1375576157/)

## Segui tre decisioni nel file

1. **Due accessi, un solo invio.** Nel registro, Pia (T-0015) entra alle 10:24 e rientra alle 10:26 con lo stesso indirizzo. Nelle iscrizioni la richiesta è *Materiali*. Il CRM conferma il permesso. Al team eventi arriva una riga.
2. **Una demo che richiede una telefonata interna prima del contatto.** Fabio (T-0006) entra alle 10:06 e annulla alle 10:54. Entrambe le fonti possono essere corrette. CRM Ops chiede al team eventi se la richiesta di demo è ancora valida.
3. **Una richiesta che esce dalla lista Sales.** Apri una copia di Excel. Nel foglio *CRM* cambia il permesso di Ada, cella C6, da *Confermato* a *Revocato*. In *Follow-up* le demo scendono da 3 a 2, il motivo diventa *Permesso revocato* e il record scompare dalla selezione filtrata su Sales dopo aver riapplicato il filtro. Ripristina il valore per ripetere la demo.

## Come si usa su un altro evento

In una copia del file, sostituisci i valori delle colonne blu nei fogli *Iscrizioni*, *Presenze* e *CRM*, righe 6–35. Conserva le formule alla loro destra. Il caso gestisce un evento alla volta, fino a 30 righe per fonte, con ticket nel formato `T-0001`. Se il partner e il team non condividono un identificativo, serve prima una corrispondenza verificata. Un nome uguale da solo non basta.

Nel foglio *Follow-up* filtra *Sales* o *Marketing* e copia le sole righe visibili. Il campo **ID CRM** indica il record da aprire e **Prossima azione** dice cosa fare. Per una pratica sospesa, controlla la colonna *Motivo* e risali alle fonti nel foglio *Riscontri*. I CSV su GitHub sono la fotografia del caso pubblicato. Excel ricalcola le decisioni quando cambiano gli input.

## Cosa dimostra la prova

**18 righe di iscrizione, 17 ingressi e 16 record CRM producono 19 pratiche.** Le unità sono diverse. Un ticket trovato in entrambe le fonti evento apre una pratica sola. Un check-in senza iscrizione apre comunque una pratica. Una riga senza ticket rimane riconoscibile tramite un riferimento provvisorio. Gli 11 casi in revisione sono stati inseriti apposta per mettere alla prova gli abbinamenti.

I nomi, gli indirizzi e i risultati appartengono a una simulazione. In Microsoft Excel sono stati verificati i 19 esiti, i destinatari, gli ID CRM, le azioni, la revoca del permesso, la correzione di un indirizzo e l’aggiunta di una pratica. Per misurare il valore economico servirebbe un evento pilota: il confronto utile è fra il tempo impiegato dall’operatore e gli errori trovati nelle liste preparate con i due metodi.
