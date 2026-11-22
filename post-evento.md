# Il foglio post-evento: come leggerlo e riusarlo

Dettagli del [secondo caso](README.md#2-dopo-levento-a-chi-passa-ogni-richiesta). **[Scarica il foglio](riconciliazione-evento.xlsx?raw=true)** · [Tre export di partenza](dati/)

Per ripetere il collaudo su Windows con Excel desktop, scarica il repository e avvia dalla sua cartella `.\verifica-post-evento.ps1` in PowerShell. Lo script apre il foglio in sola lettura, confronta i risultati con i CSV e chiude la sessione senza salvare le modifiche di prova. [Apri lo script](verifica-post-evento.ps1)

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

## Dove servirebbe davvero

Il caso riguarda un **evento occasionale gestito da un partner**. L’azienda riceve file, ha un proprio CRM e deve approvare il follow-up prima di contattare i partecipanti. Il ticket condiviso fra iscrizioni e accessi è il requisito per l’abbinamento. L’email mette poi in relazione il partecipante con il CRM.

Quando Zoom e HubSpot sono già collegati, la sincronizzazione può fornire iscrizioni e presenze. Restano da verificare casi come un ospite senza email o un’iscrizione con l’indirizzo secondario del contatto. Il foglio permette di raccoglierli in una lista da consegnare all’organizzatore. [Documentazione dell’integrazione Zoom–HubSpot](https://knowledge.hubspot.com/integrations/use-hubspot-and-zoom-webinars)

