# Il foglio post-evento: come leggerlo e riusarlo

Dettagli del [secondo caso](README.md#2-dopo-levento-a-chi-passa-ogni-richiesta). **[Scarica il foglio](riconciliazione-evento.xlsx?raw=true)** · [Tre export di partenza](dati/)

Per ripetere il collaudo su Windows con Excel desktop, scarica il repository e avvia dalla sua cartella `.\verifica-post-evento.ps1` in PowerShell. Lo script apre il foglio in sola lettura, confronta i risultati con i CSV e chiude la sessione senza salvare le modifiche di prova. [Apri lo script](verifica-post-evento.ps1)

## Segui quattro decisioni nel file

1. **Due accessi, un solo invio.** Nel registro, Pia (T-0015) entra alle 10:24 e rientra alle 10:26 con lo stesso indirizzo. Nelle iscrizioni la richiesta è *Materiali*. Il CRM conferma il permesso. Al team eventi arriva una riga.
2. **Una demo che richiede una telefonata interna prima del contatto.** Fabio (T-0006) entra alle 10:06 e annulla alle 10:54. Entrambe le fonti possono essere corrette. CRM Ops chiede al team eventi se la richiesta di demo è ancora valida.
3. **Un contatto raccolto allo stand.** Klaus (INT-0101) non è nel CRM, ha dato il consenso e chiede un’offerta: va a Elio Riva, commerciale per la Germania, con *Da creare* al posto dell’ID CRM. Joana viene dal Brasile, che non ha un commerciale nelle Regole, e va in *Coda Sales*. Hugo ha negato il consenso ed è escluso.
4. **Una richiesta che esce dalla lista Sales.** Apri una copia di Excel. Nel foglio *CRM* cambia il permesso di Ada, cella C6, da *Confermato* a *Revocato*. In *Follow-up* le richieste per Sales scendono da 7 a 6 e il motivo diventa *Permesso revocato*. Ripristina il valore per ripetere la demo.

## Come si usa su un altro evento

In una copia del file, sostituisci i valori nei fogli *Iscrizioni*, *Presenze* e *CRM* e conserva le formule alla loro destra. I contatti raccolti allo stand vanno in *Iscrizioni* con stato *Allo stand*, che vale come presenza, insieme a paese e consenso. Serve un codice comune fra iscrizioni e ingressi, per esempio il numero del badge o il codice dell’invito: senza, la pratica va in verifica. Un nome uguale da solo non basta.

Nella scheda *Regole* si adattano le categorie di richiesta con la lista e l’azione di ciascuna, i commerciali per paese, le diciture del permesso usate dal CRM e cosa fare con gli iscritti che non si sono presentati. Una richiesta o una dicitura che non è nelle Regole va in verifica, non a un commerciale.

Nel foglio *Follow-up* filtra *Sales* o *Marketing* e copia le sole righe visibili. Il campo **ID CRM** indica il record da aprire, o *Da creare*, e **Prossima azione** dice cosa fare. Per una pratica sospesa, controlla la colonna *Motivo* e risali alle fonti nel foglio *Riscontri*. I CSV su GitHub sono la fotografia del caso pubblicato. Excel ricalcola le decisioni quando cambiano gli input.

## Cosa dimostra la prova

**25 righe di iscrizione o contatto, 17 ingressi e 16 record CRM producono 26 pratiche.** Le unità sono diverse. Un ticket trovato in entrambe le fonti evento apre una pratica sola. Un check-in senza iscrizione apre comunque una pratica. Una riga senza ticket rimane riconoscibile tramite un riferimento provvisorio. I 13 casi da verificare sono stati inseriti apposta per mettere alla prova gli abbinamenti.

I nomi, gli indirizzi e i risultati appartengono a una simulazione. In Microsoft Excel sono stati verificati i 26 esiti, i destinatari, gli ID CRM, le azioni e sette modifiche: la revoca di un permesso, la correzione di un indirizzo, gli assenti mandati al team eventi, una nuova categoria di richiesta, una nuova dicitura del permesso, un paese senza commerciale e l’aggiunta di una pratica. Per misurare il valore economico servirebbe un evento pilota: il confronto utile è fra il tempo impiegato dall’operatore e gli errori trovati nelle liste preparate con i due metodi.

## Dove servirebbe davvero

Il caso riguarda un evento in sede o una fiera in cui l’azienda riceve elenchi da fonti diverse, ha un proprio CRM e deve approvare il follow-up prima di contattare i partecipanti. Il codice comune fra iscrizioni e accessi permette l’abbinamento; l’email mette poi in relazione il partecipante con il CRM.

Se iscrizioni e badge arrivano già nel CRM tramite un’integrazione, il foglio serve per i casi che l’integrazione lascia aperti, come un ospite senza email o un’iscrizione con l’indirizzo secondario del contatto.
