$ErrorActionPreference='Stop'
$file=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot 'agenda-demo.xlsx')).Path
$app=New-Object -ComObject Excel.Application
$wb=$null
$assertions=0
try {
    $app.Visible=$false
    $app.DisplayAlerts=$false
    $app.AutomationSecurity=3
    $wb=$app.Workbooks.Open($file,0,$true)
    $desk=$wb.Worksheets.Item('Cambio turno')
    $agenda=$wb.Worksheets.Item('Agenda')
    $turni=$wb.Worksheets.Item('Turni')
    function Check($sheet,$cell,$wanted) {
        $app.Calculate()
        $actual=$sheet.Range($cell).Text
        if($actual -cne [string]$wanted) {throw "$($sheet.Name)!$cell expected [$wanted] found [$actual]"}
        $script:assertions++
    }
    function Told($row) {
        # registra che visitatore e referente hanno ricevuto l'ultima versione
        $app.Calculate()
        $agenda.Range("H$row").Value2=$agenda.Range("G$row").Value2
        $agenda.Range("J$row").Value2=$agenda.Range("I$row").Value2
    }
    function Time($h,$m) { [double](($h*60+$m)/1440) }
    $app.CalculateFullRebuild()
    foreach($field in 'B6','E6') {if($desk.Range($field).Validation.Type -ne 3){throw "Missing selection in $field"}}
    foreach($field in 'H6','J6','H305','J305') {if($agenda.Range($field).Validation.Type -ne 3){throw "Missing selection in Agenda!$field"}}

    # Caso iniziale: 3 persone alle 10:00 chiedono le 10:30, dove restano 2 posti
    Check $desk 'A16' 'Posti insufficienti'
    Check $desk 'E10' 2
    Check $desk 'B9' 3
    Check $desk 'E9' 'Sala demo A'
    # la Sala demo B alle 10:45 ha posto, ma e un'altra area: si propone S03
    Check $desk 'B20' 'S03'
    Check $desk 'B21' 3
    Check $desk 'E21' 'Luca'
    Check $agenda 'K11' 'Avvisare il referente'
    Check $agenda 'K9' 'Verificare invito'
    Check $agenda 'K10' 'Chiusa'

    # Proposta, poi il si: i posti si spostano solo quando cambia l'Agenda
    $desk.Range('E6').Value2='S03'
    Check $desk 'A16' 'Cambio proponibile'
    Check $turni 'I6' 1
    Check $turni 'I8' 3
    $agenda.Range('D6').Value2='S03'
    Check $desk 'A16' 'Già in agenda'
    Check $turni 'I6' 4
    Check $turni 'I8' 0
    Check $agenda 'K6' 'Avvisare visitatore e referente'
    $agenda.Range('H6').Value2=$agenda.Range('G6').Value2
    Check $agenda 'K6' 'Avvisare il referente'
    $agenda.Range('J6').Value2=$agenda.Range('I6').Value2
    Check $agenda 'K6' 'Entrambi aggiornati'

    # Il turno slitta di 20 minuti: tutti i gruppi di quel turno tornano da avvisare
    $turni.Range('C8').Value2=[double](Time 11 20)
    Check $agenda 'K6' 'Avvisare visitatore e referente'
    Check $agenda 'K8' 'Avvisare visitatore e referente'
    $turni.Range('C8').Value2=[double](Time 11 0)
    Check $agenda 'K6' 'Entrambi aggiornati'

    # Cambia il referente: va avvisato solo il nuovo referente
    $turni.Range('G8').Value2='Marta'
    Check $agenda 'K6' 'Avvisare il referente'
    Check $agenda 'K8' 'Avvisare il referente'
    $turni.Range('G8').Value2='Luca'
    Check $agenda 'K6' 'Entrambi aggiornati'

    # Il gruppo arriva in 4: il turno va oltre capienza, ma il desk non si blocca
    $agenda.Range('C6').Value2=[double]4
    Check $turni 'I8' -1
    Check $turni 'J8' 'OK'
    Check $agenda 'K6' 'Turno oltre capienza: spostare un gruppo o alzare la capienza'
    Check $agenda 'K8' 'Turno oltre capienza: spostare un gruppo o alzare la capienza'
    Check $desk 'A16' 'Posti insufficienti'
    Check $desk 'B20' 'S05'
    # eccezione accettata alzando la capienza: resta da avvisare il referente del nuovo numero
    $turni.Range('E8').Value2=[double]5
    Check $agenda 'K6' 'Avvisare il referente'
    $turni.Range('E8').Value2=[double]4
    $agenda.Range('C6').Value2=[double]3
    Check $agenda 'K6' 'Entrambi aggiornati'

    # Altri gruppi
    $desk.Range('B6').Value2='V102'
    Check $desk 'A16' 'Posti insufficienti'
    Check $desk 'B20' 'S05'
    Check $desk 'B21' 4
    $desk.Range('B6').Value2='V104'
    Check $desk 'A16' 'Chiedere conferma al referente inviti'
    $desk.Range('B6').Value2='V105'
    Check $desk 'A16' 'Prenotazione chiusa'
    $desk.Range('B6').Value2='V101'
    $desk.Range('E6').Value2='S06'
    Check $desk 'A16' 'Posti insufficienti'
    Check $desk 'B20' 'Nessun turno utile'
    Check $desk 'A17' "Nessuna alternativa in quest’area oggi: decide il responsabile."
    $desk.Range('E6').Value2='S04'
    Check $desk 'A16' 'Turno sospeso'
    Check $desk 'B20' 'S05'
    $desk.Range('E6').Value2='S03'
    $turni.Range('G8').ClearContents() | Out-Null
    Check $desk 'A16' 'Manca il referente del turno'
    Check $agenda 'K6' 'Confermare il referente'
    $turni.Range('G8').Value2='Luca'

    # Cliente in ritardo: il turno e gia iniziato, si propone il primo successivo con posto
    $desk.Range('E4').Value2=[double](Time 11 15)
    $desk.Range('E6').Value2='S02'
    Check $desk 'A16' 'Turno già iniziato'
    Check $desk 'B20' 'S05'
    $desk.Range('E4').Value2=[double](Time 9 45)

    # Secondo giorno: le alternative restano nello stesso giorno
    $desk.Range('E6').Value2='S09'
    Check $desk 'A16' 'Posti insufficienti'
    Check $desk 'B20' 'S10'
    $desk.Range('B6').Value2='V108'
    $desk.Range('E6').Value2='S10'
    Check $desk 'A16' 'Cambio proponibile'
    Check $turni 'K15' ''

    # Dati incoerenti: il desk non propone nulla finche non si correggono
    $desk.Range('B6').Value2='V101'
    $desk.Range('E6').Value2='S02'
    $turni.Range('E8').Value2='x'
    Check $turni 'J8' 'Verificare capienza'
    Check $desk 'A16' 'Rivedere i dati di agenda e turni'
    $turni.Range('E8').Value2=[double]4
    # un secondo turno alla stessa ora, nella stessa area e giornata di S03
    $turni.Range('A16').Value2='S11'
    $turni.Range('B16').Value2=[double]$turni.Range('B8').Value2
    $turni.Range('C16').Value2=[double](Time 11 0)
    $turni.Range('D16').Value2='Sala demo A'
    $turni.Range('E16').Value2=[double]4
    $turni.Range('F16').Value2='Aperto'
    Check $turni 'J16' 'Verificare orario doppio'
    Check $desk 'A16' 'Rivedere i dati di agenda e turni'
    $turni.Range('A16:G16').ClearContents() | Out-Null
    $agenda.Range('A14').Value2='V101'
    Check $desk 'A16' 'Rivedere i dati di agenda e turni'
    $agenda.Range('A14').ClearContents() | Out-Null
    $desk.Range('B6').Value2='V999'
    Check $desk 'A16' 'Codice da verificare'

    # Un gruppo senza appuntamento aggiunto al momento
    $agenda.Range('A14').Value2='V109'
    $agenda.Range('B14').Value2='Gruppo di prova'
    $agenda.Range('C14').Value2=[double]1
    $agenda.Range('D14').Value2='S01'
    $agenda.Range('E14').Value2='Confermata'
    $agenda.Range('F14').Value2='Da verificare'
    $desk.Range('B6').Value2='V109'
    Check $desk 'A16' 'Chiedere conferma al referente inviti'
    $agenda.Range('F14').Value2='Verificato'
    Check $desk 'A16' 'Cambio proponibile'
    Check $agenda 'K14' 'Avvisare visitatore e referente'
    Told 14
    Check $agenda 'K14' 'Entrambi aggiornati'
    $desk.Range('E4').ClearContents() | Out-Null
    Check $desk 'A16' 'Giorno e ora al desk da verificare'
    "Excel $($app.Version) verified $assertions assertions and the selectors"
} finally {
    if($wb){$wb.Close($false)}
    $app.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
