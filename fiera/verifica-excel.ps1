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
    $app.CalculateFullRebuild()
    foreach($field in 'B6','E6') {if($desk.Range($field).Validation.Type -ne 3){throw "Missing selection in $field"}}
    Check $desk 'A15' 'Posti insufficienti'
    Check $desk 'E9' 2
    Check $desk 'B9' 3
    Check $desk 'B19' 'S03'
    Check $desk 'B20' 3
    Check $desk 'E20' 'Luca'
    $desk.Range('E6').Value2='S03'
    Check $desk 'A15' 'Cambio proponibile'
    Check $turni 'G6' 1
    Check $turni 'G8' 3
    $agenda.Range('D6').Value2='S03'
    Check $desk 'A15' 'Già in agenda'
    Check $turni 'G6' 4
    Check $turni 'G8' 0
    Check $agenda 'I6' 'Avvisare visitatore e referente'
    $agenda.Range('G6').Value2='S03'
    Check $agenda 'I6' 'Avvisare il referente'
    $agenda.Range('H6').Value2='S03'
    Check $agenda 'I6' 'Entrambi aggiornati'
    $desk.Range('B6').Value2='V102'
    Check $desk 'A15' 'Posti insufficienti'
    Check $desk 'B19' 'S05'
    $desk.Range('B6').Value2='V104'
    Check $desk 'A15' 'Chiedere conferma al referente inviti'
    $desk.Range('B6').Value2='V105'
    Check $desk 'A15' 'Prenotazione chiusa'
    $desk.Range('B6').Value2='V101'
    $desk.Range('E6').Value2='S06'
    Check $desk 'B19' 'Nessun turno utile'
    $desk.Range('E6').Value2='S04'
    Check $desk 'A15' 'Turno sospeso'
    $desk.Range('E6').Value2='S03'
    $turni.Range('E8').ClearContents() | Out-Null
    Check $desk 'A15' 'Manca il referente del turno'
    Check $agenda 'I6' 'Confermare il referente'
    $turni.Range('E8').Value2='Luca'
    $desk.Range('B4').Value2=[double](11.25/24)
    $desk.Range('E6').Value2='S02'
    Check $desk 'A15' 'Turno già iniziato'
    Check $desk 'B19' 'Nessun turno utile'
    $desk.Range('B4').Value2=[double](9.75/24)
    $turni.Range('C8').Value2=[double]2
    Check $desk 'A15' 'Rivedere i dati di agenda e turni'
    $turni.Range('C8').Value2=[double]4
    $agenda.Range('A12').Value2='V101'
    Check $desk 'A15' 'Rivedere i dati di agenda e turni'
    $agenda.Range('A12').ClearContents() | Out-Null
    $desk.Range('B6').Value2='V999'
    Check $desk 'A15' 'Codice da verificare'
    $desk.Range('B6').Value2='V107'
    $agenda.Range('A12').Value2='V107'
    $agenda.Range('B12').Value2='Gruppo di prova'
    $agenda.Range('C12').Value2=[double]1
    $agenda.Range('D12').Value2='S01'
    $agenda.Range('E12').Value2='Confermata'
    $agenda.Range('F12').Value2='Verificato'
    Check $desk 'A15' 'Cambio proponibile'
    $desk.Range('B4').ClearContents() | Out-Null
    Check $desk 'A15' 'Ora al desk da verificare'
    "Excel $($app.Version) verified $assertions assertions and both selectors"
} finally {
    if($wb){$wb.Close($false)}
    $app.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
