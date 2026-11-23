$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
$source = Join-Path $folder 'riconciliazione-evento.xlsx'
$app = New-Object -ComObject Excel.Application
$book = $null
$checks = 0
try {
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $app.AutomationSecurity = 3
    $book = $app.Workbooks.Open($source, 0, $true)
    $app.CalculateFullRebuild()
    $summary = $book.Worksheets.Item('Follow-up')
    $reg = $book.Worksheets.Item('Iscrizioni')
    $att = $book.Worksheets.Item('Presenze')
    $crm = $book.Worksheets.Item('CRM')
    $rules = $book.Worksheets.Item('Regole')
    function Check([string]$label, $actual, $wanted) {
        if ([string]$actual -cne [string]$wanted) { throw "$label : expected [$wanted], actual [$actual]" }
        $script:checks++
    }
    function Row([int]$r) { $app.Calculate(); return @(1..9 | ForEach-Object { $summary.Cells.Item($r, $_).Text }) }
    Check 'Sheets' $book.Worksheets.Count 6
    Check 'Filter' $summary.ListObjects.Item(1).ShowAutoFilter $true

    # Le liste pubblicate in liste/ coincidono con il foglio
    $expected = @{}
    foreach ($file in 'follow-up.csv','revisione.csv','esclusi.csv') {
        foreach ($row in (Import-Csv -Encoding UTF8 -LiteralPath (Join-Path $folder "liste/$file"))) {
            if ($expected.ContainsKey($row.Ticket)) { throw "Duplicate ticket in output lists: $($row.Ticket)" }
            $expected[$row.Ticket] = $row
        }
    }
    $fields = 'Email di riferimento','Esito','A chi passa','Motivo','ID CRM','Azienda','Paese','Prossima azione'
    $teamLists = @{'richieste-per-sales.csv'='Sales'; 'materiali-da-inviare.csv'='Marketing'}
    foreach ($file in $teamLists.Keys) {
        $team = $teamLists[$file]
        $remaining = @{}
        foreach ($row in $expected.Values) { if ($row.Esito -ceq $team) { $remaining[$row.Ticket] = $row } }
        foreach ($row in (Import-Csv -Encoding UTF8 -LiteralPath (Join-Path $folder "liste/$file"))) {
            $other = $remaining[$row.Ticket]
            if ($null -eq $other) { throw "$file contains an unexpected ticket: $($row.Ticket)" }
            foreach ($field in $fields) { Check "$file $($row.Ticket) / $field" $row.$field $other.$field }
            $remaining.Remove($row.Ticket)
        }
        Check "$file missing tickets" $remaining.Count 0
    }
    for ($i = 9; $i -le 34; $i++) {
        $key = $summary.Cells.Item($i, 1).Text
        $row = $expected[$key]
        if ($null -eq $row) { throw "Unknown key $key" }
        for ($c = 2; $c -le 9; $c++) { Check "$key / $($fields[$c-2])" $summary.Cells.Item($i, $c).Text $row.($fields[$c-2]) }
        $expected.Remove($key)
    }
    Check 'Remaining records' $expected.Count 0
    Check 'Next empty' $summary.Range('A35').Text ''
    Check 'Total' $summary.Range('A6').Text '26'
    Check 'Sales' $summary.Range('B6').Text '7'
    Check 'Marketing' $summary.Range('C6').Text '2'
    Check 'Review' $summary.Range('D6').Text '13'
    Check 'Excluded' $summary.Range('E6').Text '4'

    # Casi chiave, scritti a mano per non dipendere solo dai CSV
    Check 'Existing contact without owner goes to the country owner' (Row 10)[3] 'Paola Gatti'
    $klaus = Row 26
    Check 'Stand contact with consent' "$($klaus[2])|$($klaus[3])|$($klaus[5])" 'Sales|Elio Riva|Da creare'
    Check 'Stand contact action' $klaus[8] "Creare il contatto nel CRM. Preparare l’offerta richiesta."
    Check 'Stand contact who said no' "$((Row 28)[2])|$((Row 28)[4])" 'Escluso|Senza consenso'
    Check 'Country without an owner' (Row 29)[3] 'Coda Sales'
    Check 'Stand contact already in the CRM' "$((Row 31)[3])|$((Row 31)[5])" 'Sara Neri|C-016'
    Check 'Unknown request' (Row 32)[4] 'Richiesta da verificare'
    Check 'Trimmed email' (Row 19)[1] 'marta@example.org'

    # Permesso revocato nel CRM: esce dalla lista Sales
    $crm.Range('C6').Value2 = 'Revocato'
    $ada = Row 9
    Check 'Revocation' "$($ada[2])|$($ada[5])|$($ada[8])" 'Escluso||Escludere da questo follow-up.'
    Check 'Sales count' $summary.Range('B6').Text '6'
    $crm.Range('C6').Value2 = 'Confermato'

    # Email corretta in Presenze: la pratica passa a Sales, al commerciale del paese
    $att.Range('B10').Value2 = 'giada@example.org'
    Check 'Confirmed email' "$((Row 15)[2])|$((Row 15)[3])" 'Sales|Sara Neri'
    $att.Range('B10').Value2 = 'giada.ufficio@example.org'

    # Regola assenti: invio della registrazione a chi aveva il permesso
    $rules.Range('A28').Value2 = 'Marketing'
    $dario = Row 12
    Check 'Absent to Marketing' "$($dario[2])|$($dario[3])|$($dario[8])" 'Marketing|Team eventi|Inviare la registrazione a chi non ha partecipato.'
    Check 'Cancelled stays out' (Row 13)[2] 'Escluso'
    $rules.Range('A28').Value2 = 'Escluso'

    # Nuova categoria di richiesta aggiunta nelle Regole
    $rules.Range('A10').Value2 = 'Assistenza'; $rules.Range('B10').Value2 = 'Sales'
    $rules.Range('C10').Value2 = 'Service'; $rules.Range('D10').Value2 = 'Inoltrare la richiesta al service.'
    $teo = Row 32
    Check 'New request type' "$($teo[2])|$($teo[3])|$($teo[5])|$($teo[8])" 'Sales|Service|Da creare|Creare il contatto nel CRM. Inoltrare la richiesta al service.'
    $rules.Range('A10:D10').ClearContents() | Out-Null

    # Dicitura del permesso diversa nel CRM: prima si verifica, poi la si insegna alle Regole
    $crm.Range('C18').Value2 = 'Opt-in'
    Check 'Unknown permission wording' (Row 22)[4] 'Permesso da verificare'
    $rules.Range('I11').Value2 = 'Opt-in'; $rules.Range('J11').Value2 = 'Confermato'
    Check 'Mapped permission wording' "$((Row 22)[2])|$((Row 22)[5])" 'Marketing|C-013'
    $rules.Range('I11:J11').ClearContents() | Out-Null
    $crm.Range('C18').Value2 = 'Sconosciuto'

    # Paese senza commerciale
    $reg.Range('H7').Value2 = 'Brasile'
    $bruno = Row 10
    Check 'Queue when no country owner' "$($bruno[3])|$($bruno[8])" 'Coda Sales|Assegnare un commerciale. Proporre due orari per la demo.'
    $reg.Range('H7').Value2 = 'Spagna'

    # Nuova pratica incollata da un export: paese scritto in minuscolo e con spazi
    $reg.Range('A31').Value2 = 'T-0020'
    $reg.Range('B31').Value2 = 'Alex'
    $reg.Range('C31').Value2 = 'alex@example.org'
    $reg.Range('D31').Value2 = 'Iscritto'
    $reg.Range('E31').Value2 = 'Demo'
    $reg.Range('H31').Value2 = '  germania '
    $att.Range('A23').Value2 = 'T-0020'
    $att.Range('B23').Value2 = 'alex@example.org'
    $crm.Range('A22').Value2 = 'C-017'
    $crm.Range('B22').Value2 = 'alex@example.org'
    $crm.Range('C22').Value2 = 'Confermato'
    $app.CalculateFullRebuild()
    $alex = Row 33
    Check 'Append' "$($alex[0])|$($alex[2])|$($alex[3])|$($alex[5])" 'T-0020|Sales|Elio Riva|C-017'
    Check 'Append total' $summary.Range('A6').Text '27'
    Write-Output "Excel $($app.Version): $checks controlli superati su 26 pratiche, 5 liste CSV e 7 modifiche (revoca, email corretta, assenti, nuova richiesta, nuova dicitura del permesso, paese, nuova pratica)."
} finally {
    if ($null -ne $book) { $book.Close($false) }
    $app.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
