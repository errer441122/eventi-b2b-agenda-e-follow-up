$ErrorActionPreference = 'Stop'
$folder = $PSScriptRoot
$source = Join-Path $folder 'riconciliazione-evento.xlsx'
$app = New-Object -ComObject Excel.Application
$book = $null
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
    function Check([string]$label, $actual, $wanted) {
        if ([string]$actual -cne [string]$wanted) { throw "$label : expected [$wanted], actual [$actual]" }
    }
    Check 'Sheets' $book.Worksheets.Count 5
    Check 'Filter' $summary.ListObjects.Item(1).ShowAutoFilter $true
    $expected = @{}
    foreach ($file in 'follow-up.csv','revisione.csv','esclusi.csv') {
        foreach ($row in (Import-Csv -Encoding UTF8 -LiteralPath (Join-Path $folder "liste/$file"))) {
            if ($expected.ContainsKey($row.Ticket)) { throw "Duplicate ticket in output lists: $($row.Ticket)" }
            $expected[$row.Ticket] = $row
        }
    }
    $fields = 'Email di riferimento','Esito','A chi passa','Motivo','ID CRM','Prossima azione'
    $teamLists = @{'demo-per-sales.csv'='Sales'; 'materiali-da-inviare.csv'='Marketing'}
    foreach ($file in $teamLists.Keys) {
        $team = $teamLists[$file]
        $remaining = @{}
        foreach ($row in $expected.Values) {
            if ($row.Esito -ceq $team) { $remaining[$row.Ticket] = $row }
        }
        foreach ($row in (Import-Csv -Encoding UTF8 -LiteralPath (Join-Path $folder "liste/$file"))) {
            $other = $remaining[$row.Ticket]
            if ($null -eq $other) { throw "$file contains an unexpected ticket: $($row.Ticket)" }
            foreach ($field in $fields) { Check "$file $($row.Ticket) / $field" $row.$field $other.$field }
            $remaining.Remove($row.Ticket)
        }
        Check "$file missing tickets" $remaining.Count 0
    }
    for ($i=9; $i -le 27; $i++) {
        $key = $summary.Cells.Item($i,1).Text
        $row = $expected[$key]
        if ($null -eq $row) { throw "Unknown key $key" }
        Check "$key email" $summary.Cells.Item($i,2).Text $row.'Email di riferimento'
        Check "$key status" $summary.Cells.Item($i,3).Text $row.Esito
        Check "$key handoff" $summary.Cells.Item($i,4).Text $row.'A chi passa'
        Check "$key reason" $summary.Cells.Item($i,5).Text $row.Motivo
        Check "$key CRM ID" $summary.Cells.Item($i,6).Text $row.'ID CRM'
        Check "$key action" $summary.Cells.Item($i,7).Text $row.'Prossima azione'
        $expected.Remove($key)
    }
    Check 'Remaining records' $expected.Count 0
    Check 'Baseline total' $summary.Range('A6').Text '19'
    Check 'Sales' $summary.Range('B6').Text '3'
    Check 'Marketing' $summary.Range('C6').Text '2'
    Check 'Review' $summary.Range('D6').Text '11'
    Check 'Excluded' $summary.Range('E6').Text '3'
    Check 'Next empty' $summary.Range('A28').Text ''
    $crm.Range('C6').Value2 = 'Revocato'
    $app.Calculate()
    Check 'Revocation' $summary.Range('C9').Text 'Escluso'
    Check 'Sales count' $summary.Range('B6').Text '2'
    Check 'Revoked contact removed from handoff' $summary.Range('F9').Text ''
    Check 'Revoked action' $summary.Range('G9').Text 'Escludere da questo follow-up.'
    $crm.Range('C6').Value2 = 'Confermato'
    $att.Range('B10').Value2 = 'giada@example.org'
    $app.Calculate()
    Check 'Confirmed email' $summary.Range('C15').Text 'Sales'
    $reg.Range('A24').Value2 = 'T-0020'
    $reg.Range('B24').Value2 = 'Alex'
    $reg.Range('C24').Value2 = 'alex@example.org'
    $reg.Range('D24').Value2 = 'Iscritto'
    $reg.Range('E24').Value2 = 'Demo'
    $att.Range('A23').Value2 = 'T-0020'
    $att.Range('B23').Value2 = 'alex@example.org'
    $crm.Range('A22').Value2 = 'C-017'
    $crm.Range('B22').Value2 = 'alex@example.org'
    $crm.Range('C22').Value2 = 'Confermato'
    $app.CalculateFullRebuild()
    Check 'Append' $summary.Range('A26').Text 'T-0020'
    Check 'Append sales' $summary.Range('C26').Text 'Sales'
    Check 'Append total' $summary.Range('A6').Text '20'
    Write-Output "Excel $($app.Version): verificati 19 casi, 5 liste CSV, filtro e 3 modifiche (revoca, email corretta, nuova pratica)."
} finally {
    if ($null -ne $book) { $book.Close($false) }
    $app.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
