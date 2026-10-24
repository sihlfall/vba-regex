# Tests a built add-in from a separate workbook, including its load events.
$ErrorActionPreference = 'Stop'
$excel = $null
$addin = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $addin = $excel.Workbooks.Open((Join-Path $PSScriptRoot 'build\VbaRegexAddIn.xlam'))
    $book = $excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    $sheet.Range('A1').Formula2 = "='VbaRegexAddIn.xlam'!RxTest(""ABC"",""abc"",""i"")"
    $sheet.Range('A2').Formula2 = "='VbaRegexAddIn.xlam'!RxMatch(""a1b2"",""\d"",,2)"
    $sheet.Range('A3').Formula2 = "='VbaRegexAddIn.xlam'!RxReplace(""a1b2"",""\d"",""X"",1)"
    $sheet.Range('C1').Formula2 = '=''VbaRegexAddIn.xlam''!RxMatches("AB-12 CD-34","([A-Z]+)-(\d+)",{"$1","$2"})'
    $sheet.Range('F1').Formula2 = "='VbaRegexAddIn.xlam'!RxSplit(""red, green;blue"",""[,;]\s*"")"
    $sheet.Range('A5').Formula2 = "='VbaRegexAddIn.xlam'!RxFind(""ab12cd345"",""\d+"",2)"
    $sheet.Range('A6').Formula2 = "='VbaRegexAddIn.xlam'!RxLength(""ab12cd345"",""\d+"",2)"
    $sheet.Range('A7').Formula2 = "='VbaRegexAddIn.xlam'!RxCount(""ab12cd345"",""\d+"")"
    $sheet.Calculate()
    if ($sheet.Range('A1').Value2 -ne $true) { throw 'Packaged RxTest failed' }
    if ($sheet.Range('A2').Value2 -cne '2') { throw 'Packaged RxMatch failed' }
    if ($sheet.Range('A3').Value2 -cne 'aXb2') { throw 'Packaged RxReplace failed' }
    if ($sheet.Range('C1').Value2 -cne 'AB' -or $sheet.Range('D2').Value2 -cne '34') { throw 'Packaged RxMatches failed' }
    if ($sheet.Range('F1').Value2 -cne 'red' -or $sheet.Range('H1').Value2 -cne 'blue') { throw 'Packaged RxSplit failed' }
    if ($sheet.Range('A5').Value2 -ne 7) { throw 'Packaged RxFind failed' }
    if ($sheet.Range('A6').Value2 -ne 3) { throw 'Packaged RxLength failed' }
    if ($sheet.Range('A7').Value2 -ne 2) { throw 'Packaged RxCount failed' }
    # Exercise registration while IsAddin=True, and verify it restores that state.
    $null = $excel.Run("'" + $addin.Name + "'!VbaRegexWorksheetFunctions.RegisterRxFunctions")
    if (-not $addin.IsAddin) { throw 'Registration changed the workbook add-in state' }
    Write-Host 'Packaged add-in: all eight worksheet functions and registration passed.'
} finally {
    try {
        if ($null -ne $book) { $book.Close($false) }
    } finally {
        try {
            if ($null -ne $addin) { $addin.Close($false) }
        } finally {
            if ($null -ne $excel) { $excel.Quit() }
        }
    }
}
