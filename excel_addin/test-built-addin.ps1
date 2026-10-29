param([string]$AddinFileName = 'VbaRegexAddIn.xlam')

# Tests a built add-in from a separate workbook, including its load events.
$ErrorActionPreference = 'Stop'
$excel = $null
$addin = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $addin = $excel.Workbooks.Open((Join-Path (Join-Path $PSScriptRoot 'build') $AddinFileName))
    $formulaPrefix = "='" + $addin.Name.Replace("'", "''") + "'!"
    $book = $excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    $sheet.Range('A1').Formula2 = $formulaPrefix + 'RxTest("ABC","(?i)abc")'
    $sheet.Range('A2').Formula2 = $formulaPrefix + 'RxMatch("a1b2","\d",,2)'
    $sheet.Range('A3').Formula2 = $formulaPrefix + 'RxReplace("a1b2","\d","X",1)'
    $sheet.Range('C1').Formula2 = $formulaPrefix + 'RxMatches("AB-12 CD-34","([A-Z]+)-(\d+)",{"$1","$2"})'
    $sheet.Range('F1').Formula2 = $formulaPrefix + 'RxSplit("red, green;blue","[,;]\s*")'
    $sheet.Range('A5').Formula2 = $formulaPrefix + 'RxFind("ab12cd345","\d+",2)'
    $sheet.Range('A6').Formula2 = $formulaPrefix + 'RxLength("ab12cd345","\d+",2)'
    $sheet.Range('A7').Formula2 = $formulaPrefix + 'RxCount("ab12cd345","\d+")'
    $sheet.Range('A10').Value2 = 'a,b,c'
    $sheet.Range('A11').Value2 = 'd'
    $sheet.Range('A12').Value2 = 'e,f'
    $sheet.Range('F5').Formula2 = $formulaPrefix + 'RxSplit(A10:A12,",")'
    $sheet.Range('J1').Formula2 = $formulaPrefix + 'RxTest("a1b2","\d",3)'
    $sheet.Range('J2').Formula2 = $formulaPrefix + 'RxMatch("a1b2","\d",,,3)'
    $sheet.Range('J3').Formula2 = $formulaPrefix + 'RxReplace("a1b2","\d","X",,3)'
    $sheet.Range('J5').Formula2 = $formulaPrefix + 'RxFind("a1b2","\d",,3)'
    $sheet.Range('J6').Formula2 = $formulaPrefix + 'RxLength("a1b22","\d+",,3)'
    $sheet.Range('J7').Formula2 = $formulaPrefix + 'RxCount("a1b2","\d",3)'
    $sheet.Range('L1').Formula2 = $formulaPrefix + 'RxMatches("a1b2","\d",,3)'
    $sheet.Range('L5').Formula2 = $formulaPrefix + 'RxSplit("aXbxc","(?i)x")'
    $sheet.Calculate()
    if ($sheet.Range('A1').Value2 -ne $true) { throw 'Packaged RxTest failed' }
    if ($sheet.Range('A2').Value2 -cne '2') { throw 'Packaged RxMatch failed' }
    if ($sheet.Range('A3').Value2 -cne 'aXb2') { throw 'Packaged RxReplace failed' }
    if ($sheet.Range('C1').Value2 -cne 'AB' -or $sheet.Range('D2').Value2 -cne '34') { throw 'Packaged RxMatches failed' }
    if ($sheet.Range('F1').Value2 -cne 'red' -or $sheet.Range('H1').Value2 -cne 'blue') { throw 'Packaged RxSplit failed' }
    if ($sheet.Range('A5').Value2 -ne 7) { throw 'Packaged RxFind failed' }
    if ($sheet.Range('A6').Value2 -ne 3) { throw 'Packaged RxLength failed' }
    if ($sheet.Range('A7').Value2 -ne 2) { throw 'Packaged RxCount failed' }
    if ($sheet.Range('F5').Value2 -cne 'a' -or $sheet.Range('H5').Value2 -cne 'c') { throw 'Packaged RxSplit column first row failed' }
    if ($sheet.Range('F6').Value2 -cne 'd' -or -not $sheet.Evaluate('ISNA(G6)')) { throw 'Packaged RxSplit column padding failed' }
    if ($sheet.Range('F7').Value2 -cne 'e' -or $sheet.Range('G7').Value2 -cne 'f') { throw 'Packaged RxSplit column final row failed' }
    if ($sheet.Range('J1').Value2 -ne $true) { throw 'Packaged RxTest startAt failed' }
    if ($sheet.Range('J2').Value2 -cne '2') { throw 'Packaged RxMatch startAt failed' }
    if ($sheet.Range('J3').Value2 -cne 'a1bX') { throw 'Packaged RxReplace startAt failed' }
    if ($sheet.Range('J5').Value2 -ne 4) { throw 'Packaged RxFind startAt failed' }
    if ($sheet.Range('J6').Value2 -ne 2) { throw 'Packaged RxLength startAt failed' }
    if ($sheet.Range('J7').Value2 -ne 1) { throw 'Packaged RxCount startAt failed' }
    if ($sheet.Range('L1').Value2 -cne '2') { throw 'Packaged RxMatches startAt failed' }
    if ($sheet.Range('L5').Value2 -cne 'a' -or $sheet.Range('N5').Value2 -cne 'c') { throw 'Packaged RxSplit modifier failed' }
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
