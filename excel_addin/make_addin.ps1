param(
    [bool]$createExampleFile = $false
)

$ErrorActionPreference = 'Stop'
$outName = if ($createExampleFile) { 'WorksheetFunctionsExample.xlsm' } else { 'VbaRegexAddIn.xlam' }
$outDir = Join-Path $PSScriptRoot 'build'
$outFile = Join-Path $outDir $outName

# The engine generator writes into the calling directory's build folder.
Push-Location $PSScriptRoot
try {
    & (Join-Path $PSScriptRoot '..\aio\make_aio.ps1') -outModuleName StaticRegexSingle
} finally {
    Pop-Location
}

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $project = $workbook.VBProject
    if ($null -eq $project) {
        throw "Enable 'Trust access to the VBA project object model' in Excel's Trust Center."
    }
    foreach ($file in @('build\StaticRegexSingle.bas', 'VbaRegexWorksheetFunctions.bas')) {
        $null = $project.VBComponents.Import((Join-Path $PSScriptRoot $file))
    }
    # Registration restores the function category and argument help when loaded.
    $project.VBComponents.Item($workbook.CodeName).CodeModule.AddFromString(@"
Private Sub Workbook_Open()
    VbaRegexWorksheetFunctions.RegisterRxFunctions
End Sub

Private Sub Workbook_AddinInstall()
    VbaRegexWorksheetFunctions.RegisterRxFunctions
End Sub
"@)
    $format = if ($createExampleFile) { 52 } else { 55 }
    $workbook.SaveAs($outFile, $format)
    $null = $excel.Run("'" + $workbook.Name + "'!VbaRegexWorksheetFunctions.RegisterRxFunctions")
    $workbook.Save()
    Write-Host "Built $outFile"
} finally {
    try {
        if ($null -ne $workbook) { $workbook.Close($false) }
    } finally {
        if ($null -ne $excel) { $excel.Quit() }
    }
}
