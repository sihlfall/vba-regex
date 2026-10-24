param([switch]$TestPackage)

# Requires Windows Excel and "Trust access to the VBA project object model".
$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
$excel = $null
$workbook = $null
try {
    & ..\aio\make_aio.ps1 -outModuleName StaticRegexSingle
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    foreach ($file in @('build\StaticRegexSingle.bas', 'VbaRegexWorksheetFunctions.bas', 'TestWorksheetDimensions.bas', 'TestWorksheetFunctions.bas', 'TestRxFunctions.bas', 'TestRxDimensions.bas')) {
        $null = $workbook.VBProject.VBComponents.Import((Join-Path $PSScriptRoot $file))
    }
    $null = $excel.Run("'" + $workbook.Name + "'!VbaRegexWorksheetFunctions.RegisterRxFunctions")
    $checks = $excel.Run("'" + $workbook.Name + "'!TestWorksheetDimensions.RunDimensionTests")
    if ($checks -ne 936) { throw "Dimension tests failed: $checks" }
    Write-Host "Passed $checks dimension checks."
    $basicChecks = $excel.Run("'" + $workbook.Name + "'!TestWorksheetFunctions.RunBasicTests")
    if ($basicChecks -is [string] -or $basicChecks -le 0) { throw "Basic tests failed: $basicChecks" }
    Write-Host "Passed $basicChecks basic functionality checks."
    $rxChecks = $excel.Run("'" + $workbook.Name + "'!TestRxFunctions.RunRxTests")
    if ($rxChecks -is [string] -or $rxChecks -le 0) { throw "Rx tests failed: $rxChecks" }
    Write-Host "Passed $rxChecks Rx API checks."
    $rxDimensions = $excel.Run("'" + $workbook.Name + "'!TestRxDimensions.RunRxDimensionTests")
    if ($rxDimensions -ne 2232) { throw "Rx dimension tests failed: $rxDimensions" }
    Write-Host "Passed $rxDimensions Rx dimension checks."
} finally {
    try {
        if ($null -ne $workbook) { $workbook.Close($false) }
    } finally {
        try {
            if ($null -ne $excel) { $excel.Quit() }
        } finally {
            Pop-Location
        }
    }
}
if ($TestPackage) {
    & (Join-Path $PSScriptRoot 'make_addin.ps1')
    & (Join-Path $PSScriptRoot 'test-built-addin.ps1')
}
