param(
    [switch]$TestPackage,
    [string]$AddinFileName = 'VbaRegexAddIn.xlam',
    [ValidateSet('Python', 'Excel')]
    [string]$BuildMethod = 'Python',
    [string]$PythonExecutable = ''
)

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
    if ($rxDimensions -ne 1476) { throw "Rx dimension tests failed: $rxDimensions" }
    Write-Host "Passed $rxDimensions Rx dimension checks."

    # Set an explicit budget only in this unsaved workbook. Production builds
    # and the normal suites above retain the engine's default execution limit.
    $engine = $workbook.VBProject.VBComponents.Item('StaticRegexSingle').CodeModule
    $limitLine = 0
    for ($line = 1; $line -le $engine.CountOfLines; $line++) {
        if ($engine.Lines($line, 1).Trim() -eq 'regex.stepsLimit = DEFAULT_STEPS_LIMIT') {
            if ($limitLine -ne 0) { throw 'Multiple regex step-limit initializers found.' }
            $limitLine = $line
        }
    }
    if ($limitLine -eq 0) { throw 'Regex step-limit initializer not found.' }
    $originalLine = $engine.Lines($limitLine, 1)
    try {
        $engine.ReplaceLine($limitLine, '    regex.stepsLimit = 10000')
        $limitChecks = $excel.Run("'" + $workbook.Name + "'!TestRxFunctions.RunRxExecutionLimitTests")
        if ($limitChecks -is [string] -or $limitChecks -le 0) { throw "Rx execution-limit tests failed: $limitChecks" }
        Write-Host "Passed $limitChecks Rx execution-limit checks with a 10,000-step budget."
    } finally {
        $engine.ReplaceLine($limitLine, $originalLine)
    }
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
    & (Join-Path $PSScriptRoot 'make_addin.ps1') -outFileName $AddinFileName -BuildMethod $BuildMethod -PythonExecutable $PythonExecutable
    & (Join-Path $PSScriptRoot 'test-built-addin.ps1') -AddinFileName $AddinFileName
}
