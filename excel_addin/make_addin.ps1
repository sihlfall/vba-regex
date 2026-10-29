param(
    [bool]$createExampleFile = $false,
    [string]$outFileName = "",
    [string]$PythonExecutable = "",
    [ValidateSet('Python', 'Excel')]
    [string]$BuildMethod = 'Python'
)

$ErrorActionPreference = 'Stop'
$outName = if ($createExampleFile) { 'WorksheetFunctionsExample.xlsm' } else { 'VbaRegexAddIn.xlam' }
if ($outFileName) {
    if ([IO.Path]::GetFileName($outFileName) -ne $outFileName -or
        $outFileName.Contains('/') -or $outFileName.Contains('\')) {
        throw 'outFileName must be a file name within the build directory.'
    }
    $extension = if ($createExampleFile) { '.xlsm' } else { '.xlam' }
    if ([IO.Path]::GetExtension($outFileName) -ne $extension) {
        throw "outFileName must use the $extension extension."
    }
    $outName = $outFileName
}
$outDir = Join-Path $PSScriptRoot 'build'
$outFile = Join-Path $outDir $outName

# Use the Windows Python launcher when available; an explicit interpreter also
# supports virtual environments and PowerShell on Linux/macOS.
if ($BuildMethod -eq 'Python') {
    $pythonArguments = @()
    if ($PythonExecutable) {
        $python = (Get-Command $PythonExecutable -CommandType Application -ErrorAction Stop).Source
    } elseif ($env:OS -eq 'Windows_NT' -and (Get-Command py -CommandType Application -ErrorAction SilentlyContinue)) {
        $python = (Get-Command py -CommandType Application).Source
        $pythonArguments = @('-3')
    } else {
        $python = $null
        foreach ($candidate in @('python3', 'python')) {
            $command = Get-Command $candidate -CommandType Application -ErrorAction SilentlyContinue
            if ($command) {
                $python = $command.Source
                break
            }
        }
        if (-not $python) {
            throw 'Install Python 3.10 or newer and requirements-build.txt, or use -BuildMethod Excel.'
        }
    }
    $builder = Join-Path $PSScriptRoot 'build_addin.py'
    & $python @pythonArguments $builder --check
    if ($LASTEXITCODE -ne 0) {
        throw 'The Python build dependencies are unavailable. See the error above, or use -BuildMethod Excel.'
    }
}

# The engine generator writes into the calling directory's build folder.
Push-Location $PSScriptRoot
try {
    & (Join-Path $PSScriptRoot '..\aio\make_aio.ps1') -outModuleName StaticRegexSingle
} finally {
    Pop-Location
}

if ($BuildMethod -eq 'Python') {
    & $python @pythonArguments $builder `
        --engine (Join-Path $outDir 'StaticRegexSingle.bas') `
        --output $outFile
    if ($LASTEXITCODE -ne 0) {
        throw 'Add-in packaging failed. See the Python error above.'
    }
    return
}

# The Excel build has no Python dependency. Use it when Excel COM automation
# works and Trust access to the VBA project object model is enabled.
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
    $events = Get-Content -Raw -Encoding UTF8 (Join-Path $PSScriptRoot 'WorkbookEvents.vba')
    $code = $project.VBComponents.Item($workbook.CodeName).CodeModule
    # A user's VBE settings may have inserted Option Explicit automatically.
    if ($code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
    $code.AddFromString($events)
    $format = if ($createExampleFile) { 52 } else { 55 }
    $workbook.SaveAs($outFile, $format)
    $null = $excel.Run("'" + $workbook.Name.Replace("'", "''") + "'!VbaRegexWorksheetFunctions.RegisterRxFunctions")
    $workbook.Save()
    Write-Host "Built $outFile (using Excel)."
} finally {
    # Cleanup must not replace the original build exception (for example, a
    # type-library failure can make both the build and Quit throw).
    try {
        if ($null -ne $workbook) { $workbook.Close($false) }
    } catch {
        Write-Warning "Could not close the build workbook: $($_.Exception.Message)" -WarningAction Continue
    }
    try {
        if ($null -ne $excel) { $excel.Quit() }
    } catch {
        Write-Warning "Could not quit Excel: $($_.Exception.Message)" -WarningAction Continue
    }
}
