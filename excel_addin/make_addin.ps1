param(
    [bool]$createExampleFile = $false
)

$OUT_FILE_NAME = $( if($createExamplefile) { "WorksheetFunctionsExample.xlsm" } else { "VbaRegexAddIn.xlam" } )

$SOURCE_FILES = ".\build\StaticRegexSingle.bas", ".\VbaRegexWorksheetFunctions.bas"

$scriptBuildSinglePath = "..\aio\make_aio.ps1"

$outDirPath = ".\build"

$scriptPath = Split-Path $MyInvocation.MyCommand.Path -Parent
$projectRootPath = $scriptPath # Join-Path -Path $scriptPath -ChildPath ".."

$fullSourceFileNames = $SOURCE_FILES | % { Join-Path $projectRootPath $_ }
$fullOutFileName = Join-Path -Path $scriptPath -ChildPath $outDirPath | Join-Path -ChildPath $OUT_FILE_NAME

Write-Host $fullOutFileName

# Create build directory, if necessary
if (-not (Test-Path -Path $outDirPath)) {
    New-Item -Path $outDirPath -ItemType Directory > $null
}

# Remove existing add-in file

Remove-Item -Path $fullOutFileName -ErrorAction SilentlyContinue

# Build single-file version of the regex engine

& $scriptBuildSinglePath -outModuleName "StaticRegexSingle"

# Build new file

$excelApp = New-Object -ComObject Excel.Application

try {
    $wb = $excelApp.Workbooks.Add()

    foreach ($fullSourceFileName in $fullSourceFileNames) {
        $vbProject = $excelApp.ActiveWorkbook.VBProject
        if ($vbProject -eq $null) {
            throw "You must enable 'Trust access to the VBA project object model' in the Excel Trust Center in order to be able to run this script"
        }
        $vbProject.VBComponents.Import($fullSourceFileName)
    }

    # file type: xlam = 55, xlsm = 52
    # see https://learn.microsoft.com/en-us/office/vba/api/excel.xlfileformat
    $excelApp.ActiveWorkbook.SaveAs($fullOutFileName, $( if ($createExampleFile) { 52 } else { 55 } ))
} finally {
    $excelApp.Quit()
}