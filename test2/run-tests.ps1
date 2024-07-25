$testCasesToExecute = @(
    "007_EgrepLiterals_MixedCase",
    "008_EgrepLiterals_FoldCase",
    "009_EgrepLiterals_FoldCase",
    "010_EgrepLiterals_FoldCase",
    "011_EgrepLiterals_FoldCase",
    "016_Repetition_Simple",
    "017_Repetition_Simple",
    "018_Repetition_Capturing",
    "019_Repetition_Capturing",
    "023_CharacterClasses_Exhaustive",
    "024_CharacterClasses_ExhaustiveAB"
)

$testCaseDirName = "..\..\regex-test-cases\test-cases"
$testResultDirName = ".\test-results"
$tempDirName = ".\temp"

$runTestsExe = "build\RegexTest2Project.exe"

$timestamp = Get-Date -Format FileDateTimeUniversal

$testCaseDir = Get-Item $testCaseDirName
$testResultDir = if (Test-Path $testResultDirName) {
    Get-Item -Path $testResultDirName
} else {
    New-Item -Path $testResultDirName -ItemType "directory"
}
$tempDir = if (Test-Path $tempDirName) {
    Get-Item -Path $tempDirName
} else {
    New-Item -Path $tempDirName -ItemType "directory"
}



$testOutputDir = New-Item -Path $testResultDir -Name $timestamp -ItemType "directory"

# $testCases = Get-ChildItem $testCaseDir


$testCasesToExecute | %{

    $testCaseFileName = "$_.json"
    $testCaseFullName = Join-Path $testCaseDirName -ChildPath $testCaseFileName
    $tempFileFullName = Join-Path (Get-Item $tempDirName).FullName -ChildPath $testCaseFileName

    Write-Host `Processing $testCaseFullName ...`
    $json = Get-Content -Path $testCaseFullName | ConvertFrom-Json
    $json.regexs = $json.regexs | %{ $_.replace("\C", ".") }

    Write-Host `Writing to $tempFileFullName ...`

    $UTF8Only = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines($tempFileFullName, @($json | ConvertTo-Json), $UTF8Only)

    Start-Process -FilePath $runTestsExe -ArgumentList $tempFileFullName,(Join-Path $testOutputDir.FullName -ChildPath $testCaseFileName) -Wait
}
