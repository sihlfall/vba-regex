param(
    [ValidateSet('dfs', 'nfa', 'dfa')]
    [string[]]$Engines = @('dfs', 'nfa', 'dfa'),
    [string]$TestCasesRoot = (Join-Path $PSScriptRoot '../../regex-test-cases'),
    [string[]]$TestCases = @(
        '000_EgrepLiterals_Lowercase', '001_EgrepLiterals_Lowercase',
        '002_EgrepLiterals_Lowercase', '003_EgrepLiterals_Lowercase',
        '004_EgrepLiterals_MixedCase', '005_EgrepLiterals_MixedCase',
        '006_EgrepLiterals_MixedCase', '007_EgrepLiterals_MixedCase',
        '008_EgrepLiterals_FoldCase', '009_EgrepLiterals_FoldCase',
        '010_EgrepLiterals_FoldCase', '011_EgrepLiterals_FoldCase',
        '016_Repetition_Simple', '017_Repetition_Simple',
        '018_Repetition_Capturing', '019_Repetition_Capturing',
        '020_EmptyString_Exhaustive', '021_Punctuation_Literals',
        '023_CharacterClasses_Exhaustive', '024_CharacterClasses_ExhaustiveAB'
    )
)

$ErrorActionPreference = 'Stop'
$runTestsExe = Join-Path $PSScriptRoot 'build/RegexTest2Project.exe'
if (!(Test-Path -LiteralPath $runTestsExe)) { throw 'Build test2 with make.ps1 first.' }
$timestamp = Get-Date -Format FileDateTimeUniversal
$testOutputDir = New-Item -ItemType Directory -Force -Path (Join-Path $PSScriptRoot "test-results/$timestamp")
@{ Engines = $Engines; TestCases = $TestCases } | ConvertTo-Json |
    Set-Content -Encoding UTF8 -Path (Join-Path $testOutputDir 'run.json')
$inputDir = New-Item -ItemType Directory -Path (Join-Path $testOutputDir 'inputs')
foreach ($engine in $Engines) {
    [void](New-Item -ItemType Directory -Path (Join-Path $testOutputDir $engine))
}

foreach ($testCase in $TestCases) {
    $filename = "$testCase.json"
    $source = Join-Path $TestCasesRoot "test-cases/$filename"
    $inputPath = Join-Path $inputDir $filename
    $json = Get-Content -Raw -LiteralPath $source | ConvertFrom-Json
    # The existing subset is ASCII; RE2's byte wildcard corresponds to our dot.
    $json.regexs = @($json.regexs | ForEach-Object { $_.Replace('\C', '.') })
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($inputPath, ($json | ConvertTo-Json -Depth 10), $utf8)

    foreach ($engine in $Engines) {
        Write-Host "$engine/$testCase"
        $outputPath = Join-Path $testOutputDir "$engine/$filename"
        $process = Start-Process -FilePath $runTestsExe -ArgumentList @(
            ('"{0}"' -f $inputPath), ('"{0}"' -f $outputPath), $engine
        ) -Wait -PassThru
        if ($process.ExitCode -ne 0 -or !(Test-Path -LiteralPath $outputPath)) {
            throw "Test execution failed; see $outputPath.error.txt"
        }
    }
}
Write-Output $testOutputDir.FullName
