param(
    [string]$testRun = '',
    [string]$TestCasesRoot = (Join-Path $PSScriptRoot '../../regex-test-cases')
)

$ErrorActionPreference = 'Stop'
if (!$testRun) {
    $testRun = (Get-ChildItem (Join-Path $PSScriptRoot 'test-results') -Directory |
        Sort-Object Name | Select-Object -Last 1).FullName
} elseif (!(Test-Path -LiteralPath $testRun)) {
    $testRun = Join-Path $PSScriptRoot "test-results/$testRun"
}

function Compare-Offsets($expected, $actual) {
    $count = [Math]::Max($expected.Count, $actual.Count)
    for ($i = 0; $i -lt $count; ++$i) {
        $e = if ($i -lt $expected.Count) { $expected[$i] } else { -1 }
        $a = if ($i -lt $actual.Count) { $actual[$i] } else { -1 }
        if ($e -ne $a) { return $false }
    }
    return $true
}

$summary = @()
$totalFailed = 0
Write-Host "Analysing files in $testRun"
$run = Get-Content -Raw (Join-Path $testRun 'run.json') | ConvertFrom-Json
$totalSuites = @($run.Engines).Count * @($run.TestCases).Count
$suiteIndex = 0
foreach ($engine in $run.Engines) {
    $engineDir = Join-Path $testRun $engine
    $files = @($run.TestCases | ForEach-Object { Get-Item -LiteralPath (Join-Path $engineDir "$_.json") })
    if (!$files.Count) { throw "No test results for $engine" }
    foreach ($file in $files) {
        ++$suiteIndex
        $suiteLabel = "$engine/$($file.BaseName)"
        $suiteTimer = [System.Diagnostics.Stopwatch]::StartNew()
        Write-Host "[$suiteIndex/$totalSuites] ${suiteLabel}: loading test cases"
        $cases = Get-Content -Raw (Join-Path $TestCasesRoot "test-cases/$($file.Name)") | ConvertFrom-Json
        Write-Host "${suiteLabel}: loading RE2 reference results"
        $expected = Get-Content -Raw (Join-Path $TestCasesRoot "test-results/re2/$($file.Name)") | ConvertFrom-Json
        Write-Host "${suiteLabel}: loading matcher results"
        $actual = Get-Content -Raw -LiteralPath $file.FullName | ConvertFrom-Json
        if ($actual.Count -ne $cases.regexs.Count -or $expected.Count -ne $cases.regexs.Count) {
            throw "Wrong pattern count: $($file.FullName)"
        }
        $passed = 0; $failed = 0; $skippedPatterns = 0
        $caseCount = $cases.regexs.Count * $cases.strs.Count
        Write-Host "${suiteLabel}: comparing $caseCount cases ($($cases.regexs.Count) patterns x $($cases.strs.Count) strings)"
        $nextProgressMs = $suiteTimer.ElapsedMilliseconds + 5000
        for ($r = 0; $r -lt $cases.regexs.Count; ++$r) {
            # Keep progress visible in consoles and captured logs without adding
            # output or a clock check to every pattern/string comparison.
            if ($suiteTimer.ElapsedMilliseconds -ge $nextProgressMs) {
                Write-Host "${suiteLabel}: patterns $r/$($cases.regexs.Count); passed=$passed failed=$failed skippedPatterns=$skippedPatterns; $([Math]::Round($suiteTimer.Elapsed.TotalSeconds, 1))s elapsed"
                $nextProgressMs = $suiteTimer.ElapsedMilliseconds + 5000
            }
            if ($null -eq $actual[$r]) {
                if ($engine -eq 'dfs') { throw "Unexpected DFS skip at pattern $r" }
                ++$skippedPatterns
                continue
            }
            if ($actual[$r].Count -ne $cases.strs.Count) { throw "Wrong string count at pattern $r" }
            for ($s = 0; $s -lt $cases.strs.Count; ++$s) {
                $e = @($expected[$r][$s][1]) # RE2, unanchored, leftmost-first.
                $a = @($actual[$r][$s])
                if ($engine -eq 'dfa') {
                    if ($a.Count -ne 2) { throw 'DFA must return exactly one span.' }
                    if ($e.Count -gt 2) { $e = $e[0..1] }
                }
                if (Compare-Offsets $e $a) {
                    ++$passed
                } else {
                    ++$failed
                    if ($failed -le 10) {
                        Write-Host "$engine/$($file.BaseName) regex[$r]=$($cases.regexs[$r]) string[$s]=$($cases.strs[$s]) expected=[$e] actual=[$a]"
                    }
                }
            }
        }
        $totalFailed += $failed
        $summary += [pscustomobject]@{
            Engine = $engine; Suite = $file.BaseName; Passed = $passed; Failed = $failed
            SkippedPatterns = $skippedPatterns; SkippedCases = $skippedPatterns * $cases.strs.Count
        }
        Write-Host "${suiteLabel}: complete; passed=$passed failed=$failed skippedPatterns=$skippedPatterns; $([Math]::Round($suiteTimer.Elapsed.TotalSeconds, 1))s elapsed"
    }
}
if (!$summary.Count) { throw "No engine results in $testRun" }
$summary | Export-Csv -NoTypeInformation -Path (Join-Path $testRun 'summary.csv')
$summary | Format-Table -AutoSize
if ($totalFailed) { throw "$totalFailed matcher results differ from RE2." }
