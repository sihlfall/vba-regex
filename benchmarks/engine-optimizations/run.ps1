param(
    [string]$Candidate = 'HEAD',
    [string]$Baseline = '1619cd75526903045b8c5cb8c6048f93bc4160be',
    [string]$StdVbaRoot = (Join-Path $PSScriptRoot '..\..\..\stdVBA'),
    [string]$OutputDir = (Join-Path $PSScriptRoot 'results\latest'),
    [int]$Samples = 7,
    [double]$TargetSeconds = 0.025,
    [switch]$TestsOnly
)
$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$tempRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
$workDir = Join-Path $tempRoot ('regex-engine-bench-' + [guid]::NewGuid().ToString('N'))
& python (Join-Path $PSScriptRoot 'prepare.py') --stdvba-root $StdVbaRoot --baseline $Baseline --candidate $Candidate --output $workDir
if ($LASTEXITCODE -ne 0) { throw 'Preparing benchmark classes failed.' }
$office = $null
$document = $null
$rows = [System.Collections.Generic.List[object]]::new()
$checks = [System.Collections.Generic.List[object]]::new()
$nativeFailures = @()
try {
    $office = New-Object -ComObject Word.Application
    $office.Visible = $false
    $office.DisplayAlerts = 0
    $document = $office.Documents.Add()
    $project = $document.GetType().InvokeMember('VBProject', [System.Reflection.BindingFlags]::GetProperty, $null, $document, $null)
    foreach ($path in (Get-ChildItem (Join-Path $workDir 'native-src') -Filter '*.bas').FullName) { $null = $project.VBComponents.Import($path) }
    foreach ($path in (Get-ChildItem $workDir -File | Where-Object Extension -In '.bas','.cls').FullName) { $null = $project.VBComponents.Import($path) }
    $modules = Get-Content (Join-Path $workDir 'test-modules.json') -Raw | ConvertFrom-Json
    foreach ($module in $modules) {
        Write-Output ('Checking ' + $module)
        $result = [string]$office.Run($module + '.RunPortableTests')
        $parts = $result -split "`n", 2
        $checks.Add([pscustomobject]@{ suite=$module; result=$result })
        if ($parts.Count -gt 1 -and $parts[1].Trim().Length -gt 0) { $nativeFailures += ($module + ': ' + $parts[1].Trim()) }
    }
    if ($nativeFailures.Count -gt 0) { throw ($nativeFailures -join "`n") }
    $result = $office.Run('StdRegexTestRunner.RunTests')
    if ($result -is [string]) { throw $result }
    $checks.Add([pscustomobject]@{suite='stdVBA'; result=$result})
    $edge = $office.Run('EngineBench.EdgeChecks')
    if ($edge -is [string]) { throw $edge }
    $checks.Add([pscustomobject]@{suite='differential'; result=$edge})
    Write-Output ('Passed stdVBA ' + $result + ', differential ' + $edge + '; native suites recorded.')
    if (-not $TestsOnly) {
        for ($id = 0; $id -lt 14; $id++) {
            $values = $office.Run('EngineBench.RunCase', [ref]$id, [ref]$Samples, [ref]$TargetSeconds)
            if ($values -is [string]) { throw $values }
            for ($i = 1; $i -le $values.GetLength(0); $i++) {
                $rows.Add([pscustomobject]@{
                    case=($values.GetValue($i, 1)); case_id=[int]($values.GetValue($i, 2)); operation=($values.GetValue($i, 3)); variant=($values.GetValue($i, 4))
                    sample=[int]($values.GetValue($i, 5)); iterations=[int]($values.GetValue($i, 6)); ms=[double]($values.GetValue($i, 7)); characters=[int]($values.GetValue($i, 8))
                })
            }
            Write-Output ('Measured ' + $Candidate + ': ' + $values[1,1])
        }
    }
    $null = New-Item -ItemType Directory -Path $OutputDir -Force
    $metadata = Get-Content (Join-Path $workDir 'metadata.json') -Raw | ConvertFrom-Json
    $metadata | Add-Member measured_at_utc ([DateTime]::UtcNow.ToString('o'))
    $metadata | Add-Member office_version ([string]$office.Version)
    $metadata | Add-Member samples $Samples
    $metadata | Add-Member wrapper_sha256 (Get-FileHash (Join-Path $StdVbaRoot 'Tools/stdRegex-build/stdRegex.wrapper.cls.in')).Hash
    $metadata | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $OutputDir 'environment.json') -Encoding utf8
    $checks | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $OutputDir 'tests.json') -Encoding utf8
    if (-not $TestsOnly) {
        $rows | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $OutputDir 'samples.json') -Encoding utf8
        $rows | Export-Csv (Join-Path $OutputDir 'samples.csv') -NoTypeInformation -Encoding utf8
        & python (Join-Path $PSScriptRoot 'summarize.py') $OutputDir
        if ($LASTEXITCODE -ne 0) { throw 'Summarizing results failed.' }
    }
} finally {
    try { if ($null -ne $document) { $document.Close(0) } }
    finally {
        try { if ($null -ne $office) { $office.Quit() } }
        finally {
            $resolved = [System.IO.Path]::GetFullPath($workDir)
            if (-not $resolved.StartsWith($tempRoot.TrimEnd('\','/') + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase) -or
                -not ([System.IO.Path]::GetFileName($resolved)).StartsWith('regex-engine-bench-')) { throw 'Unsafe scratch path.' }
            Remove-Item -LiteralPath $resolved -Recurse -Force
        }
    }
}
