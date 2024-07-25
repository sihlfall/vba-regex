param(
    [string]$testRun = ""
)

function local:runScript {
    $actualResultsRootdir = ".\test-results"

    if ($testRun -eq "") {
        $testRun = (Get-ChildItem $actualResultsRootdir) | Sort-Object -Property Name | Select-Object -Last 1
    }

    $actualResultsDir = (Join-Path $actualResultsRootdir -ChildPath $testRun)

    Write-Host "Analysing files in " $actualResultsDir

    $testCasesDir = "..\..\regex-test-cases"
    $tempDir = ".\temp"

    function compareArrays {
        param (
            $expected,
            $actual
        )

        $minAryLength = [System.Math]::Min($expected.Length, $actual.Length)

        $i = 0
        for(; $i -lt $minAryLength; ++$i) {
            if ($expected[$i] -ne $actual[$i]) { return $false }
        }
        for(; $i -lt $expected.Length; ++$i) {
            if ($expected[$i] -ne -1) { return $false }
        }
        for(; $i -lt $actual.Length; ++$i) {
            if ($actual[$i] -ne -1) { return $false }
        }
        return $true
    }

    Get-ChildItem $actualResultsDir | %{

        $testCaseFileNameBase = $_.BaseName

        Write-Host "Test case: " $testCaseFileNameBase

        $testCaseFilePath = Join-Path (Join-Path $testCasesDir -ChildPath "test-cases") -ChildPath "$testCaseFileNameBase.json"
        $expectedResultsFilePath = Join-Path (Join-Path $testCasesDir -ChildPath "test-results\re2") -ChildPath "$testCaseFileNameBase.json"
        $actualResultsFilePath = join-path $actualResultsDir -ChildPath "$testCaseFileNameBase.json"
        $tempFilePath = Join-Path $tempDir -ChildPath "$testCaseFileNameBase.json"

        $testCaseData = Get-Content -Raw $testCaseFilePath | ConvertFrom-Json
        $tempData = Get-Content -Raw $tempFilePath | ConvertFrom-Json
        $expectedResultsAllData = Get-Content -Raw $expectedResultsFilePath | ConvertFrom-Json
        $actualResultsData = Get-Content -Raw $actualResultsFilePath | ConvertFrom-Json

#        write-host (ConvertTo-Json $expectedResultsAllData[10][10])

#        exit

        for ($regexIndex = 0; $regexIndex -lt $testCaseData.regexs.Length; ++$regexIndex) {
            $regexExpected = $testCaseData.regexs[$regexIndex]
            $regexActual = $tempData.regexs[$regexIndex]

            for ($strIndex = 0; $strIndex -lt $testCaseData.strs.Length; ++$strIndex) {
                $str = $testCaseData.strs[$strIndex]

                $expected = $expectedResultsAllData[$regexIndex][$strIndex][1]
                $actual = $actualResultsData[$regexIndex][$strIndex]
                if(-not (compareArrays $expected $actual)) {
                    Write-Host $str $regexExpected "Expected: " (ConvertTo-Json -Compress $expected) $regexActual "Actual: " (ConvertTo-Json -Compress $actual)
                }
            }
        }

    }
}

runScript