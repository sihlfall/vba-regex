param(
    [string]$outModuleName = "StaticRegexSingle"
)

function SplitBeforeFirstSub {
    param (
        [string[]]$lines
    )

    $part1 = @()
    $part2 = @()

    $pattern = "^((private|public)\s+)?(sub|function)"

    $i = 0
    while ($i -lt $lines.length) {
        $line = $lines[$i]

        if ($line -match $pattern) {
            break
        }

        $part1 += $line
        $i += 1
    }

    while ($i -lt $lines.length) {
        $line = $lines[$i]
        $part2 += $line
        $i += 1
    }

    return @($part1, $part2)
}

function RemoveOptionAttributeLines {
    param (
        [string[]]$lines
    )

    $res = @()
    foreach ($line in $lines) {
        if ($line -match "^option") {
            continue
        }
        if ($line -match "^attribute") {
            continue
        }
        $res += $line
    }

    return $res
}

function MakeEverythingPrivate {
    param (
        [string[]] $lines
    )

    return $lines | ForEach-Object { [regex]::Replace($_, "^Public", "Private") }
}

function RemoveTopLevelComments {
    param (
        [string[]] $lines
    )

    return $lines | Where-Object { -not ($_ -match "^'") }
}

function ReadFileAndPerformCommonTasks {
    param (
        [string]$inFilePath
    )
    $lines = Get-Content -Path $inFilePath
    $lines = MakeEverythingPrivate($lines)
    $lines = RemoveTopLevelComments($lines)
    $parts = SplitBeforeFirstSub $lines
    return @(
        (RemoveOptionAttributeLines $parts[0]),
        $parts[1]
    )
}

function ReadAndProcessHeaderComment {
    param (
        [string]$inFilePath,
        [string]$scriptName
    )

    $lines = (Get-Content -Path $inFilePath)
    $lines = $lines | ForEach-Object { $_ -replace "__SCRIPTNAME__", $scriptName }
    return $lines
}


# This was mainly ChatGPT. Thank you, system!
function ReplaceMultipleBlankLines {
    param (
        [string[]]$inputLines
    )
    
    # Join the lines into a single string with newlines
    $inputText = $inputLines -join "`n"
    
    # Use regex to replace multiple blank lines with a single blank line
    $outputText = $inputText -replace '(\r?\n){3,}', "`n`n"
    
    # Split the resulting text back into an array of lines
    $outputLines = $outputText -split "`n"
    
    return $outputLines
}

function InsertBlankLineAfterEndSubAndEndFunction {
    param (
        [string[]]$inputLines
    )
    
    $inputText = $inputLines -join "`n"
    
    $outputText = $inputText -replace '(?m)\s*End (Sub|Function)\s*', "`$0`n"
    
    # Split the resulting text back into an array of lines
    $outputLines = $outputText -split "`n"
    
    return $outputLines
}


function CreateSingleFile {
    $outDirPath = ".\build"
    $outFileName = $outModuleName + ".bas"

    $inDirPath = Join-Path -Path $PSScriptRoot -ChildPath "..\src"


    if (-not (Test-Path -Path $outDirPath)) {
        New-Item -Path $outDirPath -ItemType Directory > $null
    }

    $single = @( @(), @() )

    function ProcessFile {
        param(
            [string]$inFileName,
            [ScriptBlock]$scriptBlock
        )
        $inFilePath = Join-Path -Path $inDirPath -ChildPath $inFileName

        $parts = ReadFileAndPerformCommonTasks($inFilePath)

        $scriptBlock.InvokeWithContext($null, [psvariable]::new('_', $parts))

        $single = Get-Variable -Scope 1 -Name "single" -ValueOnly
        $single[0] += $parts[0]
        $single[1] += $parts[1]
    }

    ProcessFile "StaticStringBuilder.bas" {
        $_[0] = $_[0] | ForEach-Object { [regex]::Replace($_, "\bTy\b", "StaticStringBuilder") }
        $_[1] = $_[1] | ForEach-Object { [regex]::Replace($_, "\bTy\b", "StaticStringBuilder") }
    }
    ProcessFile "ArrayBuffer.bas" {
        $_[0] = $_[0] | ForEach-Object { [regex]::Replace($_, "\bTy\b", "ArrayBuffer") }
        $_[1] = $_[1] | ForEach-Object { [regex]::Replace($_, "\bTy\b", "ArrayBuffer") }    
    }

    ProcessFile "RegexNumericConstants.bas" { }
    ProcessFile "RegexErrors.bas" { }
    ProcessFile "RegexBytecode.bas" { }
    ProcessFile "RegexAst.bas" { }
    ProcessFile "RegexUnicodeSupport.bas" { }
    ProcessFile "RegexRanges.bas" { }
    ProcessFile "RegexIdentifierSupport.bas" { }

    ProcessFile "RegexLexer.bas" {
        $_[0] = $_[0] | ForEach-Object { [regex]::Replace($_, "(?<!\.)\bTy\b", "LexerContext") }
        $_[1] = $_[1] | ForEach-Object { [regex]::Replace($_, "(?<!\.)\bTy\b", "LexerContext") }
    }
    ProcessFile "RegexCompiler.bas" { }

    ProcessFile "RegexDfsMatcher.bas" {
        $startIdx = -1
        $endIdx = -1
        for ($i = 0; $i -lt $_[0].Count; $i++) {
            if ($_[0][$i] -match "Type StartLengthPair") {
                $startIdx = $i;
                for (; $i -lt $_[0].Count; $i++) {
                    if ($_[0][$i] -match "End Type") {
                        $endIdx = $i
                        break
                    }
                }
                break
            }
        }
        if ( ($startIdx -ne -1) -and ($endIdx -ne -1) ) {
            $_[0] = $_[0][0..($startIdx - 1)] + $_[0][($endIdx + 1)..($_[0].Count - 1)]
        }
    }

    ProcessFile "RegexReplace.bas" { }

    ProcessFile "StaticRegex.bas" { }

    $complete = @()
    $complete += @(
        "Option Private Module",
        "Option Explicit",
        "",
        ""
    )
    $complete += $single[0]
    $complete += @("", "")
    $complete += $single[1]

    $complete = $complete | ForEach-Object {
        $_ = [regex]::Replace($_, "StaticStringBuilder\.Ty", "StaticStringBuilder")
        $_ = [regex]::Replace($_, "StaticStringBuilder\.", "")
        $_ = [regex]::Replace($_, "ArrayBuffer\.Ty", "ArrayBuffer")
        $_ = [regex]::Replace($_, "ArrayBuffer\.", "")
        $_ = [regex]::Replace($_, "RegexNumericConstants\.", "")
        $_ = [regex]::Replace($_, "RegexErrors\.", "")
        $_ = [regex]::Replace($_, "RegexBytecode\.", "")
        $_ = [regex]::Replace($_, "RegexUnicodeSupport\.", "")
        $_ = [regex]::Replace($_, "RegexBytecode\.", "")
        $_ = [regex]::Replace($_, "RegexAst\.", "")
        $_ = [regex]::Replace($_, "RegexRanges\.", "")
        $_ = [regex]::Replace($_, "RegexIdentifierSupport\.", "")
        $_ = [regex]::Replace($_, "RegexLexer\.Ty", "LexerContext")
        $_ = [regex]::Replace($_, "RegexLexer\.", "")
        $_ = [regex]::Replace($_, "RegexCompiler\.", "")
        $_ = [regex]::Replace($_, "RegexDfsMatcher\.", "")
        $_ = [regex]::Replace($_, "RegexReplace\.", "")
        $_ = [regex]::Replace($_, "Private Sub InitializeRegex\(",  "Public Sub InitializeRegex(")
        $_ = [regex]::Replace($_, "Private Function TryInitializeRegex\(",  "Public Function TryInitializeRegex(")
        $_ = [regex]::Replace($_, "Private Function Test\(",  "Public Function Test(")
        $_ = [regex]::Replace($_, "Private Function Match\(", "Public Function Match(")
        $_ = [regex]::Replace($_, "Private Function GetCapture\(", "Public Function GetCapture(")
        $_ = [regex]::Replace($_, "Private Function GetCaptureByName\(", "Public Function GetCaptureByName(")
        $_ = [regex]::Replace($_, "Private Function MatchNext\(", "Public Function MatchNext(")
        $_ = [regex]::Replace($_, "Private Function Replace\(", "Public Function Replace(")
        $_ = [regex]::Replace($_, "Private Function MatchThenJoin\(", "Public Function MatchThenJoin(")
        $_ = [regex]::Replace($_, "Private Sub MatchThenList\(", "Public Sub MatchThenList(")
        $_ = [regex]::Replace($_, "Private Sub InitializeMatcherState\(", "Public Sub InitializeMatcherState(")
        $_ = [regex]::Replace($_, "Private Sub ResetMatcherState\(", "Public Sub ResetMatcherState(")
        $_ = [regex]::Replace($_, "Private Type RegexTy", "Public Type RegexTy")
        $_ = [regex]::Replace($_, "Private Type MatcherStateTy", "Public Type MatcherStateTy")
        $_
    }

    $complete = InsertBlankLineAfterEndSubAndEndFunction($complete)
    $complete = ReplaceMultipleBlankLines($complete)

    $commentSnippetPath = Join-Path -Path $PSScriptRoot -ChildPath ".\HeaderComment.bas"
    $headerComment = ReadAndProcessHeaderComment -inFilePath $commentSnippetPath -scriptName (Split-Path -Path $MyInvocation.PSCommandPath -Leaf)
    $complete = $headerComment + $complete

    $complete = (, ("Attribute VB_Name=""" + $outModuleName + """")) + $complete

    Set-Content -Path (Join-Path -Path $outDirPath -ChildPath $outFileName) -Value $complete > $null
}

CreateSingleFile