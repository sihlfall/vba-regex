param(
    [string]$outFileName = "stdRegex3Standalone.cls",
    [string]$outDirPath = ".\build"
)


$SOURCE_FILES = ".\build\StaticRegexSingle.bas", ".\stdRegex3.cls"

$scriptBuildSinglePath = "..\aio\make_aio.ps1"

$scriptPath = Split-Path $MyInvocation.MyCommand.Path -Parent
$projectRootPath = $scriptPath

$fullSourceFileNames = $SOURCE_FILES | % { Join-Path $projectRootPath $_ }
$fullOutFileName = Join-Path -Path $scriptPath -ChildPath $outDirPath | Join-Path -ChildPath $outFileName


# Create build directory, if necessary
if (-not (Test-Path -Path $outDirPath)) {
    New-Item -Path $outDirPath -ItemType Directory > $null
}


# Build single-file version of the regex engine

& $scriptBuildSinglePath -outModuleName "StaticRegexSingle"

function ReadAndProcessHeaderComment {
    param (
        [string]$inFilePath,
        [string]$scriptName
    )

    $lines = (Get-Content -Path $inFilePath)
    $lines = $lines | ForEach-Object { $_ -replace "__SCRIPTNAME__", $scriptName }
    return $lines
}

function SplitBeforeFirstSub {
    param (
        [string[]]$lines
    )

    $pattern = "^((private|public|friend)\s+)?(sub|function|property)"

    $i = 0
    while ($i -lt $lines.length) {
        $line = $lines[$i]
        if ($line -match $pattern) { break }
        ++$i
    }
    $firstLineOfSecondPart = $i

    $i = $firstLineOfSecondPart - 1
    while ($i -ge 0) {
        $line = $lines[$i]
        if (-not ($line -match "^'")) { break }
        --$i
    }
    $firstLineOfSecondPart = $i + 1

    $part1 = $lines[0..($firstLineOfSecondPart - 1)]
    $part2 = $lines[$firstLineOfSecondPart..($lines.length - 1)]

    return @($part1, $part2)
}

function RemoveHeaderAndOptionAttributeLines {
    param (
        [string[]]$lines
    )

    $beforeWithinOrAfterHeader = 0 # before header
    $res = @()
    foreach ($line in $lines) {
        if ($beforeWithinOrAfterHeader -eq 1) {
            if ($line -match "^END") { $beforeWithinOrAfterHeader = 2 } # after header
            continue
        }
        if (($line -match "^VERSION") -and ($beforeWithinOrAfterHeader -eq 0)) {
            $beforeWithinOrAfterHeader = 1 # inside header
            continue
        }
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

function RemoveTopLevelComments {
    param (
        [string[]] $lines
    )

    return $lines | Where-Object { -not ($_ -match "^'") }
}

function MakeEverythingPrivate {
    param (
        [string[]] $lines
    )

    return $lines | ForEach-Object { [regex]::Replace($_, "^Public", "Private") }
}

function ReadFileAndPerformCommonTasks {
    param (
        [string]$inFilePath
    )
    $lines = Get-Content -Path $inFilePath
#    $lines = MakeEverythingPrivate($lines)
    $parts = SplitBeforeFirstSub $lines
    return @(
        (RemoveHeaderAndOptionAttributeLines $parts[0]),
        $parts[1]
    )
}

function RemoveAllGlobalVariableDeclarations {
    param (
        [string[]] $lines
    )

    $res = @()
    foreach ($line in $lines) {
        if ($line -match "^(Private|Public|Dim)\s+(?!(Const|Type|Sub|Function|Enum)\b)") {
            continue
        }
        $res += $line
    }

    return $res
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
    
    $outputText = $inputText -replace '(?m)^\s*End (Sub|Function)\s*', "`$0`n"
    
    # Split the resulting text back into an array of lines
    $outputLines = $outputText -split "`n"
    
    return $outputLines
}

function GetAllPublicProcedureNames {
    param (
        [string[]]$lines
    )

    $res = @()
    foreach ($line in $lines) {
        if ($line -match "Public\s+(?:Sub|Function)\s+(\w+)\b") {
            $res += $matches[1]
        }
    }
    return $res
}

function FindEndOfProcDeclarationLine {
    param (
        [string[]]$lines
    )

    $idx = 0
    foreach ($line in $lines) {
        if ($line -match "^(\)|\S.*\))(\sAs\s\w+)?\s*$") { return $idx }
        ++$idx
    }
    $idx
}

function CreateStandaloneStdRegex3 {
    if (-not $outFileName.EndsWith(".cls")) { $outFileName += ".cls" }

#    $inDirPath = Join-Path -Path $PSScriptRoot -ChildPath "..\src"


    $single = @( @(), @() )

    function ProcessFile {
        param(
            [string]$inFileName,
            [ScriptBlock]$scriptBlock
        )
        $inFilePath = Join-Path -Path $scriptPath -ChildPath $inFileName

        $parts = ReadFileAndPerformCommonTasks($inFilePath)

        $scriptBlock.InvokeWithContext($null, [psvariable]::new('_', $parts))

        $single = Get-Variable -Scope 1 -Name "single" -ValueOnly
        $single[0] += $parts[0]
        $single[1] += $parts[1]
    }

    function InProcedure {
        param(
            [string]$procedureName,
            [ScriptBlock]$scriptBlock
        )
        
        $codeLines = @($input)

        $start = -1
        $end = -1
        $i = 0
        for (; $i -lt $codeLines.Length; ++$i) {
            $line = $codeLines[$i]
            if ($line -match "^(Public|Private|Friend\s)\s*(Sub|Function)\s+$procedureName\b") {
                $start = $i
                break
            }
        }
        if ($start -lt 0) { return $codeLines }
        for (; $i -lt $codeLines.Length; ++$i) {
            $line = $codeLines[$i]
            if ($line -match "^End\s+(Sub|Function)\b") {
                $end = $i
                break
            }
        }
        if ($end -lt 0) { return $codeLines }

        $newCode = $scriptBlock.InvokeWithContext($null, [psvariable]::new('lines', $codeLines[$start..$end]))

        return $codeLines[0..($start - 1)] + $newCode + $codeLines[($end + 1)..($codeLines.Length - 1)]
    }

    function MakeProcFriend {
        param(
            [string[]]$lines,
            [string]$procName
        )
        return $lines | ForEach-Object {
            $_ -replace "^(Private|Public\s)\s*(?=(Sub|Function)\s+$procName\b)", "Friend "
        }
    }

    $publicStaticRegexProcNames = @()
    $publicStaticRegexRenameMap = @{}
    function StubbornlyRenameStaticRegexProcs {
        param(
            [string[]]$lines,
            [bool]$prefixed = $false
        )
        $prefixToRemove = ""; if ($prefixed) { $prefixToRemove = 'StaticRegex\.' }
        $prefixToAdd = ""; if ($prefixed) { $prefixToAdd = 'stdRegex3.' }
        $lines | ForEach-Object {
            foreach ($procName in $publicStaticRegexProcNames) {
                $_ = $_ -replace "\b$prefixToRemove$procName\b", ($prefixToAdd + $publicStaticRegexRenameMap[$procName])
            }
            $_
        }
    }

    ProcessFile "./build/StaticRegexSingle.bas" {
        $varPart = $_[0]; $procPart = $_[1]

        $varPart = RemoveTopLevelComments $varPart
        $varPart = RemoveAllGlobalVariableDeclarations $varPart
        $varPart = MakeEverythingPrivate $varPart

        $procPart = RemoveTopLevelComments $procPart

        # gather proc names to be renamed
        Set-Variable -scope 2 -Name "publicStaticRegexProcNames" -Value (GetAllPublicProcedureNames $procPart)

        # fill rename map
        $publicStaticRegexProcNames | ForEach-Object { $publicStaticRegexRenameMap[$_] = "protStaticRegex" + $_ }


        foreach ($procName in $publicStaticRegexProcNames) {
            $procPart = MakeProcFriend $procPart $procName
            $procPart = $procPart | InProcedure $procName {
                if ($lines.Length -gt 1) {
                    [void]($lines[0] -match "Function|Sub"); $functionOrSub = $matches[0]
                    $endOfDeclarationIdx = FindEndOfProcDeclarationLine($lines) + 1
                    [void]($lines[$endOfDeclarationIdx + 1] -match "^(\s*)")
                    $indent = $matches[0]
                    $lines = $lines[0..$endOfDeclarationIdx] +
                        ( @(
                            "If Not (Me Is stdRegex3) Then",
                            ($indent + 'Error.Raise "Method called on object, not on class"'),
                            ($indent + "Exit $functionOrSub"),
                            "End If",
                            ""                   
                        ) | ForEach-Object { $indent + $_ } ) +
                        $lines[($endOfDeclarationIdx + 1)..($lines.Length - 1)]
                }
                $lines
            }
        }

        # remove procedures which we do not *want* to see in the class module
        $procPart = $procPart |
            InProcedure "UnicodeInitialize" { @() } |
            InProcedure "RangeTablesInitialize" { @() } |
            InProcedure "AstTableInitialize" { @() }

        # remove procedures that are currently not used (but could be in the future)
        $procPart = $procPart |
            InProcedure "InitializeMatcherState" { @() } |
            InProcedure "ResetMatcherState" { @() } |
            InProcedure "GetCapture" { @() } |
            InProcedure "GetCaptureByName" { @() } |
            InProcedure "TryInitializeRegex" { @() }

        $procPart = $procPart | InProcedure "AstToBytecode" {
            $lines | ForEach-Object { if($_ -match "If Not astTableInitialized Then") { "'" + $_ } else { $_ } }
        }
        $procPart = $procPart | InProcedure "Compile" {
            $lines | ForEach-Object {
                if($_ -match "If Not (UnicodeInitialized|RangeTablesInitialized) Then") { "'" + $_ } else { $_ }
            }
        }
        $procPart = $procPart | InProcedure "PrepareStackAndBytecodeBuffer" {
            $lines | ForEach-Object { $_ -replace "StaticData", "This.regex.bytecode" }
        }
        $procPart = $procPart | InProcedure "ReCanonicalizeChar" {
            $lines | ForEach-Object { $_ -replace "StaticData", "This.regex.bytecode" }
        }
        $procPart = $procPart | InProcedure "RegexpGenerateRanges" {
            $lines | ForEach-Object { $_ -replace "StaticData", "This.regex.bytecode" }
        }
        $procPart = $procPart | InProcedure "ParseReRanges" {
            $lines | ForEach-Object { $_ -replace "StaticData", "This.regex.bytecode" }
        }
        $procPart = $procPart | InProcedure "Parse" {
            $lines | ForEach-Object { $_ -replace "StaticData", "This.regex.bytecode" }
        }

        $procPart = StubbornlyRenameStaticRegexProcs $procPart
        $_[0] = $varPart; $_[1] = $procPart
    }

    $single[1] += @(
        "",
        "",
        "Private Sub Class_Initialize()",
        "    If Me Is stdRegex3 Then",
        "        ReDim This.regex.bytecode(0 To STATIC_DATA_LENGTH - 1) As Long",
        "        InitializeUnicodeCanonLookupTable This.regex.bytecode",
        "        InitializeUnicodeCanonRunsTable This.regex.bytecode",
        "        InitializeRangeTables This.regex.bytecode",
        "        InitializeAstTable This.regex.bytecode",
        "    End If",
        "End Sub",
        "",
        ""
    )

    ProcessFile "./stdRegex3.cls" {
        $varPart = $_[0]; $procPart = $_[1]

        $varPart = RemoveTopLevelComments $varPart

        $varPart = $varPart | ForEach-Object {
            $_ -replace "\bStaticRegex\.", ""
        }

        $procPart = StubbornlyRenameStaticRegexProcs $procPart -prefixed $true
        $procPart = $procPart | ForEach-Object {
            $_ -replace "\b(StaticRegex|Regex(Bytecode|Errors|DfsMatcher))\.", ""
        }

        $_[0] = $varPart; $_[1] = $procPart
    }


    $complete = @()
    $complete += @(
        "Option Explicit",
        "",
        ""
    )
    $complete += $single[0]
    $complete += $single[1]

    $complete = InsertBlankLineAfterEndSubAndEndFunction($complete)
    $complete = ReplaceMultipleBlankLines($complete)


    $commentSnippetPath = Join-Path -Path $PSScriptRoot -ChildPath ".\HeaderCommentStdRegex3Standalone.bas"
    $headerComment = ReadAndProcessHeaderComment -inFilePath $commentSnippetPath -scriptName (Split-Path -Path $MyInvocation.PSCommandPath -Leaf)
    $complete = $headerComment + $complete

    $complete = @(
        "VERSION 1.0 CLASS",
        "BEGIN",
        "  MultiUse = -1  'True",
        "  Persistable = 0  'NotPersistable",
        "  DataBindingBehavior = 0  'vbNone",
        "  DataSourceBehavior  = 0  'vbNone",
        "  MTSTransactionMode  = 0  'NotAnMTSObject",
        "END",
        "Attribute VB_Name=""stdRegex3""",
        "Attribute VB_GlobalNameSpace = False",
        "Attribute VB_Creatable = True",
        "Attribute VB_PredeclaredId = True",
        "Attribute VB_Exposed = False"
    ) + $complete

    Set-Content -Path (Join-Path -Path $outDirPath -ChildPath $outFileName) -Value $complete > $null
}

CreateStandaloneStdRegex3