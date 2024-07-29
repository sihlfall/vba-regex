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

    Get-Content -Path $inFilePath |
        % { $_ -replace "__SCRIPTNAME__", $scriptName }
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
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)][AllowEmptyString()]
        [string]$codeLine
    )

    begin {
        $beforeWithinOrAfterHeader = 0 # before header
    }

    process {
        if ($beforeWithinOrAfterHeader -eq 1) {
            if ($codeLine -match "^END") { $beforeWithinOrAfterHeader = 2 } # after header
        } elseif (($codeLine -match "^VERSION") -and ($beforeWithinOrAfterHeader -eq 0)) {
            $beforeWithinOrAfterHeader = 1 # inside header
        } elseif ($codeLine -match "^option") {
        } elseif ($codeLine -match "^attribute") {
        } else {
            $codeLine
        }
    }
}

function RemoveTopLevelComments {
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)][AllowEmptyString()]
        [string]$codeLine
    )
    process {
        if ($codeLine -match "^'") { } else { $codeLine }
    }
}

function MakeEverythingPrivate {
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)][AllowEmptyString()]
        [string]$codeLine
    )
    process {
        [regex]::Replace($codeLine, "^Public", "Private")
    }
}

function ReadFileAndPerformCommonTasks {
    param (
        [string]$inFilePath
    )
    $lines = Get-Content -Path $inFilePath
#    $lines = MakeEverythingPrivate($lines)
    $parts = SplitBeforeFirstSub $lines
    return @(
        ($parts[0] | RemoveHeaderAndOptionAttributeLines),
        $parts[1]
    )
}

function RemoveAllGlobalVariableDeclarations {
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)][AllowEmptyString()]
        [string]$codeLine
    )
    process {
        if ($codeLine -match "^(Private|Public|Dim)\s+(?!(Const|Type|Sub|Function|Enum)\b)") { return }
        $codeLine
    }
}
        
function CollapseSubsequentBlankLines {
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)][AllowEmptyString()]
        [string]$codeLine
    )

    begin {
        $lastWasBlank = $false
    }

    process {
        if ($codeLine -match "^\s*$") {
            if (-not $lastWasBlank) {
                $lastWasBlank = $true
                ""
            }
        } else {
            $lastWasBlank = $false
            $codeLine
        }
    }
}

function InsertBlankLineAfterEndSubAndEndFunction {
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)][AllowEmptyString()]
        [string]$codeLine
    )
    
    process {
        $codeLine
        if ($codeLine -match '^\s*End (Sub|Function)') { "" }
    }
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

function WithProcedure {
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)][AllowEmptyString()]
        [string]$codeLine,
        [Parameter(Mandatory=$true,Position=0)]
        [string[]]$procedureNames,
        [Parameter(Mandatory=$true,Position=1)]
        [ScriptBlock]$scriptBlock
    )
        
    begin {
        $weAreWithinProc = $false
        $procLines = @()
        $procedureNamesDict = @{}
        foreach ($procedureName in $procedureNames) { $procedureNamesDict[$procedureName] = 1 }
    }

    process {
        if (-not $weAreWithinProc) {
            [void] ($codeLine -match "^(?:Public|Private|Friend\s)\s*(?:Sub|Function)\s+([A-Za-z0-9_]+)\b")
            $currentProcedureName = $matches[1]
            if ($procedureNamesDict[$currentProcedureName]) {
                $weAreWithinProc = $true
                $procLines = @( $codeLine )
            } else {
                $codeLine
            }
        } else {
            $procLines += $codeLine
            if ($codeLine -match "^End\s+(Sub|Function)\b") {
                $scriptBlock.InvokeWithContext($null, [psvariable]::new('_', $procLines))
                $procLines = @()
                $weAreWithinProc = $false
            }
        }
    }
}


function CreateStandaloneStdRegex3 {
    if (-not $outFileName.EndsWith(".cls")) { $outFileName += ".cls" }


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
            [Parameter(Mandatory=$true,ValueFromPipeline=$true)][AllowEmptyString()]
            [string]$codeLine,
            [switch]$prefixed = $false
        )

        begin {
            $prefixToRemove = ""; if ($prefixed) { $prefixToRemove = 'StaticRegex\.' }
            $prefixToAdd = ""; if ($prefixed) { $prefixToAdd = 'stdRegex3.' }
        }

        process {
            foreach ($procName in $publicStaticRegexProcNames) {
                $codeline = $codeLine -replace "\b$prefixToRemove$procName\b", ($prefixToAdd + $publicStaticRegexRenameMap[$procName])
            }
            $codeLIne
        }
    }

    ProcessFile "./build/StaticRegexSingle.bas" {
        $varPart = $_[0]; $procPart = $_[1]

        $varPart = $varPart |
            RemoveTopLevelComments |
            RemoveAllGlobalVariableDeclarations |
            MakeEverythingPrivate

        $procPart = $procPart | RemoveTopLevelComments

        # gather proc names to be renamed
        Set-Variable -scope 2 -Name "publicStaticRegexProcNames" -Value (GetAllPublicProcedureNames $procPart)

        # fill rename map
        $publicStaticRegexProcNames | ForEach-Object { $publicStaticRegexRenameMap[$_] = "protStaticRegex" + $_ }


        foreach ($procName in $publicStaticRegexProcNames) {
            $procPart = MakeProcFriend $procPart $procName
            $procPart = $procPart | WithProcedure $procName {
                $lines = $_

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

        $procPart = $procPart |

            # remove procedures ... 
            WithProcedure (
                # ... which we do not *want* to see in the class module
                "UnicodeInitialize",
                "RangeTablesInitialize",
                "AstTableInitialize",

                # ... or which are currently not being used (but could be in the future)
                "InitializeMatcherState",
                "ResetMatcherState",
                "GetCapture",
                "GetCaptureByName",
                "TryInitializeRegex"
            ) { } |

            # remove instructions initializing StaticData, as we will be doing this in Class_Initialize
            WithProcedure "AstToBytecode" {
                $_ | ForEach-Object { if($_ -match "If Not astTableInitialized Then") { } else { $_ } }
            } |
            WithProcedure "Compile" {
                $_ | ForEach-Object {
                    if($_ -match "If Not (UnicodeInitialized|RangeTablesInitialized) Then") { } else { $_ }
                }
            } |

            # static data will be stored in the regex.bytecode field of the class instance
            WithProcedure (
                "PrepareStackAndBytecodeBuffer",
                "ReCanonicalizeChar",
                "RegexpGenerateRanges",
                "ParseReRanges",
                "Parse"
            ) {
                $_ | ForEach-Object { $_ -replace "StaticData", "This.regex.bytecode" }
            } |

            # rename procedures of the static API, as they will conflict with the method names of our object;
            # since these procedures should not be called from lower-level functions, restrict the renaming
            # to the API procedures, to reduce artifacts
            WithProcedure $publicStaticRegexProcNames {
                $_ | StubbornlyRenameStaticRegexProcs
            }

        $_[0] = $varPart; $_[1] = $procPart
    }

    $single[1] += @(
        "",
        "",
        "Private Sub Class_Initialize()",
        "    If Me Is stdRegex3 Then",
        "        ' We use the otherwise unused .regex.bytecode field of the class instance to store the static data.",
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

        $varPart = $varPart |
            RemoveTopLevelComments |
            ForEach-Object { $_ -replace "\bStaticRegex\.", "" }
        $procPart = $procPart | StubbornlyRenameStaticRegexProcs -prefixed |
            ForEach-Object { $_ -replace "\b(StaticRegex|Regex(Bytecode|Errors|DfsMatcher))\.", "" }

        $_[0] = $varPart; $_[1] = $procPart
    }

    $commentSnippetPath = Join-Path -Path $PSScriptRoot -ChildPath ".\HeaderCommentStdRegex3Standalone.bas"
    $headerComment = ReadAndProcessHeaderComment -inFilePath $commentSnippetPath -scriptName (Split-Path -Path $MyInvocation.PSCommandPath -Leaf)

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
        )
    $complete += $headerComment
    $complete += @(
            "Option Explicit",
            "",
            ""
        )
    $complete += $single[0]
    $complete += $single[1]

    $complete = $complete |
        InsertBlankLineAfterEndSubAndEndFunction |
        CollapseSubsequentBlankLines

    Set-Content -Path (Join-Path -Path $outDirPath -ChildPath $outFileName) -Value $complete > $null
}

CreateStandaloneStdRegex3