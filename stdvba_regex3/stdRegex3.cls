VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stdRegex3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const REGEX3_DEFAULT_STEPS_LIMIT As Long = 10000

Private Type TThis
    isGlobal As Boolean
    isMultiline As Boolean
    regex As StaticRegex.RegexTy
End Type

Private This As TThis

Friend Sub protInit(ByRef pattern As String, Optional ByRef flags As String, Optional stepsLimit As Long = REGEX3_DEFAULT_STEPS_LIMIT)
    Dim caseInsensitive As Boolean
    
    With This
        DecodeFlags caseInsensitive, .isGlobal, .isMultiline, flags
        
        StaticRegex.InitializeRegex .regex, pattern, caseInsensitive
        
        .regex.stepsLimit = stepsLimit
    End With
End Sub

Private Sub DecodeFlags(ByRef isCaseInsensitive As Boolean, ByRef isGlobal As Boolean, ByRef isMultiline As Boolean, ByRef flagsString As String)
    Dim i As Long
    Const ASC_UC_G As Long = 71
    Const ASC_UC_I As Long = 73
    Const ASC_UC_M As Long = 77
    Const ASC_LC_G As Long = 103
    Const ASC_LC_I As Long = 105
    Const ASC_LC_M As Long = 109
    
    isGlobal = False
    isCaseInsensitive = False
    isMultiline = False
    For i = 1 To Len(flagsString)
        Select Case AscW(Mid$(flagsString, i, 1))
        Case ASC_UC_G, ASC_LC_G: isGlobal = True
        Case ASC_UC_I, ASC_LC_I: isCaseInsensitive = True
        Case ASC_UC_M, ASC_LC_M: isMultiline = True
        Case Else: Err.Raise RegexErrors.REGEX_ERR_INVALID_FLAG_ERR, "Invalid flag: " & Mid$(flagsString, i, 1)
        End Select
    Next
End Sub

Private Function MakeCapturesDict(ByRef captures As RegexDfsMatcher.CapturesTy, ByRef sHaystack As String) As Object
    Dim j As Long, identifier As String, referenced As Long
    Dim identifierMapEnd As Long, currentEntry As Long
    Dim oRet As Object
    
    Set oRet = CreateObject("Scripting.Dictionary")
    
    With This
        oRet("$COUNT") = captures.nNumberedCaptures
    
        With captures.entireMatch
            If .start > 0 Then oRet(0) = Mid$(sHaystack, .start, .length) Else oRet(0) = vbNullString
        End With
        For j = 1 To captures.nNumberedCaptures
            With captures.numberedCaptures(j - 1)
                If .start > 0 Then oRet(j) = Mid$(sHaystack, .start, .length) Else oRet(j) = vbNullString
            End With
        Next
        
        identifierMapEnd = RegexBytecode.BYTECODE_IDENTIFIER_MAP_BEGIN + _
            RegexBytecode.BYTECODE_IDENTIFIER_MAP_ENTRY_SIZE * .regex.bytecode(RegexBytecode.BYTECODE_IDX_N_IDENTIFIERS)
        currentEntry = RegexBytecode.BYTECODE_IDENTIFIER_MAP_BEGIN
        Do While currentEntry < identifierMapEnd
            identifier = Mid$( _
                .regex.pattern, _
                .regex.bytecode(currentEntry + RegexBytecode.BYTECODE_IDENTIFIER_MAP_ENTRY_START_IN_PATTERN), _
                .regex.bytecode(currentEntry + RegexBytecode.BYTECODE_IDENTIFIER_MAP_ENTRY_LENGTH_IN_PATTERN) _
            )
            referenced = captures.namedCaptures( _
                .regex.bytecode(currentEntry + RegexBytecode.BYTECODE_IDENTIFIER_MAP_ENTRY_ID) _
            )
            If referenced <> -1 Then
                oRet(identifier) = oRet(referenced)
            Else
                oRet(identifier) = vbNullString
            End If
            currentEntry = currentEntry + RegexBytecode.BYTECODE_IDENTIFIER_MAP_ENTRY_SIZE
        Loop
    End With
    
    Set MakeCapturesDict = oRet
End Function

'Creates a new regex object
'@param pattern - The regex pattern
'@param flags - The regex flags
'@param stepsLimit - The maximum number of processing steps the engine is allowed to make before giving up.
'@returns - The regex object
'@remark - The flags are:
' - `i` - Case insensitive
' - `m` - Multiline
' - `g` - Global
'@example
'```vb
'stdRegex.Create("A\d+","i")
'```
Public Function Create(ByRef pattern As String, Optional ByRef flags As String = vbNullString, Optional stepsLimit As Long = REGEX3_DEFAULT_STEPS_LIMIT) As stdRegex3
    Dim obj As stdRegex3
    Set Create = New stdRegex3
    Create.protInit pattern, flags, stepsLimit
End Function

'Get/Set pattern for regex
'@returns - Pattern for regex
Public Property Get pattern() As String
    If Me Is stdRegex3 Then
        Error.Raise "Property called on class, not object"
        Exit Property
    End If
    
    pattern = This.regex.pattern
End Property
Public Property Let pattern(ByRef val As String)
    If Me Is stdRegex3 Then
        Error.Raise "Property called on class, not object"
        Exit Property
    End If

    Dim caseInsensitive As Boolean
    caseInsensitive = RegexBytecode.isCaseInsensitive(This.regex.bytecode)
    StaticRegex.InitializeRegex This.regex, pattern, caseInsensitive
End Property

'Get/Set Flags for regex
'@returns - Flags for regex
'@remark
'Changing the case sensitivity of the regex (the `i` flag) is expensive,
'as it requires a recompilation of the pattern.
Public Property Get flags() As String
    If Me Is stdRegex3 Then
        Error.Raise ("Property called on class, not object")
        Exit Function
    End If
    
    Dim nFlags As Long, i As Long, caseInsensitive As Boolean
    Const UNICODE_LC_G As Long = 103
        
    caseInsensitive = RegexBytecode.isCaseInsensitive(This.regex.bytecode)
    nFlags = 0& - This.isGlobal - caseInsensitive - This.isMultiline
    If nFlags = 0 Then flags = vbNullString: Exit Property
    
    flags = String(nFlags, UNICODE_LC_G)
    i = 1
    If This.isGlobal Then i = i + 1
    If caseInsensitive Then Mid$(flags, i, 1) = "i": i = i + 1
    If This.isMultiline Then Mid$(flags, i, 1) = "m": i = i + 1
End Property
Public Property Let flags(ByRef val As String)
    If Me Is stdRegex3 Then
        Error.Raise ("Property called on class, not object")
        Exit Property
    End If
    
    Dim caseInsensitive As Boolean, isGlobal As Boolean, isMultiline As Boolean
        
    DecodeFlags caseInsensitive, isGlobal, isMultiline, val
    If RegexBytecode.isCaseInsensitive(This.regex.bytecode) <> caseInsensitive Then
        StaticRegex.InitializeRegex This.regex, This.regex.pattern, caseInsensitive:=caseInsensitive
    End If
    This.isGlobal = isGlobal
    This.isMultiline = isMultiline
End Property

'Test whether the regex matches the given string (or a substring thereof).
'@param str - The string to test
'@return - `True` if the string matches the regex, `False` otherwise
Public Function Test(ByRef str As String) As Boolean
    If Me Is stdRegex3 Then
        Error.Raise "Method called on class, not on object"
        Exit Function
    End If
    
    Test = StaticRegex.Test( _
        This.regex, str, multiline:=This.isMultiline _
    )
End Function

'Find the first matching substring and all submatches.
'@param sHaystack - The string to execute the regex against
'@return - A dictonary object containing all submatches
'@remark If dictionary dict is returned,
' - `dict(0)` is the entire match
' - `dict(n)` is the n-th submatch (n >= 1)
' - `dict(name)` is the submatch named `name`
'@example
'Regex: `(?<id>\d{5}-ST[A-Z]\d) - (?<desc>.*)`
'
'Data:
'```
'Some sites were in critical condition
'* 12345-STA1 - Large crack through pipe.
'* 12323-STB9 - Acid leakage polluting watercourse.
'and some others were largely ok:
'* 23565-STC2
'* 62346-STZ9
'```
'Result:
'```json
'{
'  0: "12345-STA1 - Large crack through pipe.",
'  "id":"12345-STA1",
'  1: "12345-STA1",
'  "desc": "Large crack through pipe."
'  2: "Large crack through pipe.",
'  "$COUNT":2
'}
'```
Public Function Match(ByRef sHaystack As String) As Object
    If Me Is stdRegex3 Then
        Error.Raise "Method called on class, not on object"
        Exit Function
    End If
    
    Dim matcherState As StaticRegex.MatcherStateTy
    
    If StaticRegex.Match(matcherState, This.regex, sHaystack, multiline:=This.isMultiline) Then
        Set Match = MakeCapturesDict(matcherState.captures, sHaystack)
    Else
        Set Match = Nothing
    End If
End Function

'Find all subsequent matching substrings and all submatches.
'@param sHaystack - The string to execute the regex against
'@return - A collection of dictonaries containing all submatches
'@remark For each dictionary dict in the collection returned,
' - `dict(0)` is the entire match
' - `dict(n)` is the n-th submatch (n >= 1)
' - `dict(name)` is the submatch named `name`
'@example
'Regex: `\d{5}-ST[A-Z]\d`
'
'Data:
'```
'Some sites were in critical condition
'* 12345-STA1 - Large crack through pipe.
'* 12323-STB9 - Acid leakage polluting watercourse.
'and some others were largely ok:
'* 23565-STC2
'* 62346-STZ9
'```
'Result:
'```json
'[
'  {0: "12345-STA1", ...},
'  {0: "12323-STB9", ...},
'  {0: "23565-STC2", ...},
'  {0: "62346-ST29", ...}
']
'```
Public Function MatchAll(ByRef sHaystack As String) As Collection
    If Me Is stdRegex3 Then
        Error.Raise "Method called on class, not on object"
        Exit Function
    End If
    
    Dim matcherState As StaticRegex.MatcherStateTy
    Dim oRet As Collection

    If Me Is stdRegex3 Then
        Error.Raise "Method called on class, not on object"
        Exit Function
    End If

    Set oRet = New Collection

    matcherState.localMatch = False
    matcherState.multiline = This.isMultiline
    Do While StaticRegex.MatchNext(matcherState, This.regex, sHaystack)
        oRet.Add MakeCapturesDict(matcherState.captures, sHaystack)
    Loop

    Set MatchAll = oRet
End Function

'Given a haystack and a replacer, replace matching substrings.
'If the "global" flag on the regex is not set, only the first matching substring will be replaced;
'if the "global" flag is set, all maching substrings will be replaced.
'@param sHaystack - The string to execute the regex against
'@param sReplacer - Format string specifying what to replace the matches with
'@returns - The resulting string
'@remark In the replacer, the following special replacement patterns are allowed:
' `$$` - Inserts a `$`
' `$&` - Inserts the matched substring
' `$`` - Inserts the portion of the string that precedes the matched substring
' `$'` - Inserts the portion of the string that follows the matched substring
' `$n` - Inserts the n-th (1-indexed) capturing group
' `$<name>` - Inserts the named capturing group where `name` is the group name
' `$~` - Does not insert anything (intended to be used as a separator, e.g. `$7$~222`)
'@example
'Regex: `(?<id>\d{5}-ST[A-Z]\d)\s+(?<count>\d+)\s+(?<date>../../....)`
'Flag: `g`
'
'Haystack:
'```
'Here is some cool data:
'12345-STA1  123    10/02/2019
'12323-STB9  2123   01/01/2005
'and here is some more:
'23565-STC2  23     ??/??/????
'62346-STZ9  5      01/05/1932
'```
'Replacer: `$<id>,$<date>,$<count>`
'
'Result:
'```
'Here is some cool data:
'12345-STA1,10/02/2019,123
'12323-STB9,01/01/2005,2123
'and here is some more:
'23565-STC2,??/??/????,23
'62346-STZ9,01/05/1932,5
'```
Public Function Replace(ByRef sHaystack As String, ByRef sReplacer As String) As String
    If Me Is stdRegex3 Then
        Error.Raise "Method called on class, not on object"
        Exit Function
    End If

    Replace = StaticRegex.Replace(This.regex, sReplacer, sHaystack, localMatch:=Not This.isGlobal, multiline:=This.isMultiline)
End Function

'Given a haystack and a format string, generate a string containing all matches in the format provided.
'If the "global" flag on the regex is not set, only the first match will be considered;
'if the "global" flag is set, all maching substrings will be considered.
'@param sHaystack - The string to execute the regex against
'@param sFormat - Format string specifying how the matches will be formatted (optional, default is `$&`)
'@param sDelimiter - Delimiter (optional, default is `vbCrLf`)
'@returns - A string containing the concatenation of all match results
'@remark In the format string, the following special patterns are allowed:
' `$$` - Inserts a `$`
' `$&` - Inserts the matched substring
' `$`` - Inserts the portion of the string that precedes the matched substring
' `$'` - Inserts the portion of the string that follows the matched substring
' `$n` - Inserts the n-th (1-indexed) capturing group
' `$<name>` - Inserts the named capturing group where `name` is the group name
' `$~` - Does not insert anything (intended to be used as a separator, e.g. `$7$~222`)
'@example
'Regex: `(?<id>\d{5}-ST[A-Z]\d)\s+(?<count>\d+)\s+(?<date>../../....)`
'Flag: `g`
'
'sHaystack:
'```
'Here is some cool data:
'12345-STA1  123    10/02/2019
'12323-STB9  2123   01/01/2005
'and here is some more:
'23565-STC2  23     ??/??/????
'62346-STZ9  5      01/05/1932
'```
'sFormat: `$<id>,$<date>,$<count>`
'
'sDelimiter: `;`
'
'Result: `12345-STA1,10/02/2019,123;12323-STB9,01/01/2005,2123;23565-STC2,??/??/????,23;62346-STZ9,01/05/1932,5`
Public Function List(ByRef sHaystack As String, Optional ByRef sFormat As String = "$&", Optional ByRef sDelimiter As String = vbCrLf) As String
    If Me Is stdRegex3 Then
        Error.Raise "Method called on class, not on object"
        Exit Function
    End If
    
    List = StaticRegex.MatchThenJoin( _
        This.regex, sHaystack, format:=sFormat, delimiter:=sDelimiter, _
        localMatch:=Not This.isGlobal, multiline:=This.isMultiline _
    )
End Function

'Supplied with a haystack to search, ListArray will generate a 2d array of data where the rows represent each match and
'the columns represent the formats passed into formats parameter
'@param results - The resulting array (output parameter)
'@param sHaystack - Haystack to search
'@param sFormats - Array of format string specifying the column formats
'@remark
'Useful especially for dumping matches to Excel.
'@example
'Regex: `(?<id>\d{5}-ST[A-Z]\d)\s+(?<count>\d+)\s+(?<date>../../....)`
'Flag: `g`
'
'sHaystack:
'```
'Here is some cool data:
'12345-STA1  123    10/02/2019
'12323-STB9  2123   01/01/2005
'and here is some more:
'23565-STC2  23     ??/??/????
'62346-STZ9  5      01/05/1932
'```
'Formats: `[ "$<date>", "$<id>" , "$<count> incidents" ]`
'
'Result:
'```
'[ [ "10/02/2019", "12345-STA1", "123 incidents" ],
'  [ "01/01/2005", "12323-STB9", "2123 incidents" ],
'  [ "??/??/????", "23565-STC2", "23 incidents" ],
'  [ "01/05/1932", "62346-STZ9", "5 incidents" ] ]
'```
Public Sub ListArray(ByRef results() As String, ByRef sHaystack As String, ByRef formats() As String)
    If Me Is stdRegex3 Then
        Error.Raise "Method called on class, not on object"
        Exit Sub
    End If

    StaticRegex.MatchThenList results, This.regex, sHaystack, formats, localMatch:=Not This.isGlobal, multiline:=This.isMultiline
End Sub
