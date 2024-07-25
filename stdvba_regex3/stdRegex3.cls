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

'Creates a new regex object
'@param pattern - The regex pattern
'@param Flags - The regex flags
'@returns - The regex object
'@remark - The flags are:
' - `i` - Case insensitive
' - `m` - Multiline
' - `g` - Global
Public Function Create(ByRef pattern As String, Optional ByRef flags As String = vbNullString, Optional stepsLimit As Long = REGEX3_DEFAULT_STEPS_LIMIT) As stdRegex3
    Dim obj As stdRegex3
    Set Create = New stdRegex3
    Create.protInit pattern, flags, stepsLimit
End Function

'Initializes the regex object
'@protected
'@param pattern - The regex pattern
'@param Flags - The regex flags
'@remark - This is a protected method and should not be called directly. Use `Create` instead.
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


Public Property Get pattern() As String
    If Me Is stdRegex3 Then
        Error.Raise ("Property called on class, not object")
        Exit Property
    End If
    
    pattern = This.regex.pattern
End Property

Public Property Let pattern(ByRef val As String)
    Dim caseInsensitive As Boolean
    
    If Me Is stdRegex3 Then
        Error.Raise ("Property called on class, not object")
        Exit Property
    End If

    caseInsensitive = RegexBytecode.isCaseInsensitive(This.regex.bytecode)
    StaticRegex.InitializeRegex This.regex, pattern, caseInsensitive
End Property


Public Property Get flags() As String
    Dim nFlags As Long, i As Long, caseInsensitive As Boolean
    Const UNICODE_LC_G As Long = 103
    
    If Me Is stdRegex3 Then
        Error.Raise ("Property called on class, not object")
        Exit Function
    End If
        
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
    Dim caseInsensitive As Boolean, isGlobal As Boolean, isMultiline As Boolean
    
    If Me Is stdRegex3 Then
        Error.Raise ("Property called on class, not object")
        Exit Property
    End If
        
    DecodeFlags caseInsensitive, isGlobal, isMultiline, val
    If RegexBytecode.isCaseInsensitive(This.regex.bytecode) <> caseInsensitive Then
        StaticRegex.InitializeRegex This.regex, This.regex.pattern, caseInsensitive:=caseInsensitive
    End If
    This.isGlobal = isGlobal
    This.isMultiline = isMultiline
End Property

'Test whether a string matches the regex
'@param str - The string to test
'@param stepsLimit - The number of steps after which the DFS matcher gives up
'@return - `True` if the string matches the regex, `False` otherwise
Public Function Test(ByRef str As String) As Boolean
    If Me Is stdRegex3 Then
        Error.Raise ("Method called on class, not object")
        Exit Function
    End If
    
    Test = StaticRegex.Test( _
        This.regex, str, multiline:=This.isMultiline _
    )
End Function

'Execute the regex against a string
'@param sHaystack - The string to execute the regex against
'@param stepsLimit - The number of steps after which the DFS matcher gives up
'@return - A collection of matches
Public Function Match(ByRef sHaystack As String) As Object
    Dim matcherState As StaticRegex.MatcherStateTy

    If Me Is stdRegex3 Then
        Error.Raise ("Method called on class, not object")
        Exit Function
    End If
    
    If StaticRegex.Match(matcherState, This.regex, sHaystack, multiline:=This.isMultiline) Then
        Set Match = MakeCapturesDict(matcherState.captures, sHaystack)
    Else
        Set Match = Nothing
    End If
End Function

Public Function MatchAll(ByRef sHaystack As String) As Collection
    Dim matcherState As StaticRegex.MatcherStateTy
    Dim oRet As Collection

    If Me Is stdRegex3 Then
        Error.Raise "Method called on class, not object"
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

Private Function MakeCapturesDict(ByRef captures As RegexDfsMatcher.CapturesTy, ByRef sHaystack As String) As Object
    Dim j As Long, identifier As String, referenced As Long
    Dim identifierMapEnd As Long, currentEntry As Long
    Dim oRet As Object
    
    Set oRet = CreateObject("Scripting.Dictionary")
    
    With This
        oRet("$COUNT") = captures.nNumberedCaptures
    
        With captures.entireMatch
            If .start > 0 Then oRet(0) = Mid$(sHaystack, .start, .Length) Else oRet(0) = vbNullString
        End With
        For j = 1 To captures.nNumberedCaptures
            With captures.numberedCaptures(j - 1)
                If .start > 0 Then oRet(j) = Mid$(sHaystack, .start, .Length) Else oRet(j) = vbNullString
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

Public Function Replace(ByRef sHaystack As String, ByRef sReplacer As String) As String
    If Me Is stdRegex3 Then
        Error.Raise "Method called on class, not object"
        Exit Function
    End If

    Replace = StaticRegex.Replace(This.regex, sReplacer, sHaystack, localMatch:=Not This.isGlobal, multiline:=This.isMultiline)
End Function

Public Function List(ByRef sHaystack As String, ByRef sFormat As String) As String
    If Me Is stdRegex3 Then
        Error.Raise "Method called on class, not object"
        Exit Function
    End If
    
    List = StaticRegex.MatchThenJoin(This.regex, sHaystack, format:=sFormat, localMatch:=Not This.isGlobal, multiline:=This.isMultiline)
End Function

Public Sub ListArray(ByRef results() As String, ByRef sHaystack As String, ByRef formats() As String)
    If Me Is stdRegex3 Then
        Error.Raise "Method called on class, not object"
        Exit Sub
    End If

    StaticRegex.MatchThenList results, This.regex, sHaystack, formats, localMatch:=Not This.isGlobal, multiline:=This.isMultiline
End Sub
