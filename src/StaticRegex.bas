Attribute VB_Name = "StaticRegex"
Option Explicit

Public Type RegexTy
    pattern As String
    bytecode() As Long
    isCaseInsensitive As Long
    stepsLimit As Long
End Type

Private Enum DfaStateStructInitConstants
    DFA_IS_INITIALIZED = 1
    DFA_MULTILINE = 2
    DFA_DOTALL = 4
    DFA_CASE_INSENSITIVE = 8
End Enum

' A format is compiled for the regex used to resolve its capture references.
Public Type FormatTy
    text As String
    parsed() As Long
End Type

Public Type MatcherStateTy
    localMatch As Boolean
    multiline As Boolean
    dotAll As Boolean
    dfaStateStructInitState As Integer
    dfaPattern As String
    current As Long
    captures As RegexMatcherCommon.CapturesTy
    context As RegexDfsMatcher.DfsMatcherContext
    dfaStateStruct As RegexDfaMatcher.StateStructure
End Type

Public Function TryInitializeRegex( _
    ByRef regex As RegexTy, _
    ByRef pattern As String, _
    Optional ByVal caseInsensitive As Boolean = False _
) As Boolean
    ' Only invalid pattern syntax is converted to False; other failures propagate.
    On Error GoTo Fail
    InitializeRegex regex, pattern, caseInsensitive
    TryInitializeRegex = True
    Exit Function

Fail:
    Select Case Err.Number
    Case RegexErrors.REGEX_ERR_INVALID_REGEXP_ESCAPE, RegexErrors.REGEX_ERR_INVALID_ESCAPE, _
        RegexErrors.REGEX_ERR_INVALID_RANGE, RegexErrors.REGEX_ERR_UNTERMINATED_CHARCLASS, _
        RegexErrors.REGEX_ERR_INVALID_REGEXP_GROUP, RegexErrors.REGEX_ERR_INVALID_REGEXP_CHARACTER, _
        RegexErrors.REGEX_ERR_INVALID_QUANTIFIER, RegexErrors.REGEX_ERR_UNEXPECTED_END_OF_PATTERN, _
        RegexErrors.REGEX_ERR_UNEXPECTED_REGEXP_TOKEN, RegexErrors.REGEX_ERR_UNEXPECTED_CLOSING_PAREN, _
        RegexErrors.REGEX_ERR_INVALID_QUANTIFIER_NO_ATOM, RegexErrors.REGEX_ERR_INVALID_IDENTIFIER, _
        RegexErrors.REGEX_ERR_INVALID_MODIFIER
        TryInitializeRegex = False
    Case Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select
End Function

Public Sub InitializeRegex( _
    ByRef regex As RegexTy, _
    ByRef pattern As String, _
    Optional ByVal caseInsensitive As Boolean = False _
)
    Dim compiledBytecode() As Long
    ' Compile in isolation, then replace bytecode before changing its metadata.
    RegexCompiler.Compile compiledBytecode, pattern, caseInsensitive:=caseInsensitive
    regex.bytecode = compiledBytecode
    regex.pattern = pattern
    regex.isCaseInsensitive = caseInsensitive
    regex.stepsLimit = RegexDfsMatcher.DEFAULT_STEPS_LIMIT
End Sub

Private Function UseDfs(ByRef regex As RegexTy, ByVal multiline As Boolean) As Boolean
    Dim anchorOp As Long
    If (regex.bytecode(RegexBytecode.BYTECODE_IDX_PROPERTIES) And BYTECODE_PROPERTY_REQUIRES_DFS_MATCHING) <> 0 Or _
        regex.stepsLimit <> RegexDfsMatcher.DEFAULT_STEPS_LIMIT Then
        UseDfs = True
    ElseIf (regex.bytecode(RegexBytecode.BYTECODE_IDX_PROPERTIES) And BYTECODE_PROPERTY_PREFER_DFS) <> 0 Then
        anchorOp = RegexMatcherCommon.LeadingStartAnchor(regex.bytecode)
        UseDfs = (anchorOp - (multiline And MODIFIER_M_WRITE) And MODIFIER_M_ACTIVE) = 0
    End If
End Function

'Test whether a string matches the regex
'@return - `True` if the string matches the regex, `False` otherwise
Public Function Test( _
    ByRef regex As RegexTy, ByRef str As String, _
    Optional ByVal multiline As Boolean = False, _
    Optional ByVal dotAll As Boolean = False _
) As Boolean
    Dim captures As RegexMatcherCommon.CapturesTy
    
    If UseDfs(regex, multiline) Then
        Test = RegexDfsMatcher.DfsMatch( _
            captures, regex.bytecode, str, stepsLimit:=regex.stepsLimit, _
            multiline:=multiline, _
            dotAll:=dotAll _
        ) <> -1
    Else
        Test = RegexDfaMatcher.DfaMatch( _
            regex.bytecode, str, True, multiline:=multiline, dotAll:=dotAll _
        ) <> -1
    End If
End Function

' Invalidate the result without releasing the capture buffers.
Private Sub ClearCaptures(ByRef captures As RegexMatcherCommon.CapturesTy)
    With captures
        .entireMatch.start = 0
        .entireMatch.Length = 0
        .nNumberedCaptures = 0
        .nNamedCaptures = 0
    End With
End Sub

Private Function ExtractCapturesAt( _
    ByRef matcherState As MatcherStateTy, ByRef regex As RegexTy, ByRef haystack As String, _
    ByRef slp As RegexMatcherCommon.StartLengthPair, ByVal multiline As Boolean, ByVal dotAll As Boolean _
) As Long
    Dim r As Long
    r = RegexNfaMatcher.NfaMatchFrom( _
        matcherState.captures, regex.bytecode, haystack, slp.start - 1, _
        multiline:=multiline, dotAll:=dotAll, matchAt:=True _
    )
    If r = -1 Or matcherState.captures.entireMatch.start <> slp.start Or _
        matcherState.captures.entireMatch.Length <> slp.Length Then
        ClearCaptures matcherState.captures
        matcherState.current = -1
        Err.Raise RegexErrors.REGEX_ERR_INTERNAL_LOGIC_ERR, "StaticRegex", "DFA and NFA match spans disagree."
    End If
    ExtractCapturesAt = r
End Function

' Execute the regex against a string
' A failed match clears the result; capture getters then return empty strings.
Public Function Match( _
    ByRef matcherState As MatcherStateTy, ByRef regex As RegexTy, ByRef haystack As String, _
    Optional ByVal multiline As Boolean = False, _
    Optional ByVal dotAll As Boolean = False, _
    Optional ByVal matchFrom As Long = 1 _
) As Boolean
    Dim expectedInitState As Long, dfaIsPreliminary As Boolean, _
        slp As RegexMatcherCommon.StartLengthPair
    If UseDfs(regex, multiline) Then
        Match = RegexDfsMatcher.DfsMatchFrom( _
            matcherState.context, matcherState.captures, regex.bytecode, haystack, _
            matchFrom - 1, stepsLimit:=regex.stepsLimit, multiline:=multiline, _
            dotAll:=dotAll _
        ) <> -1
    Else
        dfaIsPreliminary = _
            regex.bytecode(RegexBytecode.BYTECODE_IDX_MAX_PROPER_CAPTURE_SLOT) > 1 Or _
            regex.bytecode(RegexBytecode.BYTECODE_IDX_N_IDENTIFIERS) > 0
        expectedInitState = DFA_IS_INITIALIZED Or DFA_MULTILINE And multiline Or _
            DFA_DOTALL And dotAll Or DFA_CASE_INSENSITIVE And regex.isCaseInsensitive
        If matcherState.dfaStateStructInitState <> expectedInitState Or _
            matcherState.dfaPattern <> regex.pattern Then
            DfaStateStructureInitialize matcherState.dfaStateStruct, multiline, dotAll
            matcherState.dfaStateStructInitState = expectedInitState
            matcherState.dfaPattern = regex.pattern
        End If
        ' Captures also require the leftmost start. An eager DFA acceptance may
        ' belong to a later start while a higher-priority branch is still running.
        If Not RegexDfaMatcher.DfaFindMatch( _
            slp, matcherState.dfaStateStruct, regex.bytecode, haystack, matchFrom, _
            False _
        ) Then
            Match = False
        ElseIf dfaIsPreliminary Then
            ExtractCapturesAt matcherState, regex, haystack, slp, multiline, dotAll
            Match = True
        Else
            ClearCaptures matcherState.captures
            matcherState.captures.entireMatch = slp
            Match = True
        End If
    End If
    If Not Match Then ClearCaptures matcherState.captures
    matcherState.current = -1
End Function

Public Sub InitializeFormat(ByRef format As FormatTy, ByRef regex As RegexTy, ByRef text As String)
    Dim parsed As ArrayBuffer.Ty
    RegexReplace.ParseFormatString parsed, text, regex.bytecode, regex.pattern
    format.text = text
    format.parsed = parsed.Buffer
End Sub

Public Function FormatMatch(ByRef matcherState As MatcherStateTy, ByRef haystack As String, ByRef format As FormatTy) As String
    Dim builder As StaticStringBuilder.Ty
    If matcherState.captures.entireMatch.start = 0 Then Exit Function
    RegexReplace.AppendFormatted builder, haystack, matcherState.captures, format.text, format.parsed
    FormatMatch = StaticStringBuilder.GetStr(builder)
End Function

Public Function GetCapture(ByRef matcherState As MatcherStateTy, ByRef haystack As String, Optional ByVal num As Long = 0) As String
    If num = 0 Then
        With matcherState.captures.entireMatch
            If .Length > 0 Then GetCapture = Mid$(haystack, .start, .Length) Else GetCapture = vbNullString
        End With
    ElseIf num > 0 And num <= matcherState.captures.nNumberedCaptures Then
        With matcherState.captures.numberedCaptures(num - 1)
            If .Length > 0 Then GetCapture = Mid$(haystack, .start, .Length) Else GetCapture = vbNullString
        End With
    Else
        GetCapture = vbNullString
    End If
End Function

Public Function GetCaptureByName( _
    ByRef matcherState As MatcherStateTy, _
    ByRef regex As RegexTy, _
    ByRef haystack As String, _
    ByRef name As String _
) As String
    Dim identifierId As Long, captureNum As Long
    
    identifierId = RegexBytecode.GetIdentifierId(regex.bytecode, regex.pattern, name)
    If identifierId < 0 Or identifierId >= matcherState.captures.nNamedCaptures Then Exit Function
    captureNum = matcherState.captures.namedCaptures(identifierId)
    If captureNum >= 0 Then
        GetCaptureByName = GetCapture(matcherState, haystack, captureNum)
    Else
        GetCaptureByName = vbNullString
    End If
End Function

Public Function MatchNext( _
    ByRef matcherState As MatcherStateTy, ByRef regex As RegexTy, ByRef haystack As String, _
    Optional ByVal raiseOnLimit As Boolean = False _
) As Boolean
    Dim expectedInitState As Long, dfaIsPreliminary As Boolean, _
        slp As RegexMatcherCommon.StartLengthPair, r As Long
    
    If matcherState.current > Len(haystack) Then GoTo NoMatch
    If matcherState.current < 0 Then GoTo NoMatch
    
    If UseDfs(regex, matcherState.multiline) Or raiseOnLimit Then
        r = RegexDfsMatcher.DfsMatchFrom( _
            matcherState.context, matcherState.captures, regex.bytecode, haystack, _
            matcherState.current, stepsLimit:=regex.stepsLimit, _
            multiline:=matcherState.multiline, dotAll:=matcherState.dotAll, raiseOnLimit:=raiseOnLimit _
        )
    Else
        dfaIsPreliminary = _
            regex.bytecode(RegexBytecode.BYTECODE_IDX_MAX_PROPER_CAPTURE_SLOT) > 1 Or _
            regex.bytecode(RegexBytecode.BYTECODE_IDX_N_IDENTIFIERS) > 0
        expectedInitState = DFA_IS_INITIALIZED Or _
            DFA_MULTILINE And matcherState.multiline Or _
            DFA_DOTALL And matcherState.dotAll Or DFA_CASE_INSENSITIVE And regex.isCaseInsensitive
        If matcherState.dfaStateStructInitState <> expectedInitState Or _
            matcherState.dfaPattern <> regex.pattern Then
            DfaStateStructureInitialize matcherState.dfaStateStruct, _
                matcherState.multiline, matcherState.dotAll
            matcherState.dfaStateStructInitState = expectedInitState
            matcherState.dfaPattern = regex.pattern
        End If
        ' Preserve the leftmost start even when the NFA will collect captures.
        If Not RegexDfaMatcher.DfaFindMatch( _
            slp, matcherState.dfaStateStruct, regex.bytecode, haystack, _
            matcherState.current + 1, False _
        ) Then
            r = -1
        ElseIf dfaIsPreliminary Then
            r = ExtractCapturesAt(matcherState, regex, haystack, slp, matcherState.multiline, matcherState.dotAll)
        Else
            ClearCaptures matcherState.captures
            matcherState.captures.entireMatch = slp
            r = slp.start + slp.Length - 1
        End If
    End If
    If r = -1 Then GoTo NoMatch
    matcherState.current = (r - (matcherState.captures.entireMatch.Length = 0)) Or matcherState.localMatch
    MatchNext = True
    Exit Function
    
NoMatch:
    ClearCaptures matcherState.captures
    matcherState.current = -1
    MatchNext = False
End Function

Public Function Replace( _
    ByRef regex As RegexTy, _
    ByRef replacer As String, _
    ByRef haystack As String, _
    Optional ByVal localMatch As Boolean = False, _
    Optional ByVal multiline As Boolean = False, _
    Optional ByVal dotAll As Boolean = False _
) As String
    Dim parsedFormat As ArrayBuffer.Ty, matcherState As MatcherStateTy, lastEndPos As Long, resultBuilder As StaticStringBuilder.Ty

    RegexReplace.ParseFormatString parsedFormat, replacer, regex.bytecode, regex.pattern

    lastEndPos = 1
    matcherState.localMatch = localMatch
    matcherState.multiline = multiline
    matcherState.dotAll = dotAll
    Do While MatchNext(matcherState, regex, haystack)
        StaticStringBuilder.AppendStr resultBuilder, Mid$(haystack, lastEndPos, matcherState.captures.entireMatch.start - lastEndPos)
        RegexReplace.AppendFormatted resultBuilder, haystack, matcherState.captures, replacer, parsedFormat.Buffer
        
        lastEndPos = matcherState.captures.entireMatch.start + matcherState.captures.entireMatch.Length
    Loop
    
    StaticStringBuilder.AppendStr resultBuilder, Mid$(haystack, lastEndPos)
    
    Replace = StaticStringBuilder.GetStr(resultBuilder)
End Function

Public Function SplitByRegex( _
    ByRef regex As RegexTy, _
    ByRef haystack As String, _
    Optional ByVal localMatch As Boolean = False, _
    Optional ByVal multiline As Boolean = False, _
    Optional ByVal dotAll As Boolean = False _
) As Collection
    Dim matcherState As MatcherStateTy, lastEndPos As Long, colStrings As Collection
    Dim i As Long, u As Long, s As String

    Set colStrings = New Collection

    lastEndPos = 1
    matcherState.localMatch = localMatch
    matcherState.multiline = multiline
    matcherState.dotAll = dotAll
    Do While MatchNext(matcherState, regex, haystack)
        With matcherState.captures.entireMatch
            colStrings.Add Mid$(haystack, lastEndPos, .start - lastEndPos)
            lastEndPos = .start + .Length
        End With

        u = matcherState.captures.nNumberedCaptures - 1
        For i = 0 To u
            With matcherState.captures.numberedCaptures(i)
                s = vbNullString
                If .Length > 0 Then s = Mid$(haystack, .start, .Length)
                colStrings.Add s
            End With
        Next
    Loop
    colStrings.Add Mid$(haystack, lastEndPos)
    
    Set SplitByRegex = colStrings
End Function

Public Function MatchThenJoin( _
    ByRef regex As RegexTy, _
    ByRef haystack As String, _
    Optional ByRef format As String = "$&", _
    Optional ByRef delimiter As String = vbNullString, _
    Optional ByVal localMatch As Boolean = False, _
    Optional ByVal multiline As Boolean = False, _
    Optional ByVal dotAll As Boolean = False _
) As String
    Dim parsedFormat As ArrayBuffer.Ty, resultBuilder As StaticStringBuilder.Ty, matcherState As MatcherStateTy
    
    RegexReplace.ParseFormatString parsedFormat, format, regex.bytecode, regex.pattern
    
    matcherState.localMatch = localMatch
    matcherState.multiline = multiline
    matcherState.dotAll = dotAll
    If MatchNext(matcherState, regex, haystack) Then
        AppendFormatted resultBuilder, haystack, matcherState.captures, format, parsedFormat.Buffer
        Do While MatchNext(matcherState, regex, haystack)
            StaticStringBuilder.AppendStr resultBuilder, delimiter
            AppendFormatted resultBuilder, haystack, matcherState.captures, format, parsedFormat.Buffer
        Loop
    End If
    
    MatchThenJoin = StaticStringBuilder.GetStr(resultBuilder)
End Function

Public Sub MatchThenList( _
    ByRef results() As String, _
    ByRef regex As RegexTy, _
    ByRef haystack As String, _
    ByRef formatStrings() As String, _
    Optional ByVal localMatch As Boolean = False, _
    Optional ByVal multiline As Boolean = False, _
    Optional ByVal dotAll As Boolean = False _
)
    Dim cola As Long, colb As Long, j As Long, k As Long, m As Long, mm As Long, nMatches As Long
    Dim parsedFormats As ArrayBuffer.Ty
    Dim matcherState As MatcherStateTy
    Dim resultBuilder As StaticStringBuilder.Ty

    cola = LBound(formatStrings)
    colb = UBound(formatStrings)
    
    For j = cola To colb
        k = parsedFormats.Length
        ArrayBuffer.AppendLong parsedFormats, 0
        RegexReplace.ParseFormatString parsedFormats, formatStrings(j), regex.bytecode, regex.pattern
        parsedFormats.Buffer(k) = parsedFormats.Length - k
    Next
    
    nMatches = 0
    
    matcherState.localMatch = localMatch
    matcherState.multiline = multiline
    matcherState.dotAll = dotAll
    Do While MatchNext(matcherState, regex, haystack)
        k = 0
        For j = cola To colb
            m = resultBuilder.Length
            AppendFormatted resultBuilder, haystack, matcherState.captures, formatStrings(j), parsedFormats.Buffer, k + 1
            ' Store the nonnegative Long offset as two UTF-16 characters, low word first.
            StaticStringBuilder.AppendStr resultBuilder, ChrW$(m And &HFFFF&) & ChrW$(m \ &H10000&)
            k = k + parsedFormats.Buffer(k)
        Next
        
        nMatches = nMatches + 1
    Loop
    
    If nMatches = 0 Then
        ' Create an initialized, zero-length String array.
        results = Split(vbNullString)
    Else
        ReDim results(0 To nMatches - 1, cola To colb) As String
        m = resultBuilder.Length
        For j = nMatches - 1 To 0 Step -1
            For k = colb To cola Step -1
                mm = (AscW(StaticStringBuilder.GetSubstr(resultBuilder, m - 1, 1)) And &HFFFF&) + _
                    (AscW(StaticStringBuilder.GetSubstr(resultBuilder, m, 1)) And &HFFFF&) * &H10000&
                results(j, k) = StaticStringBuilder.GetSubstr(resultBuilder, mm + 1, m - mm - 2)
                m = mm
            Next
        Next
    End If
End Sub

Public Sub InitializeMatcherState( _
    ByRef matcherState As MatcherStateTy, _
    Optional ByVal localMatch As Boolean = False, _
    Optional ByVal multiline As Boolean = False, _
    Optional ByVal dotAll As Boolean = False _
)
    matcherState.current = 0
    matcherState.localMatch = localMatch
    matcherState.multiline = multiline
    matcherState.dotAll = dotAll
End Sub

Public Sub ResetMatcherState(ByRef matcherState As MatcherStateTy)
    matcherState.current = 0
End Sub
