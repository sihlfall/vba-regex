Attribute VB_Name = "StaticRegex"
Option Explicit

Public Type RegexTy
    pattern As String
    bytecode() As Long
    isCaseInsensitive As Long
    stepsLimit As Long
End Type

Public Type MatcherStateTy
    localMatch As Boolean
    multiline As Boolean
    dotAll As Boolean
    current As Long
    captures As RegexDfsMatcher.CapturesTy
    context As RegexDfsMatcher.DfsMatcherContext
End Type

Public Function TryInitializeRegex( _
    ByRef regex As RegexTy, _
    ByRef pattern As String, _
    Optional ByVal caseInsensitive As Boolean = False _
) As Boolean
    ' Todo:
    '   Actually, this is not what we want to have.
    '   We should change the compiler so that it reports syntax errors in the regex via a channel
    '     different to throwing.
    '   Then InitializeRegex should make use of TryInitializeRegex, not the other way round.
    On Error GoTo Fail
    InitializeRegex regex, pattern, caseInsensitive
    TryInitializeRegex = True
    Exit Function

Fail:
    TryInitializeRegex = False
End Function

Public Sub InitializeRegex( _
    ByRef regex As RegexTy, _
    ByRef pattern As String, _
    Optional ByVal caseInsensitive As Boolean = False _
)
    regex.pattern = pattern
    regex.isCaseInsensitive = caseInsensitive
    regex.stepsLimit = RegexDfsMatcher.DEFAULT_STEPS_LIMIT
    RegexCompiler.Compile regex.bytecode, pattern, caseInsensitive:=caseInsensitive
End Sub

'Test whether a string matches the regex
'@return - `True` if the string matches the regex, `False` otherwise
Public Function Test( _
    ByRef regex As RegexTy, ByRef str As String, _
    Optional ByVal multiline As Boolean = False, _
    Optional ByVal dotAll As Boolean = False _
) As Boolean
    Dim captures As RegexDfsMatcher.CapturesTy
    
    Test = RegexDfsMatcher.DfsMatch( _
        captures, regex.bytecode, str, stepsLimit:=regex.stepsLimit, _
        multiline:=multiline, _
        dotAll:=dotAll _
    ) <> -1
End Function

' Execute the regex against a string
Public Function Match( _
    ByRef matcherState As MatcherStateTy, ByRef regex As RegexTy, ByRef haystack As String, _
    Optional ByVal multiline As Boolean = False, _
    Optional ByVal dotAll As Boolean = False, _
    Optional ByVal matchFrom As Long = 1 _
) As Boolean
    Match = RegexDfsMatcher.DfsMatchFrom( _
        matcherState.context, matcherState.captures, regex.bytecode, haystack, matchFrom - 1, _
        stepsLimit:=regex.stepsLimit, _
        multiline:=multiline, _
        dotAll:=dotAll _
    ) <> -1
    matcherState.current = -1
End Function

Public Function GetCapture(ByRef matcherState As MatcherStateTy, ByRef haystack As String, Optional ByVal num As Long = 0) As String
    If num = 0 Then
        With matcherState.captures.entireMatch
            If .Length > 0 Then GetCapture = Mid$(haystack, .start, .Length) Else GetCapture = vbNullString
        End With
    ElseIf num <= matcherState.captures.nNumberedCaptures Then
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
    Dim identifierId As Long
    
    identifierId = RegexBytecode.GetIdentifierId(regex.bytecode, regex.pattern, name)
    If identifierId < 0 Then GetCaptureByName = vbNullString: Exit Function
    GetCaptureByName = GetCapture(matcherState, haystack, matcherState.captures.namedCaptures(identifierId))
End Function

Public Function MatchNext( _
    ByRef matcherState As MatcherStateTy, ByRef regex As RegexTy, ByRef haystack As String _
) As Boolean
    Dim r As Long, oldCurrent As Long
    
    If matcherState.current = -1 Then Exit Function ' end of string reached, return False
    
    oldCurrent = matcherState.current
    r = RegexDfsMatcher.DfsMatchFrom( _
        matcherState.context, matcherState.captures, regex.bytecode, haystack, matcherState.current, _
        stepsLimit:=regex.stepsLimit, _
        multiline:=matcherState.multiline, _
        dotAll:=matcherState.dotAll _
    )
    
    matcherState.current = (r - (oldCurrent = r)) Or matcherState.localMatch
    MatchNext = r <> -1
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

Public Function Split( _
    ByRef regex As RegexTy, _
    ByRef haystack As String, _
    Optional ByVal localMatch As Boolean = False, _
    Optional ByVal multiline As Boolean = False _
) As Collection
    Dim matcherState As MatcherStateTy
    Dim matchedIndex As Long
    Dim notMatched As String
    Dim notMatchedIndex As Long
    Dim colStrings As Collection
    Dim i As Long

    Set colStrings = New Collection
    matcherState.localMatch = localMatch
    matcherState.multiline = multiline
    
    Do While MatchNext(matcherState, regex, haystack)
        matchedIndex = matcherState.captures.entireMatch.start - 1
        notMatched = Mid$(haystack, notMatchedIndex + 1, matchedIndex - notMatchedIndex)
        notMatchedIndex = matchedIndex + matcherState.captures.entireMatch.Length
        colStrings.Add notMatched
        With matcherState.captures
            For i = 0 To .nNumberedCaptures - 1
                If .numberedCaptures(i).start > 0 Then colStrings.Add Mid$(haystack, .numberedCaptures(i).start, .numberedCaptures(i).Length)
            Next i
        End With
    Loop
    colStrings.Add Mid$(haystack, notMatchedIndex + 1, Len$(haystack) - notMatchedIndex)
    
    Set Split = colStrings
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
            StaticStringBuilder.AppendStr resultBuilder, ChrW$(m)
            k = k + parsedFormats.Buffer(k)
        Next
        
        nMatches = nMatches + 1
    Loop
    
    If nMatches = 0 Then
        ' hack to create a zero-length array
        results = Split(vbNullString)
    Else
        ReDim results(0 To nMatches - 1, cola To colb) As String
        m = resultBuilder.Length
        For j = nMatches - 1 To 0 Step -1
            For k = colb To cola Step -1
                mm = AscW(StaticStringBuilder.GetSubstr(resultBuilder, m, 1))
                results(j, k) = StaticStringBuilder.GetSubstr(resultBuilder, mm + 1, m - mm - 1)
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
