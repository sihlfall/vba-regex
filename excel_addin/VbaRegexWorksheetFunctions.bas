Attribute VB_Name = "VbaRegexWorksheetFunctions"
Option Explicit

Private Const ARG_INVALID As Integer = 0
Private Const ARG_RANGE As Integer = 1
Private Const ARG_ARRAY As Integer = 2
Private Const ARG_SCALAR As Integer = 3

Private varTypeConvertibleToBoolean(0 To 20) As Boolean
Private varTypeConvertibleToBooleanInitialized As Boolean

Private Type ArgumentInfo
    rows As Long
    columns As Long
    t As Integer
    arrayDimensions As Long
    lowerRow As Long
    lowerColumn As Long
    verticallySufficient As Long    ' must be either -1 or 0, as it will be And'ed with the row number
    horizontallySufficient As Long  ' must be either -1 or 0, as it will be And'ed with the column number
    rngCells As Range
End Type

Private Type ExpectedRegex
    regex As StaticRegexSingle.RegexTy
    errCode As Long
    isFine As Boolean
End Type

Private Type ExpectedBoolean
    errCode As Long
    b As Boolean
    isFine As Boolean
End Type

Private Type ExpectedString
    s As String
    errCode As Long
    isFine As Boolean
End Type

Private Const RX_TEST As Long = 0
Private Const RX_MATCH As Long = 1
Private Const RX_REPLACE As Long = 2
Private Const RX_FIND As Long = 3
Private Const RX_LENGTH As Long = 4
Private Const RX_COUNT As Long = 5
Private Const RX_ERR_NUM As Long = vbObjectError + 3100
' Mirrors RegexErrors.REGEX_ERR_STEP_LIMIT in the generated engine.
Private Const RX_STEP_LIMIT As Long = vbObjectError + 3051
Private Const RX_MAX_TEXT As Long = 32767
Private Const RX_MAX_ROWS As Long = 1048576
Private Const RX_MAX_COLUMNS As Long = 16384

Private Type RxFlags
    ignoreCase As Boolean
    multiline As Boolean
    dotAll As Boolean
End Type

Private Type RxCache
    regex As StaticRegexSingle.RegexTy
    format As StaticRegexSingle.FormatTy
    pattern As String
    ignoreCase As Boolean
    regexReady As Boolean
    formatReady As Boolean
End Type

' The original Regex* entry points below retain their original contracts.
' Rx* uses the same dimension handling, with its own argument/result rules.

Private Function TryGetArgumentInfo(ByRef argInfo As ArgumentInfo, ByRef a As Variant) As Boolean
    Dim rng As Range
    If IsObject(a) Then
        If Not (TypeOf a Is Range) Then TryGetArgumentInfo = False: Exit Function
        Set rng = a
        With argInfo
            .rows = rng.rows.Count: .columns = rng.columns.Count: .t = ARG_RANGE: Set .rngCells = rng.cells
        End With
    ElseIf IsArray(a) Then
        With argInfo
            ' Excel may pass literal rows as 1-D arrays; Range.Value2 is 2-D.
            .arrayDimensions = ArgumentArrayDimensions(a)
            If .arrayDimensions < 1 Or .arrayDimensions > 2 Then Exit Function
            .lowerRow = LBound(a, 1)
            If .arrayDimensions = 1 Then
                .rows = 1: .columns = UBound(a, 1) - .lowerRow + 1
            Else
                .lowerColumn = LBound(a, 2)
                .rows = UBound(a, 1) - .lowerRow + 1
                .columns = UBound(a, 2) - .lowerColumn + 1
            End If
            If .rows < 1 Or .columns < 1 Then Exit Function
            .t = ARG_ARRAY: Set .rngCells = Nothing
        End With
    Else
        With argInfo
            .rows = 1: .columns = 1: .t = ARG_SCALAR: Set .rngCells = Nothing
        End With
    End If
    TryGetArgumentInfo = True
End Function

Private Function ArgumentArrayDimensions(ByRef a As Variant) As Long
    Dim dimension As Long, bound As Long
    On Error GoTo Done
    For dimension = 1 To 3
        bound = UBound(a, dimension)
        ArgumentArrayDimensions = dimension
    Next
Done:
End Function

Private Function TryDetermineCommonDimensions(ByRef nRows As Long, ByRef nColumns As Long, ByRef argInfo() As ArgumentInfo) As Boolean
    Dim j As Long
    For j = LBound(argInfo) To UBound(argInfo)
        With argInfo(j)
            If .rows <> nRows And (.rows <> 1 Or nRows = 0) Then
                If nRows <= 1 Then nRows = .rows Else TryDetermineCommonDimensions = False: Exit Function
            End If
            If .columns <> nColumns And (.columns <> 1 Or nColumns = 0) Then
                If nColumns <= 1 Then nColumns = .columns Else TryDetermineCommonDimensions = False: Exit Function
            End If
        End With
    Next
    For j = LBound(argInfo) To UBound(argInfo)
        With argInfo(j)
            .verticallySufficient = (.rows >= nRows)
            .horizontallySufficient = (.columns >= nColumns)
        End With
    Next
    TryDetermineCommonDimensions = True
End Function

Private Function GetStringArgumentValue(ByRef argInfo As ArgumentInfo, ByRef a As Variant, ByVal row As Long, ByVal col As Long) As ExpectedString
    Select Case argInfo.t
    Case ARG_RANGE:
        col = col And argInfo.horizontallySufficient
        row = row And argInfo.verticallySufficient
        GetStringArgumentValue = ExpectString(argInfo.rngCells(row + 1, col + 1))
    Case ARG_ARRAY:
        col = col And argInfo.horizontallySufficient
        row = row And argInfo.verticallySufficient
        If argInfo.arrayDimensions = 1 Then
            GetStringArgumentValue = ExpectString(a(col + argInfo.lowerRow))
        Else
            GetStringArgumentValue = ExpectString(a(row + argInfo.lowerRow, col + argInfo.lowerColumn))
        End If
    Case ARG_SCALAR:
        GetStringArgumentValue = ExpectString(a)
    End Select
End Function

Private Function GetBooleanArgumentValue(ByRef argInfo As ArgumentInfo, ByRef a As Variant, ByVal row As Long, ByVal col As Long) As ExpectedBoolean
    Select Case argInfo.t
    Case ARG_RANGE:
        col = col And argInfo.horizontallySufficient
        row = row And argInfo.verticallySufficient
        GetBooleanArgumentValue = ExpectBoolean(argInfo.rngCells(row + 1, col + 1))
    Case ARG_ARRAY:
        col = col And argInfo.horizontallySufficient
        row = row And argInfo.verticallySufficient
        If argInfo.arrayDimensions = 1 Then
            GetBooleanArgumentValue = ExpectBoolean(a(col + argInfo.lowerRow))
        Else
            GetBooleanArgumentValue = ExpectBoolean(a(row + argInfo.lowerRow, col + argInfo.lowerColumn))
        End If
    Case ARG_SCALAR:
        GetBooleanArgumentValue = ExpectBoolean(a)
    End Select
End Function

Private Function ExpectString(ByRef s As Variant) As ExpectedString
    If VarType(s) = vbString Then
        With ExpectString: .s = s: .isFine = True: End With
    ElseIf VarType(s) = vbEmpty Then
        With ExpectString: .s = vbNullString: .isFine = True: End With
    Else
        With ExpectString
            .isFine = False
            If IsError(s) Then .errCode = CLng(s) Else .errCode = xlErrValue
        End With
    End If
End Function


Private Sub InitializeVarTypeConvertibleToBoolean()
    varTypeConvertibleToBoolean(vbEmpty) = True
    varTypeConvertibleToBoolean(vbInteger) = True
    varTypeConvertibleToBoolean(vbLong) = True
    varTypeConvertibleToBoolean(vbSingle) = True
    varTypeConvertibleToBoolean(vbDouble) = True
    varTypeConvertibleToBoolean(vbCurrency) = True
    varTypeConvertibleToBoolean(vbDate) = True
    varTypeConvertibleToBoolean(vbBoolean) = True
    varTypeConvertibleToBoolean(vbDecimal) = True
    varTypeConvertibleToBoolean(vbByte) = True
    varTypeConvertibleToBoolean(20) = True ' vbLongLong
    varTypeConvertibleToBooleanInitialized = True
End Sub

Private Function ExpectBoolean(ByRef b As Variant) As ExpectedBoolean
    Dim vt As Long
    vt = VarType(b)
    If vt > 20 Then GoTo ValueError
    If vt = vbError Then With ExpectBoolean: .errCode = CLng(b): .isFine = False: End With: Exit Function
    If Not varTypeConvertibleToBooleanInitialized Then InitializeVarTypeConvertibleToBoolean
    If Not varTypeConvertibleToBoolean(vt) Then GoTo ValueError
    With ExpectBoolean: .b = b: .isFine = True: End With
    Exit Function
    
ValueError:
    With ExpectBoolean: .errCode = xlErrValue: .isFine = False: End With
End Function

Private Function Lift_InitializeRegex( _
    ByRef regex As ExpectedRegex, _
    ByRef pattern As ExpectedString, _
    ByRef caseInsensitive As ExpectedBoolean _
)
    If Not (pattern.isFine And caseInsensitive.isFine) Then GoTo ErrorBranch
    
    If StaticRegexSingle.TryInitializeRegex(regex.regex, pattern.s, caseInsensitive:=caseInsensitive.b) Then
        regex.isFine = True
        Lift_InitializeRegex = True
    Else
        regex.isFine = False
        regex.errCode = xlErrValue
    End If
    Exit Function
    
ErrorBranch:
    Const MAX_LONG As Long = &H7FFFFFFF
    With regex
        .isFine = False
        .errCode = MAX_LONG
        
        If Not pattern.isFine Then If pattern.errCode < .errCode Then .errCode = pattern.errCode
        If Not caseInsensitive.isFine Then If caseInsensitive.errCode < .errCode Then .errCode = caseInsensitive.errCode
    End With
    Lift_InitializeRegex = True
End Function

Private Function Lift_StaticRegexTest( _
    ByRef regex As ExpectedRegex, _
    ByRef s As ExpectedString, _
    ByRef multiline As ExpectedBoolean _
) As Variant
    Dim errCode As Long
        
    If Not (regex.isFine And s.isFine And multiline.isFine) Then GoTo ErrorBranch
        
    Lift_StaticRegexTest = StaticRegexSingle.Test(regex.regex, s.s, multiline.b)
    
    Exit Function
    
ErrorBranch:
    Const MAX_LONG As Long = &H7FFFFFFF
    errCode = MAX_LONG
    If Not regex.isFine Then If regex.errCode < errCode Then errCode = regex.errCode
    If Not s.isFine Then If s.errCode < errCode Then errCode = s.errCode
    If Not multiline.isFine Then If multiline.errCode < errCode Then errCode = multiline.errCode
    Lift_StaticRegexTest = CVErr(errCode)
End Function

Private Function Lift_StaticRegexMatchThenJoin( _
    ByRef regex As ExpectedRegex, _
    ByRef s As ExpectedString, _
    ByRef format As ExpectedString, _
    ByRef multiline As ExpectedBoolean _
) As Variant
    Dim errCode As Long
        
    If Not (regex.isFine And s.isFine And format.isFine And multiline.isFine) Then GoTo ErrorBranch
        
    Lift_StaticRegexMatchThenJoin = StaticRegexSingle.MatchThenJoin( _
        regex.regex, _
        s.s, _
        format:=format.s, _
        localMatch:=True, _
        multiline:=multiline.b _
    )
    
    Exit Function
    
ErrorBranch:
    Const MAX_LONG As Long = &H7FFFFFFF
    errCode = MAX_LONG
    If Not regex.isFine Then If regex.errCode < errCode Then errCode = regex.errCode
    If Not s.isFine Then If s.errCode < errCode Then errCode = s.errCode
    If Not format.isFine Then If format.errCode < errCode Then errCode = format.errCode
    If Not multiline.isFine Then If multiline.errCode < errCode Then errCode = multiline.errCode
    Lift_StaticRegexMatchThenJoin = CVErr(errCode)
End Function

Private Function Lift_StaticRegexReplace( _
    ByRef regex As ExpectedRegex, _
    ByRef s As ExpectedString, _
    ByRef replaceBy As ExpectedString, _
    ByRef multiline As ExpectedBoolean _
) As Variant
    Dim errCode As Long
        
    If Not (regex.isFine And s.isFine And replaceBy.isFine And multiline.isFine) Then GoTo ErrorBranch
        
    Lift_StaticRegexReplace = StaticRegexSingle.Replace( _
        regex.regex, _
        replaceBy.s, _
        s.s, _
        multiline:=multiline.b _
    )
    
    Exit Function
    
ErrorBranch:
    Const MAX_LONG As Long = &H7FFFFFFF
    errCode = MAX_LONG
    If Not regex.isFine Then If regex.errCode < errCode Then errCode = regex.errCode
    If Not s.isFine Then If s.errCode < errCode Then errCode = s.errCode
    If Not replaceBy.isFine Then If replaceBy.errCode < errCode Then errCode = replaceBy.errCode
    If Not multiline.isFine Then If multiline.errCode < errCode Then errCode = multiline.errCode
    Lift_StaticRegexReplace = CVErr(errCode)
End Function


Private Sub RegexTestImpl( _
    ByRef b() As Variant, _
    ByVal nRows As Long, _
    ByVal nColumns As Long, _
    ByRef ai_s As ArgumentInfo, ByRef s As Variant, _
    ByRef ai_pattern As ArgumentInfo, ByRef pattern As Variant, _
    ByRef ai_caseInsensitive As ArgumentInfo, ByRef caseInsensitive As Variant, _
    ByRef ai_multiline As ArgumentInfo, ByRef multiline As Variant _
)
    Dim i As Long, j As Long
    Dim horizontallyDistinctRegexs As Boolean, verticallyDistinctRegexs As Boolean
    Dim r As ExpectedRegex
    
    ReDim b(0 To nRows - 1, 0 To nColumns - 1) As Variant
    
    horizontallyDistinctRegexs = ai_pattern.horizontallySufficient Or ai_caseInsensitive.horizontallySufficient
    verticallyDistinctRegexs = ai_pattern.verticallySufficient Or ai_caseInsensitive.verticallySufficient
      
    If horizontallyDistinctRegexs Then
        If verticallyDistinctRegexs Then
            For i = 0 To nRows - 1
                For j = 0 To nColumns - 1
                    Lift_InitializeRegex _
                        r, _
                        GetStringArgumentValue(ai_pattern, pattern, i, j), _
                        GetBooleanArgumentValue(ai_caseInsensitive, caseInsensitive, i, j)
                   
                    b(i, j) = Lift_StaticRegexTest(r, _
                        GetStringArgumentValue(ai_s, s, i, j), _
                        GetBooleanArgumentValue(ai_multiline, multiline, i, j) _
                    )
                Next
            Next
        Else
            For j = 0 To nColumns - 1
                Lift_InitializeRegex _
                    r, _
                    GetStringArgumentValue(ai_pattern, pattern, 0, j), _
                    GetBooleanArgumentValue(ai_caseInsensitive, caseInsensitive, 0, j)
                
                For i = 0 To nRows - 1
                    b(i, j) = Lift_StaticRegexTest(r, _
                        GetStringArgumentValue(ai_s, s, i, j), _
                        GetBooleanArgumentValue(ai_multiline, multiline, i, j) _
                    )
                Next
            Next
        End If
    Else
        If verticallyDistinctRegexs Then
            For i = 0 To nRows - 1
                Lift_InitializeRegex _
                    r, _
                    GetStringArgumentValue(ai_pattern, pattern, i, 0), _
                    GetBooleanArgumentValue(ai_caseInsensitive, caseInsensitive, i, 0) _
                
                For j = 0 To nColumns - 1
                    b(i, j) = Lift_StaticRegexTest(r, _
                        GetStringArgumentValue(ai_s, s, i, j), _
                        GetBooleanArgumentValue(ai_multiline, multiline, i, j) _
                    )
                Next
            Next
        Else
            Lift_InitializeRegex _
                r, _
                GetStringArgumentValue(ai_pattern, pattern, 0, 0), _
                GetBooleanArgumentValue(ai_caseInsensitive, caseInsensitive, 0, 0)
            
            For i = 0 To nRows - 1
                For j = 0 To nColumns - 1
                    b(i, j) = Lift_StaticRegexTest(r, _
                        GetStringArgumentValue(ai_s, s, i, j), _
                        GetBooleanArgumentValue(ai_multiline, multiline, i, j) _
                    )
                Next
            Next
        End If
    End If
End Sub

Public Function RegexTest( _
    ByRef s As Variant, _
    ByRef pattern As Variant, _
    Optional ByRef caseInsensitive As Variant = False, _
    Optional ByRef multiline As Variant = False _
) As Variant
    Const P_S As Long = 0
    Const P_PATTERN As Long = 1
    Const P_CASE_INSENSITIVE As Long = 2
    Const P_MULTILINE As Long = 3
    Dim argInfo(0 To 3) As ArgumentInfo
    
    Dim nRows As Long, nColumns As Long
    
    Dim b() As Variant
    
    On Error GoTo ValueError
    
    If Not TryGetArgumentInfo(argInfo(P_S), s) Then GoTo ValueError
    If Not TryGetArgumentInfo(argInfo(P_PATTERN), pattern) Then GoTo ValueError
    If Not TryGetArgumentInfo(argInfo(P_CASE_INSENSITIVE), caseInsensitive) Then GoTo ValueError
    If Not TryGetArgumentInfo(argInfo(P_MULTILINE), multiline) Then GoTo ValueError
    If Not TryDetermineCommonDimensions(nRows, nColumns, argInfo) Then GoTo ValueError
    
    RegexTestImpl _
        b, _
        nRows, nColumns, _
        argInfo(P_S), s, _
        argInfo(P_PATTERN), pattern, _
        argInfo(P_CASE_INSENSITIVE), caseInsensitive, _
        argInfo(P_MULTILINE), multiline
        
    RegexTest = b
    
    Exit Function
    
ValueError:
    RegexTest = CVErr(xlErrValue)
End Function

Private Sub RegexMatchImpl( _
    ByRef out() As Variant, _
    ByVal nRows As Long, _
    ByVal nColumns As Long, _
    ByRef ai_s As ArgumentInfo, ByRef s As Variant, _
    ByRef ai_pattern As ArgumentInfo, ByRef pattern As Variant, _
    ByRef ai_format As ArgumentInfo, ByRef format As Variant, _
    ByRef ai_caseInsensitive As ArgumentInfo, ByRef caseInsensitive As Variant, _
    ByRef ai_multiline As ArgumentInfo, ByRef multiline As Variant _
)
    Dim i As Long, j As Long
    Dim horizontallyDistinctRegexs As Boolean, verticallyDistinctRegexs As Boolean
    Dim r As ExpectedRegex
    
    ReDim out(0 To nRows - 1, 0 To nColumns - 1) As Variant
    
    horizontallyDistinctRegexs = ai_pattern.horizontallySufficient Or ai_caseInsensitive.horizontallySufficient
    verticallyDistinctRegexs = ai_pattern.verticallySufficient Or ai_caseInsensitive.verticallySufficient

    ' Todo: pre-compile format strings
    
    If horizontallyDistinctRegexs Then
        If verticallyDistinctRegexs Then
            For i = 0 To nRows - 1
                For j = 0 To nColumns - 1
                    Lift_InitializeRegex _
                        r, _
                        GetStringArgumentValue(ai_pattern, pattern, i, j), _
                        GetBooleanArgumentValue(ai_caseInsensitive, caseInsensitive, i, j)
                   
                    out(i, j) = Lift_StaticRegexMatchThenJoin(r, _
                        GetStringArgumentValue(ai_s, s, i, j), _
                        GetStringArgumentValue(ai_format, format, i, j), _
                        GetBooleanArgumentValue(ai_multiline, multiline, i, j) _
                    )
                Next
            Next
        Else
            For j = 0 To nColumns - 1
                Lift_InitializeRegex _
                    r, _
                    GetStringArgumentValue(ai_pattern, pattern, 0, j), _
                    GetBooleanArgumentValue(ai_caseInsensitive, caseInsensitive, 0, j)
                
                For i = 0 To nRows - 1
                    out(i, j) = Lift_StaticRegexMatchThenJoin(r, _
                        GetStringArgumentValue(ai_s, s, i, j), _
                        GetStringArgumentValue(ai_format, format, i, j), _
                        GetBooleanArgumentValue(ai_multiline, multiline, i, j) _
                    )
                Next
            Next
        End If
    Else
        If verticallyDistinctRegexs Then
            For i = 0 To nRows - 1
                Lift_InitializeRegex _
                    r, _
                    GetStringArgumentValue(ai_pattern, pattern, i, 0), _
                    GetBooleanArgumentValue(ai_caseInsensitive, caseInsensitive, i, 0) _
                
                For j = 0 To nColumns - 1
                    out(i, j) = Lift_StaticRegexMatchThenJoin(r, _
                        GetStringArgumentValue(ai_s, s, i, j), _
                        GetStringArgumentValue(ai_format, format, i, j), _
                        GetBooleanArgumentValue(ai_multiline, multiline, i, j) _
                    )
                Next
            Next
        Else
            Lift_InitializeRegex _
                r, _
                GetStringArgumentValue(ai_pattern, pattern, 0, 0), _
                GetBooleanArgumentValue(ai_caseInsensitive, caseInsensitive, 0, 0)
            
            For i = 0 To nRows - 1
                For j = 0 To nColumns - 1
                    out(i, j) = Lift_StaticRegexMatchThenJoin(r, _
                        GetStringArgumentValue(ai_s, s, i, j), _
                        GetStringArgumentValue(ai_format, format, i, j), _
                        GetBooleanArgumentValue(ai_multiline, multiline, i, j) _
                    )
                Next
            Next
        End If
    End If
End Sub


Public Function RegexMatch( _
    ByRef s As Variant, _
    ByRef pattern As Variant, _
    Optional ByRef format As Variant = "$&", _
    Optional ByRef caseInsensitive As Variant = False, _
    Optional ByRef multiline As Variant = False _
) As Variant
    Const P_S As Long = 0
    Const P_PATTERN As Long = 1
    Const P_FORMAT As Long = 2
    Const P_CASE_INSENSITIVE As Long = 3
    Const P_MULTILINE As Long = 4
    Dim argInfo(0 To 4) As ArgumentInfo
    
    Dim nRows As Long, nColumns As Long
    
    Dim out() As Variant
    
    On Error GoTo ValueError
    
    If Not TryGetArgumentInfo(argInfo(P_S), s) Then GoTo ValueError
    If Not TryGetArgumentInfo(argInfo(P_PATTERN), pattern) Then GoTo ValueError
    If Not TryGetArgumentInfo(argInfo(P_FORMAT), format) Then GoTo ValueError
    If Not TryGetArgumentInfo(argInfo(P_CASE_INSENSITIVE), caseInsensitive) Then GoTo ValueError
    If Not TryGetArgumentInfo(argInfo(P_MULTILINE), multiline) Then GoTo ValueError
    If Not TryDetermineCommonDimensions(nRows, nColumns, argInfo) Then GoTo ValueError
    
    RegexMatchImpl _
        out, _
        nRows, nColumns, _
        argInfo(P_S), s, _
        argInfo(P_PATTERN), pattern, _
        argInfo(P_FORMAT), format, _
        argInfo(P_CASE_INSENSITIVE), caseInsensitive, _
        argInfo(P_MULTILINE), multiline
        
    RegexMatch = out
    
    Exit Function
    
ValueError:
    RegexMatch = CVErr(xlErrValue)
End Function

Private Sub RegexReplaceImpl( _
    ByRef out() As Variant, _
    ByVal nRows As Long, _
    ByVal nColumns As Long, _
    ByRef ai_s As ArgumentInfo, ByRef s As Variant, _
    ByRef ai_pattern As ArgumentInfo, ByRef pattern As Variant, _
    ByRef ai_replaceBy As ArgumentInfo, ByRef replaceBy As Variant, _
    ByRef ai_caseInsensitive As ArgumentInfo, ByRef caseInsensitive As Variant, _
    ByRef ai_multiline As ArgumentInfo, ByRef multiline As Variant _
)
    Dim i As Long, j As Long
    Dim horizontallyDistinctRegexs As Boolean, verticallyDistinctRegexs As Boolean
    Dim r As ExpectedRegex
    
    ReDim out(0 To nRows - 1, 0 To nColumns - 1) As Variant
    
    horizontallyDistinctRegexs = ai_pattern.horizontallySufficient Or ai_caseInsensitive.horizontallySufficient
    verticallyDistinctRegexs = ai_pattern.verticallySufficient Or ai_caseInsensitive.verticallySufficient

    ' Todo: pre-compile format strings
    
    If horizontallyDistinctRegexs Then
        If verticallyDistinctRegexs Then
            For i = 0 To nRows - 1
                For j = 0 To nColumns - 1
                    Lift_InitializeRegex _
                        r, _
                        GetStringArgumentValue(ai_pattern, pattern, i, j), _
                        GetBooleanArgumentValue(ai_caseInsensitive, caseInsensitive, i, j)
                   
                    out(i, j) = Lift_StaticRegexReplace(r, _
                        GetStringArgumentValue(ai_s, s, i, j), _
                        GetStringArgumentValue(ai_replaceBy, replaceBy, i, j), _
                        GetBooleanArgumentValue(ai_multiline, multiline, i, j) _
                    )
                Next
            Next
        Else
            For j = 0 To nColumns - 1
                Lift_InitializeRegex _
                    r, _
                    GetStringArgumentValue(ai_pattern, pattern, 0, j), _
                    GetBooleanArgumentValue(ai_caseInsensitive, caseInsensitive, 0, j)
                
                For i = 0 To nRows - 1
                    out(i, j) = Lift_StaticRegexReplace(r, _
                        GetStringArgumentValue(ai_s, s, i, j), _
                        GetStringArgumentValue(ai_replaceBy, replaceBy, i, j), _
                        GetBooleanArgumentValue(ai_multiline, multiline, i, j) _
                    )
                Next
            Next
        End If
    Else
        If verticallyDistinctRegexs Then
            For i = 0 To nRows - 1
                Lift_InitializeRegex _
                    r, _
                    GetStringArgumentValue(ai_pattern, pattern, i, 0), _
                    GetBooleanArgumentValue(ai_caseInsensitive, caseInsensitive, i, 0) _
                
                For j = 0 To nColumns - 1
                    out(i, j) = Lift_StaticRegexReplace(r, _
                        GetStringArgumentValue(ai_s, s, i, j), _
                        GetStringArgumentValue(ai_replaceBy, replaceBy, i, j), _
                        GetBooleanArgumentValue(ai_multiline, multiline, i, j) _
                    )
                Next
            Next
        Else
            Lift_InitializeRegex _
                r, _
                GetStringArgumentValue(ai_pattern, pattern, 0, 0), _
                GetBooleanArgumentValue(ai_caseInsensitive, caseInsensitive, 0, 0)
            
            For i = 0 To nRows - 1
                For j = 0 To nColumns - 1
                    out(i, j) = Lift_StaticRegexReplace(r, _
                        GetStringArgumentValue(ai_s, s, i, j), _
                        GetStringArgumentValue(ai_replaceBy, replaceBy, i, j), _
                        GetBooleanArgumentValue(ai_multiline, multiline, i, j) _
                    )
                Next
            Next
        End If
    End If
End Sub


Public Function RegexReplace( _
    ByRef s As Variant, _
    ByRef pattern As Variant, _
    ByRef replaceBy As Variant, _
    Optional ByRef caseInsensitive As Variant = False, _
    Optional ByRef multiline As Variant = False _
) As Variant
    Const P_S As Long = 0
    Const P_PATTERN As Long = 1
    Const P_REPLACE_BY As Long = 2
    Const P_CASE_INSENSITIVE As Long = 3
    Const P_MULTILINE As Long = 4
    Dim argInfo(0 To 4) As ArgumentInfo
    
    Dim nRows As Long, nColumns As Long
    
    Dim out() As Variant
    
    On Error GoTo ValueError
    
    If Not TryGetArgumentInfo(argInfo(P_S), s) Then GoTo ValueError
    If Not TryGetArgumentInfo(argInfo(P_PATTERN), pattern) Then GoTo ValueError
    If Not TryGetArgumentInfo(argInfo(P_REPLACE_BY), replaceBy) Then GoTo ValueError
    If Not TryGetArgumentInfo(argInfo(P_CASE_INSENSITIVE), caseInsensitive) Then GoTo ValueError
    If Not TryGetArgumentInfo(argInfo(P_MULTILINE), multiline) Then GoTo ValueError
    If Not TryDetermineCommonDimensions(nRows, nColumns, argInfo) Then GoTo ValueError
    
    RegexReplaceImpl _
        out, _
        nRows, nColumns, _
        argInfo(P_S), s, _
        argInfo(P_PATTERN), pattern, _
        argInfo(P_REPLACE_BY), replaceBy, _
        argInfo(P_CASE_INSENSITIVE), caseInsensitive, _
        argInfo(P_MULTILINE), multiline
        
    RegexReplace = out
    
    Exit Function
    
ValueError:
    RegexReplace = CVErr(xlErrValue)
End Function


' One result per broadcast position.
Public Function RxTest(ByRef text As Variant, ByRef pattern As Variant, _
    Optional ByRef flags As Variant = "", Optional ByRef startAt As Variant = 1) As Variant
    RxTest = RxBroadcast(RX_TEST, text, pattern, "", 1, flags, startAt)
End Function

Public Function RxMatch(ByRef text As Variant, ByRef pattern As Variant, _
    Optional ByRef format As Variant = "$&", Optional ByRef occurrence As Variant = 1, _
    Optional ByRef flags As Variant = "", Optional ByRef startAt As Variant = 1) As Variant
    RxMatch = RxBroadcast(RX_MATCH, text, pattern, format, occurrence, flags, startAt)
End Function

Public Function RxReplace(ByRef text As Variant, ByRef pattern As Variant, ByRef replacement As Variant, _
    Optional ByRef occurrence As Variant = 0, Optional ByRef flags As Variant = "", _
    Optional ByRef startAt As Variant = 1) As Variant
    RxReplace = RxBroadcast(RX_REPLACE, text, pattern, replacement, occurrence, flags, startAt)
End Function

Public Function RxFind(ByRef text As Variant, ByRef pattern As Variant, _
    Optional ByRef occurrence As Variant = 1, Optional ByRef flags As Variant = "", _
    Optional ByRef startAt As Variant = 1) As Variant
    RxFind = RxBroadcast(RX_FIND, text, pattern, "", occurrence, flags, startAt)
End Function

Public Function RxLength(ByRef text As Variant, ByRef pattern As Variant, _
    Optional ByRef occurrence As Variant = 1, Optional ByRef flags As Variant = "", _
    Optional ByRef startAt As Variant = 1) As Variant
    RxLength = RxBroadcast(RX_LENGTH, text, pattern, "", occurrence, flags, startAt)
End Function

Public Function RxCount(ByRef text As Variant, ByRef pattern As Variant, _
    Optional ByRef flags As Variant = "", Optional ByRef startAt As Variant = 1) As Variant
    RxCount = RxBroadcast(RX_COUNT, text, pattern, "", 1, flags, startAt)
End Function

Private Function RxBroadcast(ByVal operation As Long, ByRef text As Variant, ByRef pattern As Variant, _
    ByRef format As Variant, ByRef occurrence As Variant, ByRef flags As Variant, ByRef startAt As Variant) As Variant
    On Error GoTo Failed
    Dim ai(0 To 5) As ArgumentInfo, args(0 To 5) As Variant, values(0 To 5) As Variant
    Dim output() As Variant, cache As RxCache, nr As Long, nc As Long, r As Long, c As Long, k As Long
    Dim outer As Long, inner As Long, columnsFirst As Boolean
    ' Preserve Range objects instead of coercing them through their default property.
    RxAssign args(0), text: RxAssign args(1), pattern: RxAssign args(2), format
    RxAssign args(3), occurrence: RxAssign args(4), flags: RxAssign args(5), startAt
    For k = 0 To 5
        If Not TryGetArgumentInfo(ai(k), args(k)) Then GoTo Failed
        If ai(k).t = ARG_RANGE Then If args(k).Areas.Count <> 1 Then GoTo Failed
    Next
    If Not TryDetermineCommonDimensions(nr, nc, ai) Then GoTo Failed
    If nr > RX_MAX_ROWS Or nc > RX_MAX_COLUMNS Then RxBroadcast = CVErr(xlErrNum): Exit Function
    ReDim output(0 To nr - 1, 0 To nc - 1)
    ' Reuse compilation when patterns/flags vary only across columns.
    columnsFirst = (ai(1).rows = 1 And ai(4).rows = 1)
    For outer = 0 To IIf(columnsFirst, nc, nr) - 1
        For inner = 0 To IIf(columnsFirst, nr, nc) - 1
            If columnsFirst Then c = outer: r = inner Else r = outer: c = inner
            For k = 0 To 5
                values(k) = RxArgumentValue(ai(k), args(k), r, c)
            Next
            output(r, c) = RxEvaluate(operation, values, cache)
        Next
    Next
    RxBroadcast = output
    Exit Function
Failed:
    RxBroadcast = CVErr(xlErrValue)
End Function

Private Sub RxAssign(ByRef target As Variant, ByRef source As Variant)
    If IsObject(source) Then Set target = source Else target = source
End Sub

Private Function RxArgumentValue(ByRef ai As ArgumentInfo, ByRef arg As Variant, ByVal row As Long, ByVal col As Long) As Variant
    If ai.rows = 1 Then row = 0
    If ai.columns = 1 Then col = 0
    Select Case ai.t
    Case ARG_RANGE: RxArgumentValue = ai.rngCells(row + 1, col + 1).Value2
    Case ARG_ARRAY
        If ai.arrayDimensions = 1 Then
            RxArgumentValue = arg(col + ai.lowerRow)
        Else
            RxArgumentValue = arg(row + ai.lowerRow, col + ai.lowerColumn)
        End If
    Case ARG_SCALAR: RxArgumentValue = arg
    End Select
End Function

Private Function RxScalar(ByRef arg As Variant) As Variant
    Dim ai As ArgumentInfo
    If Not TryGetArgumentInfo(ai, arg) Then Err.Raise 5
    If ai.rows <> 1 Or ai.columns <> 1 Then Err.Raise 5
    If ai.t = ARG_RANGE Then If arg.Areas.Count <> 1 Then Err.Raise 5
    RxScalar = RxArgumentValue(ai, arg, 0, 0)
End Function

Private Function RxString(ByRef arg As Variant) As String
    If IsEmpty(arg) Then Exit Function
    If VarType(arg) <> vbString Then Err.Raise 5
    RxString = arg
End Function

Private Function RxInteger(ByRef arg As Variant, ByVal default As Long, ByVal minimum As Long, ByVal maximum As Long) As Long
    Dim number As Double
    If IsEmpty(arg) Then
        number = default
    Else
        Select Case VarType(arg)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, 20
            number = CDbl(arg)
        Case Else: Err.Raise 5
        End Select
    End If
    If number < minimum Or number > maximum Then Err.Raise 5
    If number <> Fix(number) Then Err.Raise 5
    RxInteger = CLng(number)
End Function

Private Function RxParseFlags(ByRef arg As Variant) As RxFlags
    Dim flags As String, i As Long
    flags = RxString(arg)
    For i = 1 To Len(flags)
        Select Case Mid$(flags, i, 1)
        Case "i": RxParseFlags.ignoreCase = True
        Case "m": RxParseFlags.multiline = True
        Case "s": RxParseFlags.dotAll = True
        Case Else: Err.Raise 5
        End Select
    Next
End Function

Private Sub RxCompile(ByRef cache As RxCache, ByRef pattern As String, ByRef flags As RxFlags)
    If cache.regexReady Then
        If cache.pattern = pattern And cache.ignoreCase = flags.ignoreCase Then Exit Sub
    End If
    cache.regexReady = False
    cache.formatReady = False
    StaticRegexSingle.InitializeRegex cache.regex, pattern, flags.ignoreCase
    cache.pattern = pattern
    cache.ignoreCase = flags.ignoreCase
    cache.regexReady = True
End Sub

Private Sub RxCompileFormat(ByRef cache As RxCache, ByRef format As String)
    If cache.formatReady Then If cache.format.text = format Then Exit Sub
    cache.formatReady = False
    StaticRegexSingle.InitializeFormat cache.format, cache.regex, format
    cache.formatReady = True
End Sub

Private Sub RxAppend(ByRef result As String, ByVal addition As String)
    If Len(result) + Len(addition) > RX_MAX_TEXT Then Err.Raise RX_ERR_NUM
    result = result & addition
End Sub

Private Function RxFailure(ByVal number As Long) As Variant
    If number = RX_STEP_LIMIT Or number = RX_ERR_NUM Then
        RxFailure = CVErr(xlErrNum)
    Else
        RxFailure = CVErr(xlErrValue)
    End If
End Function

Private Function RxEvaluate(ByVal operation As Long, ByRef values() As Variant, ByRef cache As RxCache) As Variant
    On Error GoTo Failed
    Dim text As String, pattern As String, format As String, result As String
    Dim flags As RxFlags, state As StaticRegexSingle.MatcherStateTy
    Dim occurrence As Long, position As Long, n As Long, lastEnd As Long, k As Long
    For k = 0 To 5
        If IsError(values(k)) Then RxEvaluate = values(k): Exit Function
    Next
    text = RxString(values(0)): pattern = RxString(values(1))
    flags = RxParseFlags(values(4))
    position = RxInteger(values(5), 1, 1, Len(text) + 1)
    occurrence = RxInteger(values(3), IIf(operation = RX_REPLACE, 0, 1), IIf(operation = RX_REPLACE, 0, 1), &H7FFFFFFF)
    RxCompile cache, pattern, flags
    state.current = position - 1
    state.multiline = flags.multiline: state.dotAll = flags.dotAll
    If operation = RX_MATCH Or operation = RX_REPLACE Then
        format = RxString(values(2))
        RxCompileFormat cache, format
    End If
    lastEnd = 1
    Do While StaticRegexSingle.MatchNext(state, cache.regex, text, raiseOnLimit:=True)
        n = n + 1
        Select Case operation
        Case RX_TEST
            RxEvaluate = True: Exit Function
        Case RX_MATCH
            If n = occurrence Then
                result = StaticRegexSingle.FormatMatch(state, text, cache.format)
                If Len(result) > RX_MAX_TEXT Then Err.Raise RX_ERR_NUM
                RxEvaluate = result: Exit Function
            End If
        Case RX_FIND, RX_LENGTH
            If n = occurrence Then
                If operation = RX_FIND Then
                    RxEvaluate = state.captures.entireMatch.start
                Else
                    RxEvaluate = state.captures.entireMatch.Length
                End If
                Exit Function
            End If
        Case RX_REPLACE
            If occurrence = 0 Or n = occurrence Then
                RxAppend result, Mid$(text, lastEnd, state.captures.entireMatch.start - lastEnd)
                RxAppend result, StaticRegexSingle.FormatMatch(state, text, cache.format)
                lastEnd = state.captures.entireMatch.start + state.captures.entireMatch.Length
                If occurrence > 0 Then Exit Do
            End If
        End Select
    Loop
    Select Case operation
    Case RX_TEST: RxEvaluate = False
    Case RX_MATCH, RX_FIND, RX_LENGTH: RxEvaluate = CVErr(xlErrNA)
    Case RX_COUNT: RxEvaluate = n
    Case RX_REPLACE
        RxAppend result, Mid$(text, lastEnd)
        RxEvaluate = result
    End Select
    Exit Function
Failed:
    RxEvaluate = RxFailure(Err.Number)
End Function

' Variable-size results: one input text, one pattern, scalar options.
Public Function RxMatches(ByRef text As Variant, ByRef pattern As Variant, _
    Optional ByRef format As Variant = "$&", Optional ByRef flags As Variant = "", _
    Optional ByRef startAt As Variant = 1) As Variant
    On Error GoTo Failed
    Dim values(0 To 3) As Variant, ai As ArgumentInfo, k As Long, j As Long
    Dim s As String, p As String, f As String, opts As RxFlags, position As Long
    Dim regex As StaticRegexSingle.RegexTy, state As StaticRegexSingle.MatcherStateTy
    Dim formats() As StaticRegexSingle.FormatTy, v As Variant, rows As Collection
    Dim row() As Variant, output() As Variant, nr As Long, nc As Long
    values(0) = RxScalar(text): values(1) = RxScalar(pattern)
    values(2) = RxScalar(flags): values(3) = RxScalar(startAt)
    For k = 0 To 3
        If IsError(values(k)) Then RxMatches = values(k): Exit Function
    Next
    s = RxString(values(0)): p = RxString(values(1))
    opts = RxParseFlags(values(2))
    position = RxInteger(values(3), 1, 1, Len(s) + 1)
    If Not TryGetArgumentInfo(ai, format) Then Err.Raise 5
    If ai.rows <> 1 Or ai.columns < 1 Then Err.Raise 5
    If ai.t = ARG_RANGE Then If format.Areas.Count <> 1 Then Err.Raise 5
    nc = ai.columns
    If nc > RX_MAX_COLUMNS Then Err.Raise RX_ERR_NUM
    StaticRegexSingle.InitializeRegex regex, p, opts.ignoreCase
    ReDim formats(0 To nc - 1)
    For j = 0 To nc - 1
        v = RxArgumentValue(ai, format, 0, j)
        If IsError(v) Then RxMatches = v: Exit Function
        f = RxString(v)
        StaticRegexSingle.InitializeFormat formats(j), regex, f
    Next
    state.current = position - 1
    state.multiline = opts.multiline: state.dotAll = opts.dotAll
    Set rows = New Collection
    Do While StaticRegexSingle.MatchNext(state, regex, s, raiseOnLimit:=True)
        If rows.Count >= RX_MAX_ROWS Then Err.Raise RX_ERR_NUM
        ReDim row(0 To nc - 1)
        For j = 0 To nc - 1
            row(j) = StaticRegexSingle.FormatMatch(state, s, formats(j))
            If Len(row(j)) > RX_MAX_TEXT Then Err.Raise RX_ERR_NUM
        Next
        rows.Add row
    Loop
    nr = rows.Count
    If nr = 0 Then RxMatches = CVErr(xlErrNA): Exit Function
    ReDim output(0 To nr - 1, 0 To nc - 1)
    For k = 0 To nr - 1
        v = rows(k + 1)
        For j = 0 To nc - 1
            output(k, j) = v(j)
        Next
    Next
    RxMatches = output
    Exit Function
Failed:
    RxMatches = RxFailure(Err.Number)
End Function

Public Function RxSplit(ByRef text As Variant, ByRef pattern As Variant, _
    Optional ByRef ignoreEmpty As Variant = False, Optional ByRef flags As Variant = "") As Variant
    On Error GoTo Failed
    Dim values(0 To 3) As Variant, k As Long, s As String, p As String, piece As String
    Dim opts As RxFlags, ignore As ExpectedBoolean, lastEnd As Long, pieces As Collection
    Dim regex As StaticRegexSingle.RegexTy, state As StaticRegexSingle.MatcherStateTy, output() As Variant
    values(0) = RxScalar(text): values(1) = RxScalar(pattern)
    values(2) = RxScalar(ignoreEmpty): values(3) = RxScalar(flags)
    For k = 0 To 3
        If IsError(values(k)) Then RxSplit = values(k): Exit Function
    Next
    s = RxString(values(0)): p = RxString(values(1))
    opts = RxParseFlags(values(3))
    ignore = ExpectBoolean(values(2))
    If Not ignore.isFine Then Err.Raise 5
    StaticRegexSingle.InitializeRegex regex, p, opts.ignoreCase
    state.multiline = opts.multiline: state.dotAll = opts.dotAll
    lastEnd = 1: Set pieces = New Collection
    Do While StaticRegexSingle.MatchNext(state, regex, s, raiseOnLimit:=True)
        piece = Mid$(s, lastEnd, state.captures.entireMatch.start - lastEnd)
        RxAddPiece pieces, piece, ignore.b
        lastEnd = state.captures.entireMatch.start + state.captures.entireMatch.Length
    Loop
    RxAddPiece pieces, Mid$(s, lastEnd), ignore.b
    If pieces.Count = 0 Then RxSplit = CVErr(xlErrNA): Exit Function
    ReDim output(0 To 0, 0 To pieces.Count - 1)
    For k = 1 To pieces.Count
        output(0, k - 1) = pieces(k)
    Next
    RxSplit = output
    Exit Function
Failed:
    RxSplit = RxFailure(Err.Number)
End Function

Private Sub RxAddPiece(ByRef pieces As Collection, ByVal piece As String, ByVal ignore As Boolean)
    If ignore And Len(piece) = 0 Then Exit Sub
    If pieces.Count >= RX_MAX_COLUMNS Or Len(piece) > RX_MAX_TEXT Then Err.Raise RX_ERR_NUM
    pieces.Add piece
End Sub

' Called by Workbook_Open in built add-ins, and by the test runner.
Public Sub RegisterRxFunctions()
    Dim prefix As String, wasAddin As Boolean, number As Long, description As String
    On Error GoTo Failed
    wasAddin = ThisWorkbook.IsAddin
    ' MacroOptions cannot edit a workbook while Excel marks it as an add-in.
    ThisWorkbook.IsAddin = False
    prefix = "'" & ThisWorkbook.Name & "'!"
    Application.MacroOptions Macro:=prefix & "RxTest", Category:="VbaRegex", _
        Description:="Tests whether text contains a regex match. Broadcasts arrays and ranges.", _
        ArgumentDescriptions:=Array("Text or range of text.", "Regex pattern.", "Optional flags: i, m, s. Default empty.", "1-based UTF-16 start position. Default 1.")
    Application.MacroOptions Macro:=prefix & "RxMatch", Category:="VbaRegex", _
        Description:="Returns a formatted match, or #N/A. Broadcasts arrays and ranges.", _
        ArgumentDescriptions:=Array("Text or range of text.", "Regex pattern.", "Capture format. Default $&.", "Positive match number. Default 1.", "Optional flags: i, m, s.", "1-based UTF-16 start position. Default 1.")
    Application.MacroOptions Macro:=prefix & "RxFind", Category:="VbaRegex", _
        Description:="Returns the 1-based position of a regex match, or #N/A. Broadcasts arrays and ranges.", _
        ArgumentDescriptions:=Array("Text or range of text.", "Regex pattern.", "Positive match number. Default 1.", "Optional flags: i, m, s.", "1-based UTF-16 start position. Default 1.")
    Application.MacroOptions Macro:=prefix & "RxLength", Category:="VbaRegex", _
        Description:="Returns a regex match's length in UTF-16 code units, or #N/A. Broadcasts arrays and ranges.", _
        ArgumentDescriptions:=Array("Text or range of text.", "Regex pattern.", "Positive match number. Default 1.", "Optional flags: i, m, s.", "1-based UTF-16 start position. Default 1.")
    Application.MacroOptions Macro:=prefix & "RxCount", Category:="VbaRegex", _
        Description:="Counts nonoverlapping regex matches. Returns zero for no matches. Broadcasts arrays and ranges.", _
        ArgumentDescriptions:=Array("Text or range of text.", "Regex pattern.", "Optional flags: i, m, s.", "1-based UTF-16 start position. Default 1.")
    Application.MacroOptions Macro:=prefix & "RxMatches", Category:="VbaRegex", _
        Description:="Returns all matches down rows; a row of formats produces columns.", _
        ArgumentDescriptions:=Array("One text or single cell.", "One regex pattern.", "Capture format or one row of formats. Default $&.", "Optional flags: i, m, s.", "1-based UTF-16 start position. Default 1.")
    Application.MacroOptions Macro:=prefix & "RxReplace", Category:="VbaRegex", _
        Description:="Replaces regex matches. Broadcasts arrays and ranges.", _
        ArgumentDescriptions:=Array("Text or range of text.", "Regex pattern.", "Replacement capture format.", "0 for all; positive match number for one. Default 0.", "Optional flags: i, m, s.", "1-based UTF-16 start position. Default 1.")
    Application.MacroOptions Macro:=prefix & "RxSplit", Category:="VbaRegex", _
        Description:="Splits one text across columns. Delimiter captures are excluded.", _
        ArgumentDescriptions:=Array("One text or single cell.", "Regex delimiter.", "TRUE discards empty fields. Default FALSE.", "Optional flags: i, m, s.")
    ThisWorkbook.IsAddin = wasAddin
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    ThisWorkbook.IsAddin = wasAddin
    Err.Raise number, "RegisterRxFunctions", description
End Sub
