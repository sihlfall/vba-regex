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
            ' Arrays passed to worksheet functions are 1-based
            .rows = UBound(a, 1): .columns = UBound(a, 2): .t = ARG_ARRAY: Set .rngCells = Nothing
        End With
    Else
        With argInfo
            .rows = 1: .columns = 1: .t = ARG_SCALAR: Set .rngCells = Nothing
        End With
    End If
    TryGetArgumentInfo = True
End Function

Private Function TryDetermineCommonDimensions(ByRef nRows As Long, ByRef nColumns As Long, ByRef argInfo() As ArgumentInfo) As Boolean
    Dim j As Long
    For j = LBound(argInfo) To UBound(argInfo)
        With argInfo(j)
            If .rows > nRows Then
                If nRows <= 1 Then nRows = .rows Else TryDetermineCommonDimensions = False: Exit Function
            End If
            If .columns > nColumns Then
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
        GetStringArgumentValue = ExpectString(a(row + 1, col + 1))
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
        GetBooleanArgumentValue = ExpectBoolean(a(row + 1, col + 1))
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
    Const P_CASE_INSENSITIVE As Long = 2
    Const P_MULTILINE As Long = 3
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
    Const P_CASE_INSENSITIVE As Long = 2
    Const P_MULTILINE As Long = 3
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

