Attribute VB_Name = "RegexDfsMatcher"
Option Explicit

Public Enum DfsMatcherSharedConstant
    DEFAULT_STEPS_LIMIT = 10000
End Enum

Private Enum DfsMatcherPrivateConstant
    DEFAULT_MINIMUM_THREADSTACK_CAPACITY = 16
    Q_NONE = -2
    DFS_MATCHER_STACK_MINIMUM_CAPACITY = 16
    DFS_ENDOFINPUT = -1
End Enum

Public Type StartLengthPair
    start As Long
    Length As Long
End Type

Public Type CapturesTy
    nNumberedCaptures As Long
    nNamedCaptures As Long
    entireMatch As StartLengthPair
    numberedCaptures() As StartLengthPair
    namedCaptures() As Long
End Type

Private Type DfsMatcherStackFrame
    master As Long
    capturesStackState As Long
    qStackLength As Long
    pc As Long
    sp As Long
    pcLandmark As Long
    spDelta As Long
    q As Long
    qTop As Long
End Type

Private Type DfsMatcherStack
    Buffer() As DfsMatcherStackFrame
    Capacity As Long
    Length As Long
End Type

Public Type DfsMatcherContext
    matcherStack As DfsMatcherStack
    capturesStack As ArrayBuffer.Ty
    qstack As ArrayBuffer.Ty
    nProperCapturePoints As Long
    nCapturePoints As Long ' capture points including slots for named captures
    
    master As Long
    
    capturesRequireCoW As Boolean
    qTop As Long
End Type

Public Function DfsMatch( _
    ByRef outCaptures As CapturesTy, _
    ByRef bytecode() As Long, _
    ByRef inputStr As String, _
    Optional ByVal stepsLimit As Long = DEFAULT_STEPS_LIMIT, _
    Optional ByVal multiline As Boolean = False, _
    Optional ByVal dotAll As Boolean = False _
) As Long
    Dim context As DfsMatcherContext
    DfsMatch = DfsMatchFrom( _
        context, outCaptures, bytecode, inputStr, 0, stepsLimit, _
        multiline:=multiline, dotAll:=dotAll _
    )
End Function

Public Function DfsMatchFrom( _
    ByRef context As DfsMatcherContext, _
    ByRef outCaptures As CapturesTy, _
    ByRef bytecode() As Long, _
    ByRef inputStr As String, _
    ByVal sp As Long, _
    Optional ByVal stepsLimit As Long = DEFAULT_STEPS_LIMIT, _
    Optional ByVal multiline As Boolean = False, _
    Optional ByVal dotAll As Boolean = False _
) As Long
    Dim nNamedCaptures As Long, nProperCapturePoints As Long, res As Long
    
    nProperCapturePoints = bytecode(0) + 1
    nNamedCaptures = bytecode(1)
    ' Todo: can we postpone this until we know that we will definitely need to fill outCaptures?
    With outCaptures
        .nNumberedCaptures = nProperCapturePoints \ 2 - 1
        .nNamedCaptures = bytecode(1)
        If .nNumberedCaptures > 0 Then ReDim .numberedCaptures(0 To .nNumberedCaptures - 1) As StartLengthPair
        If .nNamedCaptures > 0 Then ReDim .namedCaptures(0 To .nNamedCaptures - 1) As Long
    End With
    
    Do While sp <= Len(inputStr)
        InitializeMatcherContext context, nProperCapturePoints, nProperCapturePoints + nNamedCaptures
        res = DfsRunThreads(outCaptures, context, bytecode, inputStr, sp, stepsLimit, multiline, dotAll)
        If res <> -1 Then
            DfsMatchFrom = res
            Exit Function
        End If
        sp = sp + 1
    Loop

    DfsMatchFrom = -1
End Function

Private Function GetBc(ByRef bytecode() As Long, ByRef pc As Long) As Long
    If pc > UBound(bytecode) Then GetBc = REOP_INVALID_OPCODE: Exit Function
    GetBc = bytecode(pc)
    pc = pc + 1
End Function

Private Function GetInputCharCode(ByRef inputStr As String, ByRef sp As Long, ByVal spDelta As Long) As Long
    If sp >= Len(inputStr) Then
        GetInputCharCode = DFS_ENDOFINPUT
    ElseIf sp < 0 Then
        GetInputCharCode = DFS_ENDOFINPUT
    Else
        GetInputCharCode = AscW(Mid$(inputStr, sp + 1, 1)) And &HFFFF& ' sp is 0-based and Mid$ is 1-based
        sp = sp + spDelta
    End If
End Function

Private Function PeekInputCharCode(ByRef inputStr As String, ByRef sp As Long) As Long
    If sp >= Len(inputStr) Then
        PeekInputCharCode = DFS_ENDOFINPUT
    ElseIf sp < 0 Then
        PeekInputCharCode = DFS_ENDOFINPUT
    Else
        PeekInputCharCode = AscW(Mid$(inputStr, sp + 1, 1)) And &HFFFF& ' sp is 0-based and Mid$ is 1-based
    End If
End Function


Private Function UnicodeReIsWordchar(c As Long) As Boolean
    If c <= 90 Then ' Z
        If c >= 65 Then UnicodeReIsWordchar = True: Exit Function ' A
        If c > 57 Then UnicodeReIsWordchar = False: Exit Function ' 9
        If c < 48 Then UnicodeReIsWordchar = False: Exit Function ' 0
        UnicodeReIsWordchar = True: Exit Function
    Else
        If c > 122 Then UnicodeReIsWordchar = False: Exit Function ' z
        If c < 97 Then UnicodeReIsWordchar = c = 95: Exit Function ' a, underscore
        UnicodeReIsWordchar = True: Exit Function
    End If
End Function

Private Sub InitializeMatcherContext(ByRef context As DfsMatcherContext, ByVal nProperCapturePoints As Long, ByVal nCapturePoints As Long)
    With context
        ' Clear stacks
        .matcherStack.Length = 0
        .capturesStack.Length = 0
        .qstack.Length = 0
        
        .nProperCapturePoints = nProperCapturePoints
        .nCapturePoints = nCapturePoints
        ArrayBuffer.AppendFill .capturesStack, nCapturePoints, -1
        
        .master = -1
        .capturesRequireCoW = False
        .qTop = 0
    End With
End Sub

Private Sub PushMatcherStackFrame( _
    ByRef context As DfsMatcherContext, _
    ByVal pc As Long, _
    ByVal sp As Long, _
    ByVal pcLandmark As Long, _
    ByVal spDelta As Long, _
    ByVal q As Long _
)
    With context.matcherStack
        If .Length = .Capacity Then
            ' Increase capacity
            If .Capacity < DFS_MATCHER_STACK_MINIMUM_CAPACITY Then .Capacity = DFS_MATCHER_STACK_MINIMUM_CAPACITY Else .Capacity = .Capacity + .Capacity \ 2
            ReDim Preserve .Buffer(0 To .Capacity - 1) As DfsMatcherStackFrame
        End If
        With .Buffer(.Length)
           .master = context.master
           .capturesStackState = context.capturesStack.Length Or (context.capturesRequireCoW And RegexNumericConstants.LONG_FIRST_BIT)
           .qStackLength = context.qstack.Length
           .pc = pc
           .sp = sp
           .pcLandmark = pcLandmark
           .spDelta = spDelta
           .q = q
           .qTop = context.qTop
        End With
        .Length = .Length + 1
    End With

    context.capturesRequireCoW = True
End Sub

' If current frame is the last remaining frame, returns false
Private Function PopMatcherStackFrame( _
    ByRef context As DfsMatcherContext, _
    ByRef pc As Long, _
    ByRef sp As Long, _
    ByRef pcLandmark As Long, _
    ByRef spDelta As Long, _
    ByRef q As Long _
) As Boolean
    With context.matcherStack
        If .Length = 0 Then
            PopMatcherStackFrame = False
            Exit Function
        End If
    
        .Length = .Length - 1
        With .Buffer(.Length)
            context.master = .master
            context.capturesStack.Length = .capturesStackState And RegexNumericConstants.LONG_ALL_BUT_FIRST_BIT
            context.qstack.Length = .qStackLength
            context.capturesRequireCoW = (.capturesStackState And RegexNumericConstants.LONG_FIRST_BIT) <> 0
            pc = .pc
            sp = .sp
            pcLandmark = .pcLandmark
            spDelta = .spDelta
            q = .q
            context.qTop = .qTop
        End With
    
        PopMatcherStackFrame = True
    End With
End Function

Private Sub ReturnToMasterDiscardCaptures( _
    ByRef context As DfsMatcherContext, _
    ByRef pc As Long, _
    ByRef sp As Long, _
    ByRef pcLandmark As Long, _
    ByRef spDelta As Long, _
    ByRef q As Long _
)
    With context.matcherStack
        .Length = context.master
        With .Buffer(.Length)
            context.master = .master
            context.capturesStack.Length = .capturesStackState And RegexNumericConstants.LONG_ALL_BUT_FIRST_BIT
            context.qstack.Length = .qStackLength
            context.capturesRequireCoW = (.capturesStackState And RegexNumericConstants.LONG_FIRST_BIT) <> 0
            pc = .pc
            sp = .sp
            pcLandmark = .pcLandmark
            spDelta = .spDelta
            q = .q
            context.qTop = .qTop
        End With
    End With
End Sub

Private Sub ReturnToMasterPreserveCaptures( _
    ByRef context As DfsMatcherContext, _
    ByRef pc As Long, _
    ByRef sp As Long, _
    ByRef pcLandmark As Long, _
    ByRef spDelta As Long, _
    ByRef q As Long _
)
    Dim masterCapturesStackLength As Long, i As Long
    
    With context.matcherStack
        .Length = context.master
        With .Buffer(.Length)
            context.master = .master
            masterCapturesStackLength = .capturesStackState And RegexNumericConstants.LONG_ALL_BUT_FIRST_BIT
            context.qstack.Length = .qStackLength
            context.capturesRequireCoW = (.capturesStackState And RegexNumericConstants.LONG_FIRST_BIT) <> 0
            pc = .pc
            sp = .sp
            pcLandmark = .pcLandmark
            spDelta = .spDelta
            q = .q
            context.qTop = .qTop
        End With
    End With
    
    With context.capturesStack
        If .Length = masterCapturesStackLength Then Exit Sub

        If context.capturesRequireCoW Then
            masterCapturesStackLength = masterCapturesStackLength + context.nCapturePoints
            context.capturesRequireCoW = False
            If .Length = masterCapturesStackLength Then Exit Sub
        End If
        
        For i = 1 To context.nCapturePoints
            .Buffer(masterCapturesStackLength - i) = .Buffer(.Length - i)
        Next
        .Length = masterCapturesStackLength
    End With
End Sub

Private Sub ReturnToMasterPreserveAll( _
    ByRef context As DfsMatcherContext, _
    ByRef pc As Long, _
    ByRef sp As Long, _
    ByRef pcLandmark As Long, _
    ByRef spDelta As Long, _
    ByRef q As Long _
)
    Dim masterCapturesStackLength As Long, i As Long
    
    ' We need not adjust the q stack, since the frame we return to has the same top q
    ' element as the frame we are returning from. Hence restoring .qTop is sufficient.
    With context.matcherStack
        .Length = context.master
        With .Buffer(.Length)
            context.master = .master
            masterCapturesStackLength = .capturesStackState And RegexNumericConstants.LONG_ALL_BUT_FIRST_BIT
            context.capturesRequireCoW = (.capturesStackState And RegexNumericConstants.LONG_FIRST_BIT) <> 0
            context.qTop = .qTop
        End With
    End With
    
    With context.capturesStack
        If .Length = masterCapturesStackLength Then Exit Sub

        If context.capturesRequireCoW Then
            masterCapturesStackLength = masterCapturesStackLength + context.nCapturePoints
            context.capturesRequireCoW = False
            If .Length = masterCapturesStackLength Then Exit Sub
        End If
        
        For i = 1 To context.nCapturePoints
            .Buffer(masterCapturesStackLength - i) = .Buffer(.Length - i)
        Next
        .Length = masterCapturesStackLength
    End With
End Sub

Private Sub CopyCaptures(ByRef context As DfsMatcherContext, ByRef captures As CapturesTy)
    Dim i As Long, baseIdx As Long, pt1 As Long, pt2 As Long
    
    With context
        baseIdx = .capturesStack.Length - .nCapturePoints
        pt1 = .capturesStack.Buffer(baseIdx)
        pt2 = .capturesStack.Buffer(baseIdx + 1)
        If pt1 = -1 Then
            With captures.entireMatch: .start = 0: .Length = 0: End With
        ElseIf pt2 < pt1 Then
            With captures.entireMatch: .start = 0: .Length = 0: End With
        Else
            With captures.entireMatch: .start = pt1 + 1: .Length = pt2 - pt1: End With
        End If
            
        For i = 1 To captures.nNumberedCaptures
            pt1 = .capturesStack.Buffer(baseIdx + 2 * i)
            pt2 = .capturesStack.Buffer(baseIdx + 2 * i + 1)
            If pt1 = -1 Then
                With captures.numberedCaptures(i - 1): .start = 0: .Length = 0: End With
            ElseIf pt2 < pt1 Then
                With captures.numberedCaptures(i - 1): .start = 0: .Length = 0: End With
            Else
                With captures.numberedCaptures(i - 1): .start = pt1 + 1: .Length = pt2 - pt1: End With
            End If
        Next
        
        baseIdx = baseIdx + 2 + 2 * captures.nNumberedCaptures
        For i = 0 To captures.nNamedCaptures - 1
            captures.namedCaptures(i) = .capturesStack.Buffer(baseIdx + i)
        Next
    End With
End Sub

Private Sub SetCapturePoint(ByRef context As DfsMatcherContext, ByVal idx As Long, ByVal v As Long)
    With context
        If .capturesRequireCoW Then
            ArrayBuffer.AppendSlice .capturesStack, .capturesStack.Length - .nCapturePoints, .nCapturePoints
            .capturesRequireCoW = False
        End If
        .capturesStack.Buffer(.capturesStack.Length - .nCapturePoints + idx) = v
    End With
End Sub

Private Function GetCapturePoint(ByRef context As DfsMatcherContext, ByVal idx As Long) As Long
    With context
        GetCapturePoint = .capturesStack.Buffer( _
            .capturesStack.Length - .nCapturePoints + idx _
        )
    End With
End Function

Private Sub QStackPush(ByRef context As DfsMatcherContext, ByVal q As Long)
    With context
        ArrayBuffer.AppendThree .qstack, .qTop, .matcherStack.Length, q
        .qTop = .qstack.Length - 1
    End With
End Sub

Private Function QStackPop(ByRef context As DfsMatcherContext) As Long
    With context
        If .qTop = 0 Then
            QStackPop = Q_NONE
            Exit Function
        End If
        
        QStackPop = .qstack.Buffer(.qTop)
        If .qstack.Buffer(.qTop - 1) = .matcherStack.Length Then .qstack.Length = .qstack.Length - 3
        .qTop = .qstack.Buffer(.qTop - 2)
    End With
End Function

Private Function DfsRunThreads( _
    ByRef outCaptures As CapturesTy, _
    ByRef context As DfsMatcherContext, _
    ByRef bytecode() As Long, _
    ByRef inputStr As String, _
    ByVal sp As Long, _
    ByVal stepsLimit As Long, _
    ByVal multiline As Boolean, _
    ByVal dotAll As Boolean _
) As Long
    Dim pc As Long
    
    ' To avoid infinite loops
    Dim pcLandmark As Long
    
    Dim op As Long
    Dim c1 As Long, c2 As Long
    Dim t As Long
    Dim n As Long
    Dim successfulMatch As Boolean
    Dim r1 As Long, r2 As Long
    Dim b1 As Boolean, b2 As Boolean
    Dim aa As Long, bb As Long, mm As Long
    Dim idx As Long, off As Long
    Dim q As Long, qmin As Long, qmax As Long, qq As Long
    Dim qexact As Long
    Dim stepsCount As Long
    Dim spDelta As Long ' 1 when we walk forwards and -1 when we walk backwards
    
    pc = 3 + 3 * bytecode(RegexBytecode.BYTECODE_IDX_N_IDENTIFIERS)
    pcLandmark = -1
    stepsCount = 0
    spDelta = 1
    q = Q_NONE
    
    GoTo ContinueLoopSuccess

    ' BEGIN LOOP
ContinueLoopFail:
        If Not PopMatcherStackFrame(context, pc, sp, pcLandmark, spDelta, q) Then
            DfsRunThreads = -1
            Exit Function
        End If

ContinueLoopSuccess:
        ' TODO: Was a to prevent an infinite loop! Still necessary?
        stepsCount = stepsCount + 1
        If stepsCount >= stepsLimit Then
            DfsRunThreads = -1
            Exit Function
        End If
            
        op = GetBc(bytecode, pc)
    
        ' The following statement depends on the numerical values of our constants!
        ' The following comment line serves as a marker enabling us to locate this line of
        ' code by doing a find:
        ' __REOP_NUMERICAL_VALUES__
        
        On op And RegexBytecode.REOP_OPCODE_MASK GoTo _
            Match, L_CHAR, L_DOT, L_RANGES, L_INVRANGES, _
            L_JUMP, L_SPLIT1, L_SPLIT2, L_SAVE, L_SET_NAMED, _
            L_LOOKPOS, L_LOOKNEG, L_BACKREFERENCE, L_ASSERT_START, L_ASSERT_END, _
            L_ASSERT_WORD_BOUNDARY, L_ASSERT_NOT_WORD_BOUNDARY, _
                L_REPEAT_EXACTLY_INIT, L_REPEAT_EXACTLY_START, L_REPEAT_EXACTLY_END, _
            L_REPEAT_MAX_HUMBLE_INIT, L_REPEAT_MAX_HUMBLE_START, _
                L_REPEAT_MAX_HUMBLE_END, L_REPEAT_GREEDY_MAX_INIT, _
                L_REPEAT_GREEDY_MAX_START, _
            L_REPEAT_GREEDY_MAX_END, L_CHECK_LOOKAHEAD, L_CHECK_LOOKBEHIND, _
                L_END_LOOKPOS, L_END_LOOKNEG, _
            L_COMMIT_POSSESSIVE, ContinueLoopFail
    
        GoTo InternalError
    
L_END_LOOKPOS:
        ' Reached if pattern inside a positive lookahead matched.
        ReturnToMasterPreserveCaptures context, pc, sp, pcLandmark, spDelta, q
        ' Now we are at the REOP_LOOKPOS opcode.
        pc = pc + 1
        n = GetBc(bytecode, pc)
        pc = pc + n
        GoTo ContinueLoopSuccess
        
L_END_LOOKNEG:
        ' Reached if pattern inside a positive lookahead matched.
        ReturnToMasterDiscardCaptures context, pc, sp, pcLandmark, spDelta, q
        GoTo ContinueLoopFail
        
L_CHAR:
        pcLandmark = pc - 1
        c1 = GetBc(bytecode, pc)
        c2 = GetInputCharCode(inputStr, sp, spDelta)
        If op And RegexBytecode.MODIFIER_I_ACTIVE Then
            c2 = RegexUnicodeSupport.ReCanonicalizeChar(c2)
        End If
        If c1 <> c2 Then GoTo ContinueLoopFail
        GoTo ContinueLoopSuccess
        
L_DOT:
        pcLandmark = pc - 1
        c1 = GetInputCharCode(inputStr, sp, spDelta)
        If c1 < 0 Then GoTo ContinueLoopFail
        If op - (dotAll And RegexBytecode.MODIFIER_S_WRITE) And RegexBytecode.MODIFIER_S_ACTIVE Then GoTo ContinueLoopSuccess
        If RegexUnicodeSupport.UnicodeIsLineTerminator(c1) Then GoTo ContinueLoopFail
        GoTo ContinueLoopSuccess
        
L_RANGES:
L_INVRANGES:
        pcLandmark = pc - 1
        n = GetBc(bytecode, pc) ' assert: >= 1
        c1 = GetInputCharCode(inputStr, sp, spDelta)
        If c1 < 0 Then GoTo ContinueLoopFail
        If op And RegexBytecode.MODIFIER_I_ACTIVE Then
            c1 = RegexUnicodeSupport.ReCanonicalizeChar(c1)
        End If
        
        aa = pc - 1
        pc = pc + 2 * n
        bb = pc + 1
        
        ' We are doing a binary search here.
        Do
            mm = aa + 2 * ((bb - aa) \ 4)
            If bytecode(mm) >= c1 Then bb = mm Else aa = mm
            
            If bb - aa = 2 Then
                ' bb is the first upper bound index s.t. ary(bb)>=v
                If bb >= pc Then successfulMatch = False Else successfulMatch = bytecode(bb - 1) <= c1
                Exit Do
            End If
        Loop

        If ((op And RegexBytecode.REOP_OPCODE_MASK) = REOP_RANGES) <> successfulMatch Then GoTo ContinueLoopFail

        GoTo ContinueLoopSuccess
        
L_ASSERT_START:
        If sp <= 0 Then GoTo ContinueLoopSuccess
        If 0 = (op - (multiline And RegexBytecode.MODIFIER_M_WRITE) And RegexBytecode.MODIFIER_M_ACTIVE) Then GoTo ContinueLoopFail
        c1 = PeekInputCharCode(inputStr, sp - (spDelta + 1) \ 2)
        ' E5 Sections 15.10.2.8, 7.3
        If RegexUnicodeSupport.UnicodeIsLineTerminator(c1) Then GoTo ContinueLoopSuccess
        GoTo ContinueLoopFail
        
L_ASSERT_END:
        c1 = PeekInputCharCode(inputStr, sp - (spDelta - 1) \ 2)
        If c1 = DFS_ENDOFINPUT Then GoTo ContinueLoopSuccess
        If 0 = (op - (multiline And RegexBytecode.MODIFIER_M_WRITE) And RegexBytecode.MODIFIER_M_ACTIVE) Then GoTo ContinueLoopFail
        If RegexUnicodeSupport.UnicodeIsLineTerminator(c1) Then GoTo ContinueLoopSuccess
        GoTo ContinueLoopFail
        
L_ASSERT_WORD_BOUNDARY:
L_ASSERT_NOT_WORD_BOUNDARY:
        '
        '  E5 Section 15.10.2.6.  The previous and current character
        '  should -not- be canonicalized as they are now.  However,
        '  canonicalization does not affect the result of IsWordChar()
        '  (which depends on Unicode characters never canonicalizing
        '  into ASCII characters) so this does not matter.
        If sp <= 0 Then
            b1 = False  ' not a wordchar
        Else
            c1 = PeekInputCharCode(inputStr, sp - spDelta)
            b1 = UnicodeReIsWordchar(c1)
        End If
        If sp > Len(inputStr) Then
            b2 = False ' not a wordchar
        Else
            c1 = PeekInputCharCode(inputStr, sp)
            b2 = UnicodeReIsWordchar(c1)
        End If

        If ((op And RegexBytecode.REOP_OPCODE_MASK) = REOP_ASSERT_WORD_BOUNDARY) = (b1 = b2) Then GoTo ContinueLoopFail

        GoTo ContinueLoopSuccess
        
L_JUMP:
        n = GetBc(bytecode, pc)
        If n > 0 Then
            ' forward jump (disjunction)
            pc = pc + n
            GoTo ContinueLoopSuccess
        Else
            ' backward jump (end of loop)
            t = pc + n
            If pcLandmark <= t Then GoTo ContinueLoopSuccess ' empty match
                            
            pc = t: pcLandmark = t
            GoTo ContinueLoopSuccess
        End If
        
L_SPLIT1:
        ' split1: prefer direct execution (no jump)
        n = GetBc(bytecode, pc)
        PushMatcherStackFrame context, pc + n, sp, pcLandmark, spDelta, q
        If op And REOP_FLAG_POSSESSIVE Then context.master = context.matcherStack.Length - 1
        GoTo ContinueLoopSuccess
        
L_SPLIT2:
        ' split2: prefer jump execution (not direct)
        n = GetBc(bytecode, pc)
        PushMatcherStackFrame context, pc, sp, pcLandmark, spDelta, q
        pc = pc + n
        GoTo ContinueLoopSuccess
        
L_REPEAT_EXACTLY_INIT:
        QStackPush context, q
        q = 0
        GoTo ContinueLoopSuccess
        
L_REPEAT_EXACTLY_START:
        pc = pc + 2 ' skip arguments
        If op And REOP_FLAG_POSSESSIVE Then
            PushMatcherStackFrame context, pc, sp, pcLandmark, spDelta, q
            context.master = context.matcherStack.Length - 1
            pc = pc + 1 ' jump over REOP_FAIL
        End If
        GoTo ContinueLoopSuccess
        
L_REPEAT_EXACTLY_END:
        qexact = GetBc(bytecode, pc) ' quantity
        n = GetBc(bytecode, pc) ' offset
        q = q + 1
        If q < qexact Then
            t = pc - n - 3
            If pcLandmark > t Then pcLandmark = t
            pc = t
        Else
            q = QStackPop(context)
        End If
        GoTo ContinueLoopSuccess
        
L_REPEAT_MAX_HUMBLE_INIT:
L_REPEAT_GREEDY_MAX_INIT:
        QStackPush context, q
        q = -1
        GoTo ContinueLoopSuccess
        
L_REPEAT_MAX_HUMBLE_START:
        qmax = GetBc(bytecode, pc)
        n = GetBc(bytecode, pc)
        
        q = q + 1
        If q < qmax Then PushMatcherStackFrame context, pc, sp, pcLandmark, spDelta, q
        
        q = QStackPop(context)
        pc = pc + n
        GoTo ContinueLoopSuccess

L_REPEAT_MAX_HUMBLE_END:
        pc = pc + 1  ' skip first argument: quantity
        n = GetBc(bytecode, pc) ' offset
        t = pc - n - 3
        If pcLandmark <= t Then GoTo ContinueLoopFail ' empty match
        
        pc = t: pcLandmark = t
        GoTo ContinueLoopSuccess
        
L_REPEAT_GREEDY_MAX_START:
        qmax = GetBc(bytecode, pc)
        n = GetBc(bytecode, pc)
        
        q = q + 1
        If q < qmax Then
            qq = QStackPop(context)
            PushMatcherStackFrame context, pc + n, sp, pcLandmark, spDelta, qq
            QStackPush context, qq
            If op And REOP_FLAG_POSSESSIVE Then context.master = context.matcherStack.Length - 1
        Else
            pc = pc + n
            q = QStackPop(context)
        End If
        
        GoTo ContinueLoopSuccess
        
L_REPEAT_GREEDY_MAX_END:
        pc = pc + 1 ' Skip first argument: quantity
        n = GetBc(bytecode, pc) ' offset
        t = pc - n - 3
        If pcLandmark <= t Then GoTo ContinueLoopSuccess ' empty match
        
        pc = t: pcLandmark = t
        GoTo ContinueLoopSuccess
        
L_SAVE:
        idx = GetBc(bytecode, pc)
        If idx >= context.nCapturePoints Then GoTo InternalError
            ' idx is unsigned, < 0 check is not necessary
        SetCapturePoint context, idx, sp
        GoTo ContinueLoopSuccess
        
L_SET_NAMED:
        r1 = GetBc(bytecode, pc)
        If r1 >= context.nCapturePoints Then GoTo InternalError
        r2 = GetBc(bytecode, pc)
        SetCapturePoint context, context.nProperCapturePoints + r1, r2
        GoTo ContinueLoopSuccess
        
L_CHECK_LOOKAHEAD:
        PushMatcherStackFrame context, pc, sp, pcLandmark, spDelta, q
        context.master = context.matcherStack.Length - 1
        pc = pc + 2 ' jump over following REOP_LOOKPOS or REOP_LOOKNEG
        ' When we're moving forward, we are at the correct position. When we're moving backward, we have to step one towards the end.
        sp = sp + (1 - spDelta) \ 2
        spDelta = 1
        ' We could set pcLandmark to -1 again here, but we can be sure that pcLandmark < beginning of lookahead, so we can skip that
        GoTo ContinueLoopSuccess
        
L_CHECK_LOOKBEHIND:
        PushMatcherStackFrame context, pc, sp, pcLandmark, spDelta, q
        context.master = context.matcherStack.Length - 1
        pc = pc + 2 ' jump over following REOP_LOOKPOS or REOP_LOOKNEG
        ' When we're moving backward, we are at the correct position. When we're moving forward, we have to step one towards the beginning.
        sp = sp - (spDelta + 1) \ 2
        spDelta = -1
        ' We could set pcLandmark to -1 again here, but we can be sure that pcLandmark < beginning of lookahead, so we can skip that
        GoTo ContinueLoopSuccess
        
L_LOOKPOS:
        ' This point will only be reached if the pattern inside a negative lookahead/back did not match.
        n = GetBc(bytecode, pc)
        pc = pc + n
        GoTo ContinueLoopFail
        
L_LOOKNEG:
        ' This point will only be reached if the pattern inside a negative lookahead/back did not match.
        n = GetBc(bytecode, pc)
        pc = pc + n
        GoTo ContinueLoopSuccess
        
L_BACKREFERENCE:
        '
        '  Byte matching for back-references would be OK in case-
        '  sensitive matching.  In case-insensitive matching we need
        '  to canonicalize characters, so back-reference matching needs
        '  to be done with codepoints instead.  So, we just decode
        '  everything normally here, too.
        '
        '  Note: back-reference index which is 0 or higher than
        '  NCapturingParens (= number of capturing parens in the
        '  -entire- regexp) is a compile time error.  However, a
        '  backreference referring to a valid capture which has
        '  not matched anything always succeeds!  See E5 Section
        '  15.10.2.9, step 5, sub-step 3.

        pcLandmark = pc - 1
        idx = 2 * GetBc(bytecode, pc) ' backref n -> saved indices [n*2, n*2+1]
        If idx < 2 Then GoTo InternalError
        If idx + 1 >= context.nCapturePoints Then GoTo InternalError
        aa = GetCapturePoint(context, idx)
        bb = GetCapturePoint(context, idx + 1)
        If (aa >= 0) And (bb >= 0) Then
            If spDelta = 1 Then
                off = aa
                Do While off < bb
                    c1 = GetInputCharCode(inputStr, off, 1)
                    c2 = GetInputCharCode(inputStr, sp, 1)
                    ' No need for an explicit c2 < 0 check: because c1 >= 0,
                    ' the comparison will always fail if c2 < 0.
                    If c1 <> c2 Then
                        If 0 = (op And RegexBytecode.MODIFIER_I_ACTIVE) Then GoTo ContinueLoopFail
                        If RegexUnicodeSupport.ReCanonicalizeChar(c1) <> RegexUnicodeSupport.ReCanonicalizeChar(c2) Then GoTo ContinueLoopFail
                    End If
                Loop
            Else
                off = bb - 1
                Do While off >= aa
                    c1 = GetInputCharCode(inputStr, off, -1)
                    c2 = GetInputCharCode(inputStr, sp, -1)
                    ' No need for an explicit c2 < 0 check: because c1 >= 0,
                    ' the comparison will always fail if c2 < 0.
                    If c1 <> c2 Then
                        If 0 = (op And RegexBytecode.MODIFIER_I_ACTIVE) Then GoTo ContinueLoopFail
                        If RegexUnicodeSupport.ReCanonicalizeChar(c1) <> RegexUnicodeSupport.ReCanonicalizeChar(c2) Then GoTo ContinueLoopFail
                    End If
                Loop
            End If
        Else
            ' capture is 'undefined', always matches!
        End If
        GoTo ContinueLoopSuccess
        
L_COMMIT_POSSESSIVE:
        ReturnToMasterPreserveAll context, pc, sp, pcLandmark, spDelta, q
        GoTo ContinueLoopSuccess
        
    ' END LOOP
    
Match:
    CopyCaptures context, outCaptures
    DfsRunThreads = sp
    Exit Function
        
InternalError:
    ' TODO: Raise correct exception
    Err.Raise 3000
End Function

