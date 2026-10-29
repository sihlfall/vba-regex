Attribute VB_Name = "RegexDfsMatcher"
Option Explicit

Public Enum DfsMatcherSharedConstant
    DEFAULT_STEPS_LIMIT = 100000000
End Enum

Private Enum DfsMatcherPrivateConstant
    DEFAULT_MINIMUM_THREADSTACK_CAPACITY = 16
    Q_NONE = -2
    DFS_MATCHER_STACK_MINIMUM_CAPACITY = 16
End Enum

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
    ByRef outCaptures As RegexMatcherCommon.CapturesTy, _
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
    ByRef outCaptures As RegexMatcherCommon.CapturesTy, _
    ByRef bytecode() As Long, _
    ByRef inputStr As String, _
    ByVal sp As Long, _
    Optional ByVal stepsLimit As Long = DEFAULT_STEPS_LIMIT, _
    Optional ByVal multiline As Boolean = False, _
    Optional ByVal dotAll As Boolean = False, _
    Optional ByVal raiseOnLimit As Boolean = False _
) As Long
    Dim nNamedCaptures As Long, nProperCapturePoints As Long, res As Long
    
    If sp < 0 Or sp > Len(inputStr) Then DfsMatchFrom = -1: Exit Function
    
    nProperCapturePoints = bytecode(0) + 1
    nNamedCaptures = bytecode(1)
    ' Todo: can we postpone this until we know that we will definitely need to fill outCaptures?
    With outCaptures
        .nNumberedCaptures = nProperCapturePoints \ 2 - 1
        .nNamedCaptures = bytecode(1)
        If .nNumberedCaptures > 0 Then ReDim .numberedCaptures(0 To .nNumberedCaptures - 1) As RegexMatcherCommon.StartLengthPair
        If .nNamedCaptures > 0 Then ReDim .namedCaptures(0 To .nNamedCaptures - 1) As Long
    End With
    
    InitializeMatcherContext context, nProperCapturePoints, nProperCapturePoints + nNamedCaptures
    res = DfsRunThreads(outCaptures, context, bytecode, inputStr, sp, stepsLimit, multiline, dotAll, raiseOnLimit)
    If res <> -1 Then
        DfsMatchFrom = res
        Exit Function
    End If

    DfsMatchFrom = -1
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
    ByRef outCaptures As RegexMatcherCommon.CapturesTy, _
    ByRef context As DfsMatcherContext, _
    ByRef bytecode() As Long, _
    ByRef inputStr As String, _
    ByVal sp As Long, _
    ByVal stepsLimit As Long, _
    ByVal multiline As Boolean, _
    ByVal dotAll As Boolean, _
    Optional ByVal raiseOnLimit As Boolean = False _
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
    Dim searchDotPc As Long, searchChar As String, literalPc As Long
    Dim searchAnchor As Boolean, searchMultiline As Boolean
    Dim lineSearch As RegexMatcherCommon.LineStartSearchTy
    Dim inputLength As Long, byteOffset As Long, chunk As Long
    Dim inputBuffer As RegexMatcherCommon.InputBufferTy
    Dim referenceBuffer As RegexMatcherCommon.InputBufferTy, referenceBufferInitialized As Boolean
    
    pc = bytecode(RegexBytecode.BYTECODE_IDX_START_WITH_PREFIX)
    pcLandmark = -1
    stepsCount = 0
    spDelta = 1
    q = Q_NONE
    RegexMatcherCommon.InitializeInputBuffer inputBuffer, inputStr
    inputLength = inputBuffer.Length
    searchDotPc = -1
    ' Optimize only the compiler's search prefix, retaining normal backtracking
    ' within each candidate. Leave explicit VM step budgets unchanged.
    If stepsLimit = DEFAULT_STEPS_LIMIT Then
        If RegexMatcherCommon.HasSearchPrefix(bytecode) Then
            literalPc = bytecode(RegexBytecode.BYTECODE_IDX_START_ACTUAL)
            Do While bytecode(literalPc) = RegexBytecode.REOP_SAVE
                literalPc = literalPc + 2
            Loop
            op = bytecode(literalPc)
            If (op And RegexBytecode.REOP_OPCODE_MASK) = RegexBytecode.REOP_ASSERT_START Then
                searchAnchor = True
                searchMultiline = (op - (multiline And RegexBytecode.MODIFIER_M_WRITE) And RegexBytecode.MODIFIER_M_ACTIVE) <> 0
                searchDotPc = pc + 2
            ElseIf (op And RegexBytecode.MODIFIER_I_ACTIVE) = 0 Then
                Select Case op And RegexBytecode.REOP_OPCODE_MASK
                Case RegexBytecode.REOP_CHAR
                    searchChar = ChrW(bytecode(literalPc + 1))
                Case RegexBytecode.REOP_STRING
                    searchChar = ChrW(bytecode(literalPc + 2))
                End Select
                If Len(searchChar) > 0 Then searchDotPc = pc + 2
            End If
        End If
    End If
    
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
            If raiseOnLimit Then Err.Raise RegexErrors.REGEX_ERR_STEP_LIMIT
            DfsRunThreads = -1
            Exit Function
        End If
            
        op = bytecode(pc): pc = pc + 1
    
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
                L_REPEAT_MAX_HUMBLE_END, L_REPEAT_MAX_GREEDY_INIT, _
                L_REPEAT_MAX_GREEDY_START, _
            L_REPEAT_MAX_GREEDY_END, L_CHECK_LOOKAHEAD, L_CHECK_LOOKBEHIND, _
                L_END_LOOKPOS, L_END_LOOKNEG, _
            L_COMMIT_POSSESSIVE, ContinueLoopFail, L_STRING, InternalError, InternalError, _
            InternalError, L_LOOP, L_LOOP, ContinueLoopSuccess
    
        GoTo InternalError
    
L_END_LOOKPOS:
        ' Reached if pattern inside a positive lookahead matched.
        ReturnToMasterPreserveCaptures context, pc, sp, pcLandmark, spDelta, q
        ' Now we are at the REOP_LOOKPOS opcode.
        pc = pc + 1
        n = bytecode(pc): pc = pc + 1
        pc = pc + n
        GoTo ContinueLoopSuccess
        
L_END_LOOKNEG:
        ' Reached if pattern inside a positive lookahead matched.
        ReturnToMasterDiscardCaptures context, pc, sp, pcLandmark, spDelta, q
        GoTo ContinueLoopFail
        
L_CHAR:
        pcLandmark = pc - 1
        c1 = bytecode(pc): pc = pc + 1
        ' Intentionally inline the buffered UTF-16 read: VBA helper calls are
        ' expensive. Keep this sequence in sync with L_DOT and L_RANGES.
        If sp < inputBuffer.windowStart Or sp >= inputBuffer.windowEnd Then
            If sp < 0 Or sp >= inputLength Then
                c2 = RegexMatcherCommon.MATCHER_ENDOFINPUT
                GoTo CharCodeReady
            End If
            RefillInputBuffer inputBuffer, inputStr, sp, spDelta
        End If
        byteOffset = (sp - inputBuffer.windowStart) * 2
        c2 = CLng(inputBuffer.bytes(byteOffset)) Or (CLng(inputBuffer.bytes(byteOffset + 1)) * 256&)
        sp = sp + spDelta
CharCodeReady:
        If op And RegexBytecode.MODIFIER_I_ACTIVE Then
            c2 = RegexUnicodeSupport.ReCanonicalizeChar(c2)
        End If
        If c1 <> c2 Then GoTo ContinueLoopFail
        GoTo ContinueLoopSuccess
        
L_STRING:
        pcLandmark = pc - 1
        n = bytecode(pc): pc = pc + 1
        ' Check the complete span once, then avoid two helper calls per code unit.
        ' Backward execution points at the next character, not at its boundary.
        If spDelta > 0 Then
            If sp < 0 Or n > inputLength - sp Then GoTo ContinueLoopFail
        Else
            If sp >= inputLength Or n > sp + 1 Then GoTo ContinueLoopFail
        End If
        Do While n > 0
            If sp < inputBuffer.windowStart Or sp >= inputBuffer.windowEnd Then RefillInputBuffer inputBuffer, inputStr, sp, spDelta
            If spDelta > 0 Then chunk = inputBuffer.windowEnd - sp Else chunk = sp - inputBuffer.windowStart + 1
            If chunk > n Then chunk = n
            byteOffset = (sp - inputBuffer.windowStart) * 2
            If op And RegexBytecode.MODIFIER_I_ACTIVE Then
                For idx = pc To pc + chunk - 1
                    c2 = CLng(inputBuffer.bytes(byteOffset)) Or (CLng(inputBuffer.bytes(byteOffset + 1)) * 256&)
                    c2 = RegexUnicodeSupport.ReCanonicalizeChar(c2)
                    If bytecode(idx) <> c2 Then GoTo ContinueLoopFail
                    sp = sp + spDelta: byteOffset = byteOffset + 2 * spDelta
                Next
            Else
                For idx = pc To pc + chunk - 1
                    c2 = CLng(inputBuffer.bytes(byteOffset)) Or (CLng(inputBuffer.bytes(byteOffset + 1)) * 256&)
                    If bytecode(idx) <> c2 Then GoTo ContinueLoopFail
                    sp = sp + spDelta: byteOffset = byteOffset + 2 * spDelta
                Next
            End If
            pc = pc + chunk: n = n - chunk
        Loop
        GoTo ContinueLoopSuccess
        
L_DOT:
        pcLandmark = pc - 1
        If pcLandmark = searchDotPc And spDelta > 0 Then
            ' The preferred branch has already failed at sp. Resume at the next
            ' candidate, preserving the prefix's leftmost-first search order.
            If searchAnchor Then
                If Not searchMultiline Then GoTo ContinueLoopFail
                idx = RegexMatcherCommon.NextLineStart(lineSearch, inputStr, sp)
                If idx < 0 Then GoTo ContinueLoopFail
                sp = idx
            Else
                idx = InStr(sp + 2, inputStr, searchChar, vbBinaryCompare)
                If idx = 0 Then GoTo ContinueLoopFail
                sp = idx - 1
            End If
            pc = bytecode(RegexBytecode.BYTECODE_IDX_START_WITH_PREFIX)
            GoTo ContinueLoopSuccess
        End If
        ' Intentionally inline the buffered UTF-16 read: VBA helper calls are
        ' expensive. Keep this sequence in sync with L_CHAR and L_RANGES.
        If sp < inputBuffer.windowStart Or sp >= inputBuffer.windowEnd Then
            If sp < 0 Or sp >= inputLength Then GoTo ContinueLoopFail
            RefillInputBuffer inputBuffer, inputStr, sp, spDelta
        End If
        byteOffset = (sp - inputBuffer.windowStart) * 2
        c1 = CLng(inputBuffer.bytes(byteOffset)) Or (CLng(inputBuffer.bytes(byteOffset + 1)) * 256&)
        sp = sp + spDelta
        If op - (dotAll And RegexBytecode.MODIFIER_S_WRITE) And RegexBytecode.MODIFIER_S_ACTIVE Then GoTo ContinueLoopSuccess
        If RegexUnicodeSupport.UnicodeIsLineTerminator(c1) Then GoTo ContinueLoopFail
        GoTo ContinueLoopSuccess
        
L_RANGES:
L_INVRANGES:
        pcLandmark = pc - 1
        n = bytecode(pc): pc = pc + 1 ' assert: >= 1
        ' Intentionally inline the buffered UTF-16 read: VBA helper calls are
        ' expensive. Keep this sequence in sync with L_CHAR and L_DOT.
        If sp < inputBuffer.windowStart Or sp >= inputBuffer.windowEnd Then
            If sp < 0 Or sp >= inputLength Then GoTo ContinueLoopFail
            RefillInputBuffer inputBuffer, inputStr, sp, spDelta
        End If
        byteOffset = (sp - inputBuffer.windowStart) * 2
        c1 = CLng(inputBuffer.bytes(byteOffset)) Or (CLng(inputBuffer.bytes(byteOffset + 1)) * 256&)
        sp = sp + spDelta
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
        If sp + (1 - spDelta) \ 2 <= 0 Then GoTo ContinueLoopSuccess
        If 0 = (op - (multiline And RegexBytecode.MODIFIER_M_WRITE) And RegexBytecode.MODIFIER_M_ACTIVE) Then GoTo ContinueLoopFail
        c1 = PeekBufferCharCode(inputBuffer, inputStr, sp - (spDelta + 1) \ 2)
        ' E5 Sections 15.10.2.8, 7.3
        If RegexUnicodeSupport.UnicodeIsLineTerminator(c1) Then GoTo ContinueLoopSuccess
        GoTo ContinueLoopFail
        
L_ASSERT_END:
        c1 = PeekBufferCharCode(inputBuffer, inputStr, sp - (spDelta - 1) \ 2)
        If c1 = RegexMatcherCommon.MATCHER_ENDOFINPUT Then GoTo ContinueLoopSuccess
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
        aa = sp + (1 - spDelta) \ 2
        If aa <= 0 Then
            b1 = False  ' not a wordchar
        Else
            b1 = UnicodeReIsWordchar(PeekBufferCharCode(inputBuffer, inputStr, aa - 1))
        End If
        If aa >= Len(inputStr) Then
            b2 = False ' not a wordchar
        Else
            b2 = UnicodeReIsWordchar(PeekBufferCharCode(inputBuffer, inputStr, aa))
        End If

        If ((op And RegexBytecode.REOP_OPCODE_MASK) = REOP_ASSERT_WORD_BOUNDARY) = (b1 = b2) Then GoTo ContinueLoopFail

        GoTo ContinueLoopSuccess
        
L_JUMP:
        n = bytecode(pc): pc = pc + 1
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
        
L_LOOP:
        n = bytecode(pc): pc = pc + 1
        t = pc + n
        ' The body target is explicit. Its preceding entry marker gives nested
        ' repetitions distinct landmarks even when their first atom is shared.
        If pcLandmark < t Then GoTo ContinueLoopSuccess
        pcLandmark = t - 1
        If op = REOP_LOOP_GREEDY Then
            PushMatcherStackFrame context, pc, sp, pcLandmark, spDelta, q
            pc = t
        Else
            PushMatcherStackFrame context, t, sp, pcLandmark, spDelta, q
        End If
        GoTo ContinueLoopSuccess

L_SPLIT1:
        ' split1: prefer direct execution (no jump)
        n = bytecode(pc): pc = pc + 1
        PushMatcherStackFrame context, pc + n, sp, pcLandmark, spDelta, q
        If op And REOP_FLAG_POSSESSIVE Then context.master = context.matcherStack.Length - 1
        GoTo ContinueLoopSuccess
        
L_SPLIT2:
        ' split2: prefer jump execution (not direct)
        n = bytecode(pc): pc = pc + 1
        PushMatcherStackFrame context, pc, sp, pcLandmark, spDelta, q
        pc = pc + n
        GoTo ContinueLoopSuccess
        
L_REPEAT_EXACTLY_INIT:
L_REPEAT_MAX_HUMBLE_INIT:
L_REPEAT_MAX_GREEDY_INIT:
        QStackPush context, q
        q = -1
        GoTo ContinueLoopSuccess
        
L_REPEAT_EXACTLY_START:
        pc = pc + 2 ' skip arguments
        q = q + 1
        If op And REOP_FLAG_POSSESSIVE Then
            PushMatcherStackFrame context, pc, sp, pcLandmark, spDelta, q
            context.master = context.matcherStack.Length - 1
            pc = pc + 1 ' jump over REOP_FAIL
        End If
        GoTo ContinueLoopSuccess
        
L_REPEAT_EXACTLY_END:
        qexact = bytecode(pc): pc = pc + 1 ' quantity
        n = bytecode(pc): pc = pc + 1 ' offset
        If q + 1 < qexact Then
            t = pc - n - 3
            If pcLandmark > t Then pcLandmark = t
            pc = t
        Else
            q = QStackPop(context)
        End If
        GoTo ContinueLoopSuccess
                
L_REPEAT_MAX_HUMBLE_START:
        qmax = bytecode(pc): pc = pc + 1
        n = bytecode(pc): pc = pc + 1
        
        q = q + 1
        If q < qmax Then PushMatcherStackFrame context, pc, sp, pcLandmark, spDelta, q
        
        q = QStackPop(context)
        pc = pc + n
        GoTo ContinueLoopSuccess

L_REPEAT_MAX_HUMBLE_END:
        pc = pc + 1  ' skip first argument: quantity
        n = bytecode(pc): pc = pc + 1 ' offset
        t = pc - n - 3
        If pcLandmark <= t Then GoTo ContinueLoopFail ' empty match
        
        pc = t: pcLandmark = t
        GoTo ContinueLoopSuccess
        
L_REPEAT_MAX_GREEDY_START:
        qmax = bytecode(pc): pc = pc + 1
        n = bytecode(pc): pc = pc + 1
        
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
        
L_REPEAT_MAX_GREEDY_END:
        pc = pc + 1 ' Skip first argument: quantity
        n = bytecode(pc): pc = pc + 1 ' offset
        t = pc - n - 3
        If pcLandmark <= t Then GoTo ContinueLoopSuccess ' empty match
        
        pc = t: pcLandmark = t
        GoTo ContinueLoopSuccess
        
L_SAVE:
        r1 = bytecode(pc): pc = pc + 1
        If r1 >= context.nCapturePoints Then GoTo InternalError
            ' r1 is unsigned, < 0 check is not necessary
        r2 = sp + (1 - spDelta) \ 2
        idx = 0
        GoTo L_SAVE_Common
        
L_SET_NAMED:
        r1 = bytecode(pc): pc = pc + 1
        If r1 >= context.nCapturePoints Then GoTo InternalError
        r2 = bytecode(pc): pc = pc + 1
        idx = -1
        ' fallthrough

L_SAVE_Common:
        With context
            If .capturesRequireCoW Then
                ArrayBuffer.AppendSlice .capturesStack, .capturesStack.Length - .nCapturePoints, .nCapturePoints
                .capturesRequireCoW = False
            End If
            .capturesStack.Buffer( _ 
                .capturesStack.Length + (.nProperCapturePoints And idx) - .nCapturePoints + r1 _ 
            ) = r2
        End With
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
        n = bytecode(pc): pc = pc + 1
        pc = pc + n
        GoTo ContinueLoopFail
        
L_LOOKNEG:
        ' This point will only be reached if the pattern inside a negative lookahead/back did not match.
        n = bytecode(pc): pc = pc + 1
        pc = pc + n
        GoTo ContinueLoopSuccess
        
L_BACKREFERENCE:
        '
        ' Case-sensitive captures are dynamic literals: compare their complete
        ' UTF-16 spans. Case-insensitive captures require regex canonicalization.
        '
        '  Note: back-reference index which is 0 or higher than
        '  NCapturingParens (= number of capturing parens in the
        '  -entire- regexp) is a compile time error.  However, a
        '  backreference referring to a valid capture which has
        '  not matched anything always succeeds!  See E5 Section
        '  15.10.2.9, step 5, sub-step 3.

        pcLandmark = pc - 1
        idx = 2 * bytecode(pc): pc = pc + 1 ' backref n -> saved indices [n*2, n*2+1]
        If idx < 2 Then GoTo InternalError
        If idx + 1 >= context.nCapturePoints Then GoTo InternalError
        off = context.capturesStack.Length - context.nCapturePoints + idx
        aa = context.capturesStack.Buffer(off)
        bb = context.capturesStack.Buffer(off + 1)
        If (aa >= 0) And (bb >= 0) Then
            n = bb - aa
            If n > 0 Then
                If spDelta > 0 Then
                    If sp < 0 Or n > inputLength - sp Then GoTo ContinueLoopFail
                    off = sp + 1
                Else
                    If sp >= inputLength Or n > sp + 1 Then GoTo ContinueLoopFail
                    off = sp - n + 2
                End If
                If (op And RegexBytecode.MODIFIER_I_ACTIVE) = 0 Then
                    If StrComp(Mid$(inputStr, aa + 1, n), Mid$(inputStr, off, n), vbBinaryCompare) <> 0 Then GoTo ContinueLoopFail
                    sp = sp + spDelta * n
                    GoTo ContinueLoopSuccess
                End If
                ' Folded comparisons alternate between distant spans. Keep a
                ' second window, then compare contiguous chunks without calling
                ' the character reader twice for every UTF-16 code unit.
                If Not referenceBufferInitialized Then
                    InitializeInputBuffer referenceBuffer, inputStr
                    referenceBufferInitialized = True
                End If
                If spDelta > 0 Then off = aa Else off = bb - 1
                Do While n > 0
                    If off < referenceBuffer.windowStart Or off >= referenceBuffer.windowEnd Then RefillInputBuffer referenceBuffer, inputStr, off, spDelta
                    If sp < inputBuffer.windowStart Or sp >= inputBuffer.windowEnd Then RefillInputBuffer inputBuffer, inputStr, sp, spDelta
                    If spDelta > 0 Then
                        chunk = referenceBuffer.windowEnd - off
                        t = inputBuffer.windowEnd - sp
                    Else
                        chunk = off - referenceBuffer.windowStart + 1
                        t = sp - inputBuffer.windowStart + 1
                    End If
                    If chunk > t Then chunk = t
                    If chunk > n Then chunk = n
                    aa = (off - referenceBuffer.windowStart) * 2
                    bb = (sp - inputBuffer.windowStart) * 2
                    For idx = 1 To chunk
                        c1 = CLng(referenceBuffer.bytes(aa)) Or (CLng(referenceBuffer.bytes(aa + 1)) * 256&)
                        c2 = CLng(inputBuffer.bytes(bb)) Or (CLng(inputBuffer.bytes(bb + 1)) * 256&)
                        If c1 <> c2 Then
                            If RegexUnicodeSupport.ReCanonicalizeChar(c1) <> RegexUnicodeSupport.ReCanonicalizeChar(c2) Then GoTo ContinueLoopFail
                        End If
                        sp = sp + spDelta: off = off + spDelta
                        aa = aa + 2 * spDelta: bb = bb + 2 * spDelta
                    Next
                    n = n - chunk
                Loop
            End If
        Else
            ' An unmatched capture is an empty backreference.
        End If
        GoTo ContinueLoopSuccess
        
L_COMMIT_POSSESSIVE:
        ReturnToMasterPreserveAll context, pc, sp, pcLandmark, spDelta, q
        GoTo ContinueLoopSuccess
        
    ' END LOOP
    
Match:
    CopyCaptures context.capturesStack.Buffer, _
        context.capturesStack.Length - context.nCapturePoints, outCaptures
    DfsRunThreads = sp
    Exit Function
        
InternalError:
    ' TODO: Raise correct exception
    Err.Raise 3000
End Function

