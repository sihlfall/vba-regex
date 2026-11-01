Attribute VB_Name = "RegexNfaMatcher"
Option Explicit

Private Type ThreadListEntry
    nxt As Long
    pc As Long
    strOffset As Long
    captureFrame As Long
End Type

Private Type ThreadListsBuffer
    buf() As ThreadListEntry
    Capacity As Long
End Type

Private Type CaptureFrameListBuffer
    buf() As Long
    frameSize As Long
    capacityInFrames As Long
    topUnused As Long
End Type

Private Type NfaBuffers
    thrdLists As ThreadListsBuffer
    thrdSetBuffer As ThreadSetBuffer
    cflBuffer As CaptureFrameListBuffer
End Type

Private Enum NfaMatcherPrivateConstant
    ' We maintain two lists at the same time, an active one and an incative one.
    ' Hence we initially allocate buffers of 2 * DEFAULT_MINIMUM_THREADLIST_CAPACITY
    '   for both lists on the one hand and both sets on the other hand.
    DEFAULT_MINIMUM_THREADLIST_CAPACITY = 16
End Enum

Private Sub NfaAllocateBuffers(ByRef pool As NfaBuffers, ByVal nCapturePoints As Long)
    Dim cflCapacityInEntries As Long, lastStart As Long
    With pool
        With .thrdLists
            .Capacity = 2 * DEFAULT_MINIMUM_THREADLIST_CAPACITY
            ReDim .buf(0 To .Capacity - 1) As ThreadListEntry
        End With
        
        With .thrdSetBuffer
            .Capacity = 2 * DEFAULT_MINIMUM_THREADLIST_CAPACITY
            ReDim .buf(0 To .Capacity - 1) As ThreadSetNode
        End With
    
        With .cflBuffer
            If nCapturePoints = 0 Then
                .capacityInFrames = 1
            Else
                .capacityInFrames = DEFAULT_MINIMUM_THREADLIST_CAPACITY
            End If
            .frameSize = nCapturePoints + 1
            cflCapacityInEntries = .capacityInFrames * .frameSize
            ReDim .buf(0 To cflCapacityInEntries - 1) As Long
            lastStart = cflCapacityInEntries - .frameSize
            .buf(lastStart) = -1& - .frameSize - lastStart
            .topUnused = 0
        End With
    End With
End Sub

Private Sub SwitchToDuplicateCaptureFrame(ByRef cflBuffer As CaptureFrameListBuffer, ByRef captureFrameStart As Long)
    Dim oldCapacityInFrames As Long, newCapacityInEntries As Long, lastStart As Long
    Dim newCaptureFrameStart As Long, i As Long, j As Long, e As Long
    ' assert: pool.cflBuffer.buf(captureFrameStart) > 1
    With cflBuffer
        If .topUnused = -1 Then
            ' increase capacity
            oldCapacityInFrames = .capacityInFrames
            .capacityInFrames = oldCapacityInFrames + oldCapacityInFrames \ 2
            newCapacityInEntries = .capacityInFrames * .frameSize
            ReDim Preserve .buf(0 To newCapacityInEntries - 1) As Long
            lastStart = newCapacityInEntries - .frameSize
            .buf(lastStart) = -1& - .frameSize - lastStart
            .topUnused = oldCapacityInFrames * .frameSize
        End If
        newCaptureFrameStart = .topUnused
        .topUnused = .buf(newCaptureFrameStart) + .frameSize + newCaptureFrameStart
        .buf(newCaptureFrameStart) = 1 ' reference count
        .buf(captureFrameStart) = .buf(captureFrameStart) - 1
        i = newCaptureFrameStart + 1
        e = newCaptureFrameStart + .frameSize - 1
        j = captureFrameStart + 1
        Do While i <= e: .buf(i) = .buf(j): i = i + 1: j = j + 1: Loop
    End With
    captureFrameStart = newCaptureFrameStart
End Sub

Public Function NfaMatch( _
    ByRef outCaptures As RegexMatcherCommon.CapturesTy, _
    ByRef bytecode() As Long, _
    ByRef inputStr As String, _
    Optional ByVal multiline As Boolean = False, _
    Optional ByVal dotAll As Boolean = False _
) As Long
    NfaMatch = NfaMatchFrom( _
        outCaptures, bytecode, inputStr, 0, _
        multiline:=multiline, dotAll:=dotAll _
    )
End Function



Public Function NfaMatchFrom( _
    ByRef outCaptures As RegexMatcherCommon.CapturesTy, _
    ByRef bytecode() As Long, _
    ByRef inputStr As String, _
    ByVal sp As Long, _
    Optional ByVal multiline As Boolean = False, _
    Optional ByVal dotAll As Boolean = False, _
    Optional ByVal matchAt As Boolean = False _
) As Long
    Dim pool As NfaBuffers
    Dim inputBuffer As RegexMatcherCommon.InputBufferTy
    Dim i As Long
    Dim nNamedCaptures As Long, nProperCapturePoints As Long, res As Long
    
    If sp < 0 Or sp > Len(inputStr) Then NfaMatchFrom = -1: Exit Function
    
    nProperCapturePoints = bytecode(0) + 1
    nNamedCaptures = bytecode(1)
    ' Todo: can we postpone this until we know that we will definitely need to fill outCaptures?
    With outCaptures
        .nNumberedCaptures = nProperCapturePoints \ 2 - 1
        .nNamedCaptures = bytecode(1)
        If .nNumberedCaptures > 0 Then ReDim .numberedCaptures(0 To .nNumberedCaptures - 1) As RegexMatcherCommon.StartLengthPair
        If .nNamedCaptures > 0 Then ReDim .namedCaptures(0 To .nNamedCaptures - 1) As Long
    End With
    
    ' Allocates buffers. Also clears capture frame list properly.
    ' Maybe we should separate the capture frame list from the other buffers.
    NfaAllocateBuffers pool, nProperCapturePoints + nNamedCaptures
    
    With pool.cflBuffer
        ' Initialize first capture frame
        .topUnused = .frameSize
        .buf(0) = 1 ' initialize reference count to 1
        For i = 1 To .frameSize - 1: .buf(i) = -1: Next
    End With
    
    ' Set up first thread
    With pool.thrdLists
        With .buf(0)
            .nxt = -1
            ' Match at the supplied offset without the implicit search prefix.
            ' Keep the original input so anchors and boundaries retain their context.
            If matchAt Then
                .pc = bytecode(RegexBytecode.BYTECODE_IDX_START_ACTUAL)
            Else
                .pc = bytecode(RegexBytecode.BYTECODE_IDX_START_WITH_PREFIX)
            End If
            .strOffset = 0
            .captureFrame = 0
        End With
    End With

    RegexMatcherCommon.InitializeInputBuffer inputBuffer, inputStr
    res = NfaRunThreads( _
        outCaptures, pool, bytecode, inputStr, inputBuffer, sp, nProperCapturePoints, _
        nProperCapturePoints + nNamedCaptures, multiline, dotAll _
    )
    If res <> -1 Then
        NfaMatchFrom = res
        Exit Function
    End If

    NfaMatchFrom = -1
End Function

Private Function NfaRunThreads( _
    ByRef outCaptures As RegexMatcherCommon.CapturesTy, _
    ByRef pool As NfaBuffers, _
    ByRef bytecode() As Long, _
    ByRef inputStr As String, _
    ByRef inputBuffer As RegexMatcherCommon.InputBufferTy, _
    ByVal sp As Long, _
    ByVal nProperCapturePoints As Long, _
    ByVal nCapturePoints As Long, _
    ByVal multiline As Boolean, _
    ByVal dotAll As Boolean _
) As Long
    ' state variables -> describe matcher state
    Dim pc As Long, strOffset As Long, currentChar As Long, currentCharCanon As Long
    Dim captureFrame As Long
    Dim nextThread As Long
    Dim nextFreeActiveIndex As Long, nextFreeInactiveIndex As Long
    Dim thrdSetActiveRoot As Long, thrdSetInactiveRoot As Long
    Dim activeSwitch As Long ' 0 or 1
    
    Dim op As Long
    
    ' used only locally
    Dim c1 As Long, c2 As Long
    Dim aa As Long, bb As Long, mm As Long
    Dim n As Long
    Dim r1 As Long, r2 As Long
    Dim b1 As Boolean, b2 As Boolean
    Dim idx As Long

    ' Todo: Remove!!!
    Const spDelta As Long = 1
    
    activeSwitch = 0
    
    captureFrame = -1
    nextThread = 0
    nextFreeActiveIndex = nextThread + 2
    nextFreeInactiveIndex = 1
    
    ' Clear sets
    thrdSetActiveRoot = -1
    thrdSetInactiveRoot = -1
    pool.thrdSetBuffer.nUsed = 0
    
    currentChar = PeekBufferCharCode(inputBuffer, inputStr, sp): currentCharCanon = -2
    sp = sp + 1

    GoTo StartLoop
    
    ' OUTER LOOP: run engine until match or failure
    
    ' Inner loop: Find the next incomplete active thread, which may be the current one.
    ' When found, exit the loop and start processing the thread.
    ' When no unprocessed active threads are available, exit the loop and switch to
    '   inactive threads.
    
ContinueWithActive:
        ' Entry point when we definitely switch to another thread from the list of
        '   active threads (i.e. the current thread is definitely complete).
        
        ' Free capture frame or decrement refcount.
        If pool.cflBuffer.buf(captureFrame) <= 1 Then
            pool.cflBuffer.buf(captureFrame) = _
                pool.cflBuffer.topUnused - pool.cflBuffer.frameSize - captureFrame
            pool.cflBuffer.topUnused = captureFrame
        Else
            pool.cflBuffer.buf(captureFrame) = pool.cflBuffer.buf(captureFrame) - 1
        End If
        
ContinueWithActiveWithCaptureFrameReused:
        ' Entry point when we definitely switch to another thread from the list of
        '   active threads (i.e. the current thread is definitely complete), yet
        '   the capture frame is being reused.
        
        If nextThread = -1 Then GoTo ContinueWithInactive
        
StartLoop:
        With pool.thrdLists.buf(nextThread)
            nextThread = .nxt: pc = .pc: strOffset = .strOffset
            captureFrame = .captureFrame
        End With
        
ContinueWithCurrent:
        ' Entry point when the current thread is not yet complete:
        '   Check whether current thread has entered a state already seen before.
        
        If RegexNfaThreadPool.ThreadSetInsert( _
            pool.thrdSetBuffer, thrdSetActiveRoot, pc, strOffset _
        ) Then GoTo StartHandling
        
        ' State already seen before: Current thread is complete.
        GoTo ContinueWithActive
        
    ' End of inner loop (find next unprocessed active thread).
    
    ' Switch to inactive threads.
    ' When the list of inactive threads is empty, there is definitely no match, and
    '   hence exit.
ContinueWithInactive:
        activeSwitch = 1 - activeSwitch
        If nextFreeInactiveIndex < 2 Then
            ' List empty
            NfaRunThreads = -1
            Exit Function
        End If
        
        ' set marker for inactive
        pool.thrdLists.buf(nextFreeInactiveIndex - 2).nxt = -1
        
        With pool.thrdLists.buf(activeSwitch)
            nextThread = .nxt: pc = .pc: strOffset = .strOffset
            captureFrame = .captureFrame
        End With
        
        ' clear sets
        thrdSetActiveRoot = -1: thrdSetInactiveRoot = -1
        pool.thrdSetBuffer.nUsed = 0

        ' clear now inactive list
        nextFreeActiveIndex = nextFreeInactiveIndex
        nextFreeInactiveIndex = 1 - activeSwitch

        currentChar = PeekBufferCharCode(inputBuffer, inputStr, sp): currentCharCanon = -2
        sp = sp + 1
        
'        Dim dbgStr As String
'        dbgStr = "States in newly active (char: " & currentChar & "): "
'        idx = activeSwitch
'        Do
'            dbgStr = dbgStr & pool.thrdLists.buf(idx).pc & " [ "
'            For c1 = 0 To pool.cflBuffer.frameSize - 1
'                If c1 = 0 Then dbgStr = dbgStr & "*"
'                dbgStr = dbgStr & pool.cflBuffer.buf(pool.thrdLists.buf(idx).captureFrame + c1) & " "
'            Next
'            dbgStr = dbgStr & "] "
'            idx = pool.thrdLists.buf(idx).nxt
'        Loop Until idx = -1
'        Debug.Print dbgStr
        
        ' The first thread at this input position also needs a visited entry.
        GoTo ContinueWithCurrent

StartHandling:
        op = bytecode(pc): pc = pc + 1
        
        ' The following statement depends on the numerical values of our constants!
        ' The following comment line serves as a marker enabling us to locate this
        ' line of code by a find:
        ' __REOP_NUMERICAL_VALUES__
        
        On op And RegexBytecode.REOP_OPCODE_MASK GoTo _
            L_MATCH, L_CHAR, L_DOT, L_RANGES, L_INVRANGES, _
            L_JUMP, L_SPLIT1, L_SPLIT2, L_SAVE, L_SET_NAMED, _
            InternalError, InternalError, InternalError, L_ASSERT_START, L_ASSERT_END, _
            L_ASSERT_WORD_BOUNDARY, L_ASSERT_NOT_WORD_BOUNDARY, InternalError, _
                InternalError, InternalError, _
            InternalError, InternalError, InternalError, InternalError, _
                InternalError, _
            InternalError, InternalError, InternalError, InternalError, InternalError, _
            InternalError, L_FAIL, L_STRING, InternalError, InternalError, _
            InternalError, L_SPLIT2, L_SPLIT1, ContinueWithCurrent
    
        GoTo InternalError
        
L_MATCH:
        If nextFreeInactiveIndex < 2 Then GoTo ExitWithMatch
            
        ' Cleanup captures for remaining threads
        Do While nextThread <> -1
            With pool.thrdLists.buf(nextThread)
                idx = .captureFrame
                nextThread = .nxt
            End With
            If pool.cflBuffer.buf(idx) <= 1 Then
                pool.cflBuffer.buf(idx) = _
                    pool.cflBuffer.topUnused - pool.cflBuffer.frameSize - idx
                pool.cflBuffer.topUnused = idx
            Else
                pool.cflBuffer.buf(idx) = pool.cflBuffer.buf(idx) - 1
            End If
        Loop
        
        If RegexNfaThreadPool.ThreadSetInsert(pool.thrdSetBuffer, thrdSetInactiveRoot, pc - 1, strOffset) Then
            ' append to inactive
            If nextFreeInactiveIndex >= pool.thrdLists.Capacity Then
                pool.thrdLists.Capacity = _
                    pool.thrdLists.Capacity + pool.thrdLists.Capacity \ 2
                ReDim Preserve pool.thrdLists.buf( _
                    0 To pool.thrdLists.Capacity - 1 _
                ) As ThreadListEntry
            End If
            With pool.thrdLists.buf(nextFreeInactiveIndex)
                nextFreeInactiveIndex = nextFreeInactiveIndex + 2
                .nxt = nextFreeInactiveIndex: .pc = pc - 1: .strOffset = strOffset
                .captureFrame = captureFrame
            End With
        End If
        GoTo ContinueWithInactive ' We do not consider lower-priority threads anymore.
        
L_CHAR:
        c1 = bytecode(pc): pc = pc + 1
        
        If op And RegexBytecode.MODIFIER_I_ACTIVE Then
            If currentCharCanon = -2 Then _
                currentCharCanon = RegexUnicodeSupport.ReCanonicalizeChar(currentChar)
            c2 = currentCharCanon
        Else
            c2 = currentChar
        End If
        
        If c1 <> c2 Then GoTo ContinueWithActive
        
        GoTo AppendToInactiveAndContinueWithActive

L_DOT:
        If currentChar = RegexMatcherCommon.MATCHER_ENDOFINPUT Then GoTo ContinueWithActive
        If (op - (dotAll And RegexBytecode.MODIFIER_S_WRITE) And RegexBytecode.MODIFIER_S_ACTIVE) = 0 Then
            If RegexUnicodeSupport.UnicodeIsLineTerminator(currentChar) Then GoTo ContinueWithActive
        End If
        
        GoTo AppendToInactiveAndContinueWithActive
        
L_RANGES:
L_INVRANGES:
        n = bytecode(pc): pc = pc + 1 ' assert: >= 1
        If currentChar = RegexMatcherCommon.MATCHER_ENDOFINPUT Then GoTo ContinueWithActive
        If op And RegexBytecode.MODIFIER_I_ACTIVE Then
            If currentCharCanon = -2 Then
                currentCharCanon = RegexUnicodeSupport.ReCanonicalizeChar(currentChar)
            End If
            c1 = currentCharCanon
        Else
            c1 = currentChar
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
                ' b1 describes whether the match was successful
                If bb >= pc Then b1 = False Else b1 = bytecode(bb - 1) <= c1
                If ((op And RegexBytecode.REOP_OPCODE_MASK) = REOP_RANGES) <> b1 Then _
                    GoTo ContinueWithActive
                GoTo AppendToInactiveAndContinueWithActive
            End If
        Loop
        ' unreachable
        
L_ASSERT_START:
        If sp > 1 Then
            If 0 = (op - (multiline And RegexBytecode.MODIFIER_M_WRITE) And RegexBytecode.MODIFIER_M_ACTIVE) Then GoTo ContinueWithActive
            c1 = PeekBufferCharCode(inputBuffer, inputStr, sp - 2)
            ' E5 Sections 15.10.2.8, 7.3
            If Not RegexUnicodeSupport.UnicodeIsLineTerminator(c1) Then GoTo ContinueWithActive
        End If
        GoTo ContinueWithCurrent

L_ASSERT_END:
        If currentChar <> RegexMatcherCommon.MATCHER_ENDOFINPUT Then
            If 0 = (op - (multiline And RegexBytecode.MODIFIER_M_WRITE) And RegexBytecode.MODIFIER_M_ACTIVE) Then GoTo ContinueWithActive
            If Not RegexUnicodeSupport.UnicodeIsLineTerminator(currentChar) Then GoTo ContinueWithActive
        End If
        GoTo ContinueWithCurrent
        
L_ASSERT_WORD_BOUNDARY:
L_ASSERT_NOT_WORD_BOUNDARY:
        '
        '  E5 Section 15.10.2.6.  The previous and current character
        '  should -not- be canonicalized as they are now.  However,
        '  canonicalization does not affect the result of IsWordChar()
        '  (which depends on Unicode characters never canonicalizing
        '  into ASCII characters) so this does not matter.
        If sp <= 1 Then
            b1 = False  ' not a wordchar
        Else
            c1 = PeekBufferCharCode(inputBuffer, inputStr, sp - 2)
            b1 = UnicodeReIsWordchar(c1)
        End If
        If sp > inputBuffer.Length Then
            b2 = False ' not a wordchar
        Else
            c1 = PeekBufferCharCode(inputBuffer, inputStr, sp - 1)
            b2 = UnicodeReIsWordchar(c1)
        End If

        If ((op And RegexBytecode.REOP_OPCODE_MASK) = REOP_ASSERT_WORD_BOUNDARY) = (b1 = b2) Then GoTo ContinueWithActive
        GoTo ContinueWithCurrent
        
L_JUMP:
        n = bytecode(pc): pc = pc + 1
        pc = pc + n: strOffset = 0
        GoTo ContinueWithCurrent

L_SPLIT1:
        ' split1: prefer direct execution (no jump)
        n = bytecode(pc): pc = pc + 1
        c1 = pc: c2 = pc + n
        GoTo SplitCommon
        
L_SPLIT2:
        ' split2: prefer jump execution (not direct)
        n = bytecode(pc): pc = pc + 1
        c1 = pc + n: c2 = pc
        GoTo SplitCommon
        
SplitCommon:
        ' takes c1 and c2 as "parameters"
        
        ' append to active
        If nextFreeActiveIndex >= pool.thrdLists.Capacity Then
            pool.thrdLists.Capacity = _
                pool.thrdLists.Capacity + pool.thrdLists.Capacity \ 2
            ReDim Preserve pool.thrdLists.buf(0 To pool.thrdLists.Capacity - 1) _
                As ThreadListEntry
        End If
        With pool.thrdLists.buf(nextFreeActiveIndex)
            .nxt = nextThread: .pc = c2: .strOffset = 0
            .captureFrame = captureFrame
            nextThread = nextFreeActiveIndex
            nextFreeActiveIndex = nextFreeActiveIndex + 2
        End With
        ' Increase refcount for capture frame
        pool.cflBuffer.buf(captureFrame) = pool.cflBuffer.buf(captureFrame) + 1
        
        pc = c1
        GoTo ContinueWithCurrent
        
L_SAVE:
        idx = bytecode(pc): pc = pc + 1
        If idx >= nCapturePoints Then GoTo InternalError
        If pool.cflBuffer.buf(captureFrame) > 1 Then
            SwitchToDuplicateCaptureFrame pool.cflBuffer, captureFrame
        End If
        pool.cflBuffer.buf(captureFrame + 1 + idx) = sp - 1
        GoTo ContinueWithCurrent
        
L_SET_NAMED:
        r1 = bytecode(pc): pc = pc + 1
        If r1 >= nCapturePoints Then GoTo InternalError
        r2 = bytecode(pc): pc = pc + 1
        If pool.cflBuffer.buf(captureFrame) > 1 Then
            SwitchToDuplicateCaptureFrame pool.cflBuffer, captureFrame
        End If
        pool.cflBuffer.buf(captureFrame + 1 + nProperCapturePoints + r1) = r2
        GoTo ContinueWithCurrent

L_FAIL:
        GoTo ContinueWithActive
        
L_STRING:
        ' pc points to successor of REOP_STRING, since we had a GetBc call before
        c1 = bytecode(pc + 1 + strOffset)
        
        If op And RegexBytecode.MODIFIER_I_ACTIVE Then
            If currentCharCanon = -2 Then _
                currentCharCanon = RegexUnicodeSupport.ReCanonicalizeChar(currentChar)
            c2 = currentCharCanon
        Else
            c2 = currentChar
        End If
        
        If c1 <> c2 Then GoTo ContinueWithActive
        
        c1 = bytecode(pc) ' string length
        If strOffset + 1 = c1 Then ' reached end of string
            pc = pc + 1 + c1: strOffset = 0
        Else
            pc = pc - 1: strOffset = strOffset + 1
        End If
        GoTo AppendToInactiveAndContinueWithActive

AppendToInactiveAndContinueWithActive:
        If RegexNfaThreadPool.ThreadSetInsert(pool.thrdSetBuffer, thrdSetInactiveRoot, pc, strOffset) Then
            ' append to inactive
            If nextFreeInactiveIndex > UBound(pool.thrdLists.buf) Then
                pool.thrdLists.Capacity = _
                    pool.thrdLists.Capacity + pool.thrdLists.Capacity \ 2
                ReDim Preserve pool.thrdLists.buf(0 To pool.thrdLists.Capacity - 1) _
                    As ThreadListEntry
            End If
            With pool.thrdLists.buf(nextFreeInactiveIndex)
                nextFreeInactiveIndex = nextFreeInactiveIndex + 2
                .nxt = nextFreeInactiveIndex: .pc = pc: .strOffset = strOffset
                .captureFrame = captureFrame
            End With
        Else
            ' No queued thread took ownership of this capture-frame reference.
            GoTo ContinueWithActive
        End If
        GoTo ContinueWithActiveWithCaptureFrameReused

    ' END LOOP

ExitWithMatch:
    CopyCaptures pool.cflBuffer.buf, captureFrame + 1, outCaptures
    NfaRunThreads = outCaptures.entireMatch.start + outCaptures.entireMatch.Length - 1
    Exit Function
    
InternalError:
    Err.Raise RegexErrors.REGEX_ERR_INTERNAL_LOGIC_ERR
End Function

