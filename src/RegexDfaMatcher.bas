Attribute VB_Name = "RegexDfaMatcher"
Option Explicit

Private Enum DfaMatcherPrivateConstant
    ' Disjoint categories of the next character. Conditions are sets of these
    ' categories, so subtraction can retain just the contexts not yet visited.
    DFA_ASSERT_EXPECT_WORDCHAR = 1
    DFA_ASSERT_EXPECT_OTHER = 2
    DFA_ASSERT_EXPECT_LINE_TERMINATOR = 4
    DFA_ASSERT_EXPECT_EOI = 8
    DFA_ASSERT_EXPECT_NOT_WORDCHAR = 14
    DFA_ASSERT_ALL = 15
    DFA_STATE_TABLE_CAPACITY = 1024
    DFA_TRANSITION_TABLE_CAPACITY = 4096
    DFA_STATE_BUFFER_LIMIT = 65536 ' Longs; a single working state may exceed this
End Enum

Public Enum DfaMode
    DFA_MODE_NORMAL = 0
    DFA_MODE_FINISH_EAGERLY = 1
    DFA_MODE_EXTENSIVE = 2
End Enum

Private Type TransitionTableEntry
    hashRank As Long
    charCode As Long
    sourceStateId As Long
    targetStateId As Long
End Type

Private Type StateHashTableEntry
    hash As Long
    statePtr As Long
    context As Long
End Type

Public Type StateStructure
    literal As RegexMatcherCommon.LiteralMatcherTy
    tablesInitialized As Boolean
    ' Each serialized state is [stateId, wordCount, words...]. Words contain
    ' instruction positions, plus STRING offsets or pending assertion operands.
    stateListsBuffer As ArrayBuffer.Ty
    transitionTablePrime As Long
    stateHashTableMask As Long
    transitionTableCapacity As Long
    transitionTableCount As Long
    stateListPtrs() As Long ' used for mapping stateId's to stateListPtr's
    stateHashTable() As StateHashTableEntry ' used determining whether a state already exists
    transitionTable() As TransitionTableEntry
    nextFreeStateId As Long
    multiline As Boolean
    dotAll As Boolean
End Type

Private Enum ActionConstant
    ACTION_FINISH = &H80000000
    ACTION_SET_LAST_MATCH_PREV = &H40000000
    ACTION_SET_LAST_MATCH_CUR = &H20000000
    ACTION_MASK = &HE0000000
End Enum

Public Function DfaStateStructureInitialize( _
    ByRef stateStruct As StateStructure, _
    Optional ByVal multiline As Boolean = False, _
    Optional ByVal dotAll As Boolean = False _
)
    ' A structure belongs to one compiled regex. Initialize it again before using
    ' different bytecode or changing the global modifiers.
    With stateStruct
        .transitionTablePrime = 4093
        .transitionTableCapacity = DFA_TRANSITION_TABLE_CAPACITY
        .stateHashTableMask = DFA_STATE_TABLE_CAPACITY - 1
        .stateListsBuffer.Length = 0
        .nextFreeStateId = 1
        .transitionTableCount = 0
        .literal.isInitialized = False
        .literal.isLiteral = False
        .literal.text = vbNullString
        .tablesInitialized = False
        .dotAll = dotAll
        .multiline = multiline
    End With
End Function

Private Sub InitializeTables(ByRef stateStruct As StateStructure)
    With stateStruct
        ReDim .stateHashTable(0 To .stateHashTableMask) As StateHashTableEntry
        ReDim .stateListPtrs(0 To .stateHashTableMask + 1) As Long
        ReDim .transitionTable(0 To .transitionTableCapacity - 1) As TransitionTableEntry
        .tablesInitialized = True
    End With
End Sub

Private Sub ClearCache(ByRef stateStruct As StateStructure)
    Dim i As Long
    With stateStruct
        .nextFreeStateId = 1
        .transitionTableCount = 0
        .stateListsBuffer.Length = 0
        ' Only occupancy markers need clearing. Keep allocations in place; VB6
        ' may still hold references to these arrays during the active scan.
        For i = 0 To .stateHashTableMask
            .stateHashTable(i).statePtr = 0
        Next
        For i = 0 To .transitionTableCapacity - 1
            .transitionTable(i).sourceStateId = 0
        Next
    End With
End Sub

Private Sub ClearCacheKeepingState(ByRef stateStruct As StateStructure, ByRef p As Long)
    Dim i As Long, n As Long
    With stateStruct.stateListsBuffer
        n = .Length - p
        For i = 0 To n - 1
            .Buffer(i) = .Buffer(p + i)
        Next
    End With
    ClearCache stateStruct
    stateStruct.stateListsBuffer.Length = n
    p = 0
End Sub

Public Function DfaFindMatch( _
    ByRef outEntireMatch As RegexMatcherCommon.StartLengthPair, _
    ByRef stateStruct As StateStructure, _
    ByRef bytecode() As Long, _
    ByRef inputStr As String, _
    ByVal matchFrom As Long, _
    ByVal finishEagerly As Boolean _
) As Boolean
    ' Eager mode stops as soon as a branch accepts. Normal mode observes
    ' alternative and quantifier priority; both recover the leftmost start for
    ' the chosen endpoint.
    Dim endMatch As Long, lastMatch As Long, mode As DfaMode, pc As Long, anchorOp As Long
    Dim inputBuffer As RegexMatcherCommon.InputBufferTy

    If matchFrom < 1 Then GoTo NoMatch
    If matchFrom - 1 > Len(inputStr) Then GoTo NoMatch
    If Not stateStruct.literal.isInitialized Then _
        RegexMatcherCommon.InitializeLiteralMatcher stateStruct.literal, bytecode
    If stateStruct.literal.isLiteral Then
        DfaFindMatch = RegexMatcherCommon.LiteralFindMatch( _
            outEntireMatch, stateStruct.literal, inputStr, matchFrom, stateStruct.multiline)
        Exit Function
    End If
    pc = bytecode(RegexBytecode.BYTECODE_IDX_START_WITH_PREFIX)
    anchorOp = RegexMatcherCommon.LeadingStartAnchor(bytecode)
    If anchorOp <> 0 Then
        If (anchorOp - (stateStruct.multiline And RegexBytecode.MODIFIER_M_WRITE) And RegexBytecode.MODIFIER_M_ACTIVE) = 0 Then
            If matchFrom <> 1 Then GoTo NoMatch
            pc = bytecode(RegexBytecode.BYTECODE_IDX_START_ACTUAL)
        End If
    End If
    RegexMatcherCommon.InitializeInputBuffer inputBuffer, inputStr
    If finishEagerly Then mode = DFA_MODE_FINISH_EAGERLY Else mode = DFA_MODE_NORMAL

    DfaRunThreads _
        lastMatch, _
        stateStruct, bytecode, _
        pc, inputStr, inputBuffer, matchFrom - 1, 1, Len(inputStr) + 1, mode
    
    If lastMatch = -1 Then GoTo NoMatch
    
    endMatch = lastMatch
    
    DfaRunThreads _
        lastMatch, _
        stateStruct, bytecode, _
        bytecode(RegexBytecode.BYTECODE_IDX_START_DFA_BW_RUN), _
        inputStr, inputBuffer, endMatch, -1, matchFrom - 2, DFA_MODE_EXTENSIVE

    If lastMatch < matchFrom - 1 Then GoTo NoMatch
    
    With outEntireMatch
        .start = lastMatch + 1: .Length = endMatch - lastMatch
    End With
    DfaFindMatch = True
    Exit Function
    
NoMatch:
    With outEntireMatch: .start = 0: .Length = 0: End With
    DfaFindMatch = False
End Function

Public Function DfaMatch( _
    ByRef bytecode() As Long, _
    ByRef inputStr As String, _
    ByVal finishEagerly As Boolean, _
    Optional ByVal multiline As Boolean = False, _
    Optional ByVal dotAll As Boolean = False _
) As Long
    DfaMatch = DfaMatchFrom( _
        bytecode, inputStr, 0, False, finishEagerly, _
        multiline:=multiline, dotAll:=dotAll _
    )
End Function



Public Function DfaMatchFrom( _
    ByRef bytecode() As Long, _
    ByRef inputStr As String, _
    ByVal sp As Long, _
    ByVal bwMatch As Boolean, _
    ByVal finishEagerly As Boolean, _
    Optional ByVal multiline As Boolean = False, _
    Optional ByVal dotAll As Boolean = False _
) As Long
    Dim stateStruct As StateStructure
    Dim pc As Long, spDelta As Long, spLast As Long, mode As DfaMode
    Dim lastMatch As Long, slp As RegexMatcherCommon.StartLengthPair, anchorOp As Long
    Dim inputBuffer As RegexMatcherCommon.InputBufferTy

    If sp < 0 Then DfaMatchFrom = -1: Exit Function
    If sp > Len(inputStr) Then DfaMatchFrom = -1: Exit Function
    
    DfaStateStructureInitialize stateStruct, multiline:=multiline, dotAll:=dotAll

    If bwMatch Then
        spDelta = -1
        spLast = -1
        pc = bytecode(RegexBytecode.BYTECODE_IDX_START_DFA_BW_RUN)
    Else
        spDelta = 1
        spLast = Len(inputStr) + 1
        pc = bytecode(RegexBytecode.BYTECODE_IDX_START_WITH_PREFIX)
    End If

    If pc < 0 Then Err.Raise RegexErrors.REGEX_ERR_UNSUPPORTED
    If Not bwMatch Then
        RegexMatcherCommon.InitializeLiteralMatcher stateStruct.literal, bytecode
        If stateStruct.literal.isLiteral Then
            If RegexMatcherCommon.LiteralFindMatch(slp, stateStruct.literal, inputStr, sp + 1, multiline) Then
                DfaMatchFrom = slp.start + slp.Length - 1
            Else
                DfaMatchFrom = -1
            End If
            Exit Function
        End If
        anchorOp = RegexMatcherCommon.LeadingStartAnchor(bytecode)
        If anchorOp <> 0 Then
            If (anchorOp - (multiline And RegexBytecode.MODIFIER_M_WRITE) And RegexBytecode.MODIFIER_M_ACTIVE) = 0 Then
                If sp <> 0 Then DfaMatchFrom = -1: Exit Function
                pc = bytecode(RegexBytecode.BYTECODE_IDX_START_ACTUAL)
            End If
        End If
    End If
    RegexMatcherCommon.InitializeInputBuffer inputBuffer, inputStr
    If finishEagerly Then mode = DFA_MODE_FINISH_EAGERLY Else mode = DFA_MODE_NORMAL
    DfaRunThreads _
        lastMatch, stateStruct, bytecode, pc, inputStr, inputBuffer, sp, spDelta, spLast, mode
    DfaMatchFrom = lastMatch
End Function

Private Function TransitionTableLookup( _
    ByRef outStateId1 As Long, _
    ByRef stateStruct As StateStructure, _
    ByVal stateId0 As Long, _
    ByVal currentChar As Long _
) As Boolean
    ' Look up (state0, currentChar) in the transition table.
    ' Return true, if transition was found. In that case, outStateId1 is set to the
    ' stateId of the target.
    Dim i As Long, invStateId0 As Long, h As Long, hp As Long, mask As Long
    
    With stateStruct
        invStateId0 = Not stateId0
        
        ' Remember to adjust the hash function in TransitionTableInsert as well
        ' if you make changes here!
        h = (stateId0 And &H7FFF0000) Xor ((stateId0 And &HFFFF&) * &H1022&) _
            Xor ((currentChar And &HFFFF&) * &H4081&)
        hp = h Mod .transitionTablePrime
    
        i = hp
        mask = .transitionTableCapacity - 1
        Do
            With .transitionTable(i)
                If .sourceStateId = 0 Then TransitionTableLookup = False: Exit Function
                If .sourceStateId <> invStateId0 Then GoTo ContinueLoop
                If .charCode <> currentChar Then GoTo ContinueLoop
                outStateId1 = .targetStateId
                TransitionTableLookup = True
                Exit Function
            End With
ContinueLoop:
            i = (i + 1) And mask
        Loop
    End With
End Function

Private Sub TransitionTableInsert( _
    ByRef stateStruct As StateStructure, _
    ByVal stateId0 As Long, _
    ByVal charCode As Long, _
    ByVal stateId1 As Long _
)
    ' Called only below the load limit; at least one slot is always empty.
    
    Dim i As Long, invStateId0 As Long, h As Long, hp As Long, tmp As Long, _
        mask As Long
    Dim entry As TransitionTableEntry, tmpEntry As TransitionTableEntry

    invStateId0 = Not stateId0
    With stateStruct
        ' Remember to adjust the hash function in TransitionTableLookup as well
        ' if you make changes here!
        h = (stateId0 And &H7FFF0000) Xor ((stateId0 And &HFFFF&) * &H1022&) _
            Xor ((charCode And &HFFFF&) * &H4081&)
        hp = h Mod .transitionTablePrime
        
        With entry
            .hashRank = 0
            .sourceStateId = invStateId0
            .charCode = charCode
            .targetStateId = stateId1
        End With
        
        i = hp
        mask = .transitionTableCapacity - 1
        Do
            tmp = .transitionTable(i).sourceStateId
            If tmp = 0 Then
                ' Add entry and exit
                .transitionTable(i) = entry
                .transitionTableCount = .transitionTableCount + 1
                Exit Sub
            ElseIf tmp = entry.sourceStateId Then
                If .transitionTable(i).charCode = entry.charCode Then Exit Sub
            End If
            
            If .transitionTable(i).hashRank < entry.hashRank Then
                tmpEntry = .transitionTable(i)
                .transitionTable(i) = entry
                entry = tmpEntry
            End If
            
            entry.hashRank = entry.hashRank + 1
            i = (i + 1) And mask
        Loop
    End With
End Sub

' States contain an ordered sequence, not an unordered subset. The context
' distinguishes scan direction and exhaustive reverse matching. Eager and normal
' matching share transitions because their state construction is identical.
Private Sub StateHashTableInsert( _
    ByRef stateStruct As StateStructure, _
    ByVal context As Long, _
    ByRef inOutP As Long, _
    ByRef outStateId As Long _
)
    Dim i As Long, j As Long, p As Long, n As Long, stateHash As Long, bucketHash As Long
    With stateStruct
        n = .stateListsBuffer.Buffer(inOutP + 1)
        stateHash = context
        For j = inOutP + 2 To inOutP + 1 + n
            ' Rotate a 31-bit hash, then mix in the next serialized word.
            ' All intermediate values fit in a signed VBA Long.
            stateHash = ((stateHash And &HFF&) * &H800000&) Or (stateHash \ &H100&)
            stateHash = stateHash Xor .stateListsBuffer.Buffer(j)
        Next

        ' Mix both halves before selecting low bucket bits. The product stays
        ' below 2^31: 65535 * 25173 = 1649712555. Keep the full fingerprint.
        bucketHash = (stateHash Xor (stateHash \ &H10000&)) And &HFFFF&
        i = ((bucketHash * &H6255&) \ &H20&) And .stateHashTableMask
        Do
            p = .stateHashTable(i).statePtr
            If p = 0 Then Exit Do
            p = Not p
            If .stateHashTable(i).hash <> stateHash Then GoTo NextEntry
            If .stateHashTable(i).context <> context Then GoTo NextEntry
            If .stateListsBuffer.Buffer(p + 1) <> n Then GoTo NextEntry
            For j = 0 To n - 1
                If .stateListsBuffer.Buffer(p + 2 + j) <> _
                    .stateListsBuffer.Buffer(inOutP + 2 + j) Then GoTo NextEntry
            Next

            .stateListsBuffer.Length = inOutP ' discard the duplicate candidate
            inOutP = p
            outStateId = .stateListsBuffer.Buffer(p)
            Exit Sub
NextEntry:
            i = (i + 1) And .stateHashTableMask
        Loop

        With .stateHashTable(i)
            .hash = stateHash
            .context = context
            .statePtr = Not inOutP ' zero denotes an empty slot, including at p = 0
        End With
        outStateId = .nextFreeStateId
        .stateListsBuffer.Buffer(inOutP) = outStateId
        .stateListPtrs(outStateId) = inOutP
        .nextFreeStateId = outStateId + 1
    End With
End Sub

Private Sub DfaRunThreads( _
    ByRef outLastMatch As Long, _
    ByRef stateStruct As StateStructure, _
    ByRef bytecode() As Long, _
    ByVal pc As Long, _
    ByRef inputStr As String, _
    ByRef inputBuffer As RegexMatcherCommon.InputBufferTy, _
    ByVal sp As Long, _
    ByVal spDelta As Long, _
    ByVal spLast As Long, _
    ByVal mode As DfaMode _
)
    Dim stack As ArrayBuffer.Ty
    Dim n0 As Long, p0 As Long, q0 As Long, i0 As Long
    Dim pp1 As Long, stateId0 As Long, stateId1 As Long, strOffset As Long, op As Long
    Dim spEndOfInput As Long, finishMask As Long
    Dim lastMatch As Long, action As Long, context As Long
    Dim assertExpect As Long
    Dim currentChar As Long, currentCharCanon As Long, byteOffset As Long
    Dim windowByteLength As Long, inputPosition As Long
    Dim thrdSetActiveRoot As Long, thrdSetInactiveRoot As Long
    Dim thrdSetBuffer As RegexNfaThreadPool.ThreadSetBuffer
        
    ' Temporaries
    Dim c1 As Long, c2 As Long, n As Long, aa As Long, bb As Long, mm As Long
    Dim b1 As Boolean
    
    Const DEFAULT_MINIMUM_THREADLIST_CAPACITY = 16
    If Not stateStruct.tablesInitialized Then InitializeTables stateStruct
    With thrdSetBuffer
        .Capacity = 2 * DEFAULT_MINIMUM_THREADLIST_CAPACITY
        ReDim .buf(0 To .Capacity - 1) As ThreadSetNode
    End With
    
    If mode = DFA_MODE_FINISH_EAGERLY Then
        finishMask = ACTION_FINISH Or ACTION_SET_LAST_MATCH_CUR Or ACTION_SET_LAST_MATCH_PREV
    Else
        finishMask = ACTION_FINISH
    End If
    context = (mode And DFA_MODE_EXTENSIVE) - (spDelta < 0)
    lastMatch = -1

    If spDelta > 0 Then
        spEndOfInput = inputBuffer.Length + 1
        If sp = 0 Then
            currentChar = -2
        Else
            currentChar = RegexMatcherCommon.PeekBufferCharCode(inputBuffer, inputStr, sp - 1, spDelta)
        End If
    Else
        spEndOfInput = -1
        If sp = inputBuffer.Length Then
            currentChar = -2
        Else
            currentChar = RegexMatcherCommon.PeekBufferCharCode(inputBuffer, inputStr, sp, spDelta)
        End If
    End If
    currentCharCanon = -3
    byteOffset = 2 * (sp - 1 - (spDelta < 0) - inputBuffer.windowStart)
    windowByteLength = 2 * (inputBuffer.windowEnd - inputBuffer.windowStart)
    
    thrdSetActiveRoot = -1: thrdSetInactiveRoot = -1: thrdSetBuffer.nUsed = 0
    p0 = 0: i0 = 0: q0 = 0
    stateId0 = 0
    strOffset = 0
    action = 0
    pp1 = stateStruct.stateListsBuffer.Length
    ArrayBuffer.AppendUnspecified stateStruct.stateListsBuffer, 2
    GoTo CreateClosureFromCurrent

ContinueWithNextChar:
    If sp = spLast Then GoTo Finished
    sp = sp + spDelta
    byteOffset = byteOffset + 2 * spDelta
    If sp = spEndOfInput Then
        currentChar = -1
    Else
        ' Forward sp identifies the boundary after the character; backward sp
        ' identifies its zero-based index. Each UTF-16 code unit uses two bytes.
        If byteOffset < 0 Or byteOffset >= windowByteLength Then
            inputPosition = sp - 1 - (spDelta < 0)
            RegexMatcherCommon.RefillInputBuffer inputBuffer, inputStr, inputPosition, spDelta
            byteOffset = 2 * (inputPosition - inputBuffer.windowStart)
            windowByteLength = 2 * (inputBuffer.windowEnd - inputBuffer.windowStart)
        End If
        currentChar = CLng(inputBuffer.bytes(byteOffset)) Or (CLng(inputBuffer.bytes(byteOffset + 1)) * 256&)
    End If
    currentCharCanon = -3
    action = 0
    
    ' At a caller-imposed reverse bound we inspect the adjacent character for
    ' assertions, but must not consume it. Such an edge depends on the bound.
    If sp <> spLast Then
        If TransitionTableLookup(stateId1, stateStruct, stateId0, currentChar) Then
            action = stateId1 And ACTION_MASK
            stateId0 = stateId1 And Not ACTION_MASK
            GoTo ApplyAction
        End If
    End If

    stack.Length = 0
    thrdSetActiveRoot = -1: thrdSetInactiveRoot = -1: thrdSetBuffer.nUsed = 0

    p0 = stateStruct.stateListPtrs(stateId0)
    n0 = stateStruct.stateListsBuffer.Buffer(p0 + 1)
    q0 = p0 + 2 + n0

    i0 = p0 + 2

    pp1 = stateStruct.stateListsBuffer.Length
    ArrayBuffer.AppendUnspecified stateStruct.stateListsBuffer, 2

ContinueWithNextInState:
        If i0 >= q0 Then GoTo StateComplete
        pc = stateStruct.stateListsBuffer.Buffer(i0)
        i0 = i0 + 1
        strOffset = 0

ContinueWithPc:
        op = bytecode(pc)
        
        ' The following statement depends on the numerical values of our constants!
        ' The following comment line serves as a marker enabling us to locate this
        ' line of code by a find:
        ' __REOP_NUMERICAL_VALUES__
        
        On op And RegexBytecode.REOP_OPCODE_MASK GoTo _
            L_MATCH, L_CHAR, L_DOT, L_RANGES, L_INVRANGES, _
            CreateClosureFromCurrent, CreateClosureFromCurrent, _
                CreateClosureFromCurrent, CreateClosureFromCurrent, _
                CreateClosureFromCurrent, _
            InternalError, InternalError, InternalError, _
                CreateClosureFromCurrent, CreateClosureFromCurrent, _
            CreateClosureFromCurrent, CreateClosureFromCurrent, _
                InternalError, InternalError, InternalError, _
            InternalError, InternalError, InternalError, InternalError, _
                InternalError, _
            InternalError, InternalError, InternalError, InternalError, _
                InternalError, _
            InternalError, CreateClosureFromCurrent, L_STRING, L_PERFORM_ASSERT, _
            L_MATCH_PREV, L_MATCH_OLD, CreateClosureFromCurrent, _
            CreateClosureFromCurrent, CreateClosureFromCurrent
    
        GoTo InternalError

L_MATCH:
        action = action Or ACTION_SET_LAST_MATCH_CUR
        If mode = DFA_MODE_EXTENSIVE Then GoTo ContinueWithNextInState
        GoTo L_MATCH_OLD
        
L_MATCH_PREV:
        action = action Or ACTION_SET_LAST_MATCH_PREV
        If mode = DFA_MODE_EXTENSIVE Then GoTo ContinueWithNextInState
        ' else fallthrough
        
L_MATCH_OLD:
        GoTo LastBranchCompleteDueToMatch
        
L_CHAR:
        If sp = spLast Then GoTo ContinueWithNextInState
        pc = pc + 1: c1 = bytecode(pc)
        
        If op And RegexBytecode.MODIFIER_I_ACTIVE Then
            If currentCharCanon = -3 Then _
                currentCharCanon = RegexUnicodeSupport.ReCanonicalizeChar(currentChar)
            c2 = currentCharCanon
        Else
            c2 = currentChar
        End If
        
        If c1 <> c2 Then GoTo ContinueWithNextInState
        
        pc = pc + 1
        GoTo CreateClosureFromCurrent

L_DOT:
        If sp = spLast Then GoTo ContinueWithNextInState
        If currentChar = RegexMatcherCommon.MATCHER_ENDOFINPUT Then GoTo ContinueWithNextInState
        If 0 = ( _
            op - (stateStruct.dotAll And RegexBytecode.MODIFIER_S_WRITE) And _
            RegexBytecode.MODIFIER_S_ACTIVE _
        ) Then
            If RegexUnicodeSupport.UnicodeIsLineTerminator(currentChar) Then GoTo ContinueWithNextInState
        End If
        
        pc = pc + 1
        GoTo CreateClosureFromCurrent

L_RANGES:
L_INVRANGES:
        If sp = spLast Then GoTo ContinueWithNextInState
        pc = pc + 1
        n = bytecode(pc) ' assert: >= 1
        If currentChar = RegexMatcherCommon.MATCHER_ENDOFINPUT Then GoTo ContinueWithNextInState
        If op And RegexBytecode.MODIFIER_I_ACTIVE Then
            If currentCharCanon = -3 Then
                currentCharCanon = RegexUnicodeSupport.ReCanonicalizeChar(currentChar)
            End If
            c1 = currentCharCanon
        Else
            c1 = currentChar
        End If
        
        ' We are performing a binary search here.
        ' aa points to the upper bound of element -1 (if the first element is eleent
        ' number 0) and bb points to the upper bound of element n (if the last
        ' element is number n-1).
        ' Hence bb - aa = 2 * (n + 1).
        ' This works, but it could be done more efficently.
        aa = pc
        pc = pc + 2 * n + 1
        bb = pc + 1
        
        Do
            mm = aa + 2 * ((bb - aa) \ 4)
            If bytecode(mm) >= c1 Then bb = mm Else aa = mm
            
            If bb - aa = 2 Then
                ' bb is the first upper bound index s.t. ary(bb)>=v
                ' b1 describes whether the match was successful
                If bb >= pc Then b1 = False Else b1 = bytecode(bb - 1) <= c1
                If ((op And RegexBytecode.REOP_OPCODE_MASK) = REOP_RANGES) <> b1 Then _
                    GoTo ContinueWithNextInState
                GoTo CreateClosureFromCurrent
            End If
        Loop
        ' unreachable
                
L_STRING:
        strOffset = stateStruct.stateListsBuffer.Buffer(i0)
        i0 = i0 + 1

        If sp = spLast Then GoTo ContinueWithNextInState

        c1 = bytecode(pc + 2 + strOffset)
        
        If op And RegexBytecode.MODIFIER_I_ACTIVE Then
            If currentCharCanon = -3 Then _
                currentCharCanon = RegexUnicodeSupport.ReCanonicalizeChar(currentChar)
            c2 = currentCharCanon
        Else
            c2 = currentChar
        End If
        
        If c1 <> c2 Then GoTo ContinueWithNextInState
        
        c1 = bytecode(pc + 1) ' string length
        If strOffset + 1 = c1 Then ' reached end of string
            pc = pc + 2 + c1: strOffset = 0
            GoTo CreateClosureFromCurrent
        End If
        
        strOffset = strOffset + 1

        ' The following is the same as BranchComplete, but with
        ' both pc and strOffset being added to the state
        assertExpect = DFA_ASSERT_ALL
        If ThreadSetVisitConditions(thrdSetBuffer, thrdSetInactiveRoot, pc, strOffset, assertExpect) Then
            ArrayBuffer.AppendTwo stateStruct.stateListsBuffer, pc, strOffset
        End If
        GoTo ContinueWithNextInState

L_PERFORM_ASSERT:
    c1 = stateStruct.stateListsBuffer.Buffer(i0): i0 = i0 + 1
    pc = stateStruct.stateListsBuffer.Buffer(i0): i0 = i0 + 1
    If currentChar = RegexMatcherCommon.MATCHER_ENDOFINPUT Then
        c2 = DFA_ASSERT_EXPECT_EOI
    ElseIf RegexMatcherCommon.UnicodeReIsWordchar(currentChar) Then
        c2 = DFA_ASSERT_EXPECT_WORDCHAR
    ElseIf RegexUnicodeSupport.UnicodeIsLineTerminator(currentChar) Then
        c2 = DFA_ASSERT_EXPECT_LINE_TERMINATOR
    Else
        c2 = DFA_ASSERT_EXPECT_OTHER
    End If
    If (c1 And c2) = 0 Then
        i0 = i0 - ((bytecode(pc) And RegexBytecode.REOP_OPCODE_MASK) = REOP_STRING)
        GoTo ContinueWithNextInState
    End If
    GoTo ContinueWithPc

    ' unreachable
        
CreateClosureFromCurrent:
    assertExpect = DFA_ASSERT_ALL
    GoTo CreateClosureContinue
    
CreateClosureFromStack:
        If stack.Length = 0 Then GoTo ContinueWithNextInState
        stack.Length = stack.Length - 2
        assertExpect = stack.Buffer(stack.Length)
        pc = stack.Buffer(stack.Length + 1)
        
CreateClosureContinue:
        If assertExpect = 0 Then GoTo CreateClosureFromStack
        If Not ThreadSetVisitConditions( _
            thrdSetBuffer, thrdSetActiveRoot, pc, strOffset, assertExpect _
        ) Then GoTo CreateClosureFromStack

        op = bytecode(pc)

        ' The following statement depends on the numerical values of our constants!
        ' The following comment line serves as a marker enabling us to locate this
        ' line of code by a find:
        ' __REOP_NUMERICAL_VALUES__

        On op And RegexBytecode.REOP_OPCODE_MASK GoTo _
            BranchCompleteWithMatch, BranchComplete, BranchComplete, BranchComplete, BranchComplete, _
            L_JUMP, L_SPLIT1, L_SPLIT2, L_SAVE, L_SET_NAMED, _
            InternalError, InternalError, InternalError, L_ASSERT_START, L_ASSERT_END, _
            L_ASSERT_WORD_BOUNDARY, L_ASSERT_NOT_WORD_BOUNDARY, InternalError, _
                InternalError, InternalError, _
            InternalError, InternalError, InternalError, InternalError, _
                InternalError, _
            InternalError, InternalError, InternalError, InternalError, InternalError, _
            InternalError, L_FAIL, BranchComplete, InternalError, InternalError, _
                InternalError, L_SPLIT2, L_SPLIT1, L_LOOP_INIT

        GoTo InternalError

L_LOOP_INIT:
        pc = pc + 1
        GoTo CreateClosureContinue

L_JUMP:
        pc = pc + 1
        n = bytecode(pc)
        pc = pc + n + 1
        GoTo CreateClosureContinue

L_SPLIT1:
        ' split1: prefer direct execution (no jump)
        pc = pc + 1
        n = bytecode(pc)
        c1 = pc + 1: c2 = pc + 1 + n
        GoTo SplitCommon
        
L_SPLIT2:
        ' split2: prefer jump execution (not direct)
        pc = pc + 1
        n = bytecode(pc)
        c1 = pc + 1 + n: c2 = pc + 1
        GoTo SplitCommon
        
SplitCommon:
        ' takes c1 and c2 as "parameters"
        
        ArrayBuffer.AppendTwo stack, assertExpect, c2
        pc = c1
        GoTo CreateClosureContinue
        
L_SAVE:
        pc = pc + 2
        GoTo CreateClosureContinue

L_SET_NAMED:
        pc = pc + 3
        GoTo CreateClosureContinue

L_FAIL:
        GoTo CreateClosureFromStack

L_ASSERT_START:
        pc = pc + 1
        If 0 = ( _
            op - (stateStruct.multiline And RegexBytecode.MODIFIER_M_WRITE) And _
            RegexBytecode.MODIFIER_M_ACTIVE _
        ) Then
            If spDelta < 0 Then
                assertExpect = assertExpect And DFA_ASSERT_EXPECT_EOI
                GoTo CreateClosureContinue
            ElseIf currentChar = -2 Then
                GoTo CreateClosureContinue
            Else
                GoTo CreateClosureFromStack
            End If
        Else
            If spDelta < 0 Then
                assertExpect = assertExpect And (DFA_ASSERT_EXPECT_LINE_TERMINATOR Or DFA_ASSERT_EXPECT_EOI)
                GoTo CreateClosureContinue
            ElseIf currentChar = -2 Then
                GoTo CreateClosureContinue
            ElseIf RegexUnicodeSupport.UnicodeIsLineTerminator(currentChar) Then
                GoTo CreateClosureContinue
            Else
                GoTo CreateClosureFromStack
            End If
        End If
        
L_ASSERT_END:
        pc = pc + 1
        If 0 = ( _
            op - (stateStruct.multiline And RegexBytecode.MODIFIER_M_WRITE) And _
            RegexBytecode.MODIFIER_M_ACTIVE _
        ) Then
            If spDelta > 0 Then
                assertExpect = assertExpect And DFA_ASSERT_EXPECT_EOI
                GoTo CreateClosureContinue
            ElseIf currentChar = -2 Then
                GoTo CreateClosureContinue
            Else
                GoTo CreateClosureFromStack
            End If
        Else
            If spDelta > 0 Then
                assertExpect = assertExpect And (DFA_ASSERT_EXPECT_LINE_TERMINATOR Or DFA_ASSERT_EXPECT_EOI)
                GoTo CreateClosureContinue
            ElseIf currentChar = -2 Then
                GoTo CreateClosureContinue
            ElseIf RegexUnicodeSupport.UnicodeIsLineTerminator(currentChar) Then
                GoTo CreateClosureContinue
            Else
                GoTo CreateClosureFromStack
            End If
        End If

L_ASSERT_WORD_BOUNDARY:
        If RegexMatcherCommon.UnicodeReIsWordchar(currentChar) Then
            assertExpect = assertExpect And DFA_ASSERT_EXPECT_NOT_WORDCHAR
        Else
            assertExpect = assertExpect And DFA_ASSERT_EXPECT_WORDCHAR
        End If
        pc = pc + 1
        GoTo CreateClosureContinue
        
L_ASSERT_NOT_WORD_BOUNDARY:
        If RegexMatcherCommon.UnicodeReIsWordchar(currentChar) Then
            assertExpect = assertExpect And DFA_ASSERT_EXPECT_WORDCHAR
        Else
            assertExpect = assertExpect And DFA_ASSERT_EXPECT_NOT_WORDCHAR
        End If
        pc = pc + 1
        GoTo CreateClosureContinue
                
BranchCompleteWithMatch:
        If assertExpect <> DFA_ASSERT_ALL Then
            pc = RegexBytecode.BYTECODE_IDX_REOP_MATCH_PREV
            GoTo BranchComplete
        End If
        
        action = action Or ACTION_SET_LAST_MATCH_CUR
        If mode = DFA_MODE_EXTENSIVE Then GoTo CreateClosureFromStack
        ' else fallthrough
        
LastBranchCompleteDueToMatch:
        pc = RegexBytecode.BYTECODE_IDX_REOP_MATCH_OLD
        ThreadSetInsert thrdSetBuffer, thrdSetInactiveRoot, _
            RegexBytecode.BYTECODE_IDX_REOP_MATCH_OLD, 0
        ArrayBuffer.AppendLong stateStruct.stateListsBuffer, _
            RegexBytecode.BYTECODE_IDX_REOP_MATCH_OLD
        If stateStruct.stateListsBuffer.Length = pp1 + 3 Then _
            action = action Or ACTION_FINISH
        
        GoTo StateComplete
                
BranchComplete:
        ' Serialize the remaining categories, not the already-covered ones.
        ' The mask is part of the ordered state and therefore of its cache key.
        If ThreadSetVisitConditions(thrdSetBuffer, thrdSetInactiveRoot, pc, 0, assertExpect) Then
            If assertExpect <> DFA_ASSERT_ALL Then
                ArrayBuffer.AppendThree stateStruct.stateListsBuffer, _
                    RegexBytecode.BYTECODE_IDX_REOP_PERFORM_ASSERT, assertExpect, pc
            Else
                ArrayBuffer.AppendLong stateStruct.stateListsBuffer, pc
            End If
            If (op And RegexBytecode.REOP_OPCODE_MASK) = REOP_STRING Then
                ArrayBuffer.AppendLong stateStruct.stateListsBuffer, 0
            End If
        End If
        GoTo CreateClosureFromStack


StateComplete:
    n = stateStruct.stateListsBuffer.Length - pp1 - 2
    If n = 0 Then action = action Or ACTION_FINISH
    
    stateStruct.stateListsBuffer.Buffer(pp1 + 1) = n
    ' Flush only after constructing the target. The source is no longer needed,
    ' so keeping the target is sufficient to continue without rescanning input.
    With stateStruct
        b1 = .nextFreeStateId > (.stateHashTableMask + 1) * 3 \ 4 Or _
            .transitionTableCount >= .transitionTableCapacity * 3 \ 4 Or _
            .stateListsBuffer.Length > DFA_STATE_BUFFER_LIMIT
    End With
    If b1 Then
        ClearCacheKeepingState stateStruct, pp1
        stateId0 = 0 ' the old source ID no longer belongs to this cache
    End If
    StateHashTableInsert stateStruct, context, pp1, stateId1
    ' State zero is used only for the initial closure and after a cache reset.
    ' Initial closures depend on pc and surrounding input; never cache that edge.
    If stateId0 <> 0 And sp <> spLast Then _
        TransitionTableInsert stateStruct, stateId0, currentChar, stateId1 Or action
    stateId0 = stateId1

ApplyAction:
    ' Use the same action replay on cache hits and misses. In extensive mode,
    ' several branches can accept in one transition; CUR is furthest in the scan
    ' direction and takes precedence over PREV.
    If action And ACTION_SET_LAST_MATCH_CUR Then
        lastMatch = sp
    ElseIf action And ACTION_SET_LAST_MATCH_PREV Then
        lastMatch = sp - spDelta
    End If
    If action And finishMask Then GoTo Finished
    GoTo ContinueWithNextChar

Finished:
    outLastMatch = lastMatch
    Exit Sub
    
InternalError:
    Err.Raise RegexErrors.REGEX_ERR_INTERNAL_LOGIC_ERR
End Sub
