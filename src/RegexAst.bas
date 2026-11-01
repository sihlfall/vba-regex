Attribute VB_Name = "RegexAst"
Option Explicit

Public Enum AstNodeType
    ' We guarantee: All AST_ constants are >= 0.
    ' ! The parser relies on AST_ASSERT_LOOKAHEAD and LOOKBEHIND being >= 0.
    '
    ' There are places in the code where the numerical values of these constants are
    ' significant. If you make changes, please search for __AST_NUMERICAL_VALUES__
    ' to find them.
    MIN_AST_CODE = 0
    AST_STRING = 1
    AST_DISJ = 2
    AST_CONCAT = 3
    AST_CHAR = 4
    AST_CAPTURE = 5
    AST_REPEAT_EXACTLY = 6
    AST_REPEAT_EXACTLY_POSSESSIVE = 7
    AST_PERIOD = 8
    AST_ASSERT_START = 9
    AST_ASSERT_END = 10
    AST_ASSERT_WORD_BOUNDARY = 11
    AST_ASSERT_NOT_WORD_BOUNDARY = 12
    AST_MATCH = 13
    AST_ZEROONE_GREEDY = 14
    AST_ZEROONE_HUMBLE = 15
    AST_ZEROONE_POSSESSIVE = 16
    AST_STAR_GREEDY = 17
    AST_STAR_HUMBLE = 18
    AST_STAR_POSSESSIVE = 19
    AST_REPEAT_MAX_GREEDY = 20
    AST_REPEAT_MAX_HUMBLE = 21
    AST_REPEAT_MAX_POSSESSIVE = 22
    AST_RANGES = 23
    AST_INVRANGES = 24
    AST_ASSERT_POS_LOOKAHEAD = 25
    AST_ASSERT_NEG_LOOKAHEAD = 26
    AST_ASSERT_POS_LOOKBEHIND = 27
    AST_ASSERT_NEG_LOOKBEHIND = 28
    AST_FAIL = 29
    AST_BACKREFERENCE = 30
    AST_NAMED = 31
    AST_MODIFIER_SCOPE = 32
    AST_EMPTY = 33
    AST_CHAR_CONCAT = 34
    AST_STRING_CONCAT = 35
    
    ' used internally
    AST_ACTUAL = 36
    AST_PLUS_GREEDY = 37
    AST_PLUS_HUMBLE = 38
    MAX_AST_CODE = 38
End Enum

Private Enum AstNodeDescriptionConstant
    NODE_TYPE = 0
    NODE_LCHILD = 1
    NODE_RCHILD = 2
End Enum

Private Enum AstNodePropertiesConstant
    NODE_NONEMPTY = 1
    NODE_UNROLL = 2
    NODE_UNSAFE_REPEAT = 4
    NODE_LITERAL = 8 ' emits one CHAR or STRING, without saves
    NODE_NULLABLE = 16 ' may succeed without consuming, including assertions
End Enum

Public Enum AstTableDescriptionConstant
    AST_TABLE_OFFSET_BLEN = 0
    AST_TABLE_OFFSET_ESFS = 1
    AST_TABLE_ENTRY_LENGTH = 2
    AST_TABLE_LENGTH = AST_TABLE_ENTRY_LENGTH * (MAX_AST_CODE + 1)
End Enum

Private astTableInitialized As Boolean ' default-initialized to False

Public Enum AstDefaultValueConstant
    DEFAULT_QUANTIFIER_EXPANSION_LIMIT = 4096 ' additional Long words, both programs
End Enum

Public Sub AstTableInitialize()
    InitializeAstTable RegexUnicodeSupport.StaticData
End Sub

Private Sub InitializeAstTable(ByRef t() As Long)
    Const b As Long = RegexUnicodeSupport.AST_TABLE_START
    Const blen As Long = b + AST_TABLE_OFFSET_BLEN
    Const esfs As Long = b + AST_TABLE_OFFSET_ESFS
    Const e As Long = AST_TABLE_ENTRY_LENGTH
    
    ' blen: length of bytecode generated for this node (meaningful only if .nc >= 0)
    ' esfs: extra stack space required when generating bytecode for this node
    '   Only nodes with children are permitted to require extra stack space.
    '   Hence .esfs > 0 must imply .nc >= 1.
    
    ' ! When adding new entries, make sure to adjust BYTECODE_GENERATOR_INITIAL_STACK_CAPACITY below, if necessary!
    ' ! See comment on BYTECODE_GENERATOR_INITIAL_STACK_CAPACITY below.
    
    t(blen + e * AST_STRING) = 2:                       t(esfs + e * AST_STRING) = 0
    t(blen + e * AST_DISJ) = 4:                         t(esfs + e * AST_DISJ) = 1
    t(blen + e * AST_CONCAT) = 0:                       t(esfs + e * AST_CONCAT) = 0
    t(blen + e * AST_CHAR) = 2:                         t(esfs + e * AST_CHAR) = 0
    t(blen + e * AST_CAPTURE) = 4:                      t(esfs + e * AST_CAPTURE) = 0
    t(blen + e * AST_REPEAT_EXACTLY) = 7:               t(esfs + e * AST_REPEAT_EXACTLY) = 2
    t(blen + e * AST_REPEAT_EXACTLY_POSSESSIVE) = 9:    t(esfs + e * AST_REPEAT_EXACTLY_POSSESSIVE) = 1
    t(blen + e * AST_PERIOD) = 1:                       t(esfs + e * AST_PERIOD) = 0
    t(blen + e * AST_ASSERT_START) = 1:                 t(esfs + e * AST_ASSERT_START) = 0
    t(blen + e * AST_ASSERT_END) = 1:                   t(esfs + e * AST_ASSERT_END) = 0
    t(blen + e * AST_ASSERT_WORD_BOUNDARY) = 1:         t(esfs + e * AST_ASSERT_WORD_BOUNDARY) = 0
    t(blen + e * AST_ASSERT_NOT_WORD_BOUNDARY) = 1:     t(esfs + e * AST_ASSERT_NOT_WORD_BOUNDARY) = 0
    t(blen + e * AST_MATCH) = 1:                        t(esfs + e * AST_MATCH) = 0
    t(blen + e * AST_ZEROONE_GREEDY) = 2:               t(esfs + e * AST_ZEROONE_GREEDY) = 1
    t(blen + e * AST_ZEROONE_HUMBLE) = 2:               t(esfs + e * AST_ZEROONE_HUMBLE) = 1
    t(blen + e * AST_ZEROONE_POSSESSIVE) = 3:           t(esfs + e * AST_ZEROONE_POSSESSIVE) = 1
    t(blen + e * AST_STAR_GREEDY) = 4:                  t(esfs + e * AST_STAR_GREEDY) = 1
    t(blen + e * AST_STAR_HUMBLE) = 4:                  t(esfs + e * AST_STAR_HUMBLE) = 1
    t(blen + e * AST_STAR_POSSESSIVE) = 5:              t(esfs + e * AST_STAR_POSSESSIVE) = 1
    t(blen + e * AST_REPEAT_MAX_GREEDY) = 7:            t(esfs + e * AST_REPEAT_MAX_GREEDY) = 2
    t(blen + e * AST_REPEAT_MAX_HUMBLE) = 7:            t(esfs + e * AST_REPEAT_MAX_HUMBLE) = 2
    t(blen + e * AST_REPEAT_MAX_POSSESSIVE) = 8:        t(esfs + e * AST_REPEAT_MAX_POSSESSIVE) = 1
    t(blen + e * AST_RANGES) = 2:                       t(esfs + e * AST_RANGES) = 0
    t(blen + e * AST_INVRANGES) = 2:                    t(esfs + e * AST_INVRANGES) = 0
    t(blen + e * AST_ASSERT_POS_LOOKAHEAD) = 4:         t(esfs + e * AST_ASSERT_POS_LOOKAHEAD) = 2
    t(blen + e * AST_ASSERT_NEG_LOOKAHEAD) = 4:         t(esfs + e * AST_ASSERT_NEG_LOOKAHEAD) = 2
    t(blen + e * AST_ASSERT_POS_LOOKBEHIND) = 4:        t(esfs + e * AST_ASSERT_POS_LOOKBEHIND) = 2
    t(blen + e * AST_ASSERT_NEG_LOOKBEHIND) = 4:        t(esfs + e * AST_ASSERT_NEG_LOOKBEHIND) = 2
    t(blen + e * AST_FAIL) = 1:                         t(esfs + e * AST_FAIL) = 0
    t(blen + e * AST_BACKREFERENCE) = 2:                t(esfs + e * AST_BACKREFERENCE) = 0
    t(blen + e * AST_NAMED) = 3:                        t(esfs + e * AST_NAMED) = 0
    t(blen + e * AST_MODIFIER_SCOPE) = 0:               t(esfs + e * AST_MODIFIER_SCOPE) = 1
    t(blen + e * AST_EMPTY) = 0:                        t(esfs + e * AST_EMPTY) = 0
    t(blen + e * AST_CHAR_CONCAT) = 2:                  t(esfs + e * AST_CHAR_CONCAT) = 0
    t(blen + e * AST_STRING_CONCAT) = 2:                t(esfs + e * AST_STRING_CONCAT) = 0
    t(blen + e * AST_ACTUAL) = 0:                       t(esfs + e * AST_ACTUAL) = 0
    t(blen + e * AST_PLUS_GREEDY) = 3:                  t(esfs + e * AST_PLUS_GREEDY) = 1
    t(blen + e * AST_PLUS_HUMBLE) = 3:                  t(esfs + e * AST_PLUS_HUMBLE) = 1
End Sub

' A zero expansion budget retains the compact counted instructions. Otherwise
' the sizing traversal also chooses bounded, consuming repeats to expand.
Public Sub AstToBytecode( _
    ByRef ast As ArrayBuffer.Ty, _
    ByRef identifierTree As RegexIdentifierSupport.IdentifierTreeTy, _
    ByVal caseInsensitive As Boolean, ByRef bytecode() As Long, _
    Optional ByVal maxExpansion As Long = DEFAULT_QUANTIFIER_EXPANSION_LIMIT _
)
    Dim bytecodePtr As Long
    Dim curNode As Long, prevNode As Long
    Dim stack() As Long, sp As Long, nodePlan() As Long
    Dim direction As Long ' 0 = left before right, -1 = right before left
    Dim returningFromFirstChild As Long ' 0 = no, RegexNumericConstants.LONG_FIRST_BIT = yes
    Dim modifiers As Long, actualRoot As Long
    Dim requiresDfsMatching As Boolean
    Dim dfaBwRun As Boolean, singlePath As Boolean, expandRepeats As Boolean
    
    ' temporaries, do not survive over more than one iteration
    Dim opcode1 As Long, opcode2 As Long, opcode3 As Long, tmp As Long, tmpCnt As Long, _
        e As Long, j As Long, patchPos As Long, maxSave As Long
    
    If Not astTableInitialized Then AstTableInitialize
    
    ' amend AST
    tmp = ast.Length
    actualRoot = ast.Buffer(0)
    ArrayBuffer.AppendTwo ast, AST_ACTUAL, actualRoot
    
    ArrayBuffer.AppendLong ast, AST_PERIOD
    ArrayBuffer.AppendThree ast, AST_MODIFIER_SCOPE, tmp + 2, MODIFIER_S_MASK
    ArrayBuffer.AppendTwo ast, AST_STAR_HUMBLE, tmp + 3
    ArrayBuffer.AppendThree ast, AST_CONCAT, tmp + 6, tmp
    ast.Buffer(0) = tmp + 8
    
    PrepareStackAndBytecodeBuffer ast, identifierTree, caseInsensitive, _
        requiresDfsMatching, stack, bytecode, nodePlan, maxExpansion, expandRepeats
    
    modifiers = caseInsensitive And RegexBytecode.MODIFIER_I_MASK
    
    curNode = ast.Buffer(0) ' first word contains index of root
    bytecodePtr = RegexBytecode.BYTECODE_FIXED_PROLOGUE_LAST_IDX + 3 * bytecode(1) + 1
    bytecode(RegexBytecode.BYTECODE_IDX_START_WITH_PREFIX) = bytecodePtr
    maxSave = -1
    direction = 0
    
    dfaBwRun = False
    
StartBytecodeGeneration:
    sp = 0
    prevNode = -1
    returningFromFirstChild = 0

ContinueLoop:
    ' The following statement depends on the numerical values of our constants!
    ' The following comment line serves as a marker enabling us to locate this line of
    ' code by doing a find:
    ' __AST_NUMERICAL_VALUES__
    On ast.Buffer(curNode + NODE_TYPE) GoTo _
        L_STRING, L_DISJ, L_CONCAT, L_CHAR, L_CAPTURE, _
        L_REPEAT_EXACTLY, HandleRepeatExactlyPossessive, L_PERIOD, L_ASSERT_START, _
            L_ASSERT_END, _
        L_ASSERT_WORD_BOUNDARY, L_ASSERT_NOT_WORD_BOUNDARY, L_MATCH, _
            L_ZEROONE_GREEDY, L_ZEROONE_HUMBLE, _
        HandleZeroonePossessive, L_STAR_GREEDY, L_STAR_HUMBLE, HandleStarPossessive, _
            L_REPEAT_MAX_GREEDY, _
        L_REPEAT_MAX_HUMBLE, HandleRepeatQuantifiedPossessive, L_RANGES, L_INVRANGES, _
            L_ASSERT_POS_LOOKAHEAD, _
        L_ASSERT_NEG_LOOKAHEAD, L_ASSERT_POS_LOOKBEHIND, L_ASSERT_NEG_LOOKBEHIND, _
            L_FAIL, L_BACKREFERENCE, _
        L_NAMED, L_MODIFIER_SCOPE, TurnToParent, L_CHAR_CONCAT, L_STRING_CONCAT, _
        L_ACTUAL, L_PLUS, L_PLUS

    Err.Raise REGEX_ERR_INTERNAL_LOGIC_ERR ' unreachable

L_STRING:
    tmpCnt = ast.Buffer(curNode + 1) ' assert(tmpCnt >= 1)
    j = curNode + 2 + ((tmpCnt - 1) And direction)
    e = curNode + 1 + tmpCnt - ((tmpCnt - 1) And direction)
    tmp = 1 + 2 * direction
    
    bytecode(bytecodePtr) = REOP_STRING Or modifiers And RegexBytecode.MODIFIER_I_ACTIVE
    bytecodePtr = bytecodePtr + 1
    bytecode(bytecodePtr) = tmpCnt
    bytecodePtr = bytecodePtr + 1
    Do
        bytecode(bytecodePtr) = ast.Buffer(j)
        bytecodePtr = bytecodePtr + 1
        If j = e Then Exit Do
        j = j + tmp
    Loop
    GoTo TurnToParent
L_STRING_CONCAT:
    If (returningFromFirstChild <> 0) Xor direction Then
        tmpCnt = ast.Buffer(curNode + 2) ' assert(tmpCnt >= 1)
        j = curNode + 3 + ((tmpCnt - 1) And direction)
        e = curNode + 2 + tmpCnt - ((tmpCnt - 1) And direction)
        tmp = 1 + 2 * direction
        
        bytecode(bytecodePtr) = REOP_STRING Or modifiers And RegexBytecode.MODIFIER_I_ACTIVE
        bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = tmpCnt
        bytecodePtr = bytecodePtr + 1
        Do
            bytecode(bytecodePtr) = ast.Buffer(j)
            bytecodePtr = bytecodePtr + 1
            If j = e Then Exit Do
            j = j + tmp
        Loop
    End If
    If returningFromFirstChild Then
        GoTo TurnToParent
    Else ' previous was parent
        GoTo TurnToLeftChild
    End If
L_RANGES:
    opcode1 = REOP_RANGES Or modifiers And RegexBytecode.MODIFIER_I_ACTIVE
    GoTo HandleRanges
L_INVRANGES:
    opcode1 = REOP_INVRANGES Or modifiers And RegexBytecode.MODIFIER_I_ACTIVE
    GoTo HandleRanges
L_CHAR:
    bytecode(bytecodePtr) = REOP_CHAR Or modifiers And RegexBytecode.MODIFIER_I_ACTIVE
    bytecodePtr = bytecodePtr + 1
    bytecode(bytecodePtr) = ast.Buffer(curNode + 1)
    bytecodePtr = bytecodePtr + 1
    GoTo TurnToParent
L_CHAR_CONCAT:
    If (returningFromFirstChild <> 0) Xor direction Then
        bytecode(bytecodePtr) = REOP_CHAR Or modifiers And RegexBytecode.MODIFIER_I_ACTIVE
        bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = ast.Buffer(curNode + 2)
        bytecodePtr = bytecodePtr + 1
    End If
    If returningFromFirstChild Then
        GoTo TurnToParent
    Else ' previous was parent
        GoTo TurnToLeftChild
    End If
L_PERIOD:
    bytecode(bytecodePtr) = REOP_DOT Or modifiers And RegexBytecode.MODIFIER_S_MASK
    bytecodePtr = bytecodePtr + 1
    GoTo TurnToParent
L_MATCH:
    bytecode(bytecodePtr) = REOP_MATCH: bytecodePtr = bytecodePtr + 1
    GoTo TurnToParent
L_ASSERT_START:
    bytecode(bytecodePtr) = REOP_ASSERT_START Or modifiers And RegexBytecode.MODIFIER_M_MASK
    bytecodePtr = bytecodePtr + 1
    GoTo TurnToParent
L_ASSERT_END:
    bytecode(bytecodePtr) = REOP_ASSERT_END Or modifiers And RegexBytecode.MODIFIER_M_MASK
    bytecodePtr = bytecodePtr + 1
    GoTo TurnToParent
L_ASSERT_WORD_BOUNDARY:
    bytecode(bytecodePtr) = REOP_ASSERT_WORD_BOUNDARY: bytecodePtr = bytecodePtr + 1
    GoTo TurnToParent
L_ASSERT_NOT_WORD_BOUNDARY:
    bytecode(bytecodePtr) = REOP_ASSERT_NOT_WORD_BOUNDARY: bytecodePtr = bytecodePtr + 1
    GoTo TurnToParent
L_DISJ:
    singlePath = False
    If returningFromFirstChild Then ' previous was left child
        sp = sp - 1: patchPos = stack(sp)
        bytecode(bytecodePtr) = REOP_JUMP: bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        bytecode(patchPos) = bytecodePtr - patchPos - 1
        
        GoTo TurnToRightChild
    ElseIf prevNode = ast.Buffer(curNode + NODE_RCHILD) Then ' previous was right child
        sp = sp - 1: patchPos = stack(sp)
        bytecode(patchPos) = bytecodePtr - patchPos - 1
    
        GoTo TurnToParent
    Else ' previous was parent
        bytecode(bytecodePtr) = REOP_SPLIT1: bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        GoTo TurnToLeftChild
    End If
L_CONCAT:
    If returningFromFirstChild Then ' previous was first child
        GoTo TurnToSecondChild
    ElseIf prevNode = ast.Buffer(curNode + NODE_RCHILD + direction) Then ' previous was second child
        GoTo TurnToParent
    Else ' previous was parent
        GoTo TurnToFirstChild
    End If
L_CAPTURE:
    If returningFromFirstChild Then
        If dfaBwRun Then GoTo TurnToParent
        bytecode(bytecodePtr) = REOP_SAVE: bytecodePtr = bytecodePtr + 1
        tmp = ast.Buffer(curNode + 2) * 2 + 1
        If tmp > maxSave Then maxSave = tmp
        bytecode(bytecodePtr) = tmp + direction: bytecodePtr = bytecodePtr + 1
        GoTo TurnToParent
    Else
        If dfaBwRun Then GoTo TurnToLeftChild
        bytecode(bytecodePtr) = REOP_SAVE: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = ast.Buffer(curNode + 2) * 2 - direction: bytecodePtr = bytecodePtr + 1
        GoTo TurnToLeftChild
    End If
L_REPEAT_EXACTLY:
    opcode1 = REOP_REPEAT_EXACTLY_INIT: opcode2 = REOP_REPEAT_EXACTLY_START: opcode3 = REOP_REPEAT_EXACTLY_END
    GoTo HandleRepeatQuantified
L_REPEAT_MAX_GREEDY:
    opcode1 = REOP_REPEAT_MAX_GREEDY_INIT: opcode2 = REOP_REPEAT_MAX_GREEDY_START: opcode3 = REOP_REPEAT_MAX_GREEDY_END
    GoTo HandleRepeatQuantified
L_REPEAT_MAX_HUMBLE:
    opcode1 = REOP_REPEAT_MAX_HUMBLE_INIT: opcode2 = REOP_REPEAT_MAX_HUMBLE_START: opcode3 = REOP_REPEAT_MAX_HUMBLE_END
    GoTo HandleRepeatQuantified
L_ZEROONE_GREEDY:
    opcode1 = REOP_SPLIT1
    GoTo HandleZeroone
L_ZEROONE_HUMBLE:
    opcode1 = REOP_SPLIT2
    GoTo HandleZeroone
L_STAR_GREEDY:
    opcode1 = REOP_SPLIT1
    GoTo HandleStar
L_STAR_HUMBLE:
    opcode1 = REOP_SPLIT2
    GoTo HandleStar
L_ASSERT_POS_LOOKAHEAD:
    opcode1 = REOP_LOOKPOS: opcode2 = REOP_END_LOOKPOS
    GoTo HandleLookahead
L_ASSERT_NEG_LOOKAHEAD:
    opcode1 = REOP_LOOKNEG: opcode2 = REOP_END_LOOKNEG
    GoTo HandleLookahead
L_ASSERT_POS_LOOKBEHIND:
    opcode1 = REOP_LOOKPOS: opcode2 = REOP_END_LOOKPOS
    GoTo HandleLookbehind
L_ASSERT_NEG_LOOKBEHIND:
    opcode1 = REOP_LOOKNEG: opcode2 = REOP_END_LOOKNEG
    GoTo HandleLookbehind
L_FAIL:
    bytecode(bytecodePtr) = REOP_FAIL: bytecodePtr = bytecodePtr + 1
    GoTo TurnToParent
L_BACKREFERENCE:
    bytecode(bytecodePtr) = REOP_BACKREFERENCE Or modifiers And RegexBytecode.MODIFIER_I_ACTIVE
    bytecodePtr = bytecodePtr + 1
    bytecode(bytecodePtr) = ast.Buffer(curNode + 1): bytecodePtr = bytecodePtr + 1
    GoTo TurnToParent
L_NAMED:
    If returningFromFirstChild Then
        ' nothing to be done
        GoTo TurnToParent
    Else
        If dfaBwRun Then GoTo TurnToLeftChild
        bytecode(bytecodePtr) = REOP_SET_NAMED: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = ast.Buffer(curNode + 2): bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = ast.Buffer(curNode + 3): bytecodePtr = bytecodePtr + 1
        GoTo TurnToLeftChild
    End If
L_MODIFIER_SCOPE:
    If returningFromFirstChild Then
        sp = sp - 1: modifiers = stack(sp)
        GoTo TurnToParent
    Else
        stack(sp) = modifiers: sp = sp + 1
        tmp = ast.Buffer(curNode + 2)
        modifiers = RegexBytecode.MODIFIER_WRITE_MASK And (modifiers Or tmp) Or _
            RegexBytecode.MODIFIER_ACTIVE_MASK And ( _
                modifiers Xor tmp * 2 And (tmp Xor modifiers) _
            )
        GoTo TurnToLeftChild
    End If
L_ACTUAL:
    If returningFromFirstChild Then
        GoTo TurnToParent
    Else
        singlePath = True ' exclude the implicit unanchored search loop
        bytecode(RegexBytecode.BYTECODE_IDX_START_ACTUAL) = bytecodePtr
        GoTo TurnToLeftChild
    End If
        
HandleRanges: ' requires: opcode1
    tmpCnt = ast.Buffer(curNode + 1)
    bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
    j = curNode
    e = curNode + 1 + 2 * tmpCnt
    Do ' copy everything, including first word, which is the length
        j = j + 1
        bytecode(bytecodePtr) = ast.Buffer(j): bytecodePtr = bytecodePtr + 1
    Loop Until j = e
    GoTo TurnToParent

HandleRepeatQuantified: ' requires: opcode1, opcode2, opcode 3
    If ast.Buffer(curNode) <> AST_REPEAT_EXACTLY Then singlePath = False
    If expandRepeats Then
        If nodePlan(curNode) And NODE_UNROLL Then GoTo UnrollRepeat
    End If
    tmpCnt = ast.Buffer(curNode + 2)
    If returningFromFirstChild Then
        sp = sp - 1: patchPos = stack(sp)
        bytecode(bytecodePtr) = opcode3: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = tmpCnt: bytecodePtr = bytecodePtr + 1
        tmp = bytecodePtr - patchPos
        bytecode(bytecodePtr) = tmp: bytecodePtr = bytecodePtr + 1
        bytecode(patchPos) = tmp
        GoTo TurnToParent
    Else
        bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = opcode2: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = tmpCnt: bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        GoTo TurnToLeftChild
    End If

UnrollRepeat:
    If returningFromFirstChild Then
        sp = sp - 1: patchPos = stack(sp)
        If nodePlan(curNode) And NODE_LITERAL Then
            opcode1 = bytecode(patchPos)
            If (opcode1 And REOP_OPCODE_MASK) = REOP_CHAR Then
                bytecode(patchPos + 2) = bytecode(patchPos + 1)
                bytecodePtr = bytecodePtr + 1
            End If
            tmp = bytecodePtr - patchPos - 2
            bytecode(patchPos) = REOP_STRING Or (opcode1 And Not REOP_OPCODE_MASK)
            bytecode(patchPos + 1) = tmp * ast.Buffer(curNode + 2)
            patchPos = patchPos + 2 ' copy literal data, retaining one header
        End If
        tmp = bytecodePtr - patchPos
        e = bytecodePtr
        ' All control-flow offsets in the body are relative. Copy the block in
        ' its emitted direction, retaining capture slots and branch priority.
        For tmpCnt = 2 To ast.Buffer(curNode + 2)
            For j = patchPos To e - 1
                bytecode(bytecodePtr) = bytecode(j): bytecodePtr = bytecodePtr + 1
            Next
        Next
        If ast.Buffer(curNode) <> AST_REPEAT_EXACTLY Then
            ' Nested optionals: every split exits the entire remaining chain.
            For j = patchPos + 1 To bytecodePtr - 1 Step tmp
                bytecode(j) = bytecodePtr - j - 1
            Next
        End If
        GoTo TurnToParent
    Else
        stack(sp) = bytecodePtr: sp = sp + 1
        If ast.Buffer(curNode) <> AST_REPEAT_EXACTLY Then
            If ast.Buffer(curNode) = AST_REPEAT_MAX_HUMBLE Then tmp = REOP_SPLIT2 Else tmp = REOP_SPLIT1
            bytecode(bytecodePtr) = tmp: bytecodePtr = bytecodePtr + 2
        End If
        GoTo TurnToLeftChild
    End If
    
HandleRepeatQuantifiedPossessive:
    tmpCnt = ast.Buffer(curNode + 2)
    If returningFromFirstChild Then
        sp = sp - 1: patchPos = stack(sp)
        bytecode(bytecodePtr) = REOP_COMMIT_POSSESSIVE: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = REOP_REPEAT_MAX_GREEDY_END: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = tmpCnt: bytecodePtr = bytecodePtr + 1
        tmp = bytecodePtr - patchPos
        bytecode(bytecodePtr) = tmp: bytecodePtr = bytecodePtr + 1
        bytecode(patchPos) = tmp
        GoTo TurnToParent
    Else
        requiresDfsMatching = True
        bytecode(bytecodePtr) = REOP_REPEAT_MAX_GREEDY_INIT: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = REOP_REPEAT_MAX_GREEDY_START Or REOP_FLAG_POSSESSIVE: _
            bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = tmpCnt: bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        GoTo TurnToLeftChild
    End If

HandleRepeatExactlyPossessive:
    tmpCnt = ast.Buffer(curNode + 2)
    If returningFromFirstChild Then
        sp = sp - 1: patchPos = stack(sp)
        bytecode(bytecodePtr) = REOP_COMMIT_POSSESSIVE: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = REOP_REPEAT_EXACTLY_END: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = tmpCnt: bytecodePtr = bytecodePtr + 1
        tmp = bytecodePtr - patchPos
        bytecode(bytecodePtr) = tmp: bytecodePtr = bytecodePtr + 1
        bytecode(patchPos) = tmp
        GoTo TurnToParent
    Else
        requiresDfsMatching = True
        bytecode(bytecodePtr) = REOP_REPEAT_EXACTLY_INIT: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = REOP_REPEAT_EXACTLY_START Or REOP_FLAG_POSSESSIVE: _
            bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = tmpCnt: bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = REOP_FAIL: _
            bytecodePtr = bytecodePtr + 1
        GoTo TurnToLeftChild
    End If

HandleZeroone: ' requires: opcode1
    singlePath = False
    If returningFromFirstChild Then
        sp = sp - 1: patchPos = stack(sp)
        bytecode(patchPos) = bytecodePtr - patchPos - 1
        GoTo TurnToParent
    Else
        bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        GoTo TurnToLeftChild
    End If

HandleZeroonePossessive:
    If returningFromFirstChild Then
        sp = sp - 1: patchPos = stack(sp)
        bytecode(bytecodePtr) = REOP_COMMIT_POSSESSIVE: bytecodePtr = bytecodePtr + 1
        bytecode(patchPos) = bytecodePtr - patchPos - 1
        GoTo TurnToParent
    Else
        requiresDfsMatching = True
        bytecode(bytecodePtr) = REOP_SPLIT1 Or REOP_FLAG_POSSESSIVE: _
            bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        GoTo TurnToLeftChild
    End If

L_PLUS:
    singlePath = False
    If returningFromFirstChild Then
        sp = sp - 1: patchPos = stack(sp)
        If ast.Buffer(curNode) = AST_PLUS_GREEDY Then tmp = REOP_LOOP_GREEDY Else tmp = REOP_LOOP_HUMBLE
        bytecode(bytecodePtr) = tmp
        bytecode(bytecodePtr + 1) = patchPos - bytecodePtr - 2
        bytecodePtr = bytecodePtr + 2
        GoTo TurnToParent
    Else
        ' Besides distinguishing initial entry, the marker keeps DFS progress
        ' landmarks distinct when a plus begins with another plus.
        bytecode(bytecodePtr) = REOP_LOOP_INIT: bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1
        GoTo TurnToLeftChild
    End If

HandleStar:
    singlePath = False
    If returningFromFirstChild Then
        sp = sp - 1: patchPos = stack(sp)
        tmp = bytecodePtr - patchPos + 1
        If nodePlan(ast.Buffer(curNode + 1)) And NODE_NULLABLE Then
            ' Separate initial entry from repetition: (body+)? with one body.
            ' The post-body split saves the continuation with this iteration's
            ' captures before attempting another visit to the body.
            If ast.Buffer(curNode) = AST_STAR_GREEDY Then
                bytecode(bytecodePtr) = REOP_LOOP_GREEDY
            Else
                bytecode(bytecodePtr) = REOP_LOOP_HUMBLE
            End If
            bytecode(bytecodePtr + 1) = -tmp
        Else
            bytecode(bytecodePtr) = REOP_JUMP
            bytecode(bytecodePtr + 1) = -(tmp + 2)
        End If
        bytecodePtr = bytecodePtr + 2
        bytecode(patchPos) = tmp
        GoTo TurnToParent
    Else
        bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        GoTo TurnToLeftChild
    End If

HandleStarPossessive:
    If returningFromFirstChild Then
        sp = sp - 1: patchPos = stack(sp)
        bytecode(bytecodePtr) = REOP_COMMIT_POSSESSIVE: bytecodePtr = bytecodePtr + 1
        tmp = bytecodePtr - patchPos + 1
        bytecode(bytecodePtr) = REOP_JUMP: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = -(tmp + 2): bytecodePtr = bytecodePtr + 1
        bytecode(patchPos) = tmp
        GoTo TurnToParent
    Else
        requiresDfsMatching = True
        bytecode(bytecodePtr) = REOP_SPLIT1 Or REOP_FLAG_POSSESSIVE: _
            bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        GoTo TurnToLeftChild
    End If

HandleLookahead: ' requires opcode1, opcode2
    If returningFromFirstChild Then
        sp = sp - 1: direction = stack(sp)
        sp = sp - 1: patchPos = stack(sp)
        bytecode(bytecodePtr) = opcode2: bytecodePtr = bytecodePtr + 1
        bytecode(patchPos) = bytecodePtr - patchPos - 1
        GoTo TurnToParent
    Else
        bytecode(bytecodePtr) = REOP_CHECK_LOOKAHEAD: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        stack(sp) = direction: sp = sp + 1
        direction = 0
        GoTo TurnToLeftChild
    End If

HandleLookbehind: ' requires opcode1, opcode2
    If returningFromFirstChild Then
        sp = sp - 1: direction = stack(sp)
        sp = sp - 1: patchPos = stack(sp)
        bytecode(bytecodePtr) = opcode2: bytecodePtr = bytecodePtr + 1
        bytecode(patchPos) = bytecodePtr - patchPos - 1
        GoTo TurnToParent
    Else
        bytecode(bytecodePtr) = REOP_CHECK_LOOKBEHIND: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        stack(sp) = direction: sp = sp + 1
        direction = -1
        GoTo TurnToLeftChild
    End If

TurnToParent:
    prevNode = curNode
    If sp = 0 Then GoTo BreakLoop
    sp = sp - 1: tmp = stack(sp)
    curNode = tmp And RegexNumericConstants.LONG_ALL_BUT_FIRST_BIT
    returningFromFirstChild = tmp And RegexNumericConstants.LONG_FIRST_BIT
    GoTo ContinueLoop
TurnToLeftChild:
    prevNode = curNode
    stack(sp) = curNode Or RegexNumericConstants.LONG_FIRST_BIT: sp = sp + 1
    curNode = ast.Buffer(curNode + NODE_LCHILD)
    returningFromFirstChild = 0
    GoTo ContinueLoop
TurnToRightChild:
    prevNode = curNode
    stack(sp) = curNode: sp = sp + 1
    curNode = ast.Buffer(curNode + NODE_RCHILD)
    returningFromFirstChild = 0
    GoTo ContinueLoop
TurnToFirstChild:
    prevNode = curNode
    stack(sp) = curNode Or RegexNumericConstants.LONG_FIRST_BIT: sp = sp + 1
    curNode = ast.Buffer(curNode + NODE_LCHILD - direction)
    returningFromFirstChild = 0
    GoTo ContinueLoop
TurnToSecondChild:
    prevNode = curNode
    stack(sp) = curNode: sp = sp + 1
    curNode = ast.Buffer(curNode + NODE_RCHILD + direction)
    returningFromFirstChild = 0
    GoTo ContinueLoop
    
BreakLoop:
    If dfaBwRun Then
        bytecode(bytecodePtr) = REOP_MATCH
    Else
        bytecode(RegexBytecode.BYTECODE_IDX_MAX_PROPER_CAPTURE_SLOT) = maxSave
        bytecode(bytecodePtr) = REOP_MATCH: bytecodePtr = bytecodePtr + 1
        bytecode(RegexBytecode.BYTECODE_IDX_PROPERTIES) = _
            bytecode(RegexBytecode.BYTECODE_IDX_PROPERTIES) Or _
            (RegexBytecode.BYTECODE_PROPERTY_REQUIRES_DFS_MATCHING And _
                requiresDfsMatching)
        If Not requiresDfsMatching And singlePath Then
            If RegexMatcherCommon.LeadingStartAnchor(bytecode) <> 0 Then
                If Not RegexMatcherCommon.IsLiteralBytecode(bytecode) Then _
                    bytecode(RegexBytecode.BYTECODE_IDX_PROPERTIES) = bytecode(RegexBytecode.BYTECODE_IDX_PROPERTIES) Or BYTECODE_PROPERTY_PREFER_DFS
            End If
        End If
        If Not requiresDfsMatching Then
            dfaBwRun = True
            curNode = actualRoot
            direction = -1
            bytecode(RegexBytecode.BYTECODE_IDX_START_DFA_BW_RUN) = bytecodePtr
            GoTo StartBytecodeGeneration
        End If
        bytecode(RegexBytecode.BYTECODE_IDX_START_DFA_BW_RUN) = -1
    End If
End Sub

' This traversal reserves at least the stack space used by bytecode generation;
' ordinary repeats also reserve two sizing checkpoints for their child.
' This means we can abstain from checking stack capacities later.
Private Sub PrepareStackAndBytecodeBuffer( _
    ByRef ast As ArrayBuffer.Ty, _
    ByRef identifierTree As RegexIdentifierSupport.IdentifierTreeTy, _
    ByVal caseInsensitive As Boolean, _
    ByRef requiresDfsMatching As Boolean, _
    ByRef stack() As Long, _
    ByRef bytecode() As Long, ByRef nodePlan() As Long, ByVal maxExpansion As Long, _
    ByRef expandRepeats As Boolean _
)
    Dim sp As Long, prevNode As Long, curNode As Long, esfs As Long, stackCapacity As Long
    Dim tmp As Long, astTableIdx As Long, leftNode As Long, rightNode As Long, bodyNode As Long
    Dim bytecodeLength As Long, saveBytecodeLength As Long
    Dim extraLength As Long, extraSaves As Long, bodyLength As Long, splitLength As Long, expandedLength As Long
    Dim expandedRequiresDfs As Boolean, canExpand As Boolean, originalTotal As Double, expandedTotal As Double
    Dim returningFromFirstChild As Long ' 0 = no, RegexNumericConstants.LONG_FIRST_BIT = yes
    
    ' The initial stack capacity must be >= 2 * max([1 + entry.esfs for entry in AST_TABLE]),
    '   since when increasing the stack capacity, we increase by (current size \ 2) and
    '   we assume that this will be sufficient for the next stack frame.
    Const BYTECODE_GENERATOR_INITIAL_STACK_CAPACITY As Long = 8
    
    stackCapacity = BYTECODE_GENERATOR_INITIAL_STACK_CAPACITY
    ReDim stack(0 To BYTECODE_GENERATOR_INITIAL_STACK_CAPACITY - 1) As Long
    ReDim nodePlan(0 To ast.Length - 1) As Long
    canExpand = maxExpansion > 0

    sp = 0
    
    prevNode = -1
    curNode = ast.Buffer(0) ' first word contains index of root
    returningFromFirstChild = 0

    bytecodeLength = 0
    requiresDfsMatching = False
    saveBytecodeLength = 0
    
ContinueLoop:
    ' The parser represents A+ as a shared-body concatenation A A*. Normalize
    ' that representation here, after AST construction and within this existing
    ' traversal. For A{n,}, retain n-1 mandatory copies followed by A+.
    If ast.Buffer(curNode) = AST_CONCAT Then
        leftNode = ast.Buffer(curNode + 1): rightNode = ast.Buffer(curNode + 2)
        tmp = ast.Buffer(rightNode)
        If tmp = AST_STAR_GREEDY Or tmp = AST_STAR_HUMBLE Then
            bodyNode = ast.Buffer(rightNode + 1)
            If leftNode = bodyNode Then
                ast.Buffer(curNode) = AST_PLUS_GREEDY - (tmp = AST_STAR_HUMBLE)
            ElseIf ast.Buffer(leftNode) = AST_REPEAT_EXACTLY Then
                If ast.Buffer(leftNode + 1) = bodyNode Then
                    ast.Buffer(leftNode + 2) = ast.Buffer(leftNode + 2) - 1
                    ast.Buffer(rightNode) = AST_PLUS_GREEDY - (tmp = AST_STAR_HUMBLE)
                End If
            End If
        End If
    End If
    astTableIdx = ast.Buffer(curNode + NODE_TYPE) * AST_TABLE_ENTRY_LENGTH
    esfs = RegexUnicodeSupport.StaticData(AST_TABLE_START + astTableIdx + AST_TABLE_OFFSET_ESFS)
        
    On ast.Buffer(curNode + NODE_TYPE) GoTo _
        L_STRING, L_DISJ, L_CONCAT, L_CHAR, L_CAPTURE, _
        L_REPEAT_EXACTLY, L_REPEAT_EXACTLY_POSSESSIVE, L_PERIOD, L_ASSERT_START, _
            L_ASSERT_END, _
        L_ASSERT_WORD_BOUNDARY, L_ASSERT_NOT_WORD_BOUNDARY, L_MATCH, _
            L_ZEROONE_GREEDY, L_ZEROONE_HUMBLE, _
        L_ZEROONE_POSSESSIVE, L_STAR_GREEDY, L_STAR_HUMBLE, L_STAR_POSSESSIVE, _
            L_REPEAT_MAX_GREEDY, _
        L_REPEAT_MAX_HUMBLE, L_REPEAT_QUANTIFIED_POSSESSIVE, L_RANGES, L_INVRANGES, _
            L_ASSERT_POS_LOOKAHEAD, _
        L_ASSERT_NEG_LOOKAHEAD, L_ASSERT_POS_LOOKBEHIND, L_ASSERT_NEG_LOOKBEHIND, _
            L_FAIL, L_BACKREFERENCE, _
        L_NAMED, L_MODIFIER_SCOPE, L_EMPTY, L_CHAR_CONCAT, L_STRING_CONCAT, _
        L_ACTUAL, L_PLUS, L_PLUS
        
    Err.Raise REGEX_ERR_INTERNAL_LOGIC_ERR ' unreachable
  
L_NAMED:
    If returningFromFirstChild Then
        GoTo TurnToParent
    Else
        bytecodeLength = bytecodeLength + 3
        saveBytecodeLength = saveBytecodeLength + 3
        GoTo TurnToLeftChild
    End If

L_CAPTURE:
    If returningFromFirstChild Then
        GoTo TurnToParent
    Else
        bytecodeLength = bytecodeLength + 4
        saveBytecodeLength = saveBytecodeLength + 4
        GoTo TurnToLeftChild
    End If

L_ASSERT_START:
L_ASSERT_END:
L_ASSERT_WORD_BOUNDARY:
L_ASSERT_NOT_WORD_BOUNDARY:
    bytecodeLength = bytecodeLength + 1
    GoTo TurnToParent

L_STRING_CONCAT:
    If returningFromFirstChild Then
        GoTo TurnToParent
    Else
        bytecodeLength = bytecodeLength + 2 + ast.Buffer(curNode + 2)
        GoTo TurnToLeftChild
    End If

L_STRING:
    bytecodeLength = bytecodeLength + 2 + ast.Buffer(curNode + 1)
    GoTo TurnToParent

L_RANGES:
L_INVRANGES:
    bytecodeLength = bytecodeLength + 2 + 2 * ast.Buffer(curNode + 1)
    GoTo TurnToParent

L_BACKREFERENCE:
    requiresDfsMatching = True
    expandedRequiresDfs = True
    ' fallthrough

L_CHAR:
L_PERIOD:
L_MATCH:
L_FAIL:
L_EMPTY:
    bytecodeLength = bytecodeLength + RegexUnicodeSupport.StaticData(AST_TABLE_START + astTableIdx + AST_TABLE_OFFSET_BLEN)
    GoTo TurnToParent

L_REPEAT_EXACTLY_POSSESSIVE:
L_ZEROONE_POSSESSIVE:
L_STAR_POSSESSIVE:
L_REPEAT_QUANTIFIED_POSSESSIVE:
L_ASSERT_POS_LOOKAHEAD:
L_ASSERT_NEG_LOOKAHEAD:
L_ASSERT_POS_LOOKBEHIND:
L_ASSERT_NEG_LOOKBEHIND:
    requiresDfsMatching = True
    expandedRequiresDfs = True
    ' fallthrough
    
L_ZEROONE_GREEDY:
L_ZEROONE_HUMBLE:
L_STAR_GREEDY:
L_STAR_HUMBLE:
L_MODIFIER_SCOPE:
L_CHAR_CONCAT:
L_PLUS:
L_ACTUAL:
    If returningFromFirstChild Then
        GoTo TurnToParent
    Else
        bytecodeLength = bytecodeLength + RegexUnicodeSupport.StaticData(AST_TABLE_START + astTableIdx + AST_TABLE_OFFSET_BLEN)
        GoTo TurnToLeftChild
    End If

L_REPEAT_EXACTLY:
L_REPEAT_MAX_GREEDY:
L_REPEAT_MAX_HUMBLE:
    requiresDfsMatching = True
    If returningFromFirstChild Then
        If canExpand And (nodePlan(ast.Buffer(curNode + 1)) And (NODE_NONEMPTY Or NODE_UNSAFE_REPEAT)) = NODE_NONEMPTY Then
            bodyLength = bytecodeLength + extraLength - stack(sp - 2)
            splitLength = 0
            If ast.Buffer(curNode) <> AST_REPEAT_EXACTLY Then splitLength = 2
            tmp = bodyLength + splitLength
            expandedLength = 0
            If ast.Buffer(curNode) = AST_REPEAT_EXACTLY And (nodePlan(ast.Buffer(curNode + 1)) And NODE_LITERAL) <> 0 Then
                expandedLength = 2 ' one STRING header, followed by repeated data
                tmp = bodyLength - 2
                If bodyLength = 2 Then tmp = 1 ' CHAR
            End If
            ' Bound before multiplying, including already expanded children.
            If tmp <= maxExpansion - expandedLength Then
                If ast.Buffer(curNode + 2) <= (maxExpansion - expandedLength) \ tmp Then
                    nodePlan(curNode) = nodePlan(curNode) Or NODE_UNROLL
                    If expandedLength = 2 Then nodePlan(curNode) = nodePlan(curNode) Or NODE_LITERAL
                    expandedLength = expandedLength + ast.Buffer(curNode + 2) * tmp
                    tmp = ast.Buffer(curNode + 2) - 1
                    extraLength = extraLength + expandedLength - bodyLength - 7
                    extraSaves = extraSaves + tmp * (saveBytecodeLength + extraSaves - stack(sp - 1))
                    If extraLength > maxExpansion Then canExpand = False
                    GoTo TurnToParent
                End If
            End If
        End If
        expandedRequiresDfs = True
        GoTo TurnToParent
    Else
        bytecodeLength = bytecodeLength + 7
        GoTo TurnToLeftChild
    End If

L_DISJ:
L_CONCAT:
    If returningFromFirstChild Then ' previous was left child
        GoTo TurnToRightChild
    ElseIf prevNode = ast.Buffer(curNode + NODE_RCHILD) Then ' previous was right child
        GoTo TurnToParent
    Else ' previous was parent
        bytecodeLength = bytecodeLength + RegexUnicodeSupport.StaticData(AST_TABLE_START + astTableIdx + AST_TABLE_OFFSET_BLEN)
        GoTo TurnToLeftChild
    End If
    
    Err.Raise REGEX_ERR_INTERNAL_LOGIC_ERR ' unreachable

TurnToParent:
    ' Record subtree properties while returning in this same sizing traversal.
    Select Case ast.Buffer(curNode)
    Case AST_CHAR, AST_STRING
        nodePlan(curNode) = NODE_NONEMPTY Or NODE_LITERAL
    Case AST_PERIOD, AST_RANGES, AST_INVRANGES
        nodePlan(curNode) = NODE_NONEMPTY
    Case AST_CONCAT
        nodePlan(curNode) = (nodePlan(ast.Buffer(curNode + 1)) Or nodePlan(ast.Buffer(curNode + 2))) And (NODE_NONEMPTY Or NODE_UNSAFE_REPEAT)
    Case AST_DISJ
        nodePlan(curNode) = (nodePlan(ast.Buffer(curNode + 1)) And nodePlan(ast.Buffer(curNode + 2)) And NODE_NONEMPTY) Or _
            ((nodePlan(ast.Buffer(curNode + 1)) Or nodePlan(ast.Buffer(curNode + 2))) And NODE_UNSAFE_REPEAT)
    Case AST_CAPTURE, AST_NAMED, AST_ACTUAL, AST_PLUS_GREEDY, AST_PLUS_HUMBLE
        nodePlan(curNode) = (nodePlan(curNode) And NODE_UNROLL) Or (nodePlan(ast.Buffer(curNode + 1)) And (NODE_NONEMPTY Or NODE_UNSAFE_REPEAT))
        If ast.Buffer(curNode) = AST_PLUS_GREEDY Or ast.Buffer(curNode) = AST_PLUS_HUMBLE Then
            If (nodePlan(curNode) And NODE_NONEMPTY) = 0 Then nodePlan(curNode) = nodePlan(curNode) Or NODE_UNSAFE_REPEAT
        End If
    Case AST_MODIFIER_SCOPE
        nodePlan(curNode) = nodePlan(ast.Buffer(curNode + 1)) And (NODE_NONEMPTY Or NODE_UNSAFE_REPEAT Or NODE_LITERAL)
    Case AST_REPEAT_EXACTLY
        nodePlan(curNode) = (nodePlan(curNode) And (NODE_UNROLL Or NODE_LITERAL)) Or (nodePlan(ast.Buffer(curNode + 1)) And (NODE_NONEMPTY Or NODE_UNSAFE_REPEAT))
    Case AST_CHAR_CONCAT, AST_STRING_CONCAT
        nodePlan(curNode) = NODE_NONEMPTY Or (nodePlan(ast.Buffer(curNode + 1)) And NODE_UNSAFE_REPEAT)
    Case AST_ZEROONE_GREEDY, AST_ZEROONE_HUMBLE, AST_STAR_GREEDY, AST_STAR_HUMBLE, AST_REPEAT_MAX_GREEDY, AST_REPEAT_MAX_HUMBLE
        nodePlan(curNode) = (nodePlan(curNode) And NODE_UNROLL) Or (nodePlan(ast.Buffer(curNode + 1)) And NODE_UNSAFE_REPEAT)
        If (nodePlan(ast.Buffer(curNode + 1)) And NODE_NONEMPTY) = 0 Then nodePlan(curNode) = nodePlan(curNode) Or NODE_UNSAFE_REPEAT
    Case AST_BACKREFERENCE, AST_REPEAT_EXACTLY_POSSESSIVE, AST_REPEAT_MAX_POSSESSIVE, AST_ZEROONE_POSSESSIVE, AST_STAR_POSSESSIVE, _
        AST_ASSERT_POS_LOOKAHEAD, AST_ASSERT_NEG_LOOKAHEAD, AST_ASSERT_POS_LOOKBEHIND, AST_ASSERT_NEG_LOOKBEHIND
        nodePlan(curNode) = NODE_UNSAFE_REPEAT
    End Select
    ' Nullability is computed on return, alongside the existing expansion facts.
    ' FAIL is not nullable even though it also cannot consume a character.
    Select Case ast.Buffer(curNode)
    Case AST_EMPTY, AST_ASSERT_START, AST_ASSERT_END, AST_ASSERT_WORD_BOUNDARY, AST_ASSERT_NOT_WORD_BOUNDARY, _
        AST_ZEROONE_GREEDY, AST_ZEROONE_HUMBLE, AST_STAR_GREEDY, AST_STAR_HUMBLE, _
        AST_REPEAT_MAX_GREEDY, AST_REPEAT_MAX_HUMBLE
        nodePlan(curNode) = nodePlan(curNode) Or NODE_NULLABLE
    Case AST_CONCAT
        nodePlan(curNode) = nodePlan(curNode) Or (nodePlan(ast.Buffer(curNode + 1)) And nodePlan(ast.Buffer(curNode + 2)) And NODE_NULLABLE)
    Case AST_DISJ
        nodePlan(curNode) = nodePlan(curNode) Or ((nodePlan(ast.Buffer(curNode + 1)) Or nodePlan(ast.Buffer(curNode + 2))) And NODE_NULLABLE)
    Case AST_CAPTURE, AST_NAMED, AST_ACTUAL, AST_MODIFIER_SCOPE, AST_PLUS_GREEDY, AST_PLUS_HUMBLE, AST_REPEAT_EXACTLY
        nodePlan(curNode) = nodePlan(curNode) Or (nodePlan(ast.Buffer(curNode + 1)) And NODE_NULLABLE)
    End Select
    sp = sp - esfs
    If sp = 0 Then GoTo BreakLoop
    prevNode = curNode
    sp = sp - 1: tmp = stack(sp)
    returningFromFirstChild = tmp And RegexNumericConstants.LONG_FIRST_BIT: curNode = tmp And RegexNumericConstants.LONG_ALL_BUT_FIRST_BIT
    GoTo ContinueLoop
TurnToLeftChild:
    If sp >= stackCapacity - esfs Then
        stackCapacity = stackCapacity + stackCapacity \ 2
        ReDim Preserve stack(0 To stackCapacity - 1) As Long
    End If
    Select Case ast.Buffer(curNode)
    Case AST_REPEAT_EXACTLY, AST_REPEAT_MAX_GREEDY, AST_REPEAT_MAX_HUMBLE
        If canExpand Then
            stack(sp) = bytecodeLength + extraLength
            stack(sp + 1) = saveBytecodeLength + extraSaves
        End If
    End Select
    prevNode = curNode
    sp = sp + esfs: stack(sp) = curNode Or RegexNumericConstants.LONG_FIRST_BIT: sp = sp + 1
    returningFromFirstChild = 0: curNode = ast.Buffer(curNode + NODE_LCHILD)
    GoTo ContinueLoop
TurnToRightChild:
    prevNode = curNode
    stack(sp) = curNode: sp = sp + 1
    returningFromFirstChild = 0: curNode = ast.Buffer(curNode + NODE_RCHILD)
    GoTo ContinueLoop
    
BreakLoop:
    originalTotal = bytecodeLength
    ' Todo: Subtracting 4 for the prefix is a hack!
    If Not requiresDfsMatching Then _
        originalTotal = originalTotal + (bytecodeLength - saveBytecodeLength - 4)
    expandedTotal = CDbl(bytecodeLength) + extraLength
    If Not expandedRequiresDfs Then expandedTotal = expandedTotal + (expandedTotal - saveBytecodeLength - extraSaves - 4)
    If canExpand And expandedTotal <= originalTotal + maxExpansion And _
        expandedTotal <= RegexNumericConstants.LONG_MAX - 1 - RegexBytecode.BYTECODE_FIXED_PROLOGUE_LAST_IDX - 3 * identifierTree.nEntries Then
        expandRepeats = True
        bytecodeLength = CLng(expandedTotal)
        requiresDfsMatching = expandedRequiresDfs
    Else
        ' Disable tentative expansion globally, retaining nullability facts.
        ' No second traversal (or scan to clear individual plan entries) is needed.
        expandRepeats = False
        bytecodeLength = CLng(originalTotal)
    End If
        
    ' Actual bytecode length is
    '     bytecodeLength + fixed prologue length + 1 + 3*identifierTree(N_NODES)
    ' due to captures and final REOP_MATCH.
    ReDim bytecode(0 To bytecodeLength + 1 + RegexBytecode.BYTECODE_FIXED_PROLOGUE_LAST_IDX _
        + 3 * identifierTree.nEntries) As Long
    bytecode(RegexBytecode.BYTECODE_IDX_N_IDENTIFIERS) = identifierTree.nEntries
    bytecode(RegexBytecode.BYTECODE_IDX_PROPERTIES) = _
        RegexBytecode.BYTECODE_PROPERTY_CASE_INSENSITIVE And caseInsensitive
    bytecode(RegexBytecode.BYTECODE_IDX_REOP_PERFORM_ASSERT) = RegexBytecode.REOP_PERFORM_ASSERT
    bytecode(RegexBytecode.BYTECODE_IDX_REOP_MATCH_PREV) = RegexBytecode.REOP_MATCH_PREV
    bytecode(RegexBytecode.BYTECODE_IDX_REOP_MATCH_OLD) = RegexBytecode.REOP_MATCH_OLD
    RegexIdentifierSupport.RedBlackDumpTree bytecode, RegexBytecode.BYTECODE_IDENTIFIER_MAP_BEGIN, identifierTree
End Sub
