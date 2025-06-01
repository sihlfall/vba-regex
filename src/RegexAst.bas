Attribute VB_Name = "RegexAst"
Option Explicit

Public Enum AstNodeType
    ' We guarantee: All AST_ constants are >= 0.
    ' ! The parser relies on AST_ASSERT_LOOKAHEAD and LOOKBEHIND being >= 0.
    MIN_AST_CODE = 0
    AST_EMPTY = 0
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
    MAX_AST_CODE = 32
End Enum

Private Enum AstNodeDescriptionConstant
    NODE_TYPE = 0
    NODE_LCHILD = 1
    NODE_RCHILD = 2
End Enum

Public Enum AstTableDescriptionConstant
    AST_TABLE_OFFSET_NC = 0
    AST_TABLE_OFFSET_BLEN = 1
    AST_TABLE_OFFSET_ESFS = 2
    AST_TABLE_ENTRY_LENGTH = 3
    AST_TABLE_LENGTH = AST_TABLE_ENTRY_LENGTH * (MAX_AST_CODE + 1)
End Enum

Private astTableInitialized As Boolean ' default-initialized to False

Public Sub AstTableInitialize()
    InitializeAstTable RegexUnicodeSupport.StaticData
End Sub

Private Sub InitializeAstTable(ByRef t() As Long)
    Const b As Long = RegexUnicodeSupport.AST_TABLE_START
    Const nc As Long = b + AST_TABLE_OFFSET_NC
    Const blen As Long = b + AST_TABLE_OFFSET_BLEN
    Const esfs As Long = b + AST_TABLE_OFFSET_ESFS
    Const e As Long = AST_TABLE_ENTRY_LENGTH
    
    ' nc: number of children; negative values have special meaning
    '   -2: is AST_STRING
    '   -1: is AST_RANGES or AST_INVRANGES
    ' blen: length of bytecode generated for this node (meaningful only if .nc >= 0)
    ' esfs: extra stack space required when generating bytecode for this node
    '   Only nodes with children are permitted to require extra stack space.
    '   Hence .esfs > 0 must imply .nc >= 1.
    
    ' ! When adding new entries, make sure to adjust BYTECODE_GENERATOR_INITIAL_STACK_CAPACITY below, if necessary!
    ' ! See comment on BYTECODE_GENERATOR_INITIAL_STACK_CAPACITY below.
    
    t(nc + e * AST_EMPTY) = 0:                    t(blen + e * AST_EMPTY) = 0:                        t(esfs + e * AST_EMPTY) = 0
    t(nc + e * AST_STRING) = -2:                  t(blen + e * AST_STRING) = 2:                       t(esfs + e * AST_STRING) = 0
    t(nc + e * AST_DISJ) = 2:                     t(blen + e * AST_DISJ) = 4:                         t(esfs + e * AST_DISJ) = 1
    t(nc + e * AST_CONCAT) = 2:                   t(blen + e * AST_CONCAT) = 0:                       t(esfs + e * AST_CONCAT) = 0
    t(nc + e * AST_CHAR) = 0:                     t(blen + e * AST_CHAR) = 2:                         t(esfs + e * AST_CHAR) = 0
    t(nc + e * AST_CAPTURE) = 1:                  t(blen + e * AST_CAPTURE) = 4:                      t(esfs + e * AST_CAPTURE) = 0
    t(nc + e * AST_REPEAT_EXACTLY) = 1:           t(blen + e * AST_REPEAT_EXACTLY) = 7:               t(esfs + e * AST_REPEAT_EXACTLY) = 1
    t(nc + e * AST_REPEAT_EXACTLY_POSSESSIVE) = 1: t(blen + e * AST_REPEAT_EXACTLY_POSSESSIVE) = 9:   t(esfs + e * AST_REPEAT_EXACTLY_POSSESSIVE) = 1
    t(nc + e * AST_PERIOD) = 0:                   t(blen + e * AST_PERIOD) = 1:                       t(esfs + e * AST_PERIOD) = 0
    t(nc + e * AST_ASSERT_START) = 0:             t(blen + e * AST_ASSERT_START) = 1:                 t(esfs + e * AST_ASSERT_START) = 0
    t(nc + e * AST_ASSERT_END) = 0:               t(blen + e * AST_ASSERT_END) = 1:                   t(esfs + e * AST_ASSERT_END) = 0
    t(nc + e * AST_ASSERT_WORD_BOUNDARY) = 0:     t(blen + e * AST_ASSERT_WORD_BOUNDARY) = 1:         t(esfs + e * AST_ASSERT_WORD_BOUNDARY) = 0
    t(nc + e * AST_ASSERT_NOT_WORD_BOUNDARY) = 0: t(blen + e * AST_ASSERT_NOT_WORD_BOUNDARY) = 1:     t(esfs + e * AST_ASSERT_NOT_WORD_BOUNDARY) = 0
    t(nc + e * AST_MATCH) = 0:                    t(blen + e * AST_MATCH) = 1:                        t(esfs + e * AST_MATCH) = 0
    t(nc + e * AST_ZEROONE_GREEDY) = 1:           t(blen + e * AST_ZEROONE_GREEDY) = 2:               t(esfs + e * AST_ZEROONE_GREEDY) = 1
    t(nc + e * AST_ZEROONE_HUMBLE) = 1:           t(blen + e * AST_ZEROONE_HUMBLE) = 2:               t(esfs + e * AST_ZEROONE_HUMBLE) = 1
    t(nc + e * AST_ZEROONE_POSSESSIVE) = 1:       t(blen + e * AST_ZEROONE_POSSESSIVE) = 3:           t(esfs + e * AST_ZEROONE_POSSESSIVE) = 1
    t(nc + e * AST_STAR_GREEDY) = 1:              t(blen + e * AST_STAR_GREEDY) = 4:                  t(esfs + e * AST_STAR_GREEDY) = 1
    t(nc + e * AST_STAR_HUMBLE) = 1:              t(blen + e * AST_STAR_HUMBLE) = 4:                  t(esfs + e * AST_STAR_HUMBLE) = 1
    t(nc + e * AST_STAR_POSSESSIVE) = 1:          t(blen + e * AST_STAR_POSSESSIVE) = 5:              t(esfs + e * AST_STAR_POSSESSIVE) = 1
    t(nc + e * AST_REPEAT_MAX_GREEDY) = 1:        t(blen + e * AST_REPEAT_MAX_GREEDY) = 7:            t(esfs + e * AST_REPEAT_MAX_GREEDY) = 1
    t(nc + e * AST_REPEAT_MAX_HUMBLE) = 1:        t(blen + e * AST_REPEAT_MAX_HUMBLE) = 7:            t(esfs + e * AST_REPEAT_MAX_HUMBLE) = 1
    t(nc + e * AST_REPEAT_MAX_POSSESSIVE) = 1:    t(blen + e * AST_REPEAT_MAX_POSSESSIVE) = 8:        t(esfs + e * AST_REPEAT_MAX_POSSESSIVE) = 1
    t(nc + e * AST_RANGES) = -1:                  t(blen + e * AST_RANGES) = 2:                       t(esfs + e * AST_RANGES) = 0
    t(nc + e * AST_INVRANGES) = -1:               t(blen + e * AST_INVRANGES) = 2:                    t(esfs + e * AST_INVRANGES) = 0
    t(nc + e * AST_ASSERT_POS_LOOKAHEAD) = 1:     t(blen + e * AST_ASSERT_POS_LOOKAHEAD) = 4:         t(esfs + e * AST_ASSERT_POS_LOOKAHEAD) = 2
    t(nc + e * AST_ASSERT_NEG_LOOKAHEAD) = 1:     t(blen + e * AST_ASSERT_NEG_LOOKAHEAD) = 4:         t(esfs + e * AST_ASSERT_NEG_LOOKAHEAD) = 2
    t(nc + e * AST_ASSERT_POS_LOOKBEHIND) = 1:    t(blen + e * AST_ASSERT_POS_LOOKBEHIND) = 4:        t(esfs + e * AST_ASSERT_POS_LOOKBEHIND) = 2
    t(nc + e * AST_ASSERT_NEG_LOOKBEHIND) = 1:    t(blen + e * AST_ASSERT_NEG_LOOKBEHIND) = 4:        t(esfs + e * AST_ASSERT_NEG_LOOKBEHIND) = 2
    t(nc + e * AST_FAIL) = 0:                     t(blen + e * AST_FAIL) = 1:                         t(esfs + e * AST_FAIL) = 0
    t(nc + e * AST_BACKREFERENCE) = 0:            t(blen + e * AST_BACKREFERENCE) = 2:                t(esfs + e * AST_BACKREFERENCE) = 0
    t(nc + e * AST_NAMED) = 1:                    t(blen + e * AST_NAMED) = 3:                        t(esfs + e * AST_NAMED) = 0
    t(nc + e * AST_MODIFIER_SCOPE) = 1:           t(blen + e * AST_MODIFIER_SCOPE) = 0:               t(esfs + e * AST_MODIFIER_SCOPE) = 1
End Sub

Public Sub AstToBytecode(ByRef ast() As Long, ByRef identifierTree As RegexIdentifierSupport.IdentifierTreeTy, ByVal caseInsensitive As Boolean, ByRef bytecode() As Long)
    Dim bytecodePtr As Long
    Dim curNode As Long, prevNode As Long
    Dim stack() As Long, sp As Long
    Dim direction As Long ' 0 = left before right, -1 = right before left
    Dim returningFromFirstChild As Long ' 0 = no, RegexNumericConstants.LONG_FIRST_BIT = yes
    Dim modifiers As Long
    
    ' temporaries, do not survive over more than one iteration
    Dim opcode1 As Long, opcode2 As Long, opcode3 As Long, tmp As Long, tmpCnt As Long, _
        e As Long, j As Long, patchPos As Long, maxSave As Long
    
    If Not astTableInitialized Then AstTableInitialize
    
    PrepareStackAndBytecodeBuffer ast, identifierTree, caseInsensitive, stack, bytecode
    
    modifiers = caseInsensitive And RegexBytecode.MODIFIER_I_MASK
    sp = 0
    
    prevNode = -1
    curNode = ast(0) ' first word contains index of root
    bytecodePtr = 3 + 3 * bytecode(1)
    maxSave = -1
    direction = 0
    returningFromFirstChild = 0

ContinueLoop:
        Select Case ast(curNode + NODE_TYPE)
        Case AST_STRING
            tmpCnt = ast(curNode + 1) ' assert(tmpCnt >= 1)
            j = curNode + 2 + ((tmpCnt - 1) And direction)
            e = curNode + 1 + tmpCnt - ((tmpCnt - 1) And direction)
            tmp = 1 + 2 * direction
            Do
                bytecode(bytecodePtr) = REOP_CHAR Or modifiers And RegexBytecode.MODIFIER_I_ACTIVE
                bytecodePtr = bytecodePtr + 1
                bytecode(bytecodePtr) = ast(j)
                bytecodePtr = bytecodePtr + 1
                If j = e Then Exit Do
                j = j + tmp
            Loop
            GoTo TurnToParent
        Case AST_RANGES
            opcode1 = REOP_RANGES Or modifiers And RegexBytecode.MODIFIER_I_ACTIVE
            GoTo HandleRanges
        Case AST_INVRANGES
            opcode1 = REOP_INVRANGES Or modifiers And RegexBytecode.MODIFIER_I_ACTIVE
            GoTo HandleRanges
        Case AST_CHAR
            bytecode(bytecodePtr) = REOP_CHAR Or modifiers And RegexBytecode.MODIFIER_I_ACTIVE
            bytecodePtr = bytecodePtr + 1
            bytecode(bytecodePtr) = ast(curNode + 1)
            bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_PERIOD
            bytecode(bytecodePtr) = REOP_DOT Or modifiers And RegexBytecode.MODIFIER_S_MASK
            bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_MATCH
            bytecode(bytecodePtr) = REOP_MATCH: bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_ASSERT_START
            bytecode(bytecodePtr) = REOP_ASSERT_START Or modifiers And RegexBytecode.MODIFIER_M_MASK
            bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_ASSERT_END
            bytecode(bytecodePtr) = REOP_ASSERT_END Or modifiers And RegexBytecode.MODIFIER_M_MASK
            bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_ASSERT_WORD_BOUNDARY
            bytecode(bytecodePtr) = REOP_ASSERT_WORD_BOUNDARY: bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_ASSERT_NOT_WORD_BOUNDARY
            bytecode(bytecodePtr) = REOP_ASSERT_NOT_WORD_BOUNDARY: bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_DISJ
            If returningFromFirstChild Then ' previous was left child
                sp = sp - 1: patchPos = stack(sp)
                bytecode(bytecodePtr) = REOP_JUMP: bytecodePtr = bytecodePtr + 1
                stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
                bytecode(patchPos) = bytecodePtr - patchPos - 1
                
                GoTo TurnToRightChild
            ElseIf prevNode = ast(curNode + NODE_RCHILD) Then ' previous was right child
                sp = sp - 1: patchPos = stack(sp)
                bytecode(patchPos) = bytecodePtr - patchPos - 1
            
                GoTo TurnToParent
            Else ' previous was parent
                bytecode(bytecodePtr) = REOP_SPLIT1: bytecodePtr = bytecodePtr + 1
                stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
                GoTo TurnToLeftChild
            End If
        Case AST_CONCAT
            If returningFromFirstChild Then ' previous was first child
                GoTo TurnToSecondChild
            ElseIf prevNode = ast(curNode + NODE_RCHILD + direction) Then ' previous was second child
                GoTo TurnToParent
            Else ' previous was parent
                GoTo TurnToFirstChild
            End If
        Case AST_CAPTURE
            If returningFromFirstChild Then
                bytecode(bytecodePtr) = REOP_SAVE: bytecodePtr = bytecodePtr + 1
                tmp = ast(curNode + 2) * 2 + 1
                If tmp > maxSave Then maxSave = tmp
                bytecode(bytecodePtr) = tmp + direction: bytecodePtr = bytecodePtr + 1
                GoTo TurnToParent
            Else
                bytecode(bytecodePtr) = REOP_SAVE: bytecodePtr = bytecodePtr + 1
                bytecode(bytecodePtr) = ast(curNode + 2) * 2 - direction: bytecodePtr = bytecodePtr + 1
                GoTo TurnToLeftChild
            End If
        Case AST_REPEAT_EXACTLY
            opcode1 = REOP_REPEAT_EXACTLY_INIT: opcode2 = REOP_REPEAT_EXACTLY_START: opcode3 = REOP_REPEAT_EXACTLY_END
            GoTo HandleRepeatQuantified
        Case AST_REPEAT_EXACTLY_POSSESSIVE
            GoTo HandleRepeatExactlyPossessive
        Case AST_REPEAT_MAX_GREEDY
            opcode1 = REOP_REPEAT_GREEDY_MAX_INIT: opcode2 = REOP_REPEAT_GREEDY_MAX_START: opcode3 = REOP_REPEAT_GREEDY_MAX_END
            GoTo HandleRepeatQuantified
        Case AST_REPEAT_MAX_HUMBLE
            opcode1 = REOP_REPEAT_MAX_HUMBLE_INIT: opcode2 = REOP_REPEAT_MAX_HUMBLE_START: opcode3 = REOP_REPEAT_MAX_HUMBLE_END
            GoTo HandleRepeatQuantified
        Case AST_REPEAT_MAX_POSSESSIVE
            GoTo HandleRepeatQuantifiedPossessive
        Case AST_ZEROONE_GREEDY
            opcode1 = REOP_SPLIT1
            GoTo HandleZeroone
        Case AST_ZEROONE_HUMBLE
            opcode1 = REOP_SPLIT2
            GoTo HandleZeroone
        Case AST_ZEROONE_POSSESSIVE
            GoTo HandleZeroonePossessive
        Case AST_STAR_GREEDY
            opcode1 = REOP_SPLIT1
            GoTo HandleStar
        Case AST_STAR_HUMBLE
            opcode1 = REOP_SPLIT2
            GoTo HandleStar
        Case AST_STAR_POSSESSIVE
            GoTo HandleStarPossessive
        Case AST_ASSERT_POS_LOOKAHEAD
            opcode1 = REOP_LOOKPOS: opcode2 = REOP_END_LOOKPOS
            GoTo HandleLookahead
        Case AST_ASSERT_NEG_LOOKAHEAD
            opcode1 = REOP_LOOKNEG: opcode2 = REOP_END_LOOKNEG
            GoTo HandleLookahead
        Case AST_ASSERT_POS_LOOKBEHIND
            opcode1 = REOP_LOOKPOS: opcode2 = REOP_END_LOOKPOS
            GoTo HandleLookbehind
        Case AST_ASSERT_NEG_LOOKBEHIND
            opcode1 = REOP_LOOKNEG: opcode2 = REOP_END_LOOKNEG
            GoTo HandleLookbehind
        Case AST_EMPTY
            GoTo TurnToParent
        Case AST_FAIL
            bytecode(bytecodePtr) = REOP_FAIL: bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_BACKREFERENCE
            bytecode(bytecodePtr) = REOP_BACKREFERENCE Or modifiers And RegexBytecode.MODIFIER_I_ACTIVE
            bytecodePtr = bytecodePtr + 1
            bytecode(bytecodePtr) = ast(curNode + 1): bytecodePtr = bytecodePtr + 1
            GoTo TurnToParent
        Case AST_NAMED
            If returningFromFirstChild Then
                ' nothing to be done
                GoTo TurnToParent
            Else
                bytecode(bytecodePtr) = REOP_SET_NAMED: bytecodePtr = bytecodePtr + 1
                bytecode(bytecodePtr) = ast(curNode + 2): bytecodePtr = bytecodePtr + 1
                bytecode(bytecodePtr) = ast(curNode + 3): bytecodePtr = bytecodePtr + 1
                GoTo TurnToLeftChild
            End If
        Case AST_MODIFIER_SCOPE
            If returningFromFirstChild Then
                sp = sp - 1: modifiers = stack(sp)
                GoTo TurnToParent
            Else
                stack(sp) = modifiers: sp = sp + 1
                tmp = ast(curNode + 2)
                modifiers = RegexBytecode.MODIFIER_WRITE_MASK And (modifiers Or tmp) Or _
                    RegexBytecode.MODIFIER_ACTIVE_MASK And ( _
                        modifiers Xor tmp * 2 And (tmp Xor modifiers) _
                    )
                GoTo TurnToLeftChild
            End If
        End Select
        
        Err.Raise REGEX_ERR_INTERNAL_LOGIC_ERR ' unreachable
        
HandleRanges: ' requires: opcode1
        tmpCnt = ast(curNode + 1)
        bytecode(bytecodePtr) = opcode1: bytecodePtr = bytecodePtr + 1
        j = curNode
        e = curNode + 1 + 2 * tmpCnt
        Do ' copy everything, including first word, which is the length
            j = j + 1
            bytecode(bytecodePtr) = ast(j): bytecodePtr = bytecodePtr + 1
        Loop Until j = e
        GoTo TurnToParent

HandleRepeatQuantified: ' requires: opcode1, opcode2, opcode 3
        tmpCnt = ast(curNode + 2)
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
        
HandleRepeatQuantifiedPossessive:
        tmpCnt = ast(curNode + 2)
        If returningFromFirstChild Then
            sp = sp - 1: patchPos = stack(sp)
            bytecode(bytecodePtr) = REOP_COMMIT_POSSESSIVE: bytecodePtr = bytecodePtr + 1
            bytecode(bytecodePtr) = REOP_REPEAT_GREEDY_MAX_END: bytecodePtr = bytecodePtr + 1
            bytecode(bytecodePtr) = tmpCnt: bytecodePtr = bytecodePtr + 1
            tmp = bytecodePtr - patchPos
            bytecode(bytecodePtr) = tmp: bytecodePtr = bytecodePtr + 1
            bytecode(patchPos) = tmp
            GoTo TurnToParent
        Else
            bytecode(bytecodePtr) = REOP_REPEAT_GREEDY_MAX_INIT: bytecodePtr = bytecodePtr + 1
            bytecode(bytecodePtr) = REOP_REPEAT_GREEDY_MAX_START Or REOP_FLAG_POSSESSIVE: _
                bytecodePtr = bytecodePtr + 1
            bytecode(bytecodePtr) = tmpCnt: bytecodePtr = bytecodePtr + 1
            stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
            GoTo TurnToLeftChild
        End If

HandleRepeatExactlyPossessive:
        tmpCnt = ast(curNode + 2)
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
        bytecode(bytecodePtr) = REOP_SPLIT1 Or REOP_FLAG_POSSESSIVE: _
            bytecodePtr = bytecodePtr + 1
        stack(sp) = bytecodePtr: sp = sp + 1: bytecodePtr = bytecodePtr + 1
        GoTo TurnToLeftChild
    End If

HandleStar:
    If returningFromFirstChild Then
        sp = sp - 1: patchPos = stack(sp)
        tmp = bytecodePtr - patchPos + 1
        bytecode(bytecodePtr) = REOP_JUMP: bytecodePtr = bytecodePtr + 1
        bytecode(bytecodePtr) = -(tmp + 2): bytecodePtr = bytecodePtr + 1
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
    curNode = ast(curNode + NODE_LCHILD)
    returningFromFirstChild = 0
    GoTo ContinueLoop
TurnToRightChild:
    prevNode = curNode
    stack(sp) = curNode: sp = sp + 1
    curNode = ast(curNode + NODE_RCHILD)
    returningFromFirstChild = 0
    GoTo ContinueLoop
TurnToFirstChild:
    prevNode = curNode
    stack(sp) = curNode Or RegexNumericConstants.LONG_FIRST_BIT: sp = sp + 1
    curNode = ast(curNode + NODE_LCHILD - direction)
    returningFromFirstChild = 0
    GoTo ContinueLoop
TurnToSecondChild:
    prevNode = curNode
    stack(sp) = curNode: sp = sp + 1
    curNode = ast(curNode + NODE_RCHILD + direction)
    returningFromFirstChild = 0
    GoTo ContinueLoop
    
BreakLoop:
    bytecode(0) = maxSave
    bytecode(bytecodePtr) = REOP_MATCH
End Sub

' In this function, we allocate stack frames of the same size as we will in bytecode generation.
' Thus we simulate stack usage and make sure the stack is sufficiently large for bytecode generation.
' This means we can abstain from checking stack capacities later.
Private Sub PrepareStackAndBytecodeBuffer(ByRef ast() As Long, ByRef identifierTree As RegexIdentifierSupport.IdentifierTreeTy, ByVal caseInsensitive As Boolean, ByRef stack() As Long, ByRef bytecode() As Long)
    Dim sp As Long, prevNode As Long, curNode As Long, esfs As Long, stackCapacity As Long
    Dim tmp As Long, astTableIdx As Long
    Dim bytecodeLength As Long
    Dim returningFromFirstChild As Long ' 0 = no, RegexNumericConstants.LONG_FIRST_BIT = yes
    
    ' The initial stack capacity must be >= 2 * max([1 + entry.esfs for entry in AST_TABLE]),
    '   since when increasing the stack capacity, we increase by (current size \ 2) and
    '   we assume that this will be sufficient for the next stack frame.
    Const BYTECODE_GENERATOR_INITIAL_STACK_CAPACITY As Long = 8
    
    stackCapacity = BYTECODE_GENERATOR_INITIAL_STACK_CAPACITY
    ReDim stack(0 To BYTECODE_GENERATOR_INITIAL_STACK_CAPACITY - 1) As Long

    sp = 0
    
    prevNode = -1
    curNode = ast(0) ' first word contains index of root
    returningFromFirstChild = 0

    bytecodeLength = 0
    
ContinueLoop:
        astTableIdx = ast(curNode + NODE_TYPE) * AST_TABLE_ENTRY_LENGTH
        esfs = RegexUnicodeSupport.StaticData(AST_TABLE_START + astTableIdx + AST_TABLE_OFFSET_ESFS)
        
        Select Case RegexUnicodeSupport.StaticData(AST_TABLE_START + astTableIdx + AST_TABLE_OFFSET_NC)
        Case -2
            bytecodeLength = bytecodeLength + 2 * ast(curNode + 1)
            GoTo TurnToParent
        Case -1
            bytecodeLength = bytecodeLength + 2 + 2 * ast(curNode + 1)
            GoTo TurnToParent
        Case 0
            bytecodeLength = bytecodeLength + RegexUnicodeSupport.StaticData(AST_TABLE_START + astTableIdx + AST_TABLE_OFFSET_BLEN)
            GoTo TurnToParent
        Case 1
            If returningFromFirstChild Then
                GoTo TurnToParent
            Else
                bytecodeLength = bytecodeLength + RegexUnicodeSupport.StaticData(AST_TABLE_START + astTableIdx + AST_TABLE_OFFSET_BLEN)
                GoTo TurnToLeftChild
            End If
        Case 2
            If returningFromFirstChild Then ' previous was left child
                GoTo TurnToRightChild
            ElseIf prevNode = ast(curNode + NODE_RCHILD) Then ' previous was right child
                GoTo TurnToParent
            Else ' previous was parent
                bytecodeLength = bytecodeLength + RegexUnicodeSupport.StaticData(AST_TABLE_START + astTableIdx + AST_TABLE_OFFSET_BLEN)
                GoTo TurnToLeftChild
            End If
        End Select
        
        Err.Raise REGEX_ERR_INTERNAL_LOGIC_ERR ' unreachable

TurnToParent:
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
    prevNode = curNode
    sp = sp + esfs: stack(sp) = curNode Or RegexNumericConstants.LONG_FIRST_BIT: sp = sp + 1
    returningFromFirstChild = 0: curNode = ast(curNode + NODE_LCHILD)
    GoTo ContinueLoop
TurnToRightChild:
    prevNode = curNode
    stack(sp) = curNode: sp = sp + 1
    returningFromFirstChild = 0: curNode = ast(curNode + NODE_RCHILD)
    GoTo ContinueLoop
    
BreakLoop:
    ' Actual bytecode length is bytecodeLength + 4 + 3*identifierTree(N_NODES) due to intial nCaptures and final REOP_MATCH.
    ReDim bytecode(0 To bytecodeLength + 3 + 3 * identifierTree.nEntries) As Long
    bytecode(RegexBytecode.BYTECODE_IDX_N_IDENTIFIERS) = identifierTree.nEntries
    bytecode(RegexBytecode.BYTECODE_IDX_CASE_INSENSITIVE_INDICATOR) = -caseInsensitive
    RegexIdentifierSupport.RedBlackDumpTree bytecode, RegexBytecode.BYTECODE_IDENTIFIER_MAP_BEGIN, identifierTree
End Sub
