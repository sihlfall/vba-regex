Attribute VB_Name = "RegexCompiler"
Option Explicit

' The last word of each parse stack frame indicates the type t of the frame.
' t > 0 -> capturing group
'   In this case, t indicates the number of the capture.
' t = - AST_ASSERT_POS_LOOKAHEAD -> lookahead                                         |
' t = - AST_ASSERT_NEG_LOOKAHEAD -> lookbehind                                        |
' t = - AST_ASSERT_POS_LOOKBEHIND -> lookahead                                        |
' t = - AST_ASSERT_NEG_LOOKBEHIND -> lookbehind                                       |
' t = - (MAX_AST_CODE + 1) -> atomic group                     PSF_ATOMIC             |
' t = - (MAX_AST_CODE + 2) -> non-capturing group              PSF_NONCAPTURE         | PSF_MIN_EXPLICIT   | PSF_MAX_WITH_MODIFIERS
' t = - (MAX_AST_CODE + 3) -> modifier scope not ending at |   PSF_MODSCOPE_SPANNING                       |
' t = - (MAX_AST_CODE + 4) -> modifier scope ending at |       PSF_MODSCOPE_LOCAL                          |
Private Enum ParseStackFrameTypeConstant
    PSF_ATOMIC = -(RegexAst.MAX_AST_CODE + 1)
    
    PSF_NONCAPTURE = -(RegexAst.MAX_AST_CODE + 2)
    PSF_MAX_WITH_MODIFIERS = PSF_NONCAPTURE
    PSF_MIN_EXPLICIT = PSF_NONCAPTURE
    
    PSF_MODSCOPE_SPANNING = -(RegexAst.MAX_AST_CODE + 3)
    PSF_MODSCOPE_LOCAL = -(RegexAst.MAX_AST_CODE + 4)
    
    PSF_NONE = RegexNumericConstants.LONG_MIN ' not a valid stack frame type, for indicating nothing has been handled
    
    HANDLE_IMPLICIT_LOCAL = PSF_MODSCOPE_LOCAL
    HANDLE_IMPLICIT = PSF_MODSCOPE_SPANNING
    HANDLE_UP_TO_EXPLICIT = RegexNumericConstants.LONG_MAX
End Enum

Public Sub Compile(ByRef outBytecode() As Long, ByRef s As String, Optional ByVal caseInsensitive As Boolean = False)
    Dim lex As RegexLexer.Ty
    Dim ast As ArrayBuffer.Ty
    
    If Not RegexUnicodeSupport.UnicodeInitialized Then RegexUnicodeSupport.UnicodeInitialize
    If Not RegexUnicodeSupport.RangeTablesInitialized Then RegexUnicodeSupport.RangeTablesInitialize
    
    RegexLexer.Initialize lex, s
    Parse lex, caseInsensitive, ast
    RegexAst.AstToBytecode ast.Buffer, lex.identifierTree, caseInsensitive, outBytecode
End Sub

Private Sub PerformPotentialConcat( _
    ByRef ast As ArrayBuffer.Ty, ByRef potentialConcat2 As Long, ByRef potentialConcat1 As Long _
)
    Dim tmp As Long
    If potentialConcat2 <> -1 Then
        tmp = ast.Length
        ArrayBuffer.AppendThree ast, RegexAst.AST_CONCAT, potentialConcat2, potentialConcat1
        potentialConcat2 = -1
        potentialConcat1 = tmp
    End If
End Sub

' implicit groups are removed from the stack, explicit groups are left on the stack
' returns: stack frame type of last stack frame that was considered
' spilloverWriteMask is an output parameter, contains all WRITE bits that were present in the removed frames
Private Function CloseGroups( _
    ByRef ast As ArrayBuffer.Ty, _
    ByRef parseStack As ArrayBuffer.Ty, _
    ByRef pendingDisjunction As Long, _
    ByRef currentDisjunction As Long, _
    ByRef potentialConcat2 As Long, _
    ByRef potentialConcat1 As Long, _
    ByRef modifierMask As Long, _
    ByRef spilloverWriteMask As Long, _
    ByVal whatToHandle As Long _
) As Long
    Dim currentAstNode As Long ' local helper variable
    Dim stackFrameType As Long
    Dim modifierEntry As Long, modifierEntryWriteMask As Long, accumulatedWriteMask As Long
    
    CloseGroups = PSF_NONE
    spilloverWriteMask = 0
    
    Do
        If potentialConcat1 = -1 Then
            currentAstNode = ast.Length
            ArrayBuffer.AppendLong ast, AST_EMPTY
            potentialConcat1 = currentAstNode
        Else
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        End If
        
        With parseStack
            If .Length = 0 Then Exit Function

            stackFrameType = .Buffer(.Length - 1)
            If stackFrameType > whatToHandle Then Exit Function
            CloseGroups = stackFrameType
            
            ' Close pending disjunction, if there is one
            If pendingDisjunction <> -1 Then
                ast.Buffer(pendingDisjunction + 2) = potentialConcat1
                potentialConcat1 = currentDisjunction
            End If

            ' Pop stack frame and restore variables
            With parseStack
                pendingDisjunction = .Buffer(.Length - 4)
                currentDisjunction = .Buffer(.Length - 3)
                potentialConcat2 = .Buffer(.Length - 2)
            End With
            
            If stackFrameType <= PSF_MAX_WITH_MODIFIERS Then
                modifierEntry = .Buffer(.Length - 5)
                If modifierEntry <> 0 Then ' the group has modifiers
                    modifierEntryWriteMask = modifierEntry And MODIFIER_WRITE_MASK
                    currentAstNode = ast.Length
                    ArrayBuffer.AppendThree ast, AST_MODIFIER_SCOPE, potentialConcat1, _
                        modifierEntryWriteMask _
                            Or _
                        modifierEntryWriteMask * 2 And modifierMask
                    potentialConcat1 = currentAstNode
                    modifierMask = modifierMask Xor (modifierEntry And MODIFIER_ACTIVE_MASK)
                    spilloverWriteMask = spilloverWriteMask Or modifierEntryWriteMask
                End If
            End If
            
            ' Stop at explicit group; explicit groups are never popped from the stack.
            ' There is no spillover from explicit groups, hence we can exit the function.
            If stackFrameType >= PSF_MIN_EXPLICIT Then Exit Function
            
            .Length = .Length - 5
        End With
    Loop
End Function

Private Sub Parse(ByRef lex As RegexLexer.Ty, ByVal caseInsensitive As Boolean, ByRef ast As ArrayBuffer.Ty)
    ' carry information througout the function
    Dim currToken As RegexLexer.ReToken
    Dim parseStack As ArrayBuffer.Ty
    Dim modifierMask As Long
    Dim potentialConcat2 As Long, potentialConcat1 As Long, pendingDisjunction As Long, currentDisjunction As Long
    Dim nCaptures As Long
    
    ' only locally used
    Dim currentAstNode As Long, spilloverWriteMask As Long
    Dim tmp As Long, i As Long, qmin As Long, qmax As Long, n1 As Long, n2 As Long
    
    nCaptures = 0
    
    pendingDisjunction = -1
    potentialConcat2 = -1
    potentialConcat1 = -1
    currentDisjunction = -1
    modifierMask = RegexBytecode.MODIFIER_I_ACTIVE And caseInsensitive
    
    ArrayBuffer.AppendLong ast, 0 ' first word will be index of the root node, to be patched in the end

    Do
ContinueLoop:
        RegexLexer.ParseReToken lex, currToken
        
        Select Case currToken.t
        Case RETOK_DISJUNCTION
            n2 = modifierMask ' save modifier mask for later
            
            CloseGroups ast, parseStack, pendingDisjunction, currentDisjunction, potentialConcat2, potentialConcat1, _
                modifierMask, spilloverWriteMask, HANDLE_IMPLICIT_LOCAL
            
            currentAstNode = ast.Length
            ArrayBuffer.AppendThree ast, AST_DISJ, potentialConcat1, -1
            potentialConcat1 = -1

            If pendingDisjunction <> -1 Then
                ast.Buffer(pendingDisjunction + 2) = currentAstNode
            Else
                currentDisjunction = currentAstNode
            End If
        
            pendingDisjunction = currentAstNode
        
            If spilloverWriteMask Then
                ' Create a modifier scope
                ' See comment in Case RETOK_ATOM_START_NONCAPTURE_GROUP, RETOK_UNBOUNDED_MODIFIER for a description.
                ' Remember: n2 is the modifier mask at the end of the last alternative.
                ' Hence the following n1 is that ACTIVE part of the xor difference between n2 and the new modifierMask
                '   that spills over.
                n1 = spilloverWriteMask * 2 And (n2 Xor modifierMask)
                ' We push this xor difference as well as the WRITE part of the spillover.
                ArrayBuffer.AppendFive parseStack, _
                    spilloverWriteMask Or n1, _
                    pendingDisjunction, currentDisjunction, potentialConcat1, _
                    PSF_MODSCOPE_SPANNING
                
                ' We apply the spillover to the modifierMask.
                modifierMask = modifierMask Xor n1
                pendingDisjunction = -1
                potentialConcat2 = -1
                potentialConcat1 = -1
                currentDisjunction = -1
            End If
            
        Case RETOK_QUANTIFIER
            If potentialConcat1 = -1 Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER_NO_ATOM
            
            qmin = currToken.qmin
            qmax = currToken.qmax
            
            If qmin > qmax Then
                currentAstNode = ast.Length
                ArrayBuffer.AppendLong ast, AST_FAIL
                potentialConcat1 = currentAstNode
                GoTo ContinueLoop
            End If
            
            If qmin = 0 Then
                n1 = -1
            ElseIf currToken.quantifierType = TokenQuantifierType.QUANTIFIER_POSSESSIVE Then
                currentAstNode = ast.Length
                ArrayBuffer.AppendThree ast, AST_REPEAT_EXACTLY_POSSESSIVE, potentialConcat1, qmin
                n1 = currentAstNode
            ElseIf qmin = 1 Then
                n1 = potentialConcat1
            ElseIf qmin > 1 Then
                currentAstNode = ast.Length
                ArrayBuffer.AppendThree ast, AST_REPEAT_EXACTLY, potentialConcat1, qmin
                n1 = currentAstNode
            Else
                Err.Raise RegexErrors.REGEX_ERR_INTERNAL_LOGIC_ERR
            End If
            
            If qmax = RE_QUANTIFIER_INFINITE Then
                Select Case currToken.quantifierType
                Case TokenQuantifierType.QUANTIFIER_GREEDY
                    tmp = AST_STAR_GREEDY
                Case TokenQuantifierType.QUANTIFIER_HUMBLE
                    tmp = AST_STAR_HUMBLE
                Case TokenQuantifierType.QUANTIFIER_POSSESSIVE
                    tmp = AST_STAR_POSSESSIVE
                End Select
                currentAstNode = ast.Length
                ArrayBuffer.AppendTwo ast, tmp, potentialConcat1
                n2 = currentAstNode
            ElseIf qmax - qmin = 1 Then
                Select Case currToken.quantifierType
                Case TokenQuantifierType.QUANTIFIER_GREEDY
                    tmp = AST_ZEROONE_GREEDY
                Case TokenQuantifierType.QUANTIFIER_HUMBLE
                    tmp = AST_ZEROONE_HUMBLE
                Case TokenQuantifierType.QUANTIFIER_POSSESSIVE
                    tmp = AST_ZEROONE_POSSESSIVE
                End Select
                currentAstNode = ast.Length
                ArrayBuffer.AppendTwo ast, tmp, potentialConcat1
                n2 = currentAstNode
            ElseIf qmax - qmin > 1 Then
                Select Case currToken.quantifierType
                Case TokenQuantifierType.QUANTIFIER_GREEDY
                    tmp = AST_REPEAT_MAX_GREEDY
                Case TokenQuantifierType.QUANTIFIER_HUMBLE
                    tmp = AST_REPEAT_MAX_HUMBLE
                Case TokenQuantifierType.QUANTIFIER_POSSESSIVE
                    tmp = AST_REPEAT_MAX_POSSESSIVE
                End Select
                currentAstNode = ast.Length
                ArrayBuffer.AppendThree ast, tmp, potentialConcat1, qmax - qmin
                n2 = currentAstNode
            ElseIf qmax = qmin Then
                n2 = -1
            Else
                Err.Raise RegexErrors.REGEX_ERR_INTERNAL_LOGIC_ERR
            End If
            
            If n1 = -1 Then
                If n2 = -1 Then
                    currentAstNode = ast.Length
                    ArrayBuffer.AppendLong ast, AST_EMPTY
                    potentialConcat1 = currentAstNode
                Else
                    potentialConcat1 = n2
                End If
            Else
                If n2 = -1 Then
                    potentialConcat1 = n1
                Else
                    currentAstNode = ast.Length
                    ArrayBuffer.AppendThree ast, AST_CONCAT, n1, n2
                    potentialConcat1 = currentAstNode
                End If
            End If
            
        Case RETOK_ATOM_START_CAPTURE_GROUP
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            nCaptures = nCaptures + 1
            ArrayBuffer.AppendFive parseStack, _
                currToken.num, pendingDisjunction, currentDisjunction, potentialConcat1, _
                nCaptures
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
            
        Case RETOK_ATOM_START_ATOMIC_GROUP
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            ArrayBuffer.AppendFive parseStack, _
                0, pendingDisjunction, currentDisjunction, potentialConcat1, PSF_ATOMIC
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
                
        Case RETOK_ATOM_START_NONCAPTURE_GROUP, RETOK_UNBOUNDED_MODIFIER
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            ' n1 is being used to temporarily store the new modifierMask.
            ' Effectively, we set
            '   n1[x_ACTIVE] := currentToken.num[x_WRITE] ? currentToken.num[x_ACTIVE] : modifierMask[x_ACTIVE].
            ' We make use of the fact that for b1, b2, b3 in {0, 1},
            '   b1 ? b2 : b3   is equivalent to   b3 xor [b1 and (b2 xor b3)].
            '
            ' Note that modifierMask[x_WRITE] = 0, i.e. (modifierMask and MODIFIER_WRITE_MASK) = 0, and that
            ' the same holds for n1 after the assignment.
            n1 = modifierMask Xor _
                (((currToken.num And RegexBytecode.MODIFIER_WRITE_MASK) * 2) And (currToken.num Xor modifierMask))
        
            ' Since modifierMask uses only ACTIVE bit positions (and no WRITE bit positions),
            '   i.e. since (modifierMask and MODIFIER_WRITE_MASK) = 0,
            '   the WRITE bit positions are free for storing the WRITE bits of currToken.num on the stack.
            ' In the ACTIVE bit positions, we store the xor difference between the the old and the new
            '   modifier mask (i.e. the difference between modifierMask and n1).
            ' When popping from the stack, we will be able to restore the ACTIVE bits of currToken.num from
            '   the new modifierMask together with the WRITE bits popped from the stack.
            ArrayBuffer.AppendFive parseStack, _
                (currToken.num And RegexBytecode.MODIFIER_WRITE_MASK) Or (modifierMask Xor n1), _
                pendingDisjunction, currentDisjunction, potentialConcat1, _
                IIf(currToken.t = RETOK_ATOM_START_NONCAPTURE_GROUP, PSF_NONCAPTURE, PSF_MODSCOPE_LOCAL)
            
            modifierMask = n1
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
            
         Case RETOK_ASSERT_START_POS_LOOKAHEAD
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            ArrayBuffer.AppendFive parseStack, _
                -1, pendingDisjunction, currentDisjunction, potentialConcat1, _
                -AST_ASSERT_POS_LOOKAHEAD
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
        
        Case RETOK_ASSERT_START_NEG_LOOKAHEAD
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            ArrayBuffer.AppendFive parseStack, _
                -1, pendingDisjunction, currentDisjunction, potentialConcat1, _
                -AST_ASSERT_NEG_LOOKAHEAD
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
        
         Case RETOK_ASSERT_START_POS_LOOKBEHIND
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            ArrayBuffer.AppendFive parseStack, _
                -1, pendingDisjunction, currentDisjunction, potentialConcat1, _
                -AST_ASSERT_POS_LOOKBEHIND
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
        
        Case RETOK_ASSERT_START_NEG_LOOKBEHIND
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            ArrayBuffer.AppendFive parseStack, _
                -1, pendingDisjunction, currentDisjunction, potentialConcat1, _
                -AST_ASSERT_NEG_LOOKBEHIND
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
        
        Case RETOK_ATOM_END
            tmp = CloseGroups(ast, parseStack, pendingDisjunction, currentDisjunction, _
                potentialConcat2, potentialConcat1, modifierMask, 0, HANDLE_UP_TO_EXPLICIT)

            If tmp < PSF_MIN_EXPLICIT Then
                Err.Raise REGEX_ERR_UNEXPECTED_CLOSING_PAREN
            ElseIf tmp > 0 Then ' capture group
                n1 = parseStack.Buffer(parseStack.Length - 5)
                If n1 <> -1 Then
                    currentAstNode = ast.Length
                    ArrayBuffer.AppendFour ast, AST_NAMED, potentialConcat1, n1, tmp
                    potentialConcat1 = currentAstNode
                End If
                currentAstNode = ast.Length
                ArrayBuffer.AppendThree ast, AST_CAPTURE, potentialConcat1, tmp
                potentialConcat1 = currentAstNode
            ElseIf tmp = PSF_NONCAPTURE Then ' non-capture group
                ' nothing to do
            ElseIf tmp = PSF_ATOMIC Then ' atomic group
                currentAstNode = ast.Length
                ArrayBuffer.AppendThree ast, AST_REPEAT_EXACTLY_POSSESSIVE, potentialConcat1, 1
                potentialConcat1 = currentAstNode
            Else ' lookahead or lookbehind
                currentAstNode = ast.Length
                ArrayBuffer.AppendTwo ast, -tmp, potentialConcat1
                potentialConcat1 = currentAstNode
            End If
            parseStack.Length = parseStack.Length - 5 ' pop stack frame

        Case RETOK_ATOM_CHAR
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            
            currentAstNode = ast.Length
            tmp = currToken.num
            If modifierMask And RegexBytecode.MODIFIER_I_ACTIVE Then
                tmp = RegexUnicodeSupport.ReCanonicalizeChar(tmp)
            End If
            ArrayBuffer.AppendTwo ast, RegexAst.AST_CHAR, tmp
                
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_PERIOD
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            ArrayBuffer.AppendLong ast, RegexAst.AST_PERIOD
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_BACKREFERENCE
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            ArrayBuffer.AppendTwo ast, RegexAst.AST_BACKREFERENCE, currToken.num
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
        
        Case RETOK_ASSERT_START
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            ArrayBuffer.AppendLong ast, RegexAst.AST_ASSERT_START
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ASSERT_END
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            ArrayBuffer.AppendLong ast, RegexAst.AST_ASSERT_END
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_START_CHARCLASS, RETOK_ATOM_START_CHARCLASS_INVERTED
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            If currToken.t = RETOK_ATOM_START_CHARCLASS Then
                ArrayBuffer.AppendTwo ast, AST_RANGES, 0
            Else
                ArrayBuffer.AppendTwo ast, AST_INVRANGES, 0
            End If
            
            tmp = 0 ' unnecessary, tmp is an output parameter indicating the number of ranges
            ' Todo: Remove that parameter from ParseReRanges -- we can calculate it by comparing
            '   old and new buffer length.
            RegexLexer.ParseReRanges lex, ast, tmp, (modifierMask And RegexBytecode.MODIFIER_I_ACTIVE) <> 0
            
            ' patch range count
            ast.Buffer(currentAstNode + 1) = tmp
            
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ASSERT_WORD_BOUNDARY
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            ArrayBuffer.AppendLong ast, RegexAst.AST_ASSERT_WORD_BOUNDARY
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ASSERT_NOT_WORD_BOUNDARY
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            ArrayBuffer.AppendLong ast, RegexAst.AST_ASSERT_NOT_WORD_BOUNDARY
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_DIGIT
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            ArrayBuffer.AppendPrefixedPairsArray ast, AST_RANGES, RegexUnicodeSupport.StaticData, _
                RegexUnicodeSupport.RANGE_TABLE_DIGIT_START, RegexUnicodeSupport.RANGE_TABLE_DIGIT_LENGTH
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_NOT_DIGIT
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            ArrayBuffer.AppendPrefixedPairsArray ast, AST_RANGES, RegexUnicodeSupport.StaticData, _
                RegexUnicodeSupport.RANGE_TABLE_NOTDIGIT_START, RegexUnicodeSupport.RANGE_TABLE_NOTDIGIT_LENGTH
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_WHITE
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            ArrayBuffer.AppendPrefixedPairsArray ast, AST_RANGES, RegexUnicodeSupport.StaticData, _
                RegexUnicodeSupport.RANGE_TABLE_WHITE_START, RegexUnicodeSupport.RANGE_TABLE_WHITE_LENGTH
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_ATOM_NOT_WHITE
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            ArrayBuffer.AppendPrefixedPairsArray ast, AST_RANGES, RegexUnicodeSupport.StaticData, _
                RegexUnicodeSupport.RANGE_TABLE_NOTWHITE_START, RegexUnicodeSupport.RANGE_TABLE_NOTWHITE_LENGTH
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
        
        Case RETOK_ATOM_WORD_CHAR
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            ArrayBuffer.AppendPrefixedPairsArray ast, AST_RANGES, RegexUnicodeSupport.StaticData, _
                RegexUnicodeSupport.RANGE_TABLE_WORDCHAR_START, RegexUnicodeSupport.RANGE_TABLE_WORDCHAR_LENGTH
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
        
        Case RETOK_ATOM_NOT_WORD_CHAR
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            currentAstNode = ast.Length
            ArrayBuffer.AppendPrefixedPairsArray ast, AST_RANGES, RegexUnicodeSupport.StaticData, _
                RegexUnicodeSupport.RANGE_TABLE_NOTWORDCHAR_START, RegexUnicodeSupport.RANGE_TABLE_NOTWORDCHAR_LENGTH
            potentialConcat2 = potentialConcat1: potentialConcat1 = currentAstNode
            
        Case RETOK_EOF
            CloseGroups ast, parseStack, pendingDisjunction, currentDisjunction, potentialConcat2, potentialConcat1, _
                modifierMask, 0, HANDLE_IMPLICIT
                        
            If Not parseStack.Length = 0 Then Err.Raise REGEX_ERR_UNEXPECTED_END_OF_PATTERN
            
            ' Close disjunction
            If pendingDisjunction = -1 Then
                currentDisjunction = potentialConcat1
            Else
                ast.Buffer(pendingDisjunction + 2) = potentialConcat1
            End If
            
            currentAstNode = ast.Length
            ArrayBuffer.AppendThree ast, RegexAst.AST_CAPTURE, currentDisjunction, 0
            ast.Buffer(0) = currentAstNode ' patch index of root node into the first word
            
            Exit Do
        Case Else
            Err.Raise REGEX_ERR_UNEXPECTED_REGEXP_TOKEN
        End Select
    Loop
End Sub

