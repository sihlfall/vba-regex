Attribute VB_Name = "RegexCompiler"
Option Explicit

Public Sub Compile(ByRef outBytecode() As Long, ByRef s As String, Optional ByVal caseInsensitive As Boolean = False)
    Dim lex As RegexLexer.Ty
    Dim ast As ArrayBuffer.Ty
    
    If Not RegexUnicodeSupport.UnicodeInitialized Then RegexUnicodeSupport.UnicodeInitialize
    If Not RegexUnicodeSupport.RangeTablesInitialized Then RegexUnicodeSupport.RangeTablesInitialize
    
    RegexLexer.Initialize lex, s
    Parse lex, caseInsensitive, ast
    RegexAst.AstToBytecode ast.Buffer, lex.identifierTree, caseInsensitive, outBytecode
End Sub

Private Sub PerformPotentialConcat(ByRef ast As ArrayBuffer.Ty, ByRef potentialConcat2 As Long, ByRef potentialConcat1 As Long)
    Dim tmp As Long
    If potentialConcat2 <> -1 Then
        tmp = ast.Length
        ArrayBuffer.AppendThree ast, RegexAst.AST_CONCAT, potentialConcat2, potentialConcat1
        potentialConcat2 = -1
        potentialConcat1 = tmp
    End If
End Sub

Private Sub Parse(ByRef lex As RegexLexer.Ty, ByVal caseInsensitive As Boolean, ByRef ast As ArrayBuffer.Ty)
    Dim currToken As RegexLexer.ReToken
    Dim currentAstNode As Long
    Dim potentialConcat2 As Long, potentialConcat1 As Long, pendingDisjunction As Long, currentDisjunction As Long
    Dim nCaptures As Long
    Dim tmp As Long, i As Long, qmin As Long, qmax As Long, n1 As Long, n2 As Long
    Dim parseStack As ArrayBuffer.Ty
    Dim modifierMask As Long
    
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
            If potentialConcat1 = -1 Then
                currentAstNode = ast.Length
                ArrayBuffer.AppendLong ast, AST_EMPTY
                potentialConcat1 = currentAstNode
            Else
                PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            End If
            
            currentAstNode = ast.Length
            ArrayBuffer.AppendThree ast, AST_DISJ, potentialConcat1, -1
            potentialConcat1 = -1

            If pendingDisjunction <> -1 Then
                ast.Buffer(pendingDisjunction + 2) = currentAstNode
            Else
                currentDisjunction = currentAstNode
            End If
        
            pendingDisjunction = currentAstNode
        
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
                If currToken.greedy Then tmp = AST_STAR_GREEDY Else tmp = AST_STAR_HUMBLE
                currentAstNode = ast.Length
                ArrayBuffer.AppendTwo ast, tmp, potentialConcat1
                n2 = currentAstNode
            ElseIf qmax - qmin = 1 Then
                If currToken.greedy Then tmp = AST_ZEROONE_GREEDY Else tmp = AST_ZEROONE_HUMBLE
                currentAstNode = ast.Length
                ArrayBuffer.AppendTwo ast, tmp, potentialConcat1
                n2 = currentAstNode
            ElseIf qmax - qmin > 1 Then
                If currToken.greedy Then tmp = AST_REPEAT_MAX_GREEDY Else tmp = AST_REPEAT_MAX_HUMBLE
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
                nCaptures, currToken.num, pendingDisjunction, currentDisjunction, potentialConcat1
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
                
        Case RETOK_ATOM_START_NONCAPTURE_GROUP
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
                -1, _
                (currToken.num And RegexBytecode.MODIFIER_WRITE_MASK) Or (modifierMask Xor n1), _
                pendingDisjunction, currentDisjunction, potentialConcat1
            
            modifierMask = n1
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
            
         Case RETOK_ASSERT_START_POS_LOOKAHEAD
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            ArrayBuffer.AppendFive parseStack, _
                -(AST_ASSERT_POS_LOOKAHEAD - MIN_AST_CODE) - 2, -1, pendingDisjunction, currentDisjunction, potentialConcat1
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
        
        Case RETOK_ASSERT_START_NEG_LOOKAHEAD
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            nCaptures = nCaptures + 1
            ArrayBuffer.AppendFive parseStack, _
                -(AST_ASSERT_NEG_LOOKAHEAD - MIN_AST_CODE) - 2, -1, pendingDisjunction, currentDisjunction, potentialConcat1
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
        
         Case RETOK_ASSERT_START_POS_LOOKBEHIND
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            ArrayBuffer.AppendFive parseStack, _
                -(AST_ASSERT_POS_LOOKBEHIND - MIN_AST_CODE) - 2, -1, pendingDisjunction, currentDisjunction, potentialConcat1
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
        
        Case RETOK_ASSERT_START_NEG_LOOKBEHIND
            PerformPotentialConcat ast, potentialConcat2, potentialConcat1
        
            nCaptures = nCaptures + 1
            ArrayBuffer.AppendFive parseStack, _
                -(AST_ASSERT_NEG_LOOKBEHIND - MIN_AST_CODE) - 2, -1, pendingDisjunction, currentDisjunction, potentialConcat1
                
            pendingDisjunction = -1
            potentialConcat2 = -1
            potentialConcat1 = -1
            currentDisjunction = -1
        
        
        Case RETOK_ATOM_END
            If parseStack.Length = 0 Then Err.Raise REGEX_ERR_UNEXPECTED_CLOSING_PAREN
            
            ' Close disjunction
            If potentialConcat1 = -1 Then
                currentAstNode = ast.Length
                ArrayBuffer.AppendLong ast, AST_EMPTY
                potentialConcat1 = currentAstNode
            Else
                PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            End If
            
            If pendingDisjunction = -1 Then
                currentDisjunction = potentialConcat1
            Else
                ast.Buffer(pendingDisjunction + 2) = potentialConcat1
            End If

            potentialConcat1 = currentDisjunction
            
            ' Restore variables
            With parseStack
                .Length = .Length - 5
                tmp = .Buffer(.Length)
                n1 = .Buffer(.Length + 1)
                pendingDisjunction = .Buffer(.Length + 2)
                currentDisjunction = .Buffer(.Length + 3)
                potentialConcat2 = .Buffer(.Length + 4) ' This is correct, potentialConcat1 is the new node!
            End With
            
            If tmp > 0 Then ' capture group
                If n1 <> -1 Then
                    currentAstNode = ast.Length
                    ArrayBuffer.AppendFour ast, AST_NAMED, potentialConcat1, n1, tmp
                    potentialConcat1 = currentAstNode
                End If
                currentAstNode = ast.Length
                ArrayBuffer.AppendThree ast, AST_CAPTURE, potentialConcat1, tmp
                potentialConcat1 = currentAstNode
            ElseIf tmp = -1 Then ' non-capture group
                If n1 <> 0 Then ' the group has modifiers
                    currentAstNode = ast.Length
                    ArrayBuffer.AppendThree ast, AST_MODIFIER_SCOPE, potentialConcat1, _
                        (n1 And MODIFIER_WRITE_MASK) Or (((n1 And MODIFIER_WRITE_MASK) * 2) And modifierMask)
                    potentialConcat1 = currentAstNode
                    modifierMask = modifierMask Xor (n1 And MODIFIER_ACTIVE_MASK)
                End If
            Else ' lookahead or lookbehind
                currentAstNode = ast.Length
                ArrayBuffer.AppendTwo ast, -(tmp + 2) + MIN_AST_CODE, potentialConcat1
                potentialConcat1 = currentAstNode
            End If

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
            ' Todo: If Not expectEof Then Err.Raise REGEX_ERR_UNEXPECTED_END_OF_PATTERN
            
            ' Close disjunction
            If potentialConcat1 = -1 Then
                currentAstNode = ast.Length
                ArrayBuffer.AppendLong ast, AST_EMPTY
                potentialConcat1 = currentAstNode
            Else
                PerformPotentialConcat ast, potentialConcat2, potentialConcat1
            End If
            
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

