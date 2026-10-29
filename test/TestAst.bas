Attribute VB_Name = "TestAst"
Option Explicit

Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

Private emptyIdentifierTree As RegexIdentifierSupport.IdentifierTreeTy

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

Private Sub MakeArray(ByRef outAry() As Long, ParamArray p() As Variant)
    ReDim outAry(0 To UBound(p)) As Long
    Dim i As Long
    For i = 0 To UBound(p)
        outAry(i) = p(i)
    Next
End Sub

Private Sub MakeArrayBuffer(ByRef ast As ArrayBuffer.Ty, ParamArray p() As Variant)
    Dim i As Long
    For i = 0 To UBound(p)
        ArrayBuffer.AppendLong ast, p(i)
    Next
End Sub

Private Sub CompileRepeatPair(ByRef ast As ArrayBuffer.Ty, ByRef compact() As Long, ByRef expanded() As Long)
    Dim copy As ArrayBuffer.Ty
    copy.Buffer = ast.Buffer: copy.Length = ast.Length: copy.Capacity = ast.Capacity
    RegexAst.AstToBytecode copy, emptyIdentifierTree, False, compact, 0
    RegexAst.AstToBytecode ast, emptyIdentifierTree, False, expanded
    Assert.IsTrue UBound(expanded) <= UBound(compact) + DEFAULT_QUANTIFIER_EXPANSION_LIMIT
End Sub

'@TestMethod("Ast")
Private Sub Ast_RepeatExpansionAgreesWithCompactExecution()
    Dim ast As ArrayBuffer.Ty, compact() As Long, expanded() As Long, kinds() As Long
    Dim kind As Variant, count As Long, body As Long, root As Long, ambiguous As Long
    Dim sample As Long, value As Long, text As String, offset As Long, result As Long, actual As Long
    Dim context As RegexDfsMatcher.DfsMatcherContext, otherContext As RegexDfsMatcher.DfsMatcherContext
    Dim expected As RegexMatcherCommon.CapturesTy, captures As RegexMatcherCommon.CapturesTy
    Dim dfa As RegexDfaMatcher.StateStructure, slp As RegexMatcherCommon.StartLengthPair
    MakeArray kinds, AST_REPEAT_EXACTLY, AST_REPEAT_MAX_GREEDY, AST_REPEAT_MAX_HUMBLE
    For Each kind In kinds
        For count = 1 To 4
            For ambiguous = 0 To 4
                ast.Length = 0
                MakeArrayBuffer ast, 0, AST_CHAR, 97
                body = 1
                If ambiguous = 1 Then
                    MakeArrayBuffer ast, AST_STRING, 2, 97, 98
                    body = ast.Length
                    ArrayBuffer.AppendThree ast, AST_DISJ, 1, 3
                ElseIf ambiguous = 2 Then
                    MakeArrayBuffer ast, AST_REPEAT_MAX_GREEDY, 1, 2, AST_CHAR, 98
                    body = ast.Length
                    ArrayBuffer.AppendThree ast, AST_CONCAT, 3, 6
                ElseIf ambiguous = 3 Then
                    MakeArrayBuffer ast, AST_REPEAT_EXACTLY, 1, 2, AST_ZEROONE_GREEDY, 1
                    body = ast.Length
                    ArrayBuffer.AppendThree ast, AST_CONCAT, 3, 6
                ElseIf ambiguous = 4 Then
                    MakeArrayBuffer ast, AST_CHAR, 98, AST_CAPTURE, 3, 2, AST_ZEROONE_GREEDY, 5
                    body = ast.Length
                    ArrayBuffer.AppendThree ast, AST_CONCAT, 1, 8
                End If
                root = ast.Length
                ArrayBuffer.AppendThree ast, AST_CAPTURE, body, 1
                body = ast.Length
                ArrayBuffer.AppendThree ast, CLng(kind), root, count
                root = ast.Length
                ArrayBuffer.AppendTwo ast, AST_CHAR, 98
                ArrayBuffer.AppendThree ast, AST_CONCAT, body, root
                ast.Buffer(0) = ast.Length
                ArrayBuffer.AppendThree ast, AST_CAPTURE, ast.Length - 3, 0
                CompileRepeatPair ast, compact, expanded
                Assert.AreEqual 0&, expanded(BYTECODE_IDX_PROPERTIES) And BYTECODE_PROPERTY_REQUIRES_DFS_MATCHING
                RegexDfaMatcher.DfaStateStructureInitialize dfa
                For sample = 0 To 62 ' all strings over a/b through length five
                    text = vbNullString: value = sample
                    Do While value > 0
                        value = value - 1
                        text = text & ChrW(97 + value Mod 2): value = value \ 2
                    Loop
                    For offset = 0 To Len(text)
                        result = RegexDfsMatcher.DfsMatchFrom(context, expected, compact, text, offset)
                        actual = RegexDfsMatcher.DfsMatchFrom(otherContext, captures, expanded, text, offset)
                        Assert.AreEqual result, actual
                        If result <> -1 Then CompareRepeatCaptures expected, captures
                        actual = RegexNfaMatcher.NfaMatchFrom(captures, expanded, text, offset)
                        Assert.AreEqual result, actual
                        If result <> -1 Then CompareRepeatCaptures expected, captures
                        Assert.AreEqual (result <> -1), RegexDfaMatcher.DfaFindMatch(slp, dfa, expanded, text, offset + 1, False)
                        If result <> -1 Then
                            Assert.AreEqual expected.entireMatch.start, slp.start
                            Assert.AreEqual expected.entireMatch.Length, slp.Length
                        End If
                    Next
                Next
            Next
        Next
    Next
End Sub

Private Sub CompareRepeatCaptures(ByRef expected As RegexMatcherCommon.CapturesTy, ByRef actual As RegexMatcherCommon.CapturesTy)
    Dim i As Long
    Assert.AreEqual expected.entireMatch.start, actual.entireMatch.start
    Assert.AreEqual expected.entireMatch.Length, actual.entireMatch.Length
    Assert.AreEqual expected.nNumberedCaptures, actual.nNumberedCaptures
    For i = 0 To expected.nNumberedCaptures - 1
        Assert.AreEqual expected.numberedCaptures(i).start, actual.numberedCaptures(i).start
        Assert.AreEqual expected.numberedCaptures(i).Length, actual.numberedCaptures(i).Length
    Next
End Sub

'@TestMethod("Ast")
Private Sub Ast_RepeatExpansionIsBounded()
    Dim ast As ArrayBuffer.Ty, compact() As Long, expanded() As Long
    Dim n As Long, body As Long, root As Long
    ' Large bounds are checked before multiplication.
    MakeArrayBuffer ast, 3, AST_CHAR, 97, AST_REPEAT_EXACTLY, 1, 1000000000
    CompileRepeatPair ast, compact, expanded
    Assert.SequenceEquals compact, expanded
    ' Account for the reverse program exactly at the whole-pattern limit.
    For n = 2050 To 2051
        ast.Length = 0
        MakeArrayBuffer ast, 3, AST_CHAR, 97, AST_REPEAT_EXACTLY, 1, n
        CompileRepeatPair ast, compact, expanded
        If n = 2050 Then
            Assert.AreEqual UBound(compact) + DEFAULT_QUANTIFIER_EXPANSION_LIMIT, UBound(expanded)
            Assert.AreEqual 0&, expanded(BYTECODE_IDX_PROPERTIES) And BYTECODE_PROPERTY_REQUIRES_DFS_MATCHING
        Else
            Assert.SequenceEquals compact, expanded
        End If
    Next
    ' Nested expansions use the expanded child size, not the original AST size.
    ast.Length = 0
    MakeArrayBuffer ast, 0, AST_CHAR, 97
    body = 1
    For n = 1 To 8
        root = ast.Length
        ArrayBuffer.AppendThree ast, AST_REPEAT_EXACTLY, body, 16
        body = root
    Next
    ast.Buffer(0) = body
    CompileRepeatPair ast, compact, expanded
    Assert.IsTrue UBound(expanded) < 1024
    Assert.IsTrue (expanded(BYTECODE_IDX_PROPERTIES) And BYTECODE_PROPERTY_REQUIRES_DFS_MATCHING) <> 0
    ' Shared subtrees and many individually small repeats obey one global budget.
    ast.Length = 0
    MakeArrayBuffer ast, 0, AST_CHAR, 97, AST_REPEAT_EXACTLY, 1, 100
    body = 3
    For n = 1 To 30
        root = ast.Length
        ArrayBuffer.AppendThree ast, AST_CONCAT, body, 3
        body = root
    Next
    ast.Buffer(0) = body
    CompileRepeatPair ast, compact, expanded
    Assert.SequenceEquals compact, expanded
    ' Enabling a reverse program must fit the same total growth budget.
    ast.Length = 0
    MakeArrayBuffer ast, 0, AST_CHAR, 97, AST_REPEAT_EXACTLY, 1, 2
    body = ast.Length
    ArrayBuffer.AppendTwo ast, AST_STRING, 5000
    For n = 1 To 5000: ArrayBuffer.AppendLong ast, 98: Next
    ast.Buffer(0) = ast.Length
    ArrayBuffer.AppendThree ast, AST_CONCAT, body, 3
    CompileRepeatPair ast, compact, expanded
    Assert.SequenceEquals compact, expanded
End Sub

'@TestMethod("Ast")
Private Sub Ast_RepeatExpansionLeavesNullableAndPossessiveBodies()
    Dim ast As ArrayBuffer.Ty, compact() As Long, expanded() As Long
    MakeArrayBuffer ast, 5, AST_CHAR, 97, AST_ZEROONE_GREEDY, 1, AST_REPEAT_EXACTLY, 3, 4
    CompileRepeatPair ast, compact, expanded
    Assert.SequenceEquals compact, expanded
    ast.Length = 0
    MakeArrayBuffer ast, 3, AST_CHAR, 97, AST_REPEAT_EXACTLY_POSSESSIVE, 1, 4
    CompileRepeatPair ast, compact, expanded
    Assert.SequenceEquals compact, expanded
    ' A consuming suffix does not hide a nullable loop in the repeated body.
    ast.Length = 0
    MakeArrayBuffer ast, 12, AST_CHAR, 97, AST_ZEROONE_GREEDY, 1, AST_STAR_GREEDY, 3, _
        AST_CHAR, 98, AST_CONCAT, 5, 7, AST_REPEAT_EXACTLY, 9, 2
    CompileRepeatPair ast, compact, expanded
    Assert.SequenceEquals compact, expanded
End Sub

'@TestMethod("Ast")
Private Sub Ast_NullableStarsSeparateEntryFromIteration()
    Dim ast As ArrayBuffer.Ty, actual() As Long, expected() As Long
    Dim captures As RegexMatcherCommon.CapturesTy, span As RegexMatcherCommon.StartLengthPair
    Dim state As RegexDfaMatcher.StateStructure
    Dim humble As Long, expansion As Long, splitOp As Long, loopOp As Long
    For humble = 0 To 1
        splitOp = REOP_SPLIT1 + humble
        loopOp = REOP_LOOP_GREEDY + humble
        For expansion = 0 To 1
            ast.Length = 0
            ' Capture 0 around (?:|a)* or (?:|a)*?. Build the AST first.
            MakeArrayBuffer ast, 9, AST_EMPTY, AST_CHAR, 97, AST_DISJ, 1, 2, _
                AST_STAR_GREEDY + humble, 4, AST_CAPTURE, 7, 0
            RegexAst.AstToBytecode ast, emptyIdentifierTree, False, actual, _
                expansion * DEFAULT_QUANTIFIER_EXPANSION_LIMIT
            ' The tail returns to the body, not the initial split. Its saved
            ' continuation must precede the body's deferred consuming alternative.
            ' Disabling expansion must retain this correctness transformation.
            MakeBytecode expected, MakeBytecodeParameters(startDfaBwRunOffset:=15), _
                REOP_SAVE, 0, splitOp, 8, _
                REOP_SPLIT1, 2, REOP_JUMP, 2, REOP_CHAR, 97, loopOp, -8, _
                REOP_SAVE, 1, REOP_MATCH, _
                splitOp, 8, REOP_SPLIT1, 2, REOP_JUMP, 2, REOP_CHAR, 97, loopOp, -8, REOP_MATCH
            Assert.SequenceEquals expected, actual
            Assert.AreEqual 0&, RegexNfaMatcher.NfaMatch(captures, actual, "aa")
            RegexDfaMatcher.DfaStateStructureInitialize state
            Assert.IsTrue RegexDfaMatcher.DfaFindMatch(span, state, actual, "aa", 1, False)
            Assert.AreEqual 0&, span.Length
        Next
    Next
End Sub

'@TestMethod("Ast")
Private Sub Ast_AstToBytecode0001()
    Dim ast As ArrayBuffer.Ty, actual() As Long, expected() As Long

    MakeArrayBuffer ast, _
        1, AST_STRING, 4, AscW("a"), AscW("b"), AscW("c"), AscW("d")
    RegexAst.AstToBytecode ast, emptyIdentifierTree, False, actual
    MakeBytecode expected, MakeBytecodeParameters(maxProperCaptureSlot:=-1, startDfaBwRunOffset:=7), _
        REOP_STRING, 4, AscW("a"), AscW("b"), AscW("c"), AscW("d"), _
        REOP_MATCH, _
        REOP_STRING, 4, AscW("d"), AscW("c"), AscW("b"), AscW("a"), _
        REOP_MATCH
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Ast")
Private Sub Ast_AstToBytecode0002()
    Dim ast As ArrayBuffer.Ty, actual() As Long, expected() As Long

    MakeArrayBuffer ast, _
        10, _
        AST_STRING, 2, AscW("a"), AscW("b"), _
        AST_STRING, 3, AscW("c"), AscW("d"), AscW("e"), _
        AST_DISJ, 1, 5
    RegexAst.AstToBytecode ast, emptyIdentifierTree, False, actual
    MakeBytecode expected, MakeBytecodeParameters(maxProperCaptureSlot:=-1, startDfaBwRunOffset:=14), _
        REOP_SPLIT1, 6, _
        REOP_STRING, 2, AscW("a"), AscW("b"), _
        REOP_JUMP, 5, _
        REOP_STRING, 3, AscW("c"), AscW("d"), AscW("e"), _
        REOP_MATCH, _
        REOP_SPLIT1, 6, _
        REOP_STRING, 2, AscW("b"), AscW("a"), _
        REOP_JUMP, 5, _
        REOP_STRING, 3, AscW("e"), AscW("d"), AscW("c"), _
        REOP_MATCH
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Ast")
Private Sub Ast_AstToBytecode0003()
    Dim ast As ArrayBuffer.Ty, actual() As Long, expected() As Long

    MakeArrayBuffer ast, _
        16, _
        AST_STRING, 2, AscW("a"), AscW("b"), _
        AST_STRING, 3, AscW("c"), AscW("d"), AscW("e"), _
        AST_STRING, 1, AscW("f"), _
        AST_DISJ, 5, 10, _
        AST_DISJ, 1, 13
    RegexAst.AstToBytecode ast, emptyIdentifierTree, False, actual
    MakeBytecode expected, MakeBytecodeParameters(maxProperCaptureSlot:=-1, startDfaBwRunOffset:=21), _
        REOP_SPLIT1, 6, _
        REOP_STRING, 2, AscW("a"), AscW("b"), _
        REOP_JUMP, 12, _
        REOP_SPLIT1, 7, _
        REOP_STRING, 3, AscW("c"), AscW("d"), AscW("e"), _
        REOP_JUMP, 3, _
        REOP_STRING, 1, AscW("f"), _
        REOP_MATCH, _
        REOP_SPLIT1, 6, _
        REOP_STRING, 2, AscW("b"), AscW("a"), _
        REOP_JUMP, 12, _
        REOP_SPLIT1, 7, _
        REOP_STRING, 3, AscW("e"), AscW("d"), AscW("c"), _
        REOP_JUMP, 3, _
        REOP_STRING, 1, AscW("f"), _
        REOP_MATCH
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Ast")
Private Sub Ast_AstToBytecode0004()
    ' Test AST traversal works when left child = right child
    Dim ast As ArrayBuffer.Ty, actual() As Long, expected() As Long

    MakeArrayBuffer ast, _
        6, AST_CHAR, AscW("a"), AST_CONCAT, 1, 1, AST_DISJ, 3, 3
    RegexAst.AstToBytecode ast, emptyIdentifierTree, False, actual
    MakeBytecode expected, MakeBytecodeParameters(maxProperCaptureSlot:=-1, startDfaBwRunOffset:=13), _
        REOP_SPLIT1, 6, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("a"), _
        REOP_JUMP, 4, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("a"), _
        REOP_MATCH, _
        REOP_SPLIT1, 6, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("a"), _
        REOP_JUMP, 4, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("a"), _
        REOP_MATCH
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


