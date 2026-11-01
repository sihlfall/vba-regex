Attribute VB_Name = "TestDfaMatcher"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

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

Private Sub CheckNullableLoopPriority(ByVal pattern As String, ByVal text As String, ByVal expectedLength As Long)
    Dim bytecode() As Long, state As RegexDfaMatcher.StateStructure
    Dim span As RegexMatcherCommon.StartLengthPair, j As Long
    RegexCompiler.Compile bytecode, pattern
    Assert.AreEqual 0&, bytecode(BYTECODE_IDX_PROPERTIES) And BYTECODE_PROPERTY_REQUIRES_DFS_MATCHING
    RegexDfaMatcher.DfaStateStructureInitialize state
    For j = 1 To 2 ' Check fresh transitions and cache hits.
        Assert.AreEqual expectedLength <> -1, RegexDfaMatcher.DfaFindMatch(span, state, bytecode, text, 1, False), pattern
        If expectedLength <> -1 Then
            Assert.AreEqual 1&, span.start, pattern
            Assert.AreEqual expectedLength, span.Length, pattern
        End If
    Next
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_NullableLoopPreservesEmptyBranchPriority()
    CheckNullableLoopPriority "(?:|a)*", "aa", 0
    CheckNullableLoopPriority "(?:a*?)*", "aa", 0
    CheckNullableLoopPriority "(?:a??)+", "aa", 0
    CheckNullableLoopPriority "((?:|a)*)", "aa", 0
    CheckNullableLoopPriority "(?:a{0,2}?)*", "aa", 0
    CheckNullableLoopPriority "(?:a{0,2}?)+", "aa", 0
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_NullableLoopRetainsConsumingAlternatives()
    CheckNullableLoopPriority "^(?:|a)*b$", "aab", 3
    CheckNullableLoopPriority "^(?:a*?)*b$", "aab", 3
    CheckNullableLoopPriority "^(?:a|)*b$", "aab", 3
    CheckNullableLoopPriority "(?:|a)*a", "aaa", 1
    ' RE2 keeps consuming here; DFS's empty-iteration rule returns only "!".
    CheckNullableLoopPriority "(?:\b|.)*", "!aaa", 4
    CheckNullableLoopPriority "^(?:[ab]*?)*b", "abb", 2
    CheckNullableLoopPriority "(?:a|)*", "aaa", 3
    CheckNullableLoopPriority "^(?:|a)*b$", "aa", -1
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_NullableLoopAssertionPriority()
    CheckNullableLoopPriority "(?:\b|.)*", "aaa", 0
    CheckNullableLoopPriority "(?:\b|.)*b", "bb", 1
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_MatchRegexp0100()
    Dim bytecode() As Long, pattern As String, haystack As String
    Dim slp As RegexMatcherCommon.StartLengthPair, _
        expectedSlp As RegexMatcherCommon.StartLengthPair
    Dim stateStruct As RegexDfaMatcher.StateStructure
    Dim didMatch As Boolean, matchExpected As Boolean
    
    pattern = "(?:abc)*"
    RegexCompiler.Compile bytecode, pattern

    haystack = ""
    With expectedSlp: .start = 1: .Length = 0: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch

    haystack = "abc"
    With expectedSlp: .start = 1: .Length = 3: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch
    
    haystack = "abcabc"
    With expectedSlp: .start = 1: .Length = 6: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch
    
    haystack = "abcabcabc"
    With expectedSlp: .start = 1: .Length = 9: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_MatchRegexp0101()
    Dim bytecode() As Long, pattern As String, haystack As String
    Dim slp As RegexMatcherCommon.StartLengthPair, _
        expectedSlp As RegexMatcherCommon.StartLengthPair
    Dim stateStruct As RegexDfaMatcher.StateStructure
    Dim didMatch As Boolean, matchExpected As Boolean
    
    pattern = "(?:abc)*$"
    RegexCompiler.Compile bytecode, pattern

    haystack = ""
    With expectedSlp: .start = 1: .Length = 0: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch

    haystack = "abc"
    With expectedSlp: .start = 1: .Length = 3: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch
    
    haystack = "abcabc"
    With expectedSlp: .start = 1: .Length = 6: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch
    
    haystack = "abcabcabc"
    With expectedSlp: .start = 1: .Length = 9: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_MatchRegexp0102()
    Dim bytecode() As Long, pattern As String, haystack As String
    Dim slp As RegexMatcherCommon.StartLengthPair, _
        expectedSlp As RegexMatcherCommon.StartLengthPair
    Dim stateStruct As RegexDfaMatcher.StateStructure
    Dim didMatch As Boolean, matchExpected As Boolean
    
    pattern = "^(?:abc)*?$"
    RegexCompiler.Compile bytecode, pattern

    haystack = ""
    With expectedSlp: .start = 1: .Length = 0: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch

    haystack = "abc"
    With expectedSlp: .start = 1: .Length = 3: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch
    
    haystack = "abcabc"
    With expectedSlp: .start = 1: .Length = 6: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch
    
    haystack = "abcabcabc"
    With expectedSlp: .start = 1: .Length = 9: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_MatchRegexp0104()
    Dim bytecode() As Long, pattern As String, haystack As String
    Dim slp As RegexMatcherCommon.StartLengthPair, _
        expectedSlp As RegexMatcherCommon.StartLengthPair
    Dim stateStruct As RegexDfaMatcher.StateStructure
    Dim didMatch As Boolean, matchExpected As Boolean
    
    pattern = "abc$"
    RegexCompiler.Compile bytecode, pattern

    haystack = ""
    With expectedSlp: .start = 0: .Length = 0: End With: matchExpected = False
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch

    haystack = "abc"
    With expectedSlp: .start = 1: .Length = 3: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch
    
    haystack = "abcabc"
    With expectedSlp: .start = 4: .Length = 3: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch
    
    haystack = "abcabcabc"
    With expectedSlp: .start = 7: .Length = 3: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_MatchRegexp0105()
    Dim bytecode() As Long, pattern As String, haystack As String
    Dim slp As RegexMatcherCommon.StartLengthPair, _
        expectedSlp As RegexMatcherCommon.StartLengthPair
    Dim stateStruct As RegexDfaMatcher.StateStructure
    Dim didMatch As Boolean, matchExpected As Boolean
    
    pattern = "^abc"
    RegexCompiler.Compile bytecode, pattern

    haystack = ""
    With expectedSlp: .start = 0: .Length = 0: End With: matchExpected = False
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch

    haystack = "abc"
    With expectedSlp: .start = 1: .Length = 3: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch
    
    haystack = "abcabc"
    With expectedSlp: .start = 1: .Length = 3: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch
    
    haystack = "abcabcabc"
    With expectedSlp: .start = 1: .Length = 3: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_MatchRegexp0106()
    Dim bytecode() As Long, pattern As String, haystack As String
    Dim slp As RegexMatcherCommon.StartLengthPair, _
        expectedSlp As RegexMatcherCommon.StartLengthPair
    Dim stateStruct As RegexDfaMatcher.StateStructure
    Dim didMatch As Boolean, matchExpected As Boolean
    
    pattern = "ab*c"
    RegexCompiler.Compile bytecode, pattern

    haystack = "abbbbbbc"
    With expectedSlp: .start = 1: .Length = 8: End With: matchExpected = True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual expectedSlp.start, slp.start
    Assert.AreEqual expectedSlp.Length, slp.Length
    Assert.AreEqual matchExpected, didMatch
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_MatchRegexp0906()
    ' On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 3, _
        REOP_DOT, _
        REOP_JUMP, -5, _
        REOP_STRING, 6, AscW("a"), AscW("a"), AscW("b"), AscW("b"), AscW("a"), AscW("a"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    inputStr = "xxxaabbaabbaabbaabbaabbaayaabbay"
    
    Dim result As Long
    result = RegexDfaMatcher.DfaMatch(bytecode, inputStr, False)
    Assert.AreEqual 25&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_MatchRegexp0907()
    ' star
    ' On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    inputStr = "xxxxabbbbcd"
    
    Dim result As Long
    result = RegexDfaMatcher.DfaMatch(bytecode, inputStr, False)
    Assert.AreEqual 10&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_MatchRegexp0908()
    ' star
    ' On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 10, _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("c"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("d"), _
        REOP_JUMP, 2, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    inputStr = "cccccccccccx"
    
    Dim result As Long
    result = RegexDfaMatcher.DfaMatch(bytecode, inputStr, False)
    Assert.AreEqual 1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_MatchRegexp0909()
    ' On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_ASSERT_WORD_BOUNDARY, _
        REOP_STRING, 3, AscW("a"), AscW("b"), AscW("c"), _
        REOP_ASSERT_WORD_BOUNDARY, _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    inputStr = "xabc abcx abc yy"
    
    Dim result As Long
    result = RegexDfaMatcher.DfaMatch(bytecode, inputStr, False)
    Assert.AreEqual 13&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_MatchRegexp0920()
    ' Matching in reverse direction
    ' On Error GoTo TestFail
    
    Dim result As Long
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(startDfaBwRunOffset:=10), _
        REOP_SAVE, 0, _
        REOP_STRING, 3, AscW("a"), AscW("b"), AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH, _
        REOP_STRING, 3, AscW("c"), AscW("b"), AscW("a"), _
        REOP_MATCH
        
    inputStr = "abc"
    
    result = RegexDfaMatcher.DfaMatch(bytecode, inputStr, False)
    Assert.AreEqual Len(inputStr), result
    
    result = RegexDfaMatcher.DfaMatchFrom(bytecode, inputStr, result, -1, False)
    Assert.AreEqual 0&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_MatchRegexp0921()
    ' Matching in reverse direction
    ' On Error GoTo TestFail
    
    Dim result As Long
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(startDfaBwRunOffset:=12), _
        REOP_SAVE, 0, _
        REOP_ASSERT_START, _
        REOP_STRING, 3, AscW("a"), AscW("b"), AscW("c"), _
        REOP_ASSERT_END, _
        REOP_SAVE, 1, _
        REOP_MATCH, _
        REOP_ASSERT_END, _
        REOP_STRING, 3, AscW("c"), AscW("b"), AscW("a"), _
        REOP_ASSERT_START, _
        REOP_MATCH
        
    inputStr = "abc"
    
    result = RegexDfaMatcher.DfaMatch(bytecode, inputStr, False)
    Assert.AreEqual Len(inputStr), result
    
    result = RegexDfaMatcher.DfaMatchFrom(bytecode, inputStr, result, -1, False)
    Assert.AreEqual 0&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub AssertMatchSpan( _
    ByVal pattern As String, ByVal haystack As String, _
    ByVal expectedStart As Long, ByVal expectedLength As Long, _
    Optional ByVal matchFrom As Long = 1, _
    Optional ByVal multiline As Boolean = False, _
    Optional ByVal dotAll As Boolean = False, _
    Optional ByVal finishEagerly As Boolean = False _
)
    Dim bytecode() As Long, stateStruct As RegexDfaMatcher.StateStructure
    Dim slp As RegexMatcherCommon.StartLengthPair, didMatch As Boolean, run As Long
    RegexCompiler.Compile bytecode, pattern
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct, multiline, dotAll
    For run = 1 To 2 ' cold cache, then cached transitions
        didMatch = RegexDfaMatcher.DfaFindMatch( _
            slp, stateStruct, bytecode, haystack, matchFrom, finishEagerly)
        Assert.AreEqual (expectedStart > 0), didMatch, pattern
        Assert.AreEqual expectedStart, slp.start, pattern
        Assert.AreEqual expectedLength, slp.Length, pattern
    Next
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_PreservesMatchPriority()
    AssertMatchSpan "a.*z|b", "abz", 1, 3
    AssertMatchSpan "a.*z|b", "ab", 2, 1
    AssertMatchSpan "b|ab", "ab", 1, 2
    AssertMatchSpan "a|aa", "aa", 1, 1
    AssertMatchSpan "aa|a", "aa", 1, 2
    AssertMatchSpan "a*", "aaa", 1, 3
    AssertMatchSpan "a*?", "aaa", 1, 0
    AssertMatchSpan "(?:a?|b)*c", "abbac", 1, 5
    AssertMatchSpan "(?:a*?)*b", "aaab", 1, 4
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_EagerMode()
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, "a+"
    Assert.AreEqual 1&, RegexDfaMatcher.DfaMatch(bytecode, "aaa", True)
    Assert.AreEqual 3&, RegexDfaMatcher.DfaMatch(bytecode, "aaa", False)
    Assert.AreEqual 2&, RegexDfaMatcher.DfaMatchFrom(bytecode, "aaa", 3, True, True)
    Assert.AreEqual 0&, RegexDfaMatcher.DfaMatchFrom(bytecode, "aaa", 3, True, False)
    AssertMatchSpan "a+", "aaa", 1, 1, finishEagerly:=True
    AssertMatchSpan "a.*z|b", "abz", 2, 1, finishEagerly:=True
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_ReverseBounds()
    Dim bytecode() As Long
    AssertMatchSpan "a*", "aaa", 2, 2, matchFrom:=2
    AssertMatchSpan "a*", "aaa", 4, 0, matchFrom:=4
    AssertMatchSpan "\ba+", "aaa", 0, 0, matchFrom:=2
    AssertMatchSpan "\Ba+", "aaa", 2, 2, matchFrom:=2
    AssertMatchSpan "a+", "aaa", 0, 0, matchFrom:=0
    AssertMatchSpan "a*", "aaa", 0, 0, matchFrom:=5
    RegexCompiler.Compile bytecode, "a*b"
    Assert.AreEqual -1&, RegexDfaMatcher.DfaMatchFrom(bytecode, "aaa", 3, True, False)
    Assert.AreEqual -1&, RegexDfaMatcher.DfaMatchFrom(bytecode, "aaa", -1, False, False)
    Assert.AreEqual -1&, RegexDfaMatcher.DfaMatchFrom(bytecode, "aaa", 4, False, False)
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_AssertionsAndModifiers()
    AssertMatchSpan "^$", "", 1, 0
    AssertMatchSpan "^$", "", 1, 0, multiline:=True
    AssertMatchSpan "^a+$", "aaa", 1, 3, multiline:=True
    AssertMatchSpan "^a+$", "x" & vbLf & "aaa", 3, 3, multiline:=True
    AssertMatchSpan "^a+$", "aaa" & vbLf & "x", 1, 3, multiline:=True
    AssertMatchSpan "a$", "a" & vbLf, 0, 0
    AssertMatchSpan "(?m:^a+$)", "x" & vbCrLf & "aaa", 4, 3
    AssertMatchSpan "(?-m:^a+$)", "x" & vbLf & "aaa", 0, 0, multiline:=True
    AssertMatchSpan "a.b", "a" & vbLf & "b", 0, 0
    AssertMatchSpan "a.b", "a" & vbLf & "b", 1, 3, dotAll:=True
    AssertMatchSpan "(?s:a.b)", "a" & vbLf & "b", 1, 3
    AssertMatchSpan "(?-s:a.b)", "a" & vbLf & "b", 0, 0, dotAll:=True
    AssertMatchSpan "(?i:abc)[D-F]", "xAbCE", 2, 4
    AssertMatchSpan "\babc\b", "xabc abc!", 6, 3
    AssertMatchSpan "\B", "", 1, 0
    AssertMatchSpan "\b\B", "a", 0, 0
    AssertMatchSpan "a\b$", "a", 1, 1
    AssertMatchSpan "[^a]+", "aaa" & ChrW(&HFFFF&), 4, 1
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_NamedCapturesAreSkipped()
    AssertMatchSpan "(?<word>abc)+", "xabcabc!", 2, 6
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_CacheAcrossInputsAndBounds()
    Dim bytecode() As Long, stateStruct As RegexDfaMatcher.StateStructure
    Dim slp As RegexMatcherCommon.StartLengthPair, i As Long, stateCount As Long
    RegexCompiler.Compile bytecode, "\ba+|\Ba+"
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    For i = 1 To 3
        Assert.IsTrue RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, "aaa", 1, False)
        Assert.AreEqual 1&, slp.start
        Assert.AreEqual 3&, slp.Length
        Assert.IsTrue RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, "aaa", 2, False)
        Assert.AreEqual 2&, slp.start
        Assert.AreEqual 2&, slp.Length
        Assert.IsFalse RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, "bbb", 1, False)
        Assert.IsTrue RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, " aaa!", 1, False)
        Assert.AreEqual 2&, slp.start
        Assert.AreEqual 3&, slp.Length
        If i = 1 Then stateCount = stateStruct.nextFreeStateId
        Assert.AreEqual stateCount, stateStruct.nextFreeStateId
    Next
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_Reinitialization()
    Dim bytecode() As Long, stateStruct As RegexDfaMatcher.StateStructure
    Dim slp As RegexMatcherCommon.StartLengthPair
    RegexCompiler.Compile bytecode, "a*"
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    Assert.IsTrue RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, "aaa", 1, False)
    RegexCompiler.Compile bytecode, "b+"
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct, True, True
    Assert.AreEqual 0&, stateStruct.stateListsBuffer.Length
    Assert.AreEqual 0&, stateStruct.transitionTableCount
    Assert.AreEqual 1&, stateStruct.nextFreeStateId
    Assert.IsTrue stateStruct.multiline
    Assert.IsTrue stateStruct.dotAll
    Assert.IsTrue RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, "abb", 1, False)
    Assert.AreEqual 2&, slp.start
    Assert.AreEqual 2&, slp.Length
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_StateCacheSaturation()
    Dim bytecode() As Long, stateStruct As RegexDfaMatcher.StateStructure
    Dim pattern As String, haystack As String, slp As RegexMatcherCommon.StartLengthPair
    Dim run As Long
    pattern = "^(?:" & String$(2400, "a") & "|b)$"
    haystack = String$(2400, "a")
    RegexCompiler.Compile bytecode, pattern
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    For run = 1 To 2
        Assert.IsTrue RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
        Assert.AreEqual 1&, slp.start
        Assert.AreEqual 2400&, slp.Length
        Assert.IsTrue stateStruct.nextFreeStateId <= 769
    Next
    haystack = haystack & "b"
    Assert.IsFalse RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_TransitionCacheSaturation()
    Dim bytecode() As Long, stateStruct As RegexDfaMatcher.StateStructure
    Dim haystack As String, slp As RegexMatcherCommon.StartLengthPair, i As Long, run As Long
    For i = 256 To 10255
        haystack = haystack & ChrW(i)
    Next
    RegexCompiler.Compile bytecode, "z|y"
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    For run = 1 To 2
        Assert.IsFalse RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
        Assert.IsTrue stateStruct.transitionTableCount <= 3072
    Next
    haystack = haystack & "z"
    Assert.IsTrue RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual 10001&, slp.start
    Assert.AreEqual 1&, slp.Length
End Sub


'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_CacheResetPreservesPendingMatch()
    Dim haystack As String, i As Long
    haystack = "ab"
    For i = 256 To 10255
        haystack = haystack & ChrW(i)
    Next
    ' Remember b while the higher-priority branch keeps scanning through resets.
    AssertMatchSpan "a.*z|b", haystack, 2, 1, dotAll:=True
    ' If the higher-priority branch eventually succeeds, replace the old match.
    AssertMatchSpan "a.*z|b", haystack & "z", 1, Len(haystack) + 1, dotAll:=True
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_StateBufferSaturation()
    Dim bytecode() As Long, stateStruct As RegexDfaMatcher.StateStructure
    Dim pattern As String, haystack As String, slp As RegexMatcherCommon.StartLengthPair
    pattern = "a*" & String$(400, "a") & "b"
    haystack = String$(1200, "a") & "b"
    RegexCompiler.Compile bytecode, pattern
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    Assert.IsTrue RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, 1, False)
    Assert.AreEqual 1&, slp.start
    Assert.AreEqual 1201&, slp.Length
    Assert.IsTrue stateStruct.stateListsBuffer.Length <= 65536
End Sub

Private Sub CompareWithNfa(ByVal pattern As String, Optional ByVal compareDfs As Boolean = True)
    Dim bytecode() As Long, stateStruct As RegexDfaMatcher.StateStructure
    Dim captures As RegexMatcherCommon.CapturesTy, slp As RegexMatcherCommon.StartLengthPair
    Dim dfsCaptures As RegexMatcherCommon.CapturesTy, dfsContext As RegexDfsMatcher.DfsMatcherContext
    Dim dfsResult As Long
    Dim haystack As String, alphabet As String, description As String
    Dim sample As Long, n As Long, sp As Long, modifiers As Long, result As Long
    Dim didMatch As Boolean, multiline As Boolean, dotAll As Boolean
    RegexCompiler.Compile bytecode, pattern
    alphabet = "ab_" & vbLf
    For modifiers = 0 To 3
        multiline = (modifiers And 1) <> 0
        dotAll = (modifiers And 2) <> 0
        RegexDfaMatcher.DfaStateStructureInitialize stateStruct, multiline, dotAll
        For sample = 0 To 84 ' all strings of length zero through three
            haystack = "": n = sample
            Do While n > 0
                n = n - 1
                haystack = haystack & Mid$(alphabet, n Mod 4 + 1, 1)
                n = n \ 4
            Loop
            For sp = 0 To Len(haystack)
                result = RegexNfaMatcher.NfaMatchFrom(captures, bytecode, haystack, sp, multiline, dotAll)
                If compareDfs Then _
                    dfsResult = RegexDfsMatcher.DfsMatchFrom(dfsContext, dfsCaptures, bytecode, haystack, sp, multiline:=multiline, dotAll:=dotAll)
                didMatch = RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, haystack, sp + 1, False)
                description = pattern & " sample=" & sample & " sp=" & sp & " flags=" & modifiers
                If compareDfs Then
                    Assert.AreEqual result, dfsResult, description
                    If result <> -1 And dfsResult <> -1 Then
                        Assert.AreEqual captures.entireMatch.start, dfsCaptures.entireMatch.start, description
                        Assert.AreEqual captures.entireMatch.Length, dfsCaptures.entireMatch.Length, description
                    End If
                End If
                Assert.AreEqual (result <> -1), didMatch, description
                If didMatch And result <> -1 Then
                    Assert.AreEqual captures.entireMatch.start, slp.start, description
                    Assert.AreEqual captures.entireMatch.Length, slp.Length, description
                End If
            Next
        Next
    Next
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_AgreesWithNfa()
    CompareWithNfa ""
    CompareWithNfa "a"
    CompareWithNfa "ab"
    CompareWithNfa "ab$"
    CompareWithNfa "^ab"
    CompareWithNfa "^ab$"
    CompareWithNfa "(?m:^ab$)"
    CompareWithNfa "(?-m:^ab$)"
    CompareWithNfa "a*"
    CompareWithNfa "a*?"
    CompareWithNfa "a+"
    CompareWithNfa "a+?"
    CompareWithNfa "a?"
    CompareWithNfa "a??"
    CompareWithNfa "a.*b|a"
    CompareWithNfa "b|ab"
    CompareWithNfa "a|ab"
    CompareWithNfa "ab|a"
    ' The original DFS chooses different spans for nullable alternatives inside
    ' greedy loops. Keep this existing DFA/NFA check independent of that issue.
    CompareWithNfa "(?:a?|b)*", compareDfs:=False
    CompareWithNfa "(?:a*?)*b"
    CompareWithNfa "(?:|a)*"
    CompareWithNfa "^(?:[ab]*?)*b"
    ' RE2 and DFS differ on successful empty iterations involving assertions.
    CompareWithNfa "(?:\b|.)*", compareDfs:=False
    CompareWithNfa "(?:\b|\B)*a?", compareDfs:=False
    CompareWithNfa "(?:^|$|.)*", compareDfs:=False
    CompareWithNfa "(?:a|ab)*?b"
    CompareWithNfa "(?:a|b)*a"
    CompareWithNfa "^a*$"
    CompareWithNfa "^$"
    CompareWithNfa "a$|b"
    CompareWithNfa "\ba*"
    CompareWithNfa "\Ba+"
    CompareWithNfa "a\b|ab"
    CompareWithNfa "\ba\b|\Baa"
    CompareWithNfa "\b|a"
    CompareWithNfa "a|\b"
    CompareWithNfa "\b\B|a"
    CompareWithNfa "\b^a*$\b"
    CompareWithNfa "[^a]*"
    CompareWithNfa "[a-b_]+"
    CompareWithNfa "(?i:ab)+|_"
    CompareWithNfa "(?<word>a+)b?"
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_CaptureIntegrationPreservesLeftmostStart()
    Dim regex As StaticRegex.RegexTy, matcherState As StaticRegex.MatcherStateTy
    StaticRegex.InitializeRegex regex, "(?<word>a.*?z)|b"
    Assert.IsTrue StaticRegex.Match(matcherState, regex, "abz")
    Assert.AreEqual "abz", StaticRegex.GetCapture(matcherState, "abz")
    Assert.AreEqual "abz", StaticRegex.GetCaptureByName(matcherState, regex, "abz", "word")
    StaticRegex.InitializeMatcherState matcherState
    Assert.IsTrue StaticRegex.MatchNext(matcherState, regex, "abz abz")
    Assert.AreEqual "abz", StaticRegex.GetCapture(matcherState, "abz abz")
    Assert.IsTrue StaticRegex.MatchNext(matcherState, regex, "abz abz")
    Assert.AreEqual "abz", StaticRegex.GetCapture(matcherState, "abz abz")
    Assert.IsFalse StaticRegex.MatchNext(matcherState, regex, "abz abz")
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_LiteralSpans()
    AssertMatchSpan "needle", "xxneedleneedle", 3, 6
    AssertMatchSpan "needle", "xxneedleneedle", 9, 6, matchFrom:=4
    AssertMatchSpan "needle$", "xxneedle", 3, 6
    AssertMatchSpan "^needle", "xxneedle", 0, 0
    AssertMatchSpan "^needle$", "needle", 1, 6
    AssertMatchSpan "aaa$", "aaaa", 2, 3
    AssertMatchSpan "aaa$", "aaaa" & vbLf & "x", 2, 3, multiline:=True
    AssertMatchSpan "^aaa", vbCrLf & "aaa", 3, 3, multiline:=True
    AssertMatchSpan "(?m:^aaa$)", "x" & ChrW(&H2028) & "aaa" & ChrW(&H2029), 3, 3
    AssertMatchSpan "(?-m:^aaa$)", "x" & vbLf & "aaa", 0, 0, multiline:=True
    AssertMatchSpan "\^needle\$", "xx^needle$", 3, 8
    AssertMatchSpan "a\.b", "xa.b", 2, 3
    AssertMatchSpan "\u0000x", "y" & ChrW(0) & "x", 2, 2
    AssertMatchSpan "\uD83D\uDE00", "x" & ChrW(&HD83D&) & ChrW(&HDE00&), 2, 2
    AssertMatchSpan "", "abc", 2, 0, matchFrom:=2
    AssertMatchSpan "$", "abc", 4, 0
    AssertMatchSpan "^", "abc", 1, 0
    AssertMatchSpan "^$", "a" & vbCrLf & "b", 3, 0, multiline:=True
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_LiteralsAvoidTables()
    Dim stateStruct As RegexDfaMatcher.StateStructure, bytecode() As Long
    Dim slp As RegexMatcherCommon.StartLengthPair, pattern As String
    pattern = "^" & String$(2048, "a") & "$"
    RegexCompiler.Compile bytecode, pattern
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    Assert.IsFalse stateStruct.tablesInitialized
    Assert.IsTrue RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, String$(2048, "a"), 1, False)
    Assert.AreEqual 1&, slp.start
    Assert.AreEqual 2048&, slp.Length
    Assert.IsTrue stateStruct.literal.isLiteral
    Assert.IsFalse stateStruct.tablesInitialized
    Assert.AreEqual 2048&, RegexDfaMatcher.DfaMatch(bytecode, String$(2048, "a"), False)
    Assert.AreEqual 2048&, RegexDfaMatcher.DfaMatch(bytecode, String$(2048, "a"), True)
    Assert.AreEqual -1&, RegexDfaMatcher.DfaMatchFrom(bytecode, String$(2048, "a"), 1, False, False)

    ' Reinitialize a previously allocated cache, then use the literal path.
    RegexCompiler.Compile bytecode, "a+"
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    Assert.IsTrue RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, "aaa", 1, False)
    Assert.IsTrue stateStruct.tablesInitialized
    RegexCompiler.Compile bytecode, "b"
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    Assert.IsTrue RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, "ab", 1, False)
    Assert.AreEqual 2&, slp.start
    Assert.IsFalse stateStruct.tablesInitialized
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_LiteralFallbacks()
    Dim stateStruct As RegexDfaMatcher.StateStructure, bytecode() As Long
    Dim slp As RegexMatcherCommon.StartLengthPair
    RegexCompiler.Compile bytecode, "needle", caseInsensitive:=True
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    Assert.IsTrue RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, "xNEEDLE", 1, False)
    Assert.AreEqual 2&, slp.start
    Assert.IsFalse stateStruct.literal.isLiteral
    Assert.IsTrue stateStruct.tablesInitialized
    AssertMatchSpan "(?i:ab)c", "xABc", 2, 3
    AssertMatchSpan "(?i:ab)c", "xABC", 0, 0
    AssertMatchSpan "(abc)", "xabc", 2, 3
    AssertMatchSpan "abc|de", "xdeabc", 2, 2
    AssertMatchSpan "a?", "b", 1, 0
End Sub

'@TestMethod("DfaMatcher")
Private Sub InputBuffer_AllUtf16CodeUnits()
    Dim text As String, inputBuffer As RegexMatcherCommon.InputBufferTy
    Dim code As Long, sp As Long
    text = String$(65536, vbNullChar)
    For code = 0 To 65535
        Mid$(text, code + 1, 1) = ChrW(code)
    Next
    RegexMatcherCommon.InitializeInputBuffer inputBuffer, text
    Assert.AreEqual 65536&, inputBuffer.Length
    Assert.AreEqual 0&, inputBuffer.windowEnd
    sp = 0
    For code = 0 To 65535
        Assert.AreEqual code, RegexMatcherCommon.PeekBufferCharCode(inputBuffer, text, code)
        Assert.AreEqual code, RegexMatcherCommon.GetBufferCharCode(inputBuffer, text, sp, 1)
    Next
    Assert.AreEqual 65536&, sp
    Assert.AreEqual -1&, RegexMatcherCommon.GetBufferCharCode(inputBuffer, text, sp, 1)
    Assert.AreEqual 65536&, sp
    sp = 65535
    For code = 65535 To 0 Step -1
        Assert.AreEqual code, RegexMatcherCommon.GetBufferCharCode(inputBuffer, text, sp, -1)
    Next
    Assert.AreEqual -1&, sp
    Assert.AreEqual -1&, RegexMatcherCommon.GetBufferCharCode(inputBuffer, text, sp, -1)
    Assert.AreEqual -1&, sp
    Assert.AreEqual -1&, RegexMatcherCommon.PeekBufferCharCode(inputBuffer, text, &H7FFFFFFF)
    Assert.AreEqual -1&, RegexMatcherCommon.PeekBufferCharCode(inputBuffer, text, &H80000000)
End Sub

'@TestMethod("DfaMatcher")
Private Sub InputBuffer_EmptyAndReset()
    Dim text As String, inputBuffer As RegexMatcherCommon.InputBufferTy, sp As Long
    text = "ab"
    RegexMatcherCommon.InitializeInputBuffer inputBuffer, text
    Assert.AreEqual 0&, inputBuffer.windowEnd
    Assert.AreEqual 97&, RegexMatcherCommon.PeekBufferCharCode(inputBuffer, text, 0)
    text = "xy"
    RegexMatcherCommon.InitializeInputBuffer inputBuffer, text
    Assert.AreEqual 120&, RegexMatcherCommon.PeekBufferCharCode(inputBuffer, text, 0)
    text = vbNullString
    RegexMatcherCommon.InitializeInputBuffer inputBuffer, text
    Assert.AreEqual 0&, inputBuffer.Length
    Assert.AreEqual -1&, RegexMatcherCommon.PeekBufferCharCode(inputBuffer, text, 0)
    Assert.AreEqual -1&, RegexMatcherCommon.GetBufferCharCode(inputBuffer, text, sp, 1)
    Assert.AreEqual 0&, sp
    text = ""
    RegexMatcherCommon.InitializeInputBuffer inputBuffer, text
    Assert.AreEqual -1&, RegexMatcherCommon.PeekBufferCharCode(inputBuffer, text, -1)
    Assert.AreEqual -1&, RegexMatcherCommon.PeekBufferCharCode(inputBuffer, text, 1)
End Sub

'@TestMethod("DfaMatcher")
Private Sub InputBuffer_LazyGrowthAndBackwardDistance()
    Dim text As String, inputBuffer As RegexMatcherCommon.InputBufferTy
    Dim oldEnd As Long, oldCapacity As Long, target As Long, code As Long
    text = String$(100000, "a")
    Mid$(text, 49001, 1) = ChrW(&HFFFF&)
    RegexMatcherCommon.InitializeInputBuffer inputBuffer, text
    Assert.AreEqual 0&, inputBuffer.windowEnd
    Assert.AreEqual INPUT_BUFFER_INITIAL_BYTES, inputBuffer.Capacity
    Assert.AreEqual 97&, RegexMatcherCommon.PeekBufferCharCode(inputBuffer, text, 50000)
    Assert.IsTrue UBound(inputBuffer.bytes) + 1 <= INPUT_BUFFER_INITIAL_BYTES
    oldEnd = inputBuffer.windowEnd: oldCapacity = inputBuffer.Capacity
    Assert.AreEqual 65535&, RegexMatcherCommon.PeekBufferCharCode(inputBuffer, text, 49000)
    Assert.IsTrue inputBuffer.Capacity >= oldCapacity * 1.4
    Assert.IsTrue inputBuffer.Capacity >= 2 * (oldEnd - 49000)
    Assert.IsTrue inputBuffer.windowStart <= 49000
    Assert.AreEqual oldEnd, inputBuffer.windowEnd
    oldCapacity = inputBuffer.Capacity
    Assert.AreEqual 97&, RegexMatcherCommon.PeekBufferCharCode(inputBuffer, text, 50000)
    Assert.AreEqual oldCapacity, inputBuffer.Capacity
    Assert.AreEqual 97&, RegexMatcherCommon.PeekBufferCharCode(inputBuffer, text, 0)
    Assert.AreEqual INPUT_BUFFER_MAX_BYTES, inputBuffer.Capacity
    Assert.AreEqual 97&, RegexMatcherCommon.PeekBufferCharCode(inputBuffer, text, 99000)
    target = 99000
    For code = 99000 To 0 Step -1
        If code = 49000 Then oldEnd = 65535 Else oldEnd = 97
        Assert.AreEqual oldEnd, RegexMatcherCommon.GetBufferCharCode(inputBuffer, text, target, -1)
    Next
    Assert.AreEqual -1&, target
    Assert.IsTrue UBound(inputBuffer.bytes) + 1 <= INPUT_BUFFER_MAX_BYTES
    Assert.AreEqual 0&, inputBuffer.Capacity Mod 2
End Sub

'@TestMethod("DfaMatcher")
Private Sub Matcher_WindowBoundariesAndLongStrings()
    Dim bytecode() As Long, captures As RegexMatcherCommon.CapturesTy
    Dim text As String, literal As String, n As Variant, sizes() As Long
    MakeArray sizes, 510&, 511&, 512&, 513&, 32768&, 65536&
    For Each n In sizes
        text = String$(n, "x") & vbCrLf & "ab"
        AssertMatchSpan "^[ab]+$", text, n + 3, 2, multiline:=True
        RegexCompiler.Compile bytecode, "^[ab]+$"
        Assert.AreEqual Len(text), RegexDfsMatcher.DfsMatch(captures, bytecode, text, multiline:=True)
        Assert.AreEqual n + 3, captures.entireMatch.start
        Assert.AreEqual Len(text), RegexNfaMatcher.NfaMatch(captures, bytecode, text, multiline:=True)
        Assert.AreEqual n + 3, captures.entireMatch.start
        AssertMatchSpan "\bab\b", String$(n, "!") & "ab!", n + 1, 2
    Next
    literal = String$(40000, "a") & "b"
    text = "x" & literal
    RegexCompiler.Compile bytecode, literal
    Assert.AreEqual Len(text), RegexDfsMatcher.DfsMatch(captures, bytecode, text)
    RegexCompiler.Compile bytecode, "(?<=" & literal & ")c"
    text = literal & "c"
    Assert.AreEqual Len(text), RegexDfsMatcher.DfsMatch(captures, bytecode, text)
    Assert.AreEqual Len(text), captures.entireMatch.start
End Sub

'@TestMethod("DfaMatcher")
Private Sub DfaMatcher_LeadingAnchorBoundsAndModifiers()
    AssertMatchSpan "^[ab]+$", "aaa", 1, 3
    AssertMatchSpan "^[ab]+$", "aaa", 0, 0, matchFrom:=2
    AssertMatchSpan "^a|b", "xxb", 3, 1
    AssertMatchSpan "(?-m:^[ab]+$)", "x" & vbLf & "ab", 0, 0, multiline:=True
    AssertMatchSpan "(?m:^[ab]+$)", "x" & ChrW(&H2028) & "ab", 3, 2
    AssertMatchSpan "^[ab]+$", vbLf & "ab", 2, 2, matchFrom:=2, multiline:=True
    AssertMatchSpan "^a+", "aaa", 1, 1, finishEagerly:=True
    AssertMatchSpan "(?m:^a.*z|^b)", "a" & vbLf & "bz", 1, 4, dotAll:=True
    AssertMatchSpan "(?m:^a.*z|^b)", "a" & vbLf & "bz", 3, 1, dotAll:=True, finishEagerly:=True
End Sub

'@TestMethod("DfaMatcher")
Private Sub Matcher_ByteBufferUnicodeAndReverseReads()
    Dim text As String, bytecode() As Long
    Dim stateStruct As RegexDfaMatcher.StateStructure, captures As RegexMatcherCommon.CapturesTy
    Dim slp As RegexMatcherCommon.StartLengthPair
    text = ChrW(0) & ChrW(&H7FFF&) & ChrW(&H8000&) & ChrW(&HD83D&) & ChrW(&HDE00&) & ChrW(&HFFFF&)
    text = text & ChrW(&HD800&) & "z" & ChrW(&HDC00&)
    RegexCompiler.Compile bytecode, "^[^x]+$"
    RegexDfaMatcher.DfaStateStructureInitialize stateStruct
    Assert.IsTrue RegexDfaMatcher.DfaFindMatch(slp, stateStruct, bytecode, text, 1, False)
    Assert.AreEqual 1&, slp.start
    Assert.AreEqual Len(text), slp.Length
    Assert.AreEqual 0&, RegexDfaMatcher.DfaMatchFrom(bytecode, text, Len(text), True, False)
    Assert.AreEqual Len(text), RegexDfsMatcher.DfsMatch(captures, bytecode, text)
    Assert.AreEqual 1&, captures.entireMatch.start
    Assert.AreEqual Len(text), RegexNfaMatcher.NfaMatch(captures, bytecode, text)
    Assert.AreEqual 1&, captures.entireMatch.start

    RegexCompiler.Compile bytecode, "(.+)\1"
    text = text & text
    Assert.AreEqual Len(text), RegexDfsMatcher.DfsMatch(captures, bytecode, text)
    Assert.AreEqual Len(text) \ 2, captures.numberedCaptures(0).Length
    RegexCompiler.Compile bytecode, "(?<=\u0000\u8000)\uFFFF"
    text = ChrW(0) & ChrW(&H8000&) & ChrW(&HFFFF&)
    Assert.AreEqual 3&, RegexDfsMatcher.DfsMatch(captures, bytecode, text)
    Assert.AreEqual 3&, captures.entireMatch.start
    Assert.AreEqual 1&, captures.entireMatch.Length
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp0001()
    ' Matching simple two-character pattern
    'On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "ab"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp0002()
    ' Matching simple two-character string
    'On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("x"), _
        REOP_STRING, 2, AscW("a"), AscW("b"), _
        REOP_CHAR, AscW("y"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "xaby"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp010()
    ' star, matches several times
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    inputStr = "abbbc"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp011()
    ' star, matches once
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    inputStr = "abc"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp012()
    ' star, matches zero times
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    inputStr = "ac"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp013()
    ' empty star with capture group
    'On Error GoTo TestFail

    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(maxProperCaptureSlot:=3), _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 6, _
        REOP_SAVE, 2, _
        REOP_SAVE, 3, _
        REOP_JUMP, -8, _
        REOP_CHAR, AscW("x"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "x"

    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp014()
    ' star, complex pattern, matches zero times
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(maxProperCaptureSlot:=3), _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 10, _
        REOP_SAVE, 2, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 3, _
        REOP_JUMP, -12, _
        REOP_CHAR, AscW("x"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "x"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp015()
    ' star, complex pattern, matches once
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(maxProperCaptureSlot:=3), _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 10, _
        REOP_SAVE, 2, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 3, _
        REOP_JUMP, -12, _
        REOP_CHAR, AscW("x"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "abx"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp016()
    ' star, complex pattern, matches twice
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(maxProperCaptureSlot:=3), _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 10, _
        REOP_SAVE, 2, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 3, _
        REOP_JUMP, -12, _
        REOP_CHAR, AscW("x"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "ababx"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp017()
    ' star, non-greedy
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT2, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    inputStr = "abbc"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual 4&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp020()
    ' Plus
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    inputStr = "abbbc"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp021()
    ' Plus, complex pattern, matches once
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(maxProperCaptureSlot:=3), _
        REOP_SAVE, 0, _
        REOP_SAVE, 2, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 3, _
        REOP_SPLIT1, 10, _
        REOP_SAVE, 2, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 3, _
        REOP_JUMP, -12, _
        REOP_CHAR, AscW("x"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "abx"
    
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp022()
    ' Plus, complex pattern, matches twice
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(maxProperCaptureSlot:=3), _
        REOP_SAVE, 0, _
        REOP_SAVE, 2, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 3, _
        REOP_SPLIT1, 10, _
        REOP_SAVE, 2, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 3, _
        REOP_JUMP, -12, _
        REOP_CHAR, AscW("x"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "ababx"
    
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub



Private Sub SetupReCtxDisjunction(ByRef bytecode() As Long)
    MakeBytecode bytecode, MakeBytecodeParameters(maxProperCaptureSlot:=1), _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("a"), _
        REOP_JUMP, 2, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp050()
    ' Disjunction, match first alternative (single character)
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    SetupReCtxDisjunction bytecode
    inputStr = "a"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp051()
    ' Disjunction, match second alternative (single character)
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    SetupReCtxDisjunction bytecode
    inputStr = "b"
    
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp101()
    ' Parentheses
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(maxProperCaptureSlot:=3), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SAVE, 2, _
        REOP_CHAR, AscW("b"), _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 3, _
        REOP_CHAR, AscW("d"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "abcd"
    
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp102()
    ' Nested parentheses
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(maxProperCaptureSlot:=5), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SAVE, 2, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 4, _
        REOP_CHAR, AscW("c"), _
        REOP_CHAR, AscW("d"), _
        REOP_SAVE, 5, _
        REOP_CHAR, AscW("e"), _
        REOP_SAVE, 3, _
        REOP_CHAR, AscW("f"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "abcdef"
    
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp200()
    ' Caret, matching
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_ASSERT_START, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "abc"
    
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual 2&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp201()
    ' Caret, not matching
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_ASSERT_START, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "zabc"
    
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp202()
    ' Dollar, matching
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_ASSERT_END, _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "ab"
    
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual 2&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp203()
    ' Dollar, not matching
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_ASSERT_END, _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "abz"
    
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp204()
    ' Period, matching
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_DOT, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "azb"
    
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp210()
    ' Word boundary, matching
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW(" "), _
        REOP_ASSERT_WORD_BOUNDARY, _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = " a"
    
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual 1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp211()
    ' Word boundary, not matching
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW(" "), _
        REOP_ASSERT_WORD_BOUNDARY, _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "  "
    
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp212()
    ' Not word boundary, matching
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW(" "), _
        REOP_ASSERT_NOT_WORD_BOUNDARY, _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "  "
    
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual 1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp213()
    ' Not word boundary, not matching
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW(" "), _
        REOP_ASSERT_NOT_WORD_BOUNDARY, _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = " a"
    
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp300()
    ' Range
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "ABC"
    
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp301()
    ' Range, failing match
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "AbC"
    
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp302()
    ' Range
    On Error GoTo TestFail
    
    Dim bytecode() As Long, inputStr As String, outCaptures As RegexMatcherCommon.CapturesTy
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_RANGES, 2, AscW("A"), AscW("Z"), AscW("a"), AscW("z"), _
        REOP_RANGES, 2, AscW("A"), AscW("Z"), AscW("a"), AscW("z"), _
        REOP_RANGES, 2, AscW("A"), AscW("Z"), AscW("a"), AscW("z"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    inputStr = "AbC"
    
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(outCaptures, bytecode, inputStr)
    Assert.AreEqual Len(inputStr), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
