Attribute VB_Name = "TestNfaMatcher"
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

Private Sub CheckMatchAt( _
    ByVal pattern As String, ByVal text As String, ByVal sp As Long, ByVal expected As Long, _
    Optional ByVal multiline As Boolean = False, Optional ByVal dotAll As Boolean = False _
)
    Dim bytecode() As Long, captures As RegexMatcherCommon.CapturesTy
    RegexCompiler.Compile bytecode, pattern
    Assert.AreEqual expected, RegexNfaMatcher.NfaMatchFrom( _
        captures, bytecode, text, sp, multiline:=multiline, dotAll:=dotAll, matchAt:=True)
    If expected <> -1 Then
        Assert.AreEqual sp + 1, captures.entireMatch.start
        Assert.AreEqual expected - sp, captures.entireMatch.Length
    End If
End Sub

'@TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchAtKeepsOffsetAndContext()
    Dim bytecode() As Long, captures As RegexMatcherCommon.CapturesTy
    RegexCompiler.Compile bytecode, "(a+)"
    Assert.AreEqual 5&, RegexNfaMatcher.NfaMatchFrom(captures, bytecode, "xxaaa", 0)
    Assert.AreEqual 3&, captures.entireMatch.start
    CheckMatchAt "(a+)", "xxaaa", 0, -1
    CheckMatchAt "(a+)", "xxaaa", 2, 5
    CheckMatchAt "(a+)", "xxaaa", -1, -1
    CheckMatchAt "(a+)", "xxaaa", 6, -1
    CheckMatchAt "\b(a+)", "xaaa", 1, -1
    CheckMatchAt "\B(a+)", "xaaa", 1, 4
    CheckMatchAt "(^a+)", "xaa", 1, -1
    CheckMatchAt "(^a+)", "x" & vbLf & "aa", 2, -1
    CheckMatchAt "(^a+)", "x" & vbLf & "aa", 2, 4, multiline:=True
    CheckMatchAt "(?m:(^a+)$)", "x" & vbLf & "aa", 2, 4
    CheckMatchAt "(?-m:(^a+)$)", "x" & vbLf & "aa", 2, -1, multiline:=True
    CheckMatchAt "(a+)$", "xxaa!", 2, -1
    CheckMatchAt "(a*)$", "xxx", 3, 3
    CheckMatchAt "(.+)$", "x" & vbLf & "aa", 1, -1
    CheckMatchAt "(.+)$", "x" & vbLf & "aa", 1, 4, dotAll:=True
End Sub

'@TestMethod("NfaMatcher")
Private Sub NfaMatcher_NullableLoopPreservesEmptyBranchPriority()
    Dim bytecode() As Long, expected As RegexMatcherCommon.CapturesTy
    Dim actual As RegexMatcherCommon.CapturesTy, pattern As String, i As Long
    For i = 0 To 5
        Select Case i
        Case 0: pattern = "(?:|a)*"
        Case 1: pattern = "(?:a*?)*"
        Case 2: pattern = "(?:a??)+"
        Case 3: pattern = "((?:|a)*)"
        Case 4: pattern = "(?:a{0,2}?)*"
        Case 5: pattern = "(?:a{0,2}?)+"
        End Select
        RegexCompiler.Compile bytecode, pattern
        Assert.AreEqual 0&, bytecode(BYTECODE_IDX_PROPERTIES) And BYTECODE_PROPERTY_REQUIRES_DFS_MATCHING
        Assert.AreEqual 0&, RegexDfsMatcher.DfsMatch(expected, bytecode, "aa")
        Assert.AreEqual 0&, RegexNfaMatcher.NfaMatch(actual, bytecode, "aa"), pattern
        Assert.AreEqual 1&, actual.entireMatch.start
        Assert.AreEqual 0&, actual.entireMatch.Length, pattern
        If i = 3 Then
            Assert.AreEqual 1&, actual.nNumberedCaptures
            Assert.AreEqual 1&, actual.numberedCaptures(0).start
            Assert.AreEqual 0&, actual.numberedCaptures(0).Length
        End If
    Next
End Sub

'@TestMethod("NfaMatcher")
Private Sub NfaMatcher_NullableLoopRetainsConsumingAlternatives()
    Dim bytecode() As Long, captures As RegexMatcherCommon.CapturesTy
    Dim pattern As String, text As String, expectedLength As Long, i As Long
    For i = 0 To 9
        text = "aab": expectedLength = 3
        Select Case i
        Case 0: pattern = "^(?:|a)*b$"
        Case 1: pattern = "^(?:a*?)*b$"
        Case 2: pattern = "^(?:a|)*b$"
        Case 3: pattern = "(?:|a)*a": text = "aaa": expectedLength = 1
        ' RE2 keeps consuming here; DFS's empty-iteration rule returns only "!".
        Case 4: pattern = "(?:\b|.)*": text = "!aaa": expectedLength = 4
        Case 5: pattern = "(?:a|)*": text = "aaa"
        Case 6: pattern = "^(?:|a)*b$": text = "aa": expectedLength = -1
        Case 7: pattern = "^(?:[ab]*?)*b": text = "abb": expectedLength = 2
        Case 8: pattern = "(?:\b|.)*": text = "aaa": expectedLength = 0
        Case 9: pattern = "(?:\b|.)*b": text = "bb": expectedLength = 1
        End Select
        RegexCompiler.Compile bytecode, pattern
        Assert.AreEqual expectedLength, RegexNfaMatcher.NfaMatch(captures, bytecode, text), pattern
        If expectedLength <> -1 Then
            Assert.AreEqual 1&, captures.entireMatch.start
            Assert.AreEqual expectedLength, captures.entireMatch.Length, pattern
        End If
    Next
End Sub

'@TestMethod("NfaMatcher")
Private Sub NfaMatcher_NullableLoopCapturesAgreeWithRe2()
    Dim bytecode() As Long, captures As RegexMatcherCommon.CapturesTy
    Dim pattern As String, i As Long
    For i = 0 To 1
        If i = 0 Then pattern = "(a*)*" Else pattern = "(a*)+"
        RegexCompiler.Compile bytecode, pattern
        Assert.AreEqual 1&, RegexNfaMatcher.NfaMatch(captures, bytecode, "a"), pattern
        Assert.AreEqual 1&, captures.numberedCaptures(0).start, pattern
        Assert.AreEqual 1&, captures.numberedCaptures(0).Length, pattern
    Next
    RegexCompiler.Compile bytecode, "((a*?)*)b"
    Assert.AreEqual 3&, RegexNfaMatcher.NfaMatch(captures, bytecode, "aab")
    Assert.AreEqual 1&, captures.numberedCaptures(1).start
    Assert.AreEqual 2&, captures.numberedCaptures(1).Length
    RegexCompiler.Compile bytecode, "(?<word>a*)*"
    Assert.AreEqual 1&, RegexNfaMatcher.NfaMatch(captures, bytecode, "a")
    Assert.AreEqual 1&, captures.namedCaptures(0)
    Assert.AreEqual 1&, captures.numberedCaptures(0).start
    Assert.AreEqual 1&, captures.numberedCaptures(0).Length
    ' Failed empty alternatives must not lend their capture frames to 'a'.
    RegexCompiler.Compile bytecode, "^(?:()|()|a)*b"
    Assert.AreEqual 3&, RegexNfaMatcher.NfaMatch(captures, bytecode, "aab")
    Assert.AreEqual 0&, captures.numberedCaptures(0).start
    Assert.AreEqual 0&, captures.numberedCaptures(1).start
End Sub


'@TestMethod("NfaMatcher")
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


'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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


'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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


'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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

'@TestMethod("NfaMatcher")
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
