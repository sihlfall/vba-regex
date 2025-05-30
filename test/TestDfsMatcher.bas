Attribute VB_Name = "TestDfsMatcher"
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


'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp0001()
    ' Matching simple two-character pattern
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "ab"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp0002()
    ' Matching simple two-character pattern
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "bbbaab"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual 0&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp010()
    ' star, matches several times
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Dim s As String: s = "abbbc"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp011()
    ' star, matches once
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Dim s As String: s = "abc"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp012()
    ' star, matches zero times
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Dim s As String: s = "ac"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
'Private Sub DfsMatcher_MatchRegexp013()
'    ' empty star with non-capture group
'    ' with capture group, an infinite loop will be produced
'    On Error GoTo TestFail
'
'    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
'    MakeArray bytecode, _
'        REOP_SAVE, 0, _
'        REOP_SQGREEDY, 0, 1, 0, 1, _
'        REOP_MATCH, _
'        REOP_CHAR, AscW("x"), _
'        REOP_SAVE, 1, _
'        REOP_MATCH
'    reCtx.nSaved = 2
'    dim s as String: s = "x"
'
'    Dim result As Long
'    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
'    Assert.AreEqual Len(s), result
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp014()
    ' star, complex pattern, matches zero times
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        2, 0, 0, _
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
    Dim s As String: s = "x"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub DfsMatcher_MatchRegexp015()
    ' star, complex pattern, matches once
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        2, 0, _
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
    Dim s As String: s = "abx"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp016()
    ' star, complex pattern, matches twice
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        3, 0, 0, _
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
    Dim s As String: s = "ababx"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp017()
    ' star, non-greedy
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT2, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Dim s As String: s = "abbc"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual 4&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp020()
    ' Plus
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Dim s As String: s = "abbbc"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp021()
    ' Plus, complex pattern, matches once
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        3, 0, 0, _
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
    Dim s As String: s = "abx"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp022()
    ' Plus, complex pattern, matches twice
    'On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        3, 0, 0, _
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
    Dim s As String: s = "ababx"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub



Private Sub SetupReCtxDisjunction(ByRef bytecode() As Long)
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("a"), _
        REOP_JUMP, 2, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp050()
    ' Disjunction, match first alternative (single character)
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    SetupReCtxDisjunction bytecode
    Dim s As String: s = "a"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp051()
    ' Disjunction, match second alternative (single character)
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    SetupReCtxDisjunction bytecode
    Dim s As String: s = "b"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp101()
    ' Parentheses
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        3, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SAVE, 2, _
        REOP_CHAR, AscW("b"), _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 3, _
        REOP_CHAR, AscW("d"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "abcd"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp102()
    ' Nested parentheses
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        5, 0, 0, _
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
    Dim s As String: s = "abcdef"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp200()
    ' Caret, matching
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_ASSERT_START, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "abc"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual 2&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp201()
    ' Caret, not matching
    'On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_ASSERT_START, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "zabc"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp202()
    ' Dollar, matching
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_ASSERT_END, _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "ab"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual 2&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp203()
    ' Dollar, not matching
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_ASSERT_END, _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "abz"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp204()
    ' Period, matching
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_DOT, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "azb"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp210()
    ' Word boundary, matching
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW(" "), _
        REOP_ASSERT_WORD_BOUNDARY, _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = " a"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual 1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp211()
    ' Word boundary, not matching
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW(" "), _
        REOP_ASSERT_WORD_BOUNDARY, _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "  "
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp212()
    ' Not word boundary, matching
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW(" "), _
        REOP_ASSERT_NOT_WORD_BOUNDARY, _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "  "
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual 1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp213()
    ' Not word boundary, not matching
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW(" "), _
        REOP_ASSERT_NOT_WORD_BOUNDARY, _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = " a"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp300()
    ' Range
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "ABC"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp301()
    ' Range, failing match
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "AbC"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp302()
    ' Range
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_RANGES, 2, AscW("A"), AscW("Z"), AscW("a"), AscW("z"), _
        REOP_RANGES, 2, AscW("A"), AscW("Z"), AscW("a"), AscW("z"), _
        REOP_RANGES, 2, AscW("A"), AscW("Z"), AscW("a"), AscW("z"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "AbC"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp303()
    ' closing square bracket as character
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 3, _
        REOP_DOT, _
        REOP_JUMP, 6, _
        REOP_RANGES, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("]"), _
        REOP_ASSERT_END, _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Dim s As String: s = "babba"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp304()
    ' closing square bracket as first character of range
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 8, _
        REOP_RANGES, 2, AscW("]"), AscW("]"), AscW("a"), AscW("a"), _
        REOP_JUMP, -10, _
        REOP_ASSERT_END, _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Dim s As String: s = "babb"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual 4&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp305()
    ' closing square bracket as first character of range
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 8, _
        REOP_RANGES, 2, AscW("]"), AscW("]"), AscW("a"), AscW("a"), _
        REOP_JUMP, -10, _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Dim s As String: s = "a]ab"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual 3&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp400()
    ' Positive lookahead
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_CHECK_LOOKAHEAD, _
        REOP_LOOKPOS, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_CHAR, AscW("d"), _
        REOP_END_LOOKPOS, _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "abcde"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual 2&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp401()
    ' Positive lookahead, failing match
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_CHECK_LOOKAHEAD, _
        REOP_LOOKPOS, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_CHAR, AscW("d"), _
        REOP_END_LOOKPOS, _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "abef"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp402()
    ' Negative lookahead
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_CHECK_LOOKAHEAD, _
        REOP_LOOKNEG, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_CHAR, AscW("d"), _
        REOP_END_LOOKNEG, _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "abce"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual 2&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp403()
    ' Negative lookahead, failing match
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_CHECK_LOOKAHEAD, _
        REOP_LOOKNEG, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_CHAR, AscW("d"), _
        REOP_END_LOOKNEG, _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "abcd"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp404()
    ' Positive lookbehind
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_DOT, _
        REOP_DOT, _
        REOP_DOT, _
        REOP_DOT, _
        REOP_CHECK_LOOKBEHIND, _
        REOP_LOOKPOS, 5, _
        REOP_CHAR, AscW("d"), _
        REOP_CHAR, AscW("c"), _
        REOP_END_LOOKPOS, _
        REOP_CHAR, AscW("x"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "abcdx"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual 5&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp405()
    ' Negative lookbehind
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_DOT, _
        REOP_DOT, _
        REOP_DOT, _
        REOP_DOT, _
        REOP_CHECK_LOOKBEHIND, _
        REOP_LOOKNEG, 5, _
        REOP_CHAR, AscW("d"), _
        REOP_CHAR, AscW("y"), _
        REOP_END_LOOKNEG, _
        REOP_CHAR, AscW("x"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "abcdx"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual 5&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp406()
    ' Lookahead inside lookbehind
    ' On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_DOT, _
        REOP_DOT, _
        REOP_DOT, _
        REOP_DOT, _
        REOP_CHECK_LOOKBEHIND, _
        REOP_LOOKPOS, 11, _
        REOP_CHAR, AscW("d"), _
        REOP_CHECK_LOOKAHEAD, _
        REOP_LOOKPOS, 3, _
        REOP_CHAR, AscW("d"), _
        REOP_END_LOOKPOS, _
        REOP_CHAR, AscW("c"), _
        REOP_END_LOOKPOS, _
        REOP_CHAR, AscW("x"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "abcdx"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual 5&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp407()
    ' Lookbehind inside lookbehind
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_DOT, _
        REOP_DOT, _
        REOP_DOT, _
        REOP_DOT, _
        REOP_CHECK_LOOKBEHIND, _
        REOP_LOOKPOS, 11, _
        REOP_CHAR, AscW("d"), _
        REOP_CHECK_LOOKBEHIND, _
        REOP_LOOKPOS, 3, _
        REOP_CHAR, AscW("c"), _
        REOP_END_LOOKPOS, _
        REOP_CHAR, AscW("c"), _
        REOP_END_LOOKPOS, _
        REOP_CHAR, AscW("x"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "abcdx"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual 5&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp408()
    ' Positive lookahead with star inside, matching pattern
    'On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_CHECK_LOOKAHEAD, _
        REOP_LOOKPOS, 9, _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("c"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("d"), _
        REOP_END_LOOKPOS, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "abccd"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual 3&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp409()
    ' Negative lookahead with star inside, matching pattern
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_CHECK_LOOKAHEAD, _
        REOP_LOOKNEG, 9, _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("c"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("e"), _
        REOP_END_LOOKNEG, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "abccd"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual 3&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp410()
    ' Positive lookahead with star inside, failing pattern
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_CHECK_LOOKAHEAD, _
        REOP_LOOKPOS, 9, _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("c"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("e"), _
        REOP_END_LOOKPOS, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "abccd"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp411()
    ' Negative lookahead with star inside, failing pattern
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_CHECK_LOOKAHEAD, _
        REOP_LOOKNEG, 9, _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("c"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("d"), _
        REOP_END_LOOKNEG, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "abccd"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub



'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp450()
    ' Backreference
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        5, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SAVE, 2, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 3, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 4, _
        REOP_CHAR, AscW("d"), _
        REOP_SAVE, 5, _
        REOP_CHAR, AscW("e"), _
        REOP_BACKREFERENCE, 1, _
        REOP_BACKREFERENCE, 2, _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "abcdebd"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp451()
    ' Backreference
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        21, 0, 0, _
        REOP_SAVE, 0, _
        REOP_SAVE, 2, REOP_SAVE, 3, _
        REOP_SAVE, 4, REOP_SAVE, 5, _
        REOP_SAVE, 6, REOP_SAVE, 7, _
        REOP_SAVE, 8, REOP_SAVE, 9, _
        REOP_SAVE, 10, REOP_SAVE, 11, _
        REOP_SAVE, 12, REOP_SAVE, 13, _
        REOP_SAVE, 14, REOP_SAVE, 15, _
        REOP_SAVE, 16, REOP_SAVE, 17, _
        REOP_SAVE, 18, REOP_SAVE, 19, _
        REOP_SAVE, 20, REOP_CHAR, AscW("a"), REOP_SAVE, 21, _
        REOP_BACKREFERENCE, 10, _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "aa"
    
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp500()
    ' quantifier, exact number
    'On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_EXACTLY_INIT, _
        REOP_REPEAT_EXACTLY_START, 5, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_EXACTLY_END, 5, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Dim s As String: s = "abbbbbc"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp501()
    ' quantifier, exact number
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_EXACTLY_INIT, _
        REOP_REPEAT_EXACTLY_START, 5, 16, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_EXACTLY_INIT, _
        REOP_REPEAT_EXACTLY_START, 2, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_REPEAT_EXACTLY_END, 2, 5, _
        REOP_CHAR, AscW("d"), _
        REOP_REPEAT_EXACTLY_END, 5, 16, _
        REOP_CHAR, AscW("e"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Dim s As String: s = "abccdbccdbccdbccdbccde"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp510()
    ' quantifier, exact number
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_MAX_HUMBLE_INIT, _
        REOP_REPEAT_MAX_HUMBLE_START, 5, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_MAX_HUMBLE_END, 5, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Dim s As String: s = "abbbbbc"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp511()
    ' quantifier, exact number
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_MAX_HUMBLE_INIT, _
        REOP_REPEAT_MAX_HUMBLE_START, 5, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_MAX_HUMBLE_END, 5, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Dim s As String: s = "abbbbbbc"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp512()
    ' quantifier, exact number
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_MAX_HUMBLE_INIT, _
        REOP_REPEAT_MAX_HUMBLE_START, 5, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_MAX_HUMBLE_END, 5, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Dim s As String: s = "ac"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp520()
    ' quantifier, exact number
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_MAX_HUMBLE_INIT, _
        REOP_REPEAT_MAX_HUMBLE_START, 5, 16, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_MAX_HUMBLE_INIT, _
        REOP_REPEAT_MAX_HUMBLE_START, 2, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_REPEAT_MAX_HUMBLE_END, 2, 5, _
        REOP_CHAR, AscW("d"), _
        REOP_REPEAT_MAX_HUMBLE_END, 5, 16, _
        REOP_CHAR, AscW("e"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Dim s As String: s = "abccdbccdbccdbccdbccde"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp530()
    ' REPEAT_GREEDY_MAX
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_GREEDY_MAX_INIT, _
        REOP_REPEAT_GREEDY_MAX_START, 5, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_GREEDY_MAX_END, 5, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Dim s As String: s = "abbbbbc"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp531()
    ' REPEAT_GREEDY_MAX
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_GREEDY_MAX_INIT, _
        REOP_REPEAT_GREEDY_MAX_START, 5, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_GREEDY_MAX_END, 5, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Dim s As String: s = "abbbbbbc"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp532()
    ' REPEAT_GREEDY_MAX
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_GREEDY_MAX_INIT, _
        REOP_REPEAT_GREEDY_MAX_START, 5, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_GREEDY_MAX_END, 5, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Dim s As String: s = "ac"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp540()
    ' REPEAT_GREEDY_MAX
    'On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_GREEDY_MAX_INIT, _
        REOP_REPEAT_GREEDY_MAX_START, 5, 16, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_GREEDY_MAX_INIT, _
        REOP_REPEAT_GREEDY_MAX_START, 2, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_REPEAT_GREEDY_MAX_END, 2, 5, _
        REOP_CHAR, AscW("d"), _
        REOP_REPEAT_GREEDY_MAX_END, 5, 16, _
        REOP_CHAR, AscW("e"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Dim s As String: s = "abccdbccdbccdbccdbccde"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp600()
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    MakeArray bytecode, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_ASSERT_START, _
        REOP_SPLIT1, 4, _
        REOP_CHAR, 65, _
        REOP_JUMP, 4, _
        REOP_CHAR, 65, _
        REOP_CHAR, 65, _
        REOP_ASSERT_END, _
        REOP_SAVE, 1, _
        REOP_MATCH
    Dim s As String: s = "A"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
    
    Assert.AreEqual 1&, captures.entireMatch.start
    Assert.AreEqual 1&, captures.entireMatch.Length
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub




