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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
Private Sub DfsMatcher_MatchRegexp013()
    ' empty star
    On Error GoTo TestFail

    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 2, _
        REOP_JUMP, 0, _
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

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_MatchRegexp014()
    ' star, complex pattern, matches zero times
    On Error GoTo TestFail
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(maxProperCaptureSlot:=2), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(maxProperCaptureSlot:=2), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
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
Private Sub DfsMatcher_MatchRegexp030()
    ' Question mark, possessive
    'On Error GoTo TestFail
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT1 Or REOP_FLAG_POSSESSIVE, 3, _
        REOP_CHAR, AscW("b"), _
        REOP_COMMIT_POSSESSIVE, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
        
    Dim s As String: s = "abcx"
    
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
Private Sub DfsMatcher_MatchRegexp031()
    ' Question mark, possessive
    'On Error GoTo TestFail
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT1 Or REOP_FLAG_POSSESSIVE, 3, _
        REOP_CHAR, AscW("b"), _
        REOP_COMMIT_POSSESSIVE, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
        
    Dim s As String: s = "abcx"
    
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
Private Sub DfsMatcher_MatchRegexp032()
    ' Question mark, possessive
    'On Error GoTo TestFail
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT1 Or REOP_FLAG_POSSESSIVE, 3, _
        REOP_CHAR, AscW("b"), _
        REOP_COMMIT_POSSESSIVE, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
        
    Dim s As String: s = "abbcx"
    
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
Private Sub DfsMatcher_MatchRegexp033()
    ' Question mark, possessive, inside REPEAT_MAX_GREEDY
    'On Error GoTo TestFail
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_MAX_GREEDY_INIT, _
        REOP_REPEAT_MAX_GREEDY_START, 3, 10, _
        REOP_SPLIT1 Or REOP_FLAG_POSSESSIVE, 3, _
        REOP_CHAR, AscW("b"), _
        REOP_COMMIT_POSSESSIVE, _
        REOP_CHAR, AscW("c"), _
        REOP_REPEAT_MAX_GREEDY_END, 3, 10, _
        REOP_CHAR, AscW("d"), _
        REOP_SAVE, 1, _
        REOP_MATCH
        
    Dim s As String: s = "acbcbcd"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual 7&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub DfsMatcher_MatchRegexp034()
    ' REPEAT_MAX_GREEDY, possessive, inside REPEAT_MAX_GREEDY
    'On Error GoTo TestFail
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_MAX_GREEDY_INIT, _
        REOP_REPEAT_MAX_GREEDY_START, 3, 15, _
        REOP_REPEAT_MAX_GREEDY_INIT, _
        REOP_REPEAT_MAX_GREEDY_START Or REOP_FLAG_POSSESSIVE, 4, 6, _
        REOP_CHAR, AscW("b"), _
        REOP_COMMIT_POSSESSIVE, _
        REOP_REPEAT_MAX_GREEDY_END, 4, 6, _
        REOP_CHAR, AscW("c"), _
        REOP_REPEAT_MAX_GREEDY_END, 3, 15, _
        REOP_CHAR, AscW("d"), _
        REOP_SAVE, 1, _
        REOP_MATCH
        
    Dim s As String: s = "acbcbbcd"
    
    Dim result As Long
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual 7&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub SetupReCtxDisjunction(ByRef bytecode() As Long)
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(maxProperCaptureSlot:=5), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(maxProperCaptureSlot:=21), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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
    ' REPEAT_MAX_GREEDY
    On Error GoTo TestFail
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_MAX_GREEDY_INIT, _
        REOP_REPEAT_MAX_GREEDY_START, 5, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_MAX_GREEDY_END, 5, 5, _
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
    ' REPEAT_MAX_GREEDY
    On Error GoTo TestFail
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_MAX_GREEDY_INIT, _
        REOP_REPEAT_MAX_GREEDY_START, 5, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_MAX_GREEDY_END, 5, 5, _
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
    ' REPEAT_MAX_GREEDY
    On Error GoTo TestFail
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_MAX_GREEDY_INIT, _
        REOP_REPEAT_MAX_GREEDY_START, 5, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_MAX_GREEDY_END, 5, 5, _
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
    ' REPEAT_MAX_GREEDY
    'On Error GoTo TestFail
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_MAX_GREEDY_INIT, _
        REOP_REPEAT_MAX_GREEDY_START, 5, 16, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_MAX_GREEDY_INIT, _
        REOP_REPEAT_MAX_GREEDY_START, 2, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_REPEAT_MAX_GREEDY_END, 2, 5, _
        REOP_CHAR, AscW("d"), _
        REOP_REPEAT_MAX_GREEDY_END, 5, 16, _
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
    
    Dim captures As RegexMatcherCommon.CapturesTy, bytecode() As Long
    MakeBytecode bytecode, MakeBytecodeParameters(), _
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

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_LookbehindWordBoundary_AtStart()
    On Error GoTo TestFail

    Dim captures As RegexMatcherCommon.CapturesTy
    Dim bytecode() As Long
    Dim s As String
    Dim result As Long

    ' (?<=\b)x
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHECK_LOOKBEHIND, _
        REOP_LOOKPOS, 2, _
        REOP_ASSERT_WORD_BOUNDARY, _
        REOP_END_LOOKPOS, _
        REOP_CHAR, AscW("x"), _
        REOP_SAVE, 1, _
        REOP_MATCH

    s = "xyz"
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)

    Assert.AreEqual 1&, result
    Assert.AreEqual 1&, captures.entireMatch.start
    Assert.AreEqual 1&, captures.entireMatch.Length

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

' Issue: https://github.com/sihlfall/vba-regex/issues/16
'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_LookbehindWordBoundary_BetweenWordCharacters()
    On Error GoTo TestFail

    Dim captures As RegexMatcherCommon.CapturesTy
    Dim bytecode() As Long
    Dim s As String
    Dim result As Long

    ' (?<=\b)b
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHECK_LOOKBEHIND, _
        REOP_LOOKPOS, 2, _
        REOP_ASSERT_WORD_BOUNDARY, _
        REOP_END_LOOKPOS, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH

    s = "ab"
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)

    Assert.AreEqual -1&, result

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

' Issue: https://github.com/sihlfall/vba-regex/issues/15
'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_NegativeLookbehindStart_UsesBoundaryPosition()
    On Error GoTo TestFail

    Dim captures As RegexMatcherCommon.CapturesTy
    Dim bytecode() As Long
    Dim s As String
    Dim result As Long

    ' (?<!^)b
    MakeBytecode bytecode, MakeBytecodeParameters(), _
        REOP_SAVE, 0, _
        REOP_CHECK_LOOKBEHIND, _
        REOP_LOOKNEG, 2, _
        REOP_ASSERT_START, _
        REOP_END_LOOKNEG, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH

    ' Boundary 1 is not the start of the string.
    s = "abc"
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)

    Assert.AreEqual 2&, result
    Assert.AreEqual 2&, captures.entireMatch.start
    Assert.AreEqual 1&, captures.entireMatch.Length

    ' Boundary 0 is the start, so the negative lookbehind must fail.
    s = "b"
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)

    Assert.AreEqual -1&, result

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub




'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_LiteralPrefixSearch()
    Dim bytecode() As Long, captures As RegexMatcherCommon.CapturesTy
    Dim context As RegexDfsMatcher.DfsMatcherContext, haystack As String
    RegexCompiler.Compile bytecode, "(ab)+c"
    haystack = "xxaaababc"
    Assert.AreEqual 9&, RegexDfsMatcher.DfsMatch(captures, bytecode, haystack)
    Assert.AreEqual 5&, captures.entireMatch.start
    Assert.AreEqual 5&, captures.entireMatch.Length
    Assert.AreEqual 7&, captures.numberedCaptures(0).start
    Assert.AreEqual 2&, captures.numberedCaptures(0).Length
    Assert.AreEqual -1&, RegexDfsMatcher.DfsMatchFrom(context, captures, bytecode, haystack, 7)

    RegexCompiler.Compile bytecode, "ab(?=c)"
    Assert.AreEqual 5&, RegexDfsMatcher.DfsMatch(captures, bytecode, "abxabc")
    Assert.AreEqual 4&, captures.entireMatch.start
    Assert.AreEqual 2&, captures.entireMatch.Length
    RegexCompiler.Compile bytecode, "\u0000xy"
    Assert.AreEqual 4&, RegexDfsMatcher.DfsMatch(captures, bytecode, "z" & ChrW(0) & "xy")
    Assert.AreEqual 2&, captures.entireMatch.start
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_StringDirectionsAndCaseFolding()
    Dim bytecode() As Long, captures As RegexMatcherCommon.CapturesTy
    RegexCompiler.Compile bytecode, "(?<=abc)d"
    Assert.AreEqual 6&, RegexDfsMatcher.DfsMatch(captures, bytecode, "xxabcd")
    Assert.AreEqual 6&, captures.entireMatch.start
    RegexCompiler.Compile bytecode, "(?i:abc)"
    Assert.AreEqual 4&, RegexDfsMatcher.DfsMatch(captures, bytecode, "xAbC")
    Assert.AreEqual 2&, captures.entireMatch.start
    Assert.AreEqual -1&, RegexDfsMatcher.DfsMatch(captures, bytecode, "xab")
    RegexCompiler.Compile bytecode, "(?i:(?<=abc)d)"
    Assert.AreEqual 6&, RegexDfsMatcher.DfsMatch(captures, bytecode, "xxAbCd")
    RegexCompiler.Compile bytecode, "(abc)\1"
    Assert.AreEqual 7&, RegexDfsMatcher.DfsMatch(captures, bytecode, "xabcabc")
    Assert.AreEqual 2&, captures.entireMatch.start
    Assert.AreEqual 6&, captures.entireMatch.Length
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_LiteralSearchBoundsAndBudget()
    Dim bytecode() As Long, captures As RegexMatcherCommon.CapturesTy
    Dim context As RegexDfsMatcher.DfsMatcherContext, haystack As String
    RegexCompiler.Compile bytecode, "needle"
    haystack = String$(1024, "x") & "needle"
    Assert.AreEqual Len(haystack), RegexDfsMatcher.DfsMatch(captures, bytecode, haystack)
    Assert.AreEqual 1025&, captures.entireMatch.start
    Assert.AreEqual -1&, RegexDfsMatcher.DfsMatch(captures, bytecode, haystack, stepsLimit:=20)
    Assert.AreEqual -1&, RegexDfsMatcher.DfsMatch(captures, bytecode, "needle", stepsLimit:=1)
    Assert.AreEqual -1&, RegexDfsMatcher.DfsMatchFrom(context, captures, bytecode, "needle", -1)
    Assert.AreEqual -1&, RegexDfsMatcher.DfsMatchFrom(context, captures, bytecode, "needle", 7)
    ' A custom entry point has no unanchored-search prefix to optimize.
    bytecode(BYTECODE_IDX_START_WITH_PREFIX) = bytecode(BYTECODE_IDX_START_ACTUAL)
    Assert.AreEqual -1&, RegexDfsMatcher.DfsMatch(captures, bytecode, "xneedle")
    Assert.AreEqual 6&, RegexDfsMatcher.DfsMatch(captures, bytecode, "needle")
End Sub

Private Sub AssertAnchoredSearch(ByVal pattern As String, ByVal text As String, ByVal start As Long, ByVal length As Long, Optional ByVal multiline As Boolean = False)
    Dim bytecode() As Long, captures As RegexMatcherCommon.CapturesTy
    Dim result As Long, expected As Long
    RegexCompiler.Compile bytecode, pattern
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, text, multiline:=multiline)
    If start = 0 Then expected = -1 Else expected = start + length - 1
    Assert.AreEqual expected, result, pattern
    If start > 0 Then
        Assert.AreEqual start, captures.entireMatch.start, pattern
        Assert.AreEqual length, captures.entireMatch.Length, pattern
    End If
    ' An explicit budget disables search shortcuts, giving a VM reference run.
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, text, stepsLimit:=DEFAULT_STEPS_LIMIT - 1, multiline:=multiline)
    Assert.AreEqual expected, result, pattern
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_LeadingAnchorSearch()
    Dim terminator As Variant, bytecode() As Long, captures As RegexMatcherCommon.CapturesTy
    Dim context As RegexDfsMatcher.DfsMatcherContext, text As String
    Dim terminators(0 To 3) As String, i As Long
    AssertAnchoredSearch "^needle$", String$(16384, "x"), 0, 0
    AssertAnchoredSearch "^needle$", "needle", 1, 6
    AssertAnchoredSearch "^a|b", "xxb", 3, 1
    AssertAnchoredSearch "(^a)?b", "xxb", 3, 1
    AssertAnchoredSearch "\^needle", "x^needle", 2, 7
    AssertAnchoredSearch "(^needle)$", "x" & vbLf & "needle", 3, 6, True
    AssertAnchoredSearch "(?m:^needle$)", "x" & vbLf & "needle", 3, 6
    AssertAnchoredSearch "(?-m:^needle$)", "x" & vbLf & "needle", 0, 0, True
    AssertAnchoredSearch "(?m:^$)", "x" & vbCrLf & "y", 3, 0
    AssertAnchoredSearch "(?m:^$)", "x" & vbLf, 3, 0
    terminators(0) = vbCr: terminators(1) = vbLf
    terminators(2) = ChrW(&H2028): terminators(3) = ChrW(&H2029)
    For Each terminator In terminators
        AssertAnchoredSearch "^needle$", "x" & terminator & "needle", 3, 6, True
    Next
    text = "x" & vbCr & "x" & vbLf & "x" & ChrW(&H2028) & "x" & ChrW(&H2029) & "needle"
    AssertAnchoredSearch "^needle$", text, 9, 6, True
    text = vbNullString
    For i = 1 To 4096: text = text & "x" & vbLf: Next
    text = text & "needle"
    AssertAnchoredSearch "^needle$", text, 8193, 6, True
    AssertAnchoredSearch "^a[\s\S]*z", "a" & vbLf & "x" & vbLf & "z", 1, 5, True
    RegexCompiler.Compile bytecode, "^[ab]+$"
    Assert.AreEqual -1&, RegexDfsMatcher.DfsMatchFrom(context, captures, bytecode, "ab", 1)
    Assert.AreEqual 3&, RegexDfsMatcher.DfsMatchFrom(context, captures, bytecode, vbLf & "ab", 1, multiline:=True)
    Assert.AreEqual -1&, RegexDfsMatcher.DfsMatch(captures, bytecode, "ab", stepsLimit:=1)
    bytecode(BYTECODE_IDX_START_WITH_PREFIX) = bytecode(BYTECODE_IDX_START_ACTUAL)
    Assert.AreEqual -1&, RegexDfsMatcher.DfsMatch(captures, bytecode, "x" & vbLf & "ab", multiline:=True)
End Sub

'@TestMethod("DfsMatcher")
Private Sub DfsMatcher_BackreferenceLiteralSpans()
    Dim bytecode() As Long, captures As RegexMatcherCommon.CapturesTy, text As String
    Dim literal As String
    literal = ChrW(0) & ChrW(&H8000&) & ChrW(&HD800&) & "a" & ChrW(&HDC00&)
    text = literal & literal
    RegexCompiler.Compile bytecode, "(.+)\1"
    Assert.AreEqual Len(text), RegexDfsMatcher.DfsMatch(captures, bytecode, text)
    Assert.AreEqual Len(literal), captures.numberedCaptures(0).Length
    AssertAnchoredSearch "^(a)?\1b$", "b", 1, 1
    AssertAnchoredSearch "^()\1b$", "b", 1, 1
    AssertAnchoredSearch "^(abc)\1$", "abcab", 0, 0
    AssertAnchoredSearch "^(abc)\1$", "abcabd", 0, 0
    AssertAnchoredSearch "(?<=\1(abc))d", "abcabcd", 7, 1
    AssertAnchoredSearch "(?<=\1(abc))d", "abxabcd", 0, 0
    AssertAnchoredSearch "(?<=\1(abc))d", "abcd", 0, 0
    AssertAnchoredSearch "(abc)(?<=\1)d", "abcd", 1, 4
    RegexCompiler.Compile bytecode, "(?<=(abc))d"
    Assert.AreEqual 4&, RegexDfsMatcher.DfsMatch(captures, bytecode, "abcd")
    Assert.AreEqual 1&, captures.numberedCaptures(0).start
    Assert.AreEqual 3&, captures.numberedCaptures(0).Length
    AssertAnchoredSearch "(?i:^(abc)\1$)", "abcaBc", 1, 6
    AssertAnchoredSearch "(?i:^(\u00E4)\1$)", ChrW(&HE4&) & ChrW(&HC4&), 1, 2
    AssertAnchoredSearch "^(a\u0000b)\1$", "a" & ChrW(0) & "ba" & ChrW(0) & "c", 0, 0
    AssertAnchoredSearch "^(\u00E4)\1$", ChrW(&HE4&) & ChrW(&HC4&), 0, 0
    AssertAnchoredSearch "(?i:(?<=\1(abc))d)", "aBcaBcd", 7, 1
    ' Positions exceed the maximum window and remain far apart while comparing.
    text = String$(40000, "a")
    RegexCompiler.Compile bytecode, "^(a{40000})\1$"
    Assert.AreEqual 80000&, RegexDfsMatcher.DfsMatch(captures, bytecode, text & text)
    RegexCompiler.Compile bytecode, "(?i:^(a{40000})\1$)"
    Assert.AreEqual 80000&, RegexDfsMatcher.DfsMatch(captures, bytecode, text & String$(40000, "A"))
End Sub
