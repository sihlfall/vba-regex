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


'@TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp0001()
    ' Matching simple two-character pattern
    'On Error GoTo TestFail
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = "ab"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.stepsLimit = 10
    reCtx.input = "abbbc"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.stepsLimit = 10
    reCtx.input = "abc"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.stepsLimit = 10
    reCtx.input = "ac"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp013()
    ' empty star with capture group
    On Error GoTo TestFail

    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        3, _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 6, _
        REOP_SAVE, 2, _
        REOP_SAVE, 3, _
        REOP_JUMP, -8, _
        REOP_CHAR, AscW("x"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = "x"

    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        3, _
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
    reCtx.input = "x"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        3, _
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
    reCtx.input = "abx"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        3, _
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
    reCtx.input = "ababx"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT2, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.stepsLimit = 10
    reCtx.input = "abbc"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.stepsLimit = 10
    reCtx.input = "abbbc"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        3, _
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
    reCtx.input = "abx"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        3, _
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
    reCtx.input = "ababx"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub



Private Sub SetupReCtxDisjunction(ByRef reCtx As RegexNfaMatcher.NfaMatcher)
    MakeArray reCtx.bytecode, _
        1, _
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    SetupReCtxDisjunction reCtx
    reCtx.input = "a"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    SetupReCtxDisjunction reCtx
    reCtx.input = "b"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        3, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SAVE, 2, _
        REOP_CHAR, AscW("b"), _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 3, _
        REOP_CHAR, AscW("d"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = "abcd"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        5, _
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
    reCtx.input = "abcdef"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_ASSERT_START, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = "abc"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_ASSERT_START, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = "zabc"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_ASSERT_END, _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = "ab"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_ASSERT_END, _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = "abz"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_PERIOD, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = "azb"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW(" "), _
        REOP_ASSERT_WORD_BOUNDARY, _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = " a"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW(" "), _
        REOP_ASSERT_WORD_BOUNDARY, _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = "  "
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW(" "), _
        REOP_ASSERT_NOT_WORD_BOUNDARY, _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = "  "
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW(" "), _
        REOP_ASSERT_NOT_WORD_BOUNDARY, _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = " a"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = "ABC"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = "AbC"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
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
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_RANGES, 2, AscW("A"), AscW("Z"), AscW("a"), AscW("z"), _
        REOP_RANGES, 2, AscW("A"), AscW("Z"), AscW("a"), AscW("z"), _
        REOP_RANGES, 2, AscW("A"), AscW("Z"), AscW("a"), AscW("z"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = "AbC"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp400()
    ' Positive lookahead
    On Error GoTo TestFail
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_CHECK_LOOKAHEAD, _
        REOP_LOOKPOS, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_CHAR, AscW("d"), _
        REOP_MATCH, _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = "abcde"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual 2&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp401()
    ' Positive lookahead, failing match
    On Error GoTo TestFail
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_CHECK_LOOKAHEAD, _
        REOP_LOOKPOS, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_CHAR, AscW("d"), _
        REOP_MATCH, _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = "abef"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp402()
    ' Negative lookahead
    On Error GoTo TestFail
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_CHECK_LOOKAHEAD, _
        REOP_LOOKNEG, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_CHAR, AscW("d"), _
        REOP_MATCH, _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = "abce"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual 2&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp403()
    ' Negative lookahead, failing match
    On Error GoTo TestFail
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_CHECK_LOOKAHEAD, _
        REOP_LOOKNEG, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_CHAR, AscW("d"), _
        REOP_MATCH, _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.input = "abcd"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp450()
    ' Backreference
    On Error GoTo TestFail
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        5, _
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
    reCtx.input = "abcdebd"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp451()
    ' Backreference
    On Error GoTo TestFail
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        21, _
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
    reCtx.input = "aa"
    reCtx.stepsLimit = 100
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp500()
    ' quantifier, exact number
    On Error GoTo TestFail
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_EXACTLY_INIT, _
        REOP_REPEAT_EXACTLY_START, 5, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_EXACTLY_END, 5, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.stepsLimit = 10
    reCtx.input = "abbbbbc"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp501()
    ' quantifier, exact number
    On Error GoTo TestFail
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
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
    reCtx.stepsLimit = 10
    reCtx.input = "abccdbccdbccdbccdbccde"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp510()
    ' quantifier, exact number
    On Error GoTo TestFail
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_MAX_HUMBLE_INIT, _
        REOP_REPEAT_MAX_HUMBLE_START, 5, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_MAX_HUMBLE_END, 5, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.stepsLimit = 10
    reCtx.input = "abbbbbc"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp511()
    ' quantifier, exact number
    On Error GoTo TestFail
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_MAX_HUMBLE_INIT, _
        REOP_REPEAT_MAX_HUMBLE_START, 5, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_MAX_HUMBLE_END, 5, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.stepsLimit = 10
    reCtx.input = "abbbbbbc"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp512()
    ' quantifier, exact number
    On Error GoTo TestFail
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_MAX_HUMBLE_INIT, _
        REOP_REPEAT_MAX_HUMBLE_START, 5, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_MAX_HUMBLE_END, 5, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.stepsLimit = 10
    reCtx.input = "ac"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp520()
    ' quantifier, exact number
    On Error GoTo TestFail
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
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
    reCtx.stepsLimit = 10
    reCtx.input = "abccdbccdbccdbccdbccde"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp530()
    ' REPEAT_GREEDY_MAX
    On Error GoTo TestFail
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_GREEDY_MAX_INIT, _
        REOP_REPEAT_GREEDY_MAX_START, 5, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_GREEDY_MAX_END, 5, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.stepsLimit = 10
    reCtx.input = "abbbbbc"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp531()
    ' REPEAT_GREEDY_MAX
    On Error GoTo TestFail
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_GREEDY_MAX_INIT, _
        REOP_REPEAT_GREEDY_MAX_START, 5, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_GREEDY_MAX_END, 5, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.stepsLimit = 10
    reCtx.input = "abbbbbbc"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp532()
    ' REPEAT_GREEDY_MAX
    On Error GoTo TestFail
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_GREEDY_MAX_INIT, _
        REOP_REPEAT_GREEDY_MAX_START, 5, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_GREEDY_MAX_END, 5, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    reCtx.stepsLimit = 10
    reCtx.input = "ac"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'DISABLED_TestMethod("NfaMatcher")
Private Sub NfaMatcher_MatchRegexp540()
    ' REPEAT_GREEDY_MAX
    On Error GoTo TestFail
    
    Dim reCtx As RegexNfaMatcher.NfaMatcher
    MakeArray reCtx.bytecode, _
        1, _
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
    reCtx.stepsLimit = 10
    reCtx.input = "abccdbccdbccdbccdbccde"
    
    Dim result As Long
    result = RegexNfaMatcher.NfaMatch(reCtx)
    Assert.AreEqual Len(reCtx.input), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub







