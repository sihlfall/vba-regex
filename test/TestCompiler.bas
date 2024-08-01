Attribute VB_Name = "TestCompiler"
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

'@TestMethod("Compiler")
Private Sub Compiler_Compile0001()
    ' simple two-character pattern
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "ab"
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0002()
    ' simple four-character pattern
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "abcd"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_CHAR, AscW("c"), _
        REOP_CHAR, AscW("d"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0003()
    ' empty pattern
    On Error GoTo TestFail
    
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, ""
    
    Dim expected() As Long
    Dim actual() As Long
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_SAVE, 1, _
        REOP_MATCH
    Assert.SequenceEquals expected, bytecode
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0004()
    ' one-character pattern
    On Error GoTo TestFail
    
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, "a"
    
    Dim expected() As Long
    Dim actual() As Long
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Assert.SequenceEquals expected, bytecode
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Compiler")
Private Sub Compiler_Compile0005()
    ' three-character pattern
    On Error GoTo TestFail
    
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, "abc"
    
    Dim expected() As Long
    Dim actual() As Long
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Assert.SequenceEquals expected, bytecode
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0010()
    ' Kleene star
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "ab*c"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0011()
    ' Empty Kleene star
    On Error GoTo TestFail

    Dim actual() As Long
    RegexCompiler.Compile actual, "()*x"
    

    Dim expected() As Long
    
    MakeArray expected, _
        3, 0, 0, _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 6, _
        REOP_SAVE, 2, _
        REOP_SAVE, 3, _
        REOP_JUMP, -8, _
        REOP_CHAR, AscW("x"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0012()
    ' Kleene star, complex pattern
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "(ab)*x"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
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
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0013()
    ' Empty Kleene star with noncapture group
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "(?:)*x"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 2, _
        REOP_JUMP, -4, _
        REOP_CHAR, AscW("x"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0014()
    ' Kleene star, non-greedy
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "ab*?c"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT2, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0015()
    ' Another Kleene star with parentheses
    On Error GoTo TestFail
    
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, "a(bc)*d"
    
    Dim expected() As Long
    Dim actual() As Long
    MakeArray expected, _
        3, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT1, 10, _
        REOP_SAVE, 2, _
        REOP_CHAR, AscW("b"), _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 3, _
        REOP_JUMP, -12, _
        REOP_CHAR, AscW("d"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Assert.SequenceEquals expected, bytecode
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0020()
    ' Plus
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "ab+c"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
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
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0021()
    ' Plus, complex pattern
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "(ab)+x"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
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
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0030()
    ' Question mark
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "ab?c"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT1, 2, _
        REOP_CHAR, AscW("b"), _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0040()
    ' Quantifier, two numbers
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "ab{2,5}c"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_EXACTLY_INIT, _
        REOP_REPEAT_EXACTLY_START, 2, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_EXACTLY_END, 2, 5, _
        REOP_REPEAT_GREEDY_MAX_INIT, _
        REOP_REPEAT_GREEDY_MAX_START, 3, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_GREEDY_MAX_END, 3, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0041()
    ' Quantifier, only first number
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "ab{12,}c"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_EXACTLY_INIT, _
        REOP_REPEAT_EXACTLY_START, 12, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_EXACTLY_END, 12, 5, _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0042()
    ' Quantifier, only second number
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "ab{,20}c"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_GREEDY_MAX_INIT, _
        REOP_REPEAT_GREEDY_MAX_START, 20, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_GREEDY_MAX_END, 20, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0043()
    ' Quantifier, no comma
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "ab{20}c"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_REPEAT_EXACTLY_INIT, _
        REOP_REPEAT_EXACTLY_START, 20, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_REPEAT_EXACTLY_END, 20, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0044()
    ' Quantifier, exactly zero
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "ab{0}c"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0045()
    ' Quantifier, exactly zero
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "(?:(?:a{0}){0})"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Compiler")
Private Sub Compiler_Compile0050()
    ' Disjunction
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "a|b"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("a"), _
        REOP_JUMP, 2, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0051()
    ' Two disjunctions
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "a|b|c"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("a"), _
        REOP_JUMP, 8, _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, 2, _
        REOP_CHAR, AscW("c"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0052()
    ' single disjunction with two-character terms
    On Error GoTo TestFail
    
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, "ab|cd"
    
    Dim expected() As Long
    Dim actual() As Long
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 6, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, 4, _
        REOP_CHAR, AscW("c"), _
        REOP_CHAR, AscW("d"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Assert.SequenceEquals expected, bytecode
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0053()
    ' two disjunctions with two-character terms
    On Error GoTo TestFail
    
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, "ab|cd|ef"
    
    Dim expected() As Long
    Dim actual() As Long
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 6, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, 12, _
        REOP_SPLIT1, 6, _
        REOP_CHAR, AscW("c"), _
        REOP_CHAR, AscW("d"), _
        REOP_JUMP, 4, _
        REOP_CHAR, AscW("e"), _
        REOP_CHAR, AscW("f"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Assert.SequenceEquals expected, bytecode
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Compiler")
Private Sub Compiler_Compile0100()
    ' Parentheses
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "a(bc)d"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
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
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0101()
    ' Nested parentheses
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "a(b(cd)e)f"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
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
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0102()
    ' Two disjunctions in parentheses
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "a(b|c|d)e"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        3, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SAVE, 2, _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, 8, _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("c"), _
        REOP_JUMP, 2, _
        REOP_CHAR, AscW("d"), _
        REOP_SAVE, 3, _
        REOP_CHAR, AscW("e"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Compiler")
Private Sub Compiler_Compile0200()
    ' Caret
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "^ab"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_ASSERT_START, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0202()
    ' Dollar
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "ab$"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_ASSERT_END, _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0204()
    ' Period
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "a.b"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_DOT, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0210()
    ' Word boundary
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "\b"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_ASSERT_WORD_BOUNDARY, _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0211()
    ' Not word boundary
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "\B"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_ASSERT_NOT_WORD_BOUNDARY, _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0300()
    ' range
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "[A-Z]"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0301()
    ' range
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "[A-Za-z]"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_RANGES, 2, AscW("A"), AscW("Z"), AscW("a"), AscW("z"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0302()
    ' range
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "[A-Za-z]0"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_RANGES, 2, AscW("A"), AscW("Z"), AscW("a"), AscW("z"), _
        REOP_CHAR, AscW("0"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0303()
    ' closing square bracket as first character of range
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "(?:(?:.|[]a]))$"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 3, _
        REOP_DOT, _
        REOP_JUMP, 6, _
        REOP_RANGES, 2, AscW("]"), AscW("]"), AscW("a"), AscW("a"), _
        REOP_ASSERT_END, _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Compiler")
Private Sub Compiler_Compile0304()
    ' closing square bracket as first character of range
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "(?:(?:[]a]*))$"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_SPLIT1, 8, _
        REOP_RANGES, 2, AscW("]"), AscW("]"), AscW("a"), AscW("a"), _
        REOP_JUMP, -10, _
        REOP_ASSERT_END, _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0305()
    ' range
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "[a-zA-Z]0"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_RANGES, 2, AscW("A"), AscW("Z"), AscW("a"), AscW("z"), _
        REOP_CHAR, AscW("0"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Compiler")
Private Sub Compiler_Compile0306()
    ' range
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "[e-ix-zh-ma-fu-w]0"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_RANGES, 2, AscW("a"), AscW("m"), AscW("u"), AscW("z"), _
        REOP_CHAR, AscW("0"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0307()
    ' range
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "[987654]0"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_RANGES, 1, AscW("4"), AscW("9"), _
        REOP_CHAR, AscW("0"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0308()
    ' another range, test merging and sorting within range
    On Error GoTo TestFail
    
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, "a[abcU-Z]b"
    
    Dim expected() As Long
    Dim actual() As Long
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_RANGES, 2, AscW("U"), AscW("Z"), AscW("a"), AscW("c"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Assert.SequenceEquals expected, bytecode
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0310()
    ' inverted range
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "[^A-Z]"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_INVRANGES, 1, AscW("A"), AscW("Z"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0320()
    ' range with escape \d
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "a[\d]b"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_RANGES, 1, AscW("0"), AscW("9"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0321()
    ' range with escape \x
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "a[\xAB]b"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_RANGES, 1, &HAB&, &HAB&, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0322()
    ' range with escape \x
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "a[\x90-\xA0]b"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_RANGES, 1, &H90&, &HA0&, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0323()
    ' range with escape \uHHHH
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "a[\u0500-\u7000]b"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_RANGES, 1, &H500&, &H7000&, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0324()
    ' range with escape \u{H+}
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "a[\u{50}-\u{700}]b"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_RANGES, 1, &H50&, &H700&, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0325()
    ' range with octal escapes
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "a[\0\15\322\8\9]b"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_RANGES, 4, 0&, 0&, 13&, 13&, AscW("8"), AscW("9"), 210, 210, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0330()
    ' Canonicalization if case-insensitive
    On Error GoTo TestFail
    
    Dim actual() As Long
    
    RegexCompiler.Compile actual, "abuzäöü", caseInsensitive:=True
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("A"), _
        REOP_CHAR, AscW("B"), _
        REOP_CHAR, AscW("U"), _
        REOP_CHAR, AscW("Z"), _
        REOP_CHAR, AscW("Ä"), _
        REOP_CHAR, AscW("Ö"), _
        REOP_CHAR, AscW("Ü"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0331()
    ' Canonicalization if case-insensitive, range
    On Error GoTo TestFail
    
    Dim actual() As Long
    
    RegexCompiler.Compile actual, "a[a-zM-X]b", caseInsensitive:=True
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("A"), _
        REOP_RANGES, 1, AscW("A"), AscW("Z"), _
        REOP_CHAR, AscW("B"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0332()
    ' Canonicalization if case-insensitive, range
    On Error GoTo TestFail
    
    Dim actual() As Long
    
    RegexCompiler.Compile actual, "a[x-{]b", caseInsensitive:=True
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 1, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("A"), _
        REOP_RANGES, 2, AscW("X"), AscW("Z"), AscW("{"), AscW("{"), _
        REOP_CHAR, AscW("B"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub



'@TestMethod("Compiler")
Private Sub Compiler_Compile0350()
    ' \d
    On Error GoTo TestFail
    
    Dim actual() As Long
    
    RegexUnicodeSupport.RangeTablesInitialize
    
    RegexCompiler.Compile actual, "a\db"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_RANGES, 1, AscW("0"), AscW("9"), _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0351()
    ' \D
    On Error GoTo TestFail
    
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, "a\Db"
    
    Dim expected() As Long
    Dim actual() As Long
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_RANGES, 2, &H80000000, AscW("0") - 1, AscW("9") + 1, &H7FFFFFFF, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Assert.SequenceEquals expected, bytecode
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Compiler")
Private Sub Compiler_Compile0400()
    ' positive lookahead
    'On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "ab(?=cd)"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
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
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0401()
    ' negative lookahead
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "ab(?!cd)e"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_CHECK_LOOKAHEAD, _
        REOP_LOOKNEG, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_CHAR, AscW("d"), _
        REOP_END_LOOKNEG, _
        REOP_CHAR, AscW("e"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0402()
    ' another positive lookahead
    On Error GoTo TestFail
    
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, "a(?=bc)d"
    
    Dim expected() As Long
    Dim actual() As Long
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHECK_LOOKAHEAD, _
        REOP_LOOKPOS, 5, _
        REOP_CHAR, AscW("b"), _
        REOP_CHAR, AscW("c"), _
        REOP_END_LOOKPOS, _
        REOP_CHAR, AscW("d"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Assert.SequenceEquals expected, bytecode
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0410()
    ' non-capture group
    On Error GoTo TestFail
    
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, "a(?:bc)*d"
    
    Dim expected() As Long
    Dim actual() As Long
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SPLIT1, 6, _
        REOP_CHAR, AscW("b"), _
        REOP_CHAR, AscW("c"), _
        REOP_JUMP, -8, _
        REOP_CHAR, AscW("d"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Assert.SequenceEquals expected, bytecode
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0411()
    ' non-capture group
    On Error GoTo TestFail
    
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, "^(?:(?:A|(?:AA)))$"
    
    Dim expected() As Long
    Dim actual() As Long
    MakeArray expected, _
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
    Assert.SequenceEquals expected, bytecode
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub



'@TestMethod("Compiler")
Private Sub Compiler_Compile0420()
    ' lookbehind
    ' On Error GoTo TestFail
    
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, "a(?<=bc)d"
    
    Dim expected() As Long
    Dim actual() As Long
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHECK_LOOKBEHIND, _
        REOP_LOOKPOS, 5, _
        REOP_CHAR, AscW("c"), _
        REOP_CHAR, AscW("b"), _
        REOP_END_LOOKPOS, _
        REOP_CHAR, AscW("d"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Assert.SequenceEquals expected, bytecode
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0421()
    ' star inside lookbehind
    On Error GoTo TestFail
    
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, "a(?<=xy*z)d"
    
    Dim expected() As Long
    Dim actual() As Long
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHECK_LOOKBEHIND, _
        REOP_LOOKPOS, 11, _
        REOP_CHAR, AscW("z"), _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("y"), _
        REOP_JUMP, -6, _
        REOP_CHAR, AscW("x"), _
        REOP_END_LOOKPOS, _
        REOP_CHAR, AscW("d"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Assert.SequenceEquals expected, bytecode
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0422()
    ' disjunction inside lookbehind
    On Error GoTo TestFail
    
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, "a(?<=x(i|j)z)d"
    
    Dim expected() As Long
    Dim actual() As Long
    MakeArray expected, _
        3, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHECK_LOOKBEHIND, _
        REOP_LOOKPOS, 17, _
        REOP_CHAR, AscW("z"), _
        REOP_SAVE, 3, _
        REOP_SPLIT1, 4, _
        REOP_CHAR, AscW("i"), _
        REOP_JUMP, 2, _
        REOP_CHAR, AscW("j"), _
        REOP_SAVE, 2, _
        REOP_CHAR, AscW("x"), _
        REOP_END_LOOKPOS, _
        REOP_CHAR, AscW("d"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Assert.SequenceEquals expected, bytecode
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0423()
    ' lookahead inside lookbehind
    On Error GoTo TestFail
    
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, "a(?<=x(?=ij)z)d"
    
    Dim expected() As Long
    Dim actual() As Long
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHECK_LOOKBEHIND, _
        REOP_LOOKPOS, 13, _
        REOP_CHAR, AscW("z"), _
        REOP_CHECK_LOOKAHEAD, _
        REOP_LOOKPOS, 5, _
        REOP_CHAR, AscW("i"), _
        REOP_CHAR, AscW("j"), _
        REOP_END_LOOKPOS, _
        REOP_CHAR, AscW("x"), _
        REOP_END_LOOKPOS, _
        REOP_CHAR, AscW("d"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Assert.SequenceEquals expected, bytecode
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Compiler")
Private Sub Compiler_Compile0450()
    ' backreference
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "a(b)c(d)e\1\2"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
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
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0451()
    ' backreference, two-digit backreference number
    On Error GoTo TestFail
    
    Dim actual() As Long
    RegexCompiler.Compile actual, "()()()()()()()()()()\10"
    
    
    Dim expected() As Long
    
    MakeArray expected, _
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
        REOP_SAVE, 20, REOP_SAVE, 21, _
        REOP_BACKREFERENCE, 10, _
        REOP_SAVE, 1, _
        REOP_MATCH
    
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0452()
    ' another backreference
    On Error GoTo TestFail
    
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, "a(b)c\1d"
    
    Dim expected() As Long
    Dim actual() As Long
    MakeArray expected, _
        3, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_SAVE, 2, _
        REOP_CHAR, AscW("b"), _
        REOP_SAVE, 3, _
        REOP_CHAR, AscW("c"), _
        REOP_BACKREFERENCE, 1, _
        REOP_CHAR, AscW("d"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Assert.SequenceEquals expected, bytecode
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0500()
    ' i modifier
    On Error GoTo TestFail
    
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, "a(?i:bCdeF)gh"
    
    Dim expected() As Long
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHANGE_MODIFIERS, RegexBytecode.MODIFIER_I_WRITE Or RegexBytecode.MODIFIER_I_ACTIVE, _
        REOP_CHAR, AscW("B"), _
        REOP_CHAR, AscW("C"), _
        REOP_CHAR, AscW("D"), _
        REOP_CHAR, AscW("E"), _
        REOP_CHAR, AscW("F"), _
        REOP_RESTORE_MODIFIERS, _
        REOP_CHAR, AscW("g"), _
        REOP_CHAR, AscW("h"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Assert.SequenceEquals expected, bytecode
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Compiler")
Private Sub Compiler_Compile0501()
    ' i and -i modifiers
    On Error GoTo TestFail
    
    Dim bytecode() As Long
    RegexCompiler.Compile bytecode, "a(?i:b(?-i:cd)ef)gh"
    
    Dim expected() As Long
    MakeArray expected, _
        1, 0, 0, _
        REOP_SAVE, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHANGE_MODIFIERS, RegexBytecode.MODIFIER_I_WRITE Or RegexBytecode.MODIFIER_I_ACTIVE, _
        REOP_CHAR, AscW("B"), _
        REOP_CHANGE_MODIFIERS, RegexBytecode.MODIFIER_I_WRITE, _
        REOP_CHAR, AscW("c"), _
        REOP_CHAR, AscW("d"), _
        REOP_RESTORE_MODIFIERS, _
        REOP_CHAR, AscW("E"), _
        REOP_CHAR, AscW("F"), _
        REOP_RESTORE_MODIFIERS, _
        REOP_CHAR, AscW("g"), _
        REOP_CHAR, AscW("h"), _
        REOP_SAVE, 1, _
        REOP_MATCH
    Assert.SequenceEquals expected, bytecode
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


