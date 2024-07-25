Attribute VB_Name = "TestLexer"
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


'@TestMethod("Lexer")
Private Sub Lexer_ParseReToken001()
    ' Correctly reads empty string
    On Error GoTo TestFail
    
    Dim lex As RegexLexer.Ty
    Dim i As Long, toks(0 To 9) As RegexLexer.ReToken
    
    RegexLexer.Initialize lex, ""
    i = 0
    Do
        RegexLexer.ParseReToken lex, toks(i)
        If toks(i).t = RETOK_EOF Then Exit Do
        i = i + 1
    Loop
    
    Assert.AreEqual RETOK_EOF, toks(0).t
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Lexer")
Private Sub Lexer_ParseReToken002()
    ' Correctly reads ASCII characters
    On Error GoTo TestFail
    
    Dim lex As RegexLexer.Ty
    Dim i As Long, toks(0 To 9) As RegexLexer.ReToken
    
    RegexLexer.Initialize lex, "abc"
    i = 0
    Do
        RegexLexer.ParseReToken lex, toks(i)
        If toks(i).t = RETOK_EOF Then Exit Do
        i = i + 1
    Loop
    
    Assert.AreEqual RETOK_ATOM_CHAR, toks(0).t
    Assert.AreEqual 0& + AscW("a"), toks(0).num
    Assert.AreEqual RETOK_ATOM_CHAR, toks(1).t
    Assert.AreEqual 0& + AscW("b"), toks(1).num
    Assert.AreEqual RETOK_ATOM_CHAR, toks(2).t
    Assert.AreEqual 0& + AscW("c"), toks(2).num
    Assert.AreEqual RETOK_EOF, toks(3).t
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Lexer")
Private Sub Lexer_ParseReToken003()
    ' Correctly reads asterisk characters
    On Error GoTo TestFail
    
    Dim lex As RegexLexer.Ty
    Dim i As Long, toks(0 To 9) As RegexLexer.ReToken
    
    RegexLexer.Initialize lex, "a*b"
    i = 0
    Do
        RegexLexer.ParseReToken lex, toks(i)
        If toks(i).t = RETOK_EOF Then Exit Do
        i = i + 1
    Loop
    
    Assert.AreEqual RETOK_ATOM_CHAR, toks(0).t
    Assert.AreEqual 0& + AscW("a"), toks(0).num
    Assert.AreEqual RETOK_QUANTIFIER, toks(1).t
    Assert.AreEqual 0&, toks(1).qmin
    Assert.AreEqual RE_QUANTIFIER_INFINITE, toks(1).qmax
    Assert.AreEqual RETOK_ATOM_CHAR, toks(2).t
    Assert.AreEqual 0& + AscW("b"), toks(2).num
    Assert.AreEqual RETOK_EOF, toks(3).t
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Lexer")
Private Sub Lexer_ParseReToken004()
    ' Correctly reads \xHH escape sequence
    On Error GoTo TestFail
    
    Dim lex As RegexLexer.Ty
    Dim i As Long, toks(0 To 9) As RegexLexer.ReToken
    
    RegexLexer.Initialize lex, "a\x1bb"
    i = 0
    Do
        RegexLexer.ParseReToken lex, toks(i)
        If toks(i).t = RETOK_EOF Then Exit Do
        i = i + 1
    Loop
    
    Assert.AreEqual RETOK_ATOM_CHAR, toks(0).t
    Assert.AreEqual 0& + AscW("a"), toks(0).num
    Assert.AreEqual RETOK_ATOM_CHAR, toks(1).t
    Assert.AreEqual &H1B&, toks(1).num
    Assert.AreEqual RETOK_ATOM_CHAR, toks(2).t
    Assert.AreEqual 0& + AscW("b"), toks(2).num
    Assert.AreEqual RETOK_EOF, toks(3).t
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub Lexer_ParseReToken005()
    ' Correctly reads \uHHHH escape sequence
    On Error GoTo TestFail
    
    Dim lex As RegexLexer.Ty
    Dim i As Long, toks(0 To 9) As RegexLexer.ReToken
    
    RegexLexer.Initialize lex, "a\u1abcb"
    i = 0
    Do
        RegexLexer.ParseReToken lex, toks(i)
        If toks(i).t = RETOK_EOF Then Exit Do
        i = i + 1
    Loop
    
    Assert.AreEqual RETOK_ATOM_CHAR, toks(0).t
    Assert.AreEqual 0& + AscW("a"), toks(0).num
    Assert.AreEqual RETOK_ATOM_CHAR, toks(1).t
    Assert.AreEqual &H1ABC&, toks(1).num
    Assert.AreEqual RETOK_ATOM_CHAR, toks(2).t
    Assert.AreEqual 0& + AscW("b"), toks(2).num
    Assert.AreEqual RETOK_EOF, toks(3).t
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Lexer")
Private Sub Lexer_ParseReToken006()
    ' Correctly reads positive lookbehind
    On Error GoTo TestFail
    
    Dim lex As RegexLexer.Ty
    Dim i As Long, toks(0 To 9) As RegexLexer.ReToken
    
    RegexLexer.Initialize lex, "a(?<=b)c"
    i = 0
    Do
        RegexLexer.ParseReToken lex, toks(i)
        If toks(i).t = RETOK_EOF Then Exit Do
        i = i + 1
    Loop
    
    Assert.AreEqual RETOK_ATOM_CHAR, toks(0).t
    Assert.AreEqual 0& + AscW("a"), toks(0).num
    Assert.AreEqual RETOK_ASSERT_START_POS_LOOKBEHIND, toks(1).t
    Assert.AreEqual RETOK_ATOM_CHAR, toks(2).t
    Assert.AreEqual 0& + AscW("b"), toks(2).num
    Assert.AreEqual RETOK_ATOM_END, toks(3).t
    Assert.AreEqual RETOK_ATOM_CHAR, toks(4).t
    Assert.AreEqual 0& + AscW("c"), toks(4).num
    Assert.AreEqual RETOK_EOF, toks(5).t
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Lexer")
Private Sub Lexer_ParseReToken007()
    ' Correctly reads negative lookbehind
    On Error GoTo TestFail
    
    Dim lex As RegexLexer.Ty
    Dim i As Long, toks(0 To 9) As RegexLexer.ReToken
    
    RegexLexer.Initialize lex, "a(?<!b)c"
    i = 0
    Do
        RegexLexer.ParseReToken lex, toks(i)
        If toks(i).t = RETOK_EOF Then Exit Do
        i = i + 1
    Loop
    
    Assert.AreEqual RETOK_ATOM_CHAR, toks(0).t
    Assert.AreEqual 0& + AscW("a"), toks(0).num
    Assert.AreEqual RETOK_ASSERT_START_NEG_LOOKBEHIND, toks(1).t
    Assert.AreEqual RETOK_ATOM_CHAR, toks(2).t
    Assert.AreEqual 0& + AscW("b"), toks(2).num
    Assert.AreEqual RETOK_ATOM_END, toks(3).t
    Assert.AreEqual RETOK_ATOM_CHAR, toks(4).t
    Assert.AreEqual 0& + AscW("c"), toks(4).num
    Assert.AreEqual RETOK_EOF, toks(5).t
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Lexer")
Private Sub Lexer_ParseReToken010()
    ' Correctly reads named captures
    'On Error GoTo TestFail
    
    Dim lex As RegexLexer.Ty
    Dim i As Long, toks(0 To 19) As RegexLexer.ReToken, dump(0 To 8) As Long, expected() As Long
    
    RegexLexer.Initialize lex, "(?<XYZ>x)(?<UV>y)(?<XYZ>x)(?<A>z)"
    i = 0
    Do
        RegexLexer.ParseReToken lex, toks(i)
        If toks(i).t = RETOK_EOF Then Exit Do
        i = i + 1
    Loop
    
    RegexIdentifierSupport.RedBlackDumpTree dump, 0, lex.identifierTree
    
    MakeArray expected, _
        30, 1, 2, _
        13, 2, 1, _
        4, 3, 0
    Assert.SequenceEquals expected, dump
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
