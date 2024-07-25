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

'@TestMethod("Ast")
Private Sub Ast_AstToBytecode0001()
    Dim ast() As Long, actual() As Long, expected() As Long

    MakeArray ast, _
        1, AST_STRING, 4, AscW("a"), AscW("b"), AscW("c"), AscW("d")
    RegexAst.AstToBytecode ast, emptyIdentifierTree, False, actual
    MakeArray expected, _
        -1, 0, 0, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_CHAR, AscW("c"), _
        REOP_CHAR, AscW("d"), _
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
    Dim ast() As Long, actual() As Long, expected() As Long

    MakeArray ast, _
        10, _
        AST_STRING, 2, AscW("a"), AscW("b"), _
        AST_STRING, 3, AscW("c"), AscW("d"), AscW("e"), _
        AST_DISJ, 1, 5
    RegexAst.AstToBytecode ast, emptyIdentifierTree, False, actual
    MakeArray expected, _
        -1, 0, 0, _
        REOP_SPLIT1, 6, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, 6, _
        REOP_CHAR, AscW("c"), _
        REOP_CHAR, AscW("d"), _
        REOP_CHAR, AscW("e"), _
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
    Dim ast() As Long, actual() As Long, expected() As Long

    MakeArray ast, _
        16, _
        AST_STRING, 2, AscW("a"), AscW("b"), _
        AST_STRING, 3, AscW("c"), AscW("d"), AscW("e"), _
        AST_STRING, 1, AscW("f"), _
        AST_DISJ, 5, 10, _
        AST_DISJ, 1, 13
    RegexAst.AstToBytecode ast, emptyIdentifierTree, False, actual
    MakeArray expected, _
        -1, 0, 0, _
        REOP_SPLIT1, 6, _
        REOP_CHAR, AscW("a"), _
        REOP_CHAR, AscW("b"), _
        REOP_JUMP, 12, _
        REOP_SPLIT1, 8, _
        REOP_CHAR, AscW("c"), _
        REOP_CHAR, AscW("d"), _
        REOP_CHAR, AscW("e"), _
        REOP_JUMP, 2, _
        REOP_CHAR, AscW("f"), _
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
    Dim ast() As Long, actual() As Long, expected() As Long

    MakeArray ast, _
        6, AST_CHAR, AscW("a"), AST_CONCAT, 1, 1, AST_DISJ, 3, 3
    RegexAst.AstToBytecode ast, emptyIdentifierTree, False, actual
    MakeArray expected, _
        -1, 0, 0, _
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


