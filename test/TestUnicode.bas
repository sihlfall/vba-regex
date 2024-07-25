Attribute VB_Name = "TestUnicode"
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


'@TestMethod("Unicode")
Private Sub Unicode0001()
    Assert.IsTrue RegexUnicodeSupport.UnicodeIsLineTerminator(10)
    Assert.IsTrue RegexUnicodeSupport.UnicodeIsLineTerminator(13)
    Assert.IsTrue RegexUnicodeSupport.UnicodeIsLineTerminator(&H2028&)
    Assert.IsTrue RegexUnicodeSupport.UnicodeIsLineTerminator(&H2029&)
    Assert.IsFalse RegexUnicodeSupport.UnicodeIsLineTerminator(AscW("A"))
    Assert.IsFalse RegexUnicodeSupport.UnicodeIsLineTerminator(9)
    Assert.IsFalse RegexUnicodeSupport.UnicodeIsLineTerminator(11)
    Assert.IsFalse RegexUnicodeSupport.UnicodeIsLineTerminator(12)
    Assert.IsFalse RegexUnicodeSupport.UnicodeIsLineTerminator(14)
    Assert.IsFalse RegexUnicodeSupport.UnicodeIsLineTerminator(&H2027&)
    Assert.IsFalse RegexUnicodeSupport.UnicodeIsLineTerminator(&H202A&)
    Assert.IsFalse RegexUnicodeSupport.UnicodeIsLineTerminator(-1)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
