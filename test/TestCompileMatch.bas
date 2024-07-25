Attribute VB_Name = "TestCompileMatch"
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

'@TestMethod("CompileMatch")
Private Sub CompileMatch0001()
    ' beginning-of-string (^), matching
    On Error GoTo TestFail
    
    Dim bytecode() As Long, captures As RegexDfsMatcher.CapturesTy
    Dim sPattern As String, sHaystack As String
    
    sPattern = "^abcde"
    sHaystack = "abcdefg"
    
    RegexCompiler.Compile bytecode, sPattern
    Assert.IsTrue RegexDfsMatcher.DfsMatch(captures, bytecode, sHaystack) <> -1
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CompileMatch")
Private Sub CompileMatch0002()
    ' beginning-of-string (^), failing
    On Error GoTo TestFail
    
    Dim bytecode() As Long, captures As RegexDfsMatcher.CapturesTy
    Dim sPattern As String, sHaystack As String
    
    sPattern = "b^a+"
    sHaystack = "baaa"
    
    RegexCompiler.Compile bytecode, sPattern
    Assert.IsTrue RegexDfsMatcher.DfsMatch(captures, bytecode, sHaystack) = -1
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CompileMatch")
Private Sub CompileMatch0003()
    ' beginning-of-string (^), failing for eol in non-multiline mode
    On Error GoTo TestFail
    
    Dim bytecode() As Long, captures As RegexDfsMatcher.CapturesTy
    Dim sPattern As String, sHaystack As String
    
    sPattern = "^aaaa"
    sHaystack = vbCrLf & "aaaaaaa"
    
    RegexCompiler.Compile bytecode, sPattern
    Assert.IsTrue RegexDfsMatcher.DfsMatch(captures, bytecode, sHaystack) = -1
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CompileMatch")
Private Sub CompileMatch0004()
    ' beginning-of-string (^), matching for eol in multiline mode
    On Error GoTo TestFail
    
    Dim bytecode() As Long, captures As RegexDfsMatcher.CapturesTy
    Dim sPattern As String, sHaystack As String
    
    sPattern = "^aaaa"
    sHaystack = vbCrLf & "aaaaaaa"
    
    RegexCompiler.Compile bytecode, sPattern
    Assert.IsTrue RegexDfsMatcher.DfsMatch(captures, bytecode, sHaystack, multiline:=True) <> -1
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CompileMatch")
Private Sub CompileMatch0005()
    ' beginning-of-string (^), matching for beginning-of-string in lookbehind
    On Error GoTo TestFail
    
    Dim bytecode() As Long, captures As RegexDfsMatcher.CapturesTy
    Dim sPattern As String, sHaystack As String
    
    sPattern = "aa(?<=^aa)aa"
    sHaystack = "aaaabbb"
    
    RegexCompiler.Compile bytecode, sPattern
    Assert.IsTrue RegexDfsMatcher.DfsMatch(captures, bytecode, sHaystack) <> -1
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CompileMatch")
Private Sub CompileMatch0006()
    ' beginning-of-string (^), matching for eol in lookbehind in multiline mode
    On Error GoTo TestFail
    
    Dim bytecode() As Long, captures As RegexDfsMatcher.CapturesTy
    Dim sPattern As String, sHaystack As String
    
    sPattern = "aa(?<=^aa)aa"
    sHaystack = "b" & ChrW$(10) & "aaaabbb"
    
    RegexCompiler.Compile bytecode, sPattern
    Assert.IsTrue RegexDfsMatcher.DfsMatch(captures, bytecode, sHaystack, multiline:=True) <> -1
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CompileMatch")
Private Sub CompileMatch0010()
    ' end-of-string ($), matching for end-of-string
    On Error GoTo TestFail
    
    Dim bytecode() As Long, captures As RegexDfsMatcher.CapturesTy
    Dim sPattern As String, sHaystack As String
    
    sPattern = "aa$"
    sHaystack = "aaaaaaaa"
    
    RegexCompiler.Compile bytecode, sPattern
    Assert.IsTrue RegexDfsMatcher.DfsMatch(captures, bytecode, sHaystack) <> -1
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CompileMatch")
Private Sub CompileMatch0011()
    ' end-of-string ($), failing for eol
    On Error GoTo TestFail
    
    Dim bytecode() As Long, captures As RegexDfsMatcher.CapturesTy
    Dim sPattern As String, sHaystack As String
    
    sPattern = "aa$"
    sHaystack = "aaaaaaaa" & vbCrLf & "bb"
    
    RegexCompiler.Compile bytecode, sPattern
    Assert.IsTrue RegexDfsMatcher.DfsMatch(captures, bytecode, sHaystack) = -1
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CompileMatch")
Private Sub CompileMatch0012()
    ' end-of-string ($), matching for end-of-string in multiline mode
    On Error GoTo TestFail
    
    Dim bytecode() As Long, captures As RegexDfsMatcher.CapturesTy
    Dim sPattern As String, sHaystack As String
    
    sPattern = "aa$"
    sHaystack = "aaaaaaaa"
    
    RegexCompiler.Compile bytecode, sPattern
    Assert.IsTrue RegexDfsMatcher.DfsMatch(captures, bytecode, sHaystack, multiline:=True) <> -1
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CompileMatch")
Private Sub CompileMatch0014()
    ' end-of-string ($), matching for eol in multiline mode
    On Error GoTo TestFail
    
    Dim bytecode() As Long, captures As RegexDfsMatcher.CapturesTy
    Dim sPattern As String, sHaystack As String
    
    sPattern = "aa$"
    sHaystack = "aaaaaaaa" & vbCrLf & "bb"
    
    RegexCompiler.Compile bytecode, sPattern
    Assert.IsTrue RegexDfsMatcher.DfsMatch(captures, bytecode, sHaystack, multiline:=True) <> -1
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CompileMatch")
Private Sub CompileMatch0015()
    ' end-of-string ($) in lookbehind, matching for eol in multiline mode
    On Error GoTo TestFail
    
    Dim bytecode() As Long, captures As RegexDfsMatcher.CapturesTy
    Dim sPattern As String, sHaystack As String
    
    sPattern = "(?<=aa$[^a]*)bb"
    sHaystack = "aaaaaaaa" & vbCrLf & "bb"
    
    RegexCompiler.Compile bytecode, sPattern
    Assert.IsTrue RegexDfsMatcher.DfsMatch(captures, bytecode, sHaystack, multiline:=True) <> -1
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

