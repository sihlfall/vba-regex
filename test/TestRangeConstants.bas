Attribute VB_Name = "TestRangeConstants"
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


'@TestMethod("RangeConstants")
Private Sub RangeConstants0001()
    ' range with \d, matching
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    Dim result As Long
    RegexCompiler.Compile bytecode, "a[\d]b"
    Dim s As String: s = "a0b"
    
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RangeConstants")
Private Sub RangeConstants0002()
    ' range with \d, failing
    On Error GoTo TestFail

    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    Dim result As Long
    RegexCompiler.Compile bytecode, "a[\d]b"
    Dim s As String: s = "acb"
    
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RangeConstants")
Private Sub RangeConstants0011()
    ' range with \D, matching
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    Dim result As Long
    RegexCompiler.Compile bytecode, "a[\D]b"
    Dim s As String: s = "acb"
    
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RangeConstants")
Private Sub RangeConstants0012()
    ' range with \D, matching
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    Dim result As Long
    RegexCompiler.Compile bytecode, "a[\D]b"
    Dim s As String: s = "a7b"
    
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RangeConstants")
Private Sub RangeConstants0021()
    ' range with \s, matching
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    Dim result As Long
    RegexCompiler.Compile bytecode, "a[\s]b"
    Dim s As String: s = "a b"
    
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RangeConstants")
Private Sub RangeConstants0022()
    ' range with \s, failing
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    Dim result As Long
    RegexCompiler.Compile bytecode, "a[\s]b"
    Dim s As String: s = "acb"
    
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RangeConstants")
Private Sub RangeConstants0031()
    ' range with \S, matching
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    Dim result As Long
    RegexCompiler.Compile bytecode, "a[\S]b"
    Dim s As String: s = "a_b"
    
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RangeConstants")
Private Sub RangeConstants0032()
    ' range with \S, failing
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    Dim result As Long
    RegexCompiler.Compile bytecode, "a[\S]b"
    Dim s As String: s = "a b"
    
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RangeConstants")
Private Sub RangeConstants0041()
    ' range with \w, matching
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    Dim result As Long
    RegexCompiler.Compile bytecode, "a[\w]+b"
    Dim s As String: s = "a_abyzABYZb"
    
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RangeConstants")
Private Sub RangeConstants0042()
    ' range with \w, failing
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    Dim result As Long
    RegexCompiler.Compile bytecode, "a[\w]b"
    Dim s As String: s = "a,b"
    
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RangeConstants")
Private Sub RangeConstants0051()
    ' range with \W, matching
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    Dim result As Long
    RegexCompiler.Compile bytecode, "a[\W]+b"
    Dim s As String: s = "a,.+ ;$b"
    
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual Len(s), result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("RangeConstants")
Private Sub RangeConstants0052()
    ' range with \W, failing
    On Error GoTo TestFail
    
    Dim captures As RegexDfsMatcher.CapturesTy, bytecode() As Long
    Dim result As Long
    RegexCompiler.Compile bytecode, "a[\W]b"
    Dim s As String: s = "aCb"
    
    result = RegexDfsMatcher.DfsMatch(captures, bytecode, s)
    Assert.AreEqual -1&, result
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

