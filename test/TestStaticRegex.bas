Attribute VB_Name = "TestStaticRegex"
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

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Test_001()
    Dim r As StaticRegex.RegexTy
    On Error GoTo TestFail
    
    StaticRegex.InitializeRegex r, "abc"
    Assert.IsTrue StaticRegex.Test(r, "abc")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Test_002()
    Dim r As StaticRegex.RegexTy
    On Error GoTo TestFail
    
    StaticRegex.InitializeRegex r, "x[a-z]", caseInsensitive:=True
    Assert.IsTrue StaticRegex.Test(r, "xY")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub MakeArray(ByRef outAry() As String, ParamArray p() As Variant)
    ReDim outAry(0 To UBound(p)) As String
    Dim i As Long
    For i = 0 To UBound(p)
        outAry(i) = p(i)
    Next
End Sub

Private Sub ExtractAllNumberedCaptures(ByRef result() As String, ByRef captures As RegexDfsMatcher.CapturesTy, ByRef haystack As String)
    Dim i As Long

    ReDim result(0 To captures.nNumberedCaptures) As String
    
    result(0) = Mid$(haystack, captures.entireMatch.start, captures.entireMatch.Length)
    For i = 1 To captures.nNumberedCaptures
        With captures.numberedCaptures(i - 1)
            result(i) = Mid$(haystack, .start, .Length)
        End With
    Next
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Match_001()
    On Error GoTo TestFail
    
    Dim r As StaticRegex.RegexTy
    Dim actual() As String, expected() As String, haystack As String, matcherState As StaticRegex.MatcherStateTy
    
    MakeArray expected, "abc", "a", "b", "c"
    
    haystack = "abc"
    StaticRegex.InitializeRegex r, "(a)(b)(c)"
    
    Assert.IsTrue StaticRegex.Match(matcherState, r, haystack)
    ExtractAllNumberedCaptures actual, matcherState.captures, haystack
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Match_002()
    On Error GoTo TestFail
    
    Dim r As StaticRegex.RegexTy
    Dim actual() As String, expected() As String, haystack As String, matcherState As StaticRegex.MatcherStateTy
    
    MakeArray expected, "ccc"
    
    haystack = "abcccdef"
    StaticRegex.InitializeRegex r, "c{1,3}"
    
    Assert.IsTrue StaticRegex.Match(matcherState, r, haystack)
    ExtractAllNumberedCaptures actual, matcherState.captures, haystack
    Assert.SequenceEquals expected, actual
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Features_100()
    ' Modifiers: i
    Dim r As StaticRegex.RegexTy
    On Error GoTo TestFail
    
    StaticRegex.InitializeRegex r, "aBcD(?i:eFgH)iJkL", caseInsensitive:=False
    Assert.IsTrue StaticRegex.Test(r, "aBcDeFgHiJkL")
    Assert.IsTrue StaticRegex.Test(r, "aBcDEFGHiJkL")
    Assert.IsTrue StaticRegex.Test(r, "aBcDefghiJkL")
    
    Assert.IsFalse StaticRegex.Test(r, "aBcdeFgHiJkL")
    Assert.IsFalse StaticRegex.Test(r, "aBcdeFgHIjkL")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Features_101()
    ' Modifiers: -i
    Dim r As StaticRegex.RegexTy
    On Error GoTo TestFail
    
    StaticRegex.InitializeRegex r, "aBcD(?-i:eFgH)iJkL", caseInsensitive:=True
    Assert.IsTrue StaticRegex.Test(r, "aBcDeFgHiJkL")
    Assert.IsTrue StaticRegex.Test(r, "abcdeFgHijkl")
    Assert.IsTrue StaticRegex.Test(r, "ABCDeFgHIJKL")
    
    Assert.IsFalse StaticRegex.Test(r, "aBcDefghiJkL")
    Assert.IsFalse StaticRegex.Test(r, "aBcDEFGHiJkL")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Features_103()
    ' Modifiers: i inside -i
    Dim r As StaticRegex.RegexTy
    On Error GoTo TestFail
    
    StaticRegex.InitializeRegex r, "aBcD(?-i:e(?i:Fg)H)iJkL", caseInsensitive:=True
    Assert.IsTrue StaticRegex.Test(r, "aBcDeFGHiJkL")
    Assert.IsTrue StaticRegex.Test(r, "abcdefgHijkl")
    Assert.IsTrue StaticRegex.Test(r, "ABCDefgHIJKL")
    
    Assert.IsFalse StaticRegex.Test(r, "aBcDEfgHiJkL")
    Assert.IsFalse StaticRegex.Test(r, "aBcDeFGhiJkL")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Features_104()
    ' Modifiers: i applied to range
    Dim r As StaticRegex.RegexTy
    On Error GoTo TestFail
    
    StaticRegex.InitializeRegex r, "aBcD(?i:[A-Z]{4})iJkL", caseInsensitive:=False
    Assert.IsTrue StaticRegex.Test(r, "aBcDeFgHiJkL")
    
    Assert.IsFalse StaticRegex.Test(r, "abcdEfgHIJKL")
    Assert.IsFalse StaticRegex.Test(r, "ABCDeFGhijkl")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Features_105()
    ' Modifiers: status of i correctly restored after failing alternative
    Dim r As StaticRegex.RegexTy
    On Error GoTo TestFail
    
    StaticRegex.InitializeRegex r, "aB(?:cD(?i:eFgH)|cDxYz)", caseInsensitive:=False
    Assert.IsTrue StaticRegex.Test(r, "aBcDefgh")
    Assert.IsTrue StaticRegex.Test(r, "aBcDxYz")
    Assert.IsFalse StaticRegex.Test(r, "aBcDxyz")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Features_106()
    ' Modifiers: m and -m
    Dim r As StaticRegex.RegexTy
    On Error GoTo TestFail
    
    StaticRegex.InitializeRegex r, "(?m:^abc$)"
    Assert.IsTrue StaticRegex.Test(r, "xy" & vbCrLf & "abc" & vbCrLf & "xy", multiline:=False)
    
    StaticRegex.InitializeRegex r, "(?-m:^abc$)"
    Assert.IsFalse StaticRegex.Test(r, "xy" & vbCrLf & "abc" & vbCrLf & "xy", multiline:=True)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Features_107()
    ' Modifiers: s
    Dim r As StaticRegex.RegexTy
    On Error GoTo TestFail
    
    StaticRegex.InitializeRegex r, "^ab.+yz$"
    Assert.IsFalse StaticRegex.Test(r, "abcde" & vbCrLf & "uvwxyz", multiline:=False)
    
    StaticRegex.InitializeRegex r, "(?s:^ab.+yz$)"
    Assert.IsTrue StaticRegex.Test(r, "abcde" & vbCrLf & "uvwxyz", multiline:=False)
    
    StaticRegex.InitializeRegex r, "(?-s:^ab.+yz$)"
    Assert.IsFalse StaticRegex.Test(r, "abcde" & vbCrLf & "uvwxyz", multiline:=False, dotAll:=True)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


