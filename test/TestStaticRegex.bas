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

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Features_108()
    ' parameter dotAll is respected
    Dim r As StaticRegex.RegexTy
    On Error GoTo TestFail
    
    StaticRegex.InitializeRegex r, "a.+b", caseInsensitive:=False
    Assert.IsTrue StaticRegex.Test(r, "a" & vbCrLf & "b", dotAll:=True)
    Assert.IsFalse StaticRegex.Test(r, "a" & vbCrLf & "b", dotAll:=False)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Features_120()
    ' possessive ?+
    Dim r As StaticRegex.RegexTy
    On Error GoTo TestFail
    
    StaticRegex.InitializeRegex r, "(ab)?+(ab)+", caseInsensitive:=False
    Assert.IsFalse StaticRegex.Test(r, "ab")
    Assert.IsTrue StaticRegex.Test(r, "abab")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Features_121()
    ' possessive *+
    Dim r As StaticRegex.RegexTy
    On Error GoTo TestFail
    
    StaticRegex.InitializeRegex r, "(ab)*+a[bc]", caseInsensitive:=False
    Assert.IsTrue StaticRegex.Test(r, "ac")
    Assert.IsTrue StaticRegex.Test(r, "abac")
    Assert.IsTrue StaticRegex.Test(r, "ababac")
    Assert.IsFalse StaticRegex.Test(r, "ababab")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Features_122()
    ' possessive ++
    Dim r As StaticRegex.RegexTy
    On Error GoTo TestFail
    
    StaticRegex.InitializeRegex r, "(ab)++(a[bc])?", caseInsensitive:=False
    Assert.IsFalse StaticRegex.Test(r, "ac")
    Assert.IsTrue StaticRegex.Test(r, "abac")
    Assert.IsTrue StaticRegex.Test(r, "ababac")
    Assert.IsTrue StaticRegex.Test(r, "ababab")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Features_123()
    ' possessive *+
    Dim r As StaticRegex.RegexTy
    On Error GoTo TestFail
    
    StaticRegex.InitializeRegex r, "^(aa)*a$", caseInsensitive:=False
    Assert.IsFalse StaticRegex.Test(r, "")
    Assert.IsTrue StaticRegex.Test(r, "a")
    Assert.IsFalse StaticRegex.Test(r, "aa")
    Assert.IsTrue StaticRegex.Test(r, "aaa")
    Assert.IsFalse StaticRegex.Test(r, "aaaa")
    Assert.IsTrue StaticRegex.Test(r, "aaaaa")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Features_124()
    ' possessive {m,n}+
    Dim r As StaticRegex.RegexTy
    On Error GoTo TestFail
    
    StaticRegex.InitializeRegex r, "^(aa?){1,3}+a$", caseInsensitive:=False
    Assert.IsFalse StaticRegex.Test(r, "")
    Assert.IsFalse StaticRegex.Test(r, "a")
    Assert.IsFalse StaticRegex.Test(r, "aa")
    Assert.IsFalse StaticRegex.Test(r, "aaa")
    Assert.IsFalse StaticRegex.Test(r, "aaaa")
    Assert.IsFalse StaticRegex.Test(r, "aaaaa")
    Assert.IsFalse StaticRegex.Test(r, "aaaaaa")
    Assert.IsTrue StaticRegex.Test(r, "aaaaaaa")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Features_125()
    ' atomic group
    Dim r As StaticRegex.RegexTy
    On Error GoTo TestFail
    
    StaticRegex.InitializeRegex r, "^(?>a|ab)c", caseInsensitive:=False
    Assert.IsTrue StaticRegex.Test(r, "ac")
    Assert.IsFalse StaticRegex.Test(r, "abc")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("StaticRegex")
Private Sub StaticRegex_Replace_001()
    On Error GoTo TestFail
    
    Dim r As StaticRegex.RegexTy
    Dim haystack As String, replacer As String, expected As String, actual As String
    
    haystack = "123abc"
    replacer = "$1"
    StaticRegex.InitializeRegex r, "(\d*)(\D*)"
    expected = "123"
    
    actual = StaticRegex.Replace(r, replacer:=replacer, haystack:=haystack, localMatch:=False)
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("StaticRegex")
Private Sub StaticRegex_Replace_002()
    On Error GoTo TestFail
    
    Dim r As StaticRegex.RegexTy
    Dim haystack As String, replacer As String, expected As String, actual As String
    
    haystack = "123abc"
    replacer = "$1"
    StaticRegex.InitializeRegex r, "(\d*)(\D?)"
    expected = "123bc"
    
    actual = StaticRegex.Replace(r, replacer:=replacer, haystack:=haystack, localMatch:=True)
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Replace_003()
    On Error GoTo TestFail
    
    Dim r As StaticRegex.RegexTy
    Dim haystack As String, replacer As String, expected As String, actual As String
    
    haystack = "123abc"
    replacer = "$1"
    StaticRegex.InitializeRegex r, "(\d*)(\D?)"
    expected = "123"
    
    actual = StaticRegex.Replace(r, replacer:=replacer, haystack:=haystack, localMatch:=False)
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Replace_004()
    On Error GoTo TestFail
    Dim r As StaticRegex.RegexTy
    Dim inputString As String
    Dim expected As String
    
    inputString = "On Jul-4-1776, independence was declared. " & "On Apr-30-1789, George Washington became the first president."
    expected = "On Jul, independence was declared. On Apr, George Washington became the first president."

    StaticRegex.InitializeRegex r, "(?<month>\w{3})-(?<day>\d{1,2})-(?<year>\d{4})"
    
    Assert.AreEqual expected, StaticRegex.Replace(r, replacer:="$1", haystack:=inputString, localMatch:=False)
    
    Assert.AreEqual expected, StaticRegex.Replace(r, replacer:="$<month>", haystack:=inputString, localMatch:=False)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub MakeStringArray(ByRef ary() As String, ParamArray p() As Variant)
    Dim u As Long, i As Long
    u = UBound(p)
    ReDim ary(0 To u) As String
    For i = 0 To u
        ary(i) = p(i)
    Next
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_SplitByRegex_001()
    On Error GoTo TestFail
    
    Dim r As StaticRegex.RegexTy
    Dim haystack As String, expected() As String, actual As Collection
    Dim s As Variant, i As Long, nExpected As Long, nEqual As Long
    
    haystack = "12|3|4||5|6"
    StaticRegex.InitializeRegex r, "\|"
    MakeStringArray expected, _
        "12", "3", "4", "", "5", "6"
    
    Set actual = StaticRegex.SplitByRegex(r, haystack:=haystack, localMatch:=False)
    
    nExpected = UBound(expected) + 1
    Assert.AreEqual nExpected, actual.Count
    
    i = 0
    nEqual = 0
    For Each s In actual
        If i >= nExpected Then Exit For
        If s <> expected(i) Then Exit For
        nEqual = nEqual + 1
        i = i + 1
    Next

    Assert.AreEqual nExpected, nEqual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_SplitByRegex_002()
    On Error GoTo TestFail
    
    Dim r As StaticRegex.RegexTy
    Dim haystack As String, expected() As String, actual As Collection
    Dim s As Variant, i As Long, nExpected As Long, nEqual As Long
    
    haystack = "12"
    StaticRegex.InitializeRegex r, "\|"
    MakeStringArray expected, _
        "12"
    
    Set actual = StaticRegex.SplitByRegex(r, haystack:=haystack, localMatch:=False)
    
    nExpected = UBound(expected) + 1
    Assert.AreEqual nExpected, actual.Count
    
    i = 0
    nEqual = 0
    For Each s In actual
        If i >= nExpected Then Exit For
        If s <> expected(i) Then Exit For
        nEqual = nEqual + 1
        i = i + 1
    Next

    Assert.AreEqual nExpected, nEqual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_SplitByRegex_003()
    On Error GoTo TestFail
    
    Dim r As StaticRegex.RegexTy
    Dim haystack As String, expected() As String, actual As Collection
    Dim s As Variant, i As Long, nExpected As Long, nEqual As Long
    
    haystack = ""
    StaticRegex.InitializeRegex r, "\|"
    MakeStringArray expected, _
        ""
    
    Set actual = StaticRegex.SplitByRegex(r, haystack:=haystack, localMatch:=False)
    
    nExpected = UBound(expected) + 1
    Assert.AreEqual nExpected, actual.Count
    
    i = 0
    nEqual = 0
    For Each s In actual
        If i >= nExpected Then Exit For
        If s <> expected(i) Then Exit For
        nEqual = nEqual + 1
        i = i + 1
    Next

    Assert.AreEqual nExpected, nEqual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_SplitByRegex_004()
    On Error GoTo TestFail
    
    Dim r As StaticRegex.RegexTy
    Dim haystack As String, expected() As String, actual As Collection
    Dim s As Variant, i As Long, nExpected As Long, nEqual As Long
    
    haystack = "12|3*4|*5*6"
    StaticRegex.InitializeRegex r, "\||\*"
    MakeStringArray expected, _
        "12", "3", "4", "", "5", "6"
    
    Set actual = StaticRegex.SplitByRegex(r, haystack:=haystack, localMatch:=False)
    
    nExpected = UBound(expected) + 1
    Assert.AreEqual nExpected, actual.Count
    
    i = 0
    nEqual = 0
    For Each s In actual
        If i >= nExpected Then Exit For
        If s <> expected(i) Then Exit For
        nEqual = nEqual + 1
        i = i + 1
    Next

    Assert.AreEqual nExpected, nEqual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_SplitByRegex_005()
    On Error GoTo TestFail
    
    Dim r As StaticRegex.RegexTy
    Dim haystack As String, expected() As String, actual As Collection
    Dim s As Variant, i As Long, nExpected As Long, nEqual As Long
    
    haystack = "12|3*4|*5*6"
    StaticRegex.InitializeRegex r, "(\||\*)"
    MakeStringArray expected, _
        "12", "|", "3", "*", "4", "|", "", "*", "5", "*", "6"
    
    Set actual = StaticRegex.SplitByRegex(r, haystack:=haystack, localMatch:=False)
    
    nExpected = UBound(expected) + 1
    Assert.AreEqual nExpected, actual.Count
    
    i = 0
    nEqual = 0
    For Each s In actual
        If i >= nExpected Then Exit For
        If s <> expected(i) Then Exit For
        nEqual = nEqual + 1
        i = i + 1
    Next

    Assert.AreEqual nExpected, nEqual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

