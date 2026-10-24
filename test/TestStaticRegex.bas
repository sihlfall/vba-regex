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


'@TestMethod("StaticRegex")
Private Sub StaticRegex_Test_010()
    ' Tests for \b, \B
    ' taken from duktape
    Dim r As StaticRegex.RegexTy
    Dim pattern As String, haystack As String, expected As Boolean
    On Error GoTo TestFail
    
    pattern = ".\b.": haystack = "ab": expected = False: GoSub PerformTest
    pattern = ".\b.": haystack = "a=": expected = True: GoSub PerformTest
    pattern = ".\b.": haystack = "=a": expected = True: GoSub PerformTest
    pattern = ".\b.": haystack = "==": expected = False: GoSub PerformTest
    pattern = ".\B.": haystack = "ab": expected = True: GoSub PerformTest
    pattern = ".\B.": haystack = "a=": expected = False: GoSub PerformTest
    pattern = ".\B.": haystack = "=a": expected = False: GoSub PerformTest
    pattern = ".\B.": haystack = "==": expected = True: GoSub PerformTest

    pattern = "\b.": haystack = "a": expected = True: GoSub PerformTest
    pattern = "\b.": haystack = "=": expected = False: GoSub PerformTest
    pattern = "\B.": haystack = "a": expected = False: GoSub PerformTest
    pattern = "\B.": haystack = "=": expected = True: GoSub PerformTest

    pattern = ".\b": haystack = "a": expected = True: GoSub PerformTest
    pattern = ".\b": haystack = "=": expected = False: GoSub PerformTest
    pattern = ".\B": haystack = "a": expected = False: GoSub PerformTest
    pattern = ".\B": haystack = "=": expected = True: GoSub PerformTest

    pattern = "\b": haystack = "": expected = False: GoSub PerformTest
    pattern = "\B": haystack = "": expected = True: GoSub PerformTest

    pattern = ".\b.": haystack = "0b": expected = False: GoSub PerformTest
    pattern = ".\b.": haystack = "0=": expected = True: GoSub PerformTest
    pattern = ".\b.": haystack = "9b": expected = False: GoSub PerformTest
    pattern = ".\b.": haystack = "9=": expected = True: GoSub PerformTest
    pattern = ".\b.": haystack = "_b": expected = False: GoSub PerformTest
    pattern = ".\b.": haystack = "_=": expected = True: GoSub PerformTest
    pattern = ".\b.": haystack = "Ab": expected = False: GoSub PerformTest
    pattern = ".\b.": haystack = "A=": expected = True: GoSub PerformTest
    pattern = ".\b.": haystack = "Zb": expected = False: GoSub PerformTest
    pattern = ".\b.": haystack = "Z=": expected = True: GoSub PerformTest
    pattern = ".\b.": haystack = "/b": expected = True: GoSub PerformTest
    pattern = ".\b.": haystack = "/=": expected = False: GoSub PerformTest
    pattern = ".\b.": haystack = ":b": expected = True: GoSub PerformTest
    pattern = ".\b.": haystack = ":=": expected = False: GoSub PerformTest
    pattern = ".\b.": haystack = "@b": expected = True: GoSub PerformTest
    pattern = ".\b.": haystack = "@=": expected = False: GoSub PerformTest
    pattern = ".\b.": haystack = "\[b": expected = True: GoSub PerformTest
    pattern = ".\b.": haystack = "\[=": expected = False: GoSub PerformTest
    pattern = ".\b.": haystack = "`b": expected = True: GoSub PerformTest
    pattern = ".\b.": haystack = "`=": expected = False: GoSub PerformTest
    pattern = ".\b.": haystack = "{b": expected = True: GoSub PerformTest
    pattern = ".\b.": haystack = "{=": expected = False: GoSub PerformTest

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
PerformTest:
    StaticRegex.InitializeRegex r, pattern
    Assert.IsTrue expected = StaticRegex.Test(r, haystack)
    Return
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Test_011()
    ' Tests for \b, \B
    ' taken from https://github.com/6DiegoDiego9/VBA-dotNET-regex/blob/main/CLRRegexTest.bas
    Dim r As StaticRegex.RegexTy
    Dim pattern As String, haystack As String, expected As Boolean
    On Error GoTo TestFail
    
    pattern = "\b([a-z0-9]+)\s*=\s*([0-9.]+)\b": haystack = "val1=100, val2 = 200.5, val3= 300, val4=400. This is"
    StaticRegex.InitializeRegex r, pattern
    Assert.IsTrue StaticRegex.Test(r, haystack)
    
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

Private Sub StaticRegex_Features_130()
    ' GetCapture, GetCaptureByName
    Dim r As StaticRegex.RegexTy
    'On Error GoTo TestFail
    
    Dim inputString As String
    Dim pattern As String
    
    inputString = "Hello world. Hello universe."
    pattern = "(Hello) (?<named>world)?"
    
    StaticRegex.InitializeRegex r, pattern
    
    Dim matcherState As StaticRegex.MatcherStateTy
    matcherState.localMatch = False

    Assert.IsTrue StaticRegex.MatchNext(matcherState, r, inputString)
    Assert.AreEqual "Hello world", StaticRegex.GetCapture(matcherState, inputString, 0)
    Assert.AreEqual "Hello", StaticRegex.GetCapture(matcherState, inputString, 1)

    Assert.AreEqual "world", StaticRegex.GetCapture(matcherState, inputString, 2)

    Assert.AreEqual "world", StaticRegex.GetCaptureByName(matcherState, r, inputString, "named")

    Assert.IsTrue StaticRegex.MatchNext(matcherState, r, inputString)
    Assert.AreEqual "Hello ", StaticRegex.GetCapture(matcherState, inputString, 0)
    Assert.AreEqual "Hello", StaticRegex.GetCapture(matcherState, inputString, 1)
    Assert.AreEqual vbNullString, StaticRegex.GetCapture(matcherState, inputString, 2)
    Assert.AreEqual vbNullString, StaticRegex.GetCaptureByName(matcherState, r, inputString, "named")

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

'Issue: https://github.com/sihlfall/vba-regex/issues/20
'@TestMethod("StaticRegex")
Private Sub StaticRegex_MatchNext_SkippedEmptyMatchIsNotRepeated()
    On Error GoTo TestFail

    Dim r As StaticRegex.RegexTy
    Dim haystack As String
    Dim matcherState As StaticRegex.MatcherStateTy

    haystack = "a" & vbLf & "b"
    StaticRegex.InitializeRegex r, "^"
    StaticRegex.InitializeMatcherState matcherState, multiline:=True

    Assert.IsTrue StaticRegex.MatchNext(matcherState, r, haystack)
    Assert.AreEqual 1&, matcherState.captures.entireMatch.start
    Assert.AreEqual 0&, matcherState.captures.entireMatch.Length

    ' The search skips forward to the empty match at position 3.
    Assert.IsTrue StaticRegex.MatchNext(matcherState, r, haystack)
    Assert.AreEqual 3&, matcherState.captures.entireMatch.start
    Assert.AreEqual 0&, matcherState.captures.entireMatch.Length

    ' Only two line starts exist; position 3 must not be returned twice.
    Assert.IsFalse StaticRegex.MatchNext(matcherState, r, haystack)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'Issue: https://github.com/sihlfall/vba-regex/issues/20
'@TestMethod("StaticRegex")
Private Sub StaticRegex_MatchNext_FailureAfterEmptyMatchStaysExhausted()
    On Error GoTo TestFail

    Dim r As StaticRegex.RegexTy
    Dim haystack As String
    Dim matcherState As StaticRegex.MatcherStateTy

    haystack = "a"
    StaticRegex.InitializeRegex r, "^"
    StaticRegex.InitializeMatcherState matcherState

    Assert.IsTrue StaticRegex.MatchNext(matcherState, r, haystack)
    Assert.AreEqual 1&, matcherState.captures.entireMatch.start
    Assert.AreEqual 0&, matcherState.captures.entireMatch.Length

    ' A failed match leaves the previous empty capture in the state.
    Assert.IsFalse StaticRegex.MatchNext(matcherState, r, haystack)
    Assert.AreEqual -1&, matcherState.current

    ' Further calls must remain exhausted.
    Assert.IsFalse StaticRegex.MatchNext(matcherState, r, haystack)
    Assert.AreEqual -1&, matcherState.current

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


' Same function as MakeStringArray in README Example 6.
Private Function ReadmeStringArray(ParamArray strings() As Variant) As String()
    Dim ary() As String, i As Long
    ReDim ary(0 To UBound(strings) - LBound(strings)) As String
    For i = LBound(strings) To UBound(strings)
        ary(i - LBound(strings)) = strings(i)
    Next
    ReadmeStringArray = ary
End Function

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Readme_Example6_MatchThenList()
    On Error GoTo TestFail

    Dim regex As StaticRegex.RegexTy
    Dim exampleString As String, results() As String

    exampleString = "On Jul-4-1776, independence was declared. " & _
        "On Apr-30-1789, George Washington became the first president."
    StaticRegex.InitializeRegex regex, _
        "(?<month>\w{3})-(?<day>\d{1,2})-(?<year>\d{4})"

    StaticRegex.MatchThenList results, regex, exampleString, _
        ReadmeStringArray("$&", "$<day>", "$<month>", "$<year>")

    Assert.AreEqual 0&, LBound(results, 1)
    Assert.AreEqual 1&, UBound(results, 1)
    Assert.AreEqual 0&, LBound(results, 2)
    Assert.AreEqual 3&, UBound(results, 2)
    Assert.AreEqual "Jul-4-1776", results(0, 0)
    Assert.AreEqual "4", results(0, 1)
    Assert.AreEqual "Jul", results(0, 2)
    Assert.AreEqual "1776", results(0, 3)
    Assert.AreEqual "Apr-30-1789", results(1, 0)
    Assert.AreEqual "30", results(1, 1)
    Assert.AreEqual "Apr", results(1, 2)
    Assert.AreEqual "1789", results(1, 3)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Readme_Example7_Replace()
    On Error GoTo TestFail

    Dim regex As StaticRegex.RegexTy
    Dim exampleString As String, expected As String

    exampleString = "On Jul-4-1776, independence was declared. " & _
        "On Apr-30-1789, George Washington became the first president."
    StaticRegex.InitializeRegex regex, _
        "(?<month>\w{3})-(?<day>\d{1,2})-(?<year>\d{4})"
    expected = "On 4 Jul 1776, independence was declared. " & _
        "On 30 Apr 1789, George Washington became the first president."
    Assert.AreEqual expected, StaticRegex.Replace( _
        regex, replacer:="$<day> $<month> $<year>", haystack:=exampleString)

    expected = "On 4 Jul 1776, independence was declared. " & _
        "On Apr-30-1789, George Washington became the first president."
    Assert.AreEqual expected, StaticRegex.Replace( _
        regex, replacer:="$<day> $<month> $<year>", haystack:=exampleString, localMatch:=True)
    Assert.AreEqual "No dates here.", StaticRegex.Replace( _
        regex, replacer:="$<day> $<month> $<year>", haystack:="No dates here.")
    Assert.AreEqual "Jul-4-1776|Jul|Jul|$", StaticRegex.Replace( _
        regex, replacer:="$&|$1|$<month>|$$", haystack:="Jul-4-1776")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Readme_Example8_SplitByRegex()
    On Error GoTo TestFail

    Dim separatorRegex As StaticRegex.RegexTy
    Dim parts As Collection, part As Variant, i As Long
    Dim expected() As String

    StaticRegex.InitializeRegex separatorRegex, "[,;]\s*"
    Set parts = StaticRegex.SplitByRegex(separatorRegex, "red, green; blue")
    expected = ReadmeStringArray("red", "green", "blue")
    Assert.AreEqual 3&, parts.Count
    i = 0
    For Each part In parts
        Assert.AreEqual expected(i), CStr(part)
        i = i + 1
    Next

    Set parts = StaticRegex.SplitByRegex(separatorRegex, "red, green; blue", localMatch:=True)
    Assert.AreEqual 2&, parts.Count
    Assert.AreEqual "red", CStr(parts(1))
    Assert.AreEqual "green; blue", CStr(parts(2))
    Set parts = StaticRegex.SplitByRegex(separatorRegex, "red")
    Assert.AreEqual 1&, parts.Count
    Assert.AreEqual "red", CStr(parts(1))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Readme_Example8_CapturesAndEmptyParts()
    On Error GoTo TestFail

    Dim separatorRegex As StaticRegex.RegexTy
    Dim parts As Collection

    StaticRegex.InitializeRegex separatorRegex, "([,;])\s*"
    Set parts = StaticRegex.SplitByRegex(separatorRegex, ",red;")
    Assert.AreEqual 5&, parts.Count
    Assert.AreEqual vbNullString, CStr(parts(1))
    Assert.AreEqual ",", CStr(parts(2))
    Assert.AreEqual "red", CStr(parts(3))
    Assert.AreEqual ";", CStr(parts(4))
    Assert.AreEqual vbNullString, CStr(parts(5))

    StaticRegex.InitializeRegex separatorRegex, "(?:[,;])\s*"
    Set parts = StaticRegex.SplitByRegex(separatorRegex, ",red;")
    Assert.AreEqual 3&, parts.Count
    Assert.AreEqual vbNullString, CStr(parts(1))
    Assert.AreEqual "red", CStr(parts(2))
    Assert.AreEqual vbNullString, CStr(parts(3))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Readme_Options_CaseInsensitiveAndDotAll()
    On Error GoTo TestFail

    Dim messageRegex As StaticRegex.RegexTy
    Dim message As String

    message = "BEGIN" & vbCrLf & "END"
    StaticRegex.InitializeRegex messageRegex, "begin.*end", caseInsensitive:=True
    Assert.IsTrue StaticRegex.Test(messageRegex, message, dotAll:=True)
    ' Both options are needed for this input.
    Assert.IsFalse StaticRegex.Test(messageRegex, message)
    StaticRegex.InitializeRegex messageRegex, "begin.*end"
    Assert.IsFalse StaticRegex.Test(messageRegex, message, dotAll:=True)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Readme_Options_MatchNext()
    On Error GoTo TestFail

    Dim regex As StaticRegex.RegexTy
    Dim matcherState As StaticRegex.MatcherStateTy
    Dim exampleString As String, actual As String, nMatches As Long

    exampleString = "On Jul-4-1776, independence was declared. " & _
        "On Apr-30-1789, George Washington became the first president."
    StaticRegex.InitializeRegex regex, _
        "(?<month>\w{3})-(?<day>\d{1,2})-(?<year>\d{4})"
    StaticRegex.InitializeMatcherState matcherState, multiline:=True, dotAll:=True

    Do While StaticRegex.MatchNext(matcherState, regex, exampleString)
        If nMatches > 0 Then actual = actual & ", "
        actual = actual & StaticRegex.GetCapture(matcherState, exampleString)
        nMatches = nMatches + 1
        ' Fail instead of looping indefinitely if iteration regresses.
        Assert.IsTrue nMatches <= 2
        If nMatches > 2 Then Exit Do
    Loop
    Assert.AreEqual 2&, nMatches
    Assert.AreEqual "Jul-4-1776, Apr-30-1789", actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Readme_Options_InitializeAndResetMatcherState()
    On Error GoTo TestFail

    Dim regex As StaticRegex.RegexTy
    Dim matcherState As StaticRegex.MatcherStateTy
    Dim haystack As String

    haystack = "prefix" & vbLf & "begin" & vbLf & "end" & vbLf & _
        "begin" & vbLf & "end" & vbLf & "suffix"
    StaticRegex.InitializeRegex regex, "^begin.end$"
    ' A new state has all options off, so this pattern cannot match.
    Assert.IsFalse matcherState.localMatch
    Assert.IsFalse matcherState.multiline
    Assert.IsFalse matcherState.dotAll
    Assert.IsFalse StaticRegex.MatchNext(matcherState, regex, haystack)

    StaticRegex.InitializeMatcherState matcherState, localMatch:=True, multiline:=True, dotAll:=True
    Assert.IsTrue StaticRegex.MatchNext(matcherState, regex, haystack)
    Assert.AreEqual "begin" & vbLf & "end", StaticRegex.GetCapture(matcherState, haystack)
    Assert.AreEqual 8&, matcherState.captures.entireMatch.start
    Assert.IsFalse StaticRegex.MatchNext(matcherState, regex, haystack)

    StaticRegex.ResetMatcherState matcherState
    Assert.IsTrue matcherState.localMatch
    Assert.IsTrue matcherState.multiline
    Assert.IsTrue matcherState.dotAll
    Assert.IsTrue StaticRegex.MatchNext(matcherState, regex, haystack)
    Assert.AreEqual 8&, matcherState.captures.entireMatch.start
    Assert.IsFalse StaticRegex.MatchNext(matcherState, regex, haystack)

    ' Reinitializing also resets the options to their defaults.
    StaticRegex.InitializeMatcherState matcherState
    Assert.IsFalse matcherState.localMatch
    Assert.IsFalse matcherState.multiline
    Assert.IsFalse matcherState.dotAll
    Assert.IsFalse StaticRegex.MatchNext(matcherState, regex, haystack)

    ' With both line options enabled, the default is to find both matches.
    StaticRegex.InitializeMatcherState matcherState, multiline:=True, dotAll:=True
    Assert.IsTrue StaticRegex.MatchNext(matcherState, regex, haystack)
    Assert.AreEqual 8&, matcherState.captures.entireMatch.start
    Assert.IsTrue StaticRegex.MatchNext(matcherState, regex, haystack)
    Assert.AreEqual 18&, matcherState.captures.entireMatch.start
    Assert.IsFalse StaticRegex.MatchNext(matcherState, regex, haystack)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Readme_Options_MatchFrom()
    On Error GoTo TestFail

    Dim regex As StaticRegex.RegexTy
    Dim matcherState As StaticRegex.MatcherStateTy

    StaticRegex.InitializeRegex regex, "cat"
    Assert.IsTrue StaticRegex.Match(matcherState, regex, "cat cat")
    Assert.AreEqual 1&, matcherState.captures.entireMatch.start
    Assert.IsTrue StaticRegex.Match(matcherState, regex, "cat cat", matchFrom:=5)
    Assert.AreEqual 5&, matcherState.captures.entireMatch.start
    ' Searching may skip forward from the requested position.
    Assert.IsTrue StaticRegex.Match(matcherState, regex, "cat cat", matchFrom:=2)
    Assert.AreEqual 5&, matcherState.captures.entireMatch.start
    Assert.IsFalse StaticRegex.Match(matcherState, regex, "cat cat", matchFrom:=6)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticRegex")
Private Sub StaticRegex_Readme_Options_StepsLimit()
    On Error GoTo TestFail

    Dim regex As StaticRegex.RegexTy

    StaticRegex.InitializeRegex regex, "abc"
    Assert.AreEqual 10000&, regex.stepsLimit
    regex.stepsLimit = 1
    Assert.IsFalse StaticRegex.Test(regex, "abc")
    regex.stepsLimit = 10000
    Assert.IsTrue StaticRegex.Test(regex, "abc")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
