Attribute VB_Name = "TestWorksheetFunctions"
Option Explicit

Private checks As Long
Private currentCase As String
Private failures As String

Public Function RunBasicTests() As Variant
    On Error GoTo Failed
    checks = 0
    failures = vbNullString
    TestScalars
    TestOptions
    TestArraysAndRanges
    TestErrors
    TestWorksheetFormulas
    If Len(failures) > 0 Then RunBasicTests = failures Else RunBasicTests = checks
    Exit Function
Failed:
    RunBasicTests = currentCase & ": " & Err.Description
End Function

Private Sub TestScalars()
    AssertSingle "test finds substring", VbaRegexWorksheetFunctions.RegexTest("abc123def", "\d+"), True
    AssertSingle "test no match", VbaRegexWorksheetFunctions.RegexTest("abc", "\d+"), False
    AssertSingle "test empty input", VbaRegexWorksheetFunctions.RegexTest("", "^$"), True
    AssertSingle "test empty pattern", VbaRegexWorksheetFunctions.RegexTest("abc", ""), True
    AssertSingle "match first occurrence", VbaRegexWorksheetFunctions.RegexMatch("ab12cd34", "\d+"), "12"
    AssertSingle "match no occurrence", VbaRegexWorksheetFunctions.RegexMatch("abc", "\d+"), ""
    AssertSingle "match captures", VbaRegexWorksheetFunctions.RegexMatch("abc-123", "([a-z]+)-(\d+)", "$2/$1"), "123/abc"
    AssertSingle "match whole capture", VbaRegexWorksheetFunctions.RegexMatch("abc12", "\d+", "[$&]"), "[12]"
    AssertSingle "match empty format", VbaRegexWorksheetFunctions.RegexMatch("abc", "a", ""), ""
    AssertSingle "replace all occurrences", VbaRegexWorksheetFunctions.RegexReplace("ab12cd34", "\d+", "X"), "abXcdX"
    AssertSingle "replace no occurrence", VbaRegexWorksheetFunctions.RegexReplace("abc", "\d+", "X"), "abc"
    AssertSingle "replace captures", VbaRegexWorksheetFunctions.RegexReplace("a1 b2", "([a-z])(\d)", "$2$1"), "1a 2b"
    AssertSingle "replace deletion", VbaRegexWorksheetFunctions.RegexReplace("banana", "a", ""), "bnn"
    AssertSingle "replace literal dollar", VbaRegexWorksheetFunctions.RegexReplace("a", "a", "$$"), "$"
End Sub

Private Sub TestOptions()
    Dim text As String
    text = "x" & vbLf & "abc" & vbLf & "y"
    AssertSingle "test default case sensitive", VbaRegexWorksheetFunctions.RegexTest("ABC", "abc"), False
    AssertSingle "test ignore case", VbaRegexWorksheetFunctions.RegexTest("ABC", "abc", True), True
    AssertSingle "match default case sensitive", VbaRegexWorksheetFunctions.RegexMatch("ABC", "abc"), ""
    AssertSingle "match ignore case", VbaRegexWorksheetFunctions.RegexMatch("ABC", "abc", "$&", True), "ABC"
    AssertSingle "replace default case sensitive", VbaRegexWorksheetFunctions.RegexReplace("Aa", "a", "x"), "Ax"
    AssertSingle "replace ignore case", VbaRegexWorksheetFunctions.RegexReplace("Aa", "a", "x", True), "xx"
    AssertSingle "test default multiline", VbaRegexWorksheetFunctions.RegexTest(text, "^abc$"), False
    AssertSingle "test multiline", VbaRegexWorksheetFunctions.RegexTest(text, "^abc$", False, True), True
    AssertSingle "match default multiline", VbaRegexWorksheetFunctions.RegexMatch(text, "^abc$"), ""
    AssertSingle "match multiline", VbaRegexWorksheetFunctions.RegexMatch(text, "^abc$", "$&", False, True), "abc"
    AssertSingle "replace default multiline", VbaRegexWorksheetFunctions.RegexReplace(text, "^abc$", "z"), text
    AssertSingle "replace multiline", VbaRegexWorksheetFunctions.RegexReplace(text, "^abc$", "z", False, True), "x" & vbLf & "z" & vbLf & "y"
    AssertSingle "numeric boolean option", VbaRegexWorksheetFunctions.RegexTest("ABC", "abc", 1), True
End Sub

Private Sub TestArraysAndRanges()
    Dim inputs(1 To 2, 1 To 1) As Variant
    Dim patterns(1 To 1, 1 To 2) As Variant
    Dim flags(1 To 2, 1 To 1) As Variant
    Dim formats(1 To 1, 1 To 2) As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    inputs(1, 1) = "a1": inputs(2, 1) = "b2"
    patterns(1, 1) = "[a-z]": patterns(1, 2) = "\d"
    formats(1, 1) = "[$&]": formats(1, 2) = "<$&>"

    AssertGrid "array test cross broadcast", VbaRegexWorksheetFunctions.RegexTest(inputs, patterns), 2, 2, Array(True, True, True, True)
    AssertGrid "array match cross broadcast", VbaRegexWorksheetFunctions.RegexMatch(inputs, patterns), 2, 2, Array("a", "1", "b", "2")
    AssertGrid "array replace cross broadcast", VbaRegexWorksheetFunctions.RegexReplace(inputs, patterns, "X"), 2, 2, Array("X1", "aX", "X2", "bX")
    AssertGrid "array match formats", VbaRegexWorksheetFunctions.RegexMatch(inputs, "\d", formats), 2, 2, Array("[1]", "<1>", "[2]", "<2>")
    AssertGrid "array replace formats", VbaRegexWorksheetFunctions.RegexReplace(inputs, "\d", formats), 2, 2, Array("a[1]", "a<1>", "b[2]", "b<2>")

    ws.Range("A1:A2").Value2 = inputs
    ws.Range("C1:D1").Value2 = patterns
    AssertGrid "range test cross broadcast", VbaRegexWorksheetFunctions.RegexTest(ws.Range("A1:A2"), ws.Range("C1:D1")), 2, 2, Array(True, True, True, True)
    AssertGrid "range match cross broadcast", VbaRegexWorksheetFunctions.RegexMatch(ws.Range("A1:A2"), ws.Range("C1:D1")), 2, 2, Array("a", "1", "b", "2")
    AssertGrid "range replace cross broadcast", VbaRegexWorksheetFunctions.RegexReplace(ws.Range("A1:A2"), ws.Range("C1:D1"), "X"), 2, 2, Array("X1", "aX", "X2", "bX")
    AssertGrid "mixed range and array", VbaRegexWorksheetFunctions.RegexMatch(ws.Range("A1:A2"), patterns), 2, 2, Array("a", "1", "b", "2")
    AssertSingle "single cell range", VbaRegexWorksheetFunctions.RegexMatch(ws.Range("A1"), "\d"), "1"
    ws.Range("F1").ClearContents
    AssertSingle "empty cell string", VbaRegexWorksheetFunctions.RegexTest(ws.Range("F1"), "^$"), True

    flags(1, 1) = False: flags(2, 1) = True
    AssertGrid "per row case flags", VbaRegexWorksheetFunctions.RegexTest("A", "a", flags), 2, 1, Array(False, True)
    AssertGrid "per row multiline flags", VbaRegexWorksheetFunctions.RegexMatch("x" & vbLf & "a", "^a", "$&", False, flags), 2, 1, Array("", "a")

    inputs(1, 1) = "abc": inputs(2, 1) = "123"
    AssertGrid "elementwise test results", VbaRegexWorksheetFunctions.RegexTest(inputs, "\d+"), 2, 1, Array(False, True)
    inputs(1, 1) = CVErr(xlErrNA)
    AssertGrid "cell error stays local", VbaRegexWorksheetFunctions.RegexMatch(inputs, "\d+"), 2, 1, Array(CVErr(xlErrNA), "123")
End Sub

Private Sub TestErrors()
    AssertSingle "test invalid pattern", VbaRegexWorksheetFunctions.RegexTest("a", "["), CVErr(xlErrValue)
    AssertSingle "match invalid pattern", VbaRegexWorksheetFunctions.RegexMatch("a", "["), CVErr(xlErrValue)
    AssertSingle "replace invalid pattern", VbaRegexWorksheetFunctions.RegexReplace("a", "[", "x"), CVErr(xlErrValue)
    AssertSingle "non string input", VbaRegexWorksheetFunctions.RegexTest(123, "."), CVErr(xlErrValue)
    AssertSingle "non string pattern", VbaRegexWorksheetFunctions.RegexMatch("a", 123), CVErr(xlErrValue)
    AssertSingle "non string replacement", VbaRegexWorksheetFunctions.RegexReplace("a", "a", 123), CVErr(xlErrValue)
    AssertSingle "invalid boolean", VbaRegexWorksheetFunctions.RegexTest("a", "a", "true"), CVErr(xlErrValue)
    AssertSingle "input error propagation", VbaRegexWorksheetFunctions.RegexTest(CVErr(xlErrNA), "."), CVErr(xlErrNA)
    AssertSingle "pattern error propagation", VbaRegexWorksheetFunctions.RegexMatch("a", CVErr(xlErrRef)), CVErr(xlErrRef)
    AssertSingle "replacement error propagation", VbaRegexWorksheetFunctions.RegexReplace("a", "a", CVErr(xlErrDiv0)), CVErr(xlErrDiv0)
End Sub

Private Sub TestWorksheetFormulas()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("A1").Value2 = "abc123"
    ws.Range("H1").Formula = "=VbaRegexWorksheetFunctions.RegexTest(A1,""\d+"")"
    ws.Range("H2").Formula = "=VbaRegexWorksheetFunctions.RegexMatch(A1,""\d+"")"
    ws.Range("H3").Formula = "=VbaRegexWorksheetFunctions.RegexReplace(A1,""\d+"",""X"")"
    ws.Calculate
    currentCase = "worksheet RegexTest": AssertValue ws.Range("H1").Value2, True
    currentCase = "worksheet RegexMatch": AssertValue ws.Range("H2").Value2, "123"
    currentCase = "worksheet RegexReplace": AssertValue ws.Range("H3").Value2, "abcX"
End Sub

Private Sub AssertSingle(ByVal label As String, ByVal actual As Variant, ByVal expected As Variant)
    AssertGrid label, actual, 1, 1, Array(expected)
End Sub

Private Sub AssertGrid(ByVal label As String, ByVal actual As Variant, ByVal nr As Long, ByVal nc As Long, ByVal expected As Variant)
    Dim r As Long, c As Long
    currentCase = label
    If Not IsArray(actual) Then Err.Raise 513, , "Expected result array, got " & CStr(actual)
    If LBound(actual, 1) <> 0 Or LBound(actual, 2) <> 0 Then Err.Raise 513, , "Unexpected lower bounds"
    If UBound(actual, 1) <> nr - 1 Or UBound(actual, 2) <> nc - 1 Then Err.Raise 513, , "Wrong result dimensions"
    For r = 0 To nr - 1
        For c = 0 To nc - 1
            currentCase = label & " (" & r & ", " & c & ")"
            AssertValue actual(r, c), expected(r * nc + c)
        Next
    Next
End Sub

Private Sub AssertValue(ByVal actual As Variant, ByVal expected As Variant)
    Dim message As String
    If VarType(actual) <> VarType(expected) Then
        message = "Wrong value type: " & CStr(actual)
    ElseIf IsError(expected) Then
        If CStr(actual) <> CStr(expected) Then message = "Wrong error: " & CStr(actual)
    Else
        If actual <> expected Then message = "Expected " & CStr(expected) & ", got " & CStr(actual)
    End If
    If Len(message) > 0 Then failures = failures & currentCase & ": " & message & vbLf
    checks = checks + 1
End Sub