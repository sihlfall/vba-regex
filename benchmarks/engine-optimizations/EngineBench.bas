Attribute VB_Name = "EngineBench"
Option Explicit
#If VBA7 Then
  Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
  Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#Else
  Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
  Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#End If

Private Function Seconds() As Double
  Static frequency As Currency
  Dim value As Currency
  If frequency = 0 Then QueryPerformanceFrequency frequency
  QueryPerformanceCounter value
  Seconds = value / frequency
End Function

Private Sub CaseData(ByVal id As Long, ByRef pattern As String, ByRef text As String, ByRef flags As String, ByRef label As String)
  Dim piece As String, count As Long, padding As String, i As Long
  pattern = "(?<id>\d{5}-[A-Z]{2}\d) (?<count>\d+) (?<date>\d{2}/\d{2}/\d{4})"
  piece = "12345-AA1 123 10/02/2019" & vbCrLf
  flags = "g"
  Select Case id
    Case 0: label = "No matches, 1000 characters": text = String$(1000, "x")
    Case 1: label = "1 match, 3 named captures": count = 1
    Case 2: label = "100 matches, 3 named captures": count = 100
    Case 3: label = "1000 matches, 3 named captures": count = 1000
    Case 4: label = "100 matches, no captures": count = 100: pattern = "\d{5}-[A-Z]{2}\d"
    Case 5: label = "100 sparse matches, 3 named captures": count = 100: padding = String$(200, "x") & vbCrLf
    Case 6: label = "Literal prefix, absent": pattern = "needle(\d+)": text = String$(10000, "x")
    Case 7: label = "Literal prefix, sparse": pattern = "needle(\d+)": piece = "needle123 ": count = 100: padding = String$(200, "x")
    Case 8: label = "Absolute anchor, absent": pattern = "^needle(\d+)": text = String$(10000, "x")
    Case 9: label = "Long literal runs": pattern = "abcdefghijklmnopqrstuvwxyz012345": piece = pattern & " ": count = 100
    Case 10
      label = "UTF16 captures": piece = ChrW$(&HE9) & ChrW$(&H6F22) & ChrW$(&HD83D) & ChrW$(&HDE00)
      pattern = "(?<word>" & piece & ")(?<num>\d+)": piece = piece & "123 ": count = 100
    Case 11: label = "Caseless backreferences": pattern = "(?<word>[A-Z]+)-\1": piece = "AbC-aBc ": count = 100: flags = "gi"
    Case 12: label = "Bounded backtracking": pattern = "(a|ab)+c": piece = "ababac ": count = 100
    Case 13: label = "Short input, 12 captures": pattern = "(a)(b)(c)(d)(e)(f)(g)(h)(i)(j)(k)(l)": piece = "abcdefghijkl": count = 1
    Case Else: Err.Raise 5, , "Unknown case"
  End Select
  For i = 1 To count: text = text & padding & piece: Next
End Sub

Public Function RunCase(ByVal id As Long, ByVal samples As Long, ByVal targetSeconds As Double) As Variant
  On Error GoTo Failed
  Dim pattern As String, text As String, flags As String, label As String, format As String, formats As Variant
  CaseData id, pattern, text, flags, label
  If id <= 3 Or id = 5 Then
    format = "$id,$date,$count\r\n": formats = Array("$id-$date", "$count")
  Else
    format = "$&\r\n": formats = Array("$&", "[$&]")
  End If
  Dim before As RegexBaseline, after As RegexCandidate
  Set before = RegexBaseline.Create(pattern, flags)
  Set after = RegexCandidate.Create(pattern, flags)
  Validate before, after, text, format, formats
  Dim iterations(0 To 7, 0 To 1) As Long, elapsed As Double
  Dim op As Long, mode As Long, sample As Long, slot As Long, n As Long, row As Long
  For op = 0 To 7
    For mode = 0 To 1
      elapsed = TimeOperation(mode, op, before, after, text, format, formats, 2)
      n = 1
      Do
        elapsed = TimeOperation(mode, op, before, after, text, format, formats, n)
        If elapsed >= targetSeconds Or n >= 65536 Then Exit Do
        n = n * 2
      Loop
      iterations(op, mode) = n
    Next
  Next
  Dim result() As Variant
  ReDim result(1 To samples * 16, 1 To 8)
  For sample = 1 To samples
    For slot = 0 To 15
      op = ((slot \ 2) + sample - 1) Mod 8
      mode = (slot + sample - 1) Mod 2
      n = iterations(op, mode)
      elapsed = TimeOperation(mode, op, before, after, text, format, formats, n)
      row = row + 1
      result(row, 1) = label: result(row, 2) = id
      result(row, 3) = Choose(op + 1, "List", "ListArr", "MatchAll", "MatchAllArr", "EngineTest", "Iterate", "Replace", "Compile")
      result(row, 4) = IIf(mode = 0, "baseline", "candidate")
      result(row, 5) = sample: result(row, 6) = n
      result(row, 7) = elapsed * 1000 / n: result(row, 8) = Len(text)
    Next
  Next
  RunCase = result
  Exit Function
Failed:
  RunCase = "VBA error " & Err.Number & ": " & Err.Description
End Function

Private Function TimeOperation(ByVal candidate As Long, ByVal op As Long, ByVal before As RegexBaseline, ByVal after As RegexCandidate, _
    ByRef text As String, ByRef format As String, ByRef formats As Variant, ByVal n As Long) As Double
  Dim started As Double, i As Long, values As Variant, names As Object, matches As Collection, output As String, count As Long, found As Boolean
  started = Seconds()
  For i = 1 To n
    If candidate = 0 Then
      Select Case op
        Case 0: output = before.List(text, format): output = vbNullString
        Case 1: values = before.ListArr(text, formats): Erase values
        Case 2: Set matches = before.MatchAll(text): Set matches = Nothing
        Case 3: values = before.MatchAllArr(text, names): Erase values: Set names = Nothing
        Case 4: found = before.BenchEngineTest(text)
        Case 5: count = before.BenchIterate(text)
        Case 6: output = before.Replace(text, "[$&]"): output = vbNullString
        Case 7: before.BenchCompile
      End Select
    Else
      Select Case op
        Case 0: output = after.List(text, format): output = vbNullString
        Case 1: values = after.ListArr(text, formats): Erase values
        Case 2: Set matches = after.MatchAll(text): Set matches = Nothing
        Case 3: values = after.MatchAllArr(text, names): Erase values: Set names = Nothing
        Case 4: found = after.BenchEngineTest(text)
        Case 5: count = after.BenchIterate(text)
        Case 6: output = after.Replace(text, "[$&]"): output = vbNullString
        Case 7: after.BenchCompile
      End Select
    End If
  Next
  TimeOperation = Seconds() - started
End Function

Private Sub Validate(ByVal before As RegexBaseline, ByVal after As RegexCandidate, ByRef text As String, ByRef format As String, ByRef formats As Variant)
  If before.List(text, format) <> after.List(text, format) Then Err.Raise 5, , "List differs"
  If before.Replace(text, "[$&]") <> after.Replace(text, "[$&]") Then Err.Raise 5, , "Replace differs"
  If before.BenchEngineTest(text) <> after.BenchEngineTest(text) Then Err.Raise 5, , "EngineTest differs"
  If before.BenchIterate(text) <> after.BenchIterate(text) Then Err.Raise 5, , "Iteration differs"
  Dim oldValues As Variant, newValues As Variant, row As Long, col As Long
  oldValues = before.ListArr(text, formats): newValues = after.ListArr(text, formats)
  If UBound(oldValues) <> UBound(newValues) Then Err.Raise 5, , "ListArr rows differ"
  If UBound(oldValues) >= 1 Then
    For row = 1 To UBound(oldValues, 1)
      For col = 1 To UBound(oldValues, 2)
        If oldValues(row, col) <> newValues(row, col) Then Err.Raise 5, , "ListArr values differ"
      Next
    Next
  End If
  oldValues = before.MatchAllArr(text): newValues = after.MatchAllArr(text)
  If UBound(oldValues) <> UBound(newValues) Then Err.Raise 5, , "MatchAllArr rows differ"
  If UBound(oldValues) >= 1 Then
    For row = 1 To UBound(oldValues, 1)
      For col = 0 To UBound(oldValues, 2)
        If oldValues(row, col) <> newValues(row, col) Then Err.Raise 5, , "Capture values differ"
      Next
    Next
  End If
End Sub

Public Function EdgeChecks() As Variant
  On Error GoTo Failed
  Dim patterns As Variant, inputs As Variant, flags As Variant, p As Variant, text As Variant, flag As Variant, i As Long
  patterns = Array("", "abc", "^abc", "(?m:^abc)", "(?-m:^abc)", "[0-9]+", "[^x]+", "a?b", "a|bc", "(?=ab)ab", _
    "(?<=ab)c", "(a)(b)?", "(?<x>a)|(?<x>b)", "(ab)\1", "(?i:AB)c", "\\x00", "a{0}", "(?:ab)*c", "\bword\b")
  inputs = Array("", "abc", "xabc", "ababc", "ab", "a", "b", "123 abc", "word word", vbCrLf & "abc", ChrW$(&H2028) & "abc", ChrW$(0) & "abc", ChrW$(&HFFFF))
  flags = Array("g", "gi", "gm", "gim", "gs")
  Dim before As RegexBaseline, after As RegexCandidate, formats As Variant, format As String
  format = "$&/$0/$1/$2/$x": formats = Array("$&", "$1")
  For Each p In patterns
    For Each flag In flags
      Set before = RegexBaseline.Create(CStr(p), CStr(flag))
      Set after = RegexCandidate.Create(CStr(p), CStr(flag))
      For Each text In inputs
        Dim value As String: value = CStr(text)
        Validate before, after, value, format, formats
        i = i + 1
      Next
    Next
  Next
  'Verify byte-buffer refresh after in-place mutation and when input lengths match.
  Dim regex As StaticRegex.RegexTy, state As StaticRegex.MatcherStateTy, inputText As String
  StaticRegex.InitializeRegex regex, "(abc)"
  inputText = "abc"
  If Not StaticRegex.Match(state, regex, inputText) Then Err.Raise 5, , "Initial mutation match failed"
  Mid$(inputText, 1, 1) = "z"
  If StaticRegex.Match(state, regex, inputText) Then Err.Raise 5, , "Stale input buffer after mutation"
  inputText = "abc"
  If Not StaticRegex.Match(state, regex, inputText) Then Err.Raise 5, , "Input buffer did not refresh"
  StaticRegex.InitializeRegex regex, "(x)(y)"
  inputText = "xy"
  If Not StaticRegex.Match(state, regex, inputText) Then Err.Raise 5, , "Capture resize failed"
  If StaticRegex.GetCapture(state, inputText, 2) <> "y" Then Err.Raise 5, , "Capture resize value failed"
  inputText = "zz"
  If StaticRegex.Match(state, regex, inputText) Then Err.Raise 5, , "Unexpected failed-match result"
  If state.captures.numberedCaptures(0).start <> 0 Then Err.Raise 5, , "Failed search did not clear captures"
  Dim budget As Long, budgetPattern As Variant
  For Each budgetPattern In Array("abc", "^abc", "[a-c]+", "needle(\d+)", "(abc)", "a|bc")
    For budget = 1 To 12
      Set before = RegexBaseline.Create(CStr(budgetPattern), "g", budget)
      Set after = RegexCandidate.Create(CStr(budgetPattern), "g", budget)
      For Each text In Array("", "abc", "xxxabc")
        value = CStr(text)
        If BudgetSignature(before, value) <> BudgetSignature(after, value) Then Err.Raise 5, , "Processing budget behavior changed"
        i = i + 1
      Next
    Next
  Next
  EdgeChecks = i + 6
  Exit Function
Failed:
  EdgeChecks = "Edge check " & i & ": " & Err.Number & ": " & Err.Description
End Function

Private Function BudgetSignature(ByVal rx As Object, ByRef text As String) As String
  On Error GoTo Failed
  BudgetSignature = CStr(rx.Test(text))
  Exit Function
Failed:
  BudgetSignature = "error=" & Err.Number
End Function
