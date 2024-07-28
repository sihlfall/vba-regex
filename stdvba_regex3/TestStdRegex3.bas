Attribute VB_Name = "TestStdRegex3"
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

Private Sub MakeStringArray(ByRef arr() As String, ParamArray p() As Variant)
    Dim j As Long, lb As Long, ub As Long
    lb = LBound(p)
    ub = UBound(p)
    ReDim arr(lb To ub) As String
    For j = lb To ub
        arr(j) = p(j)
    Next
End Sub


Private Sub SetPatternAndHaystack01(ByRef sPattern As String, ByRef sHaystack As String)
    sPattern = "(([A-Z])-(?:\d{2}-(\d[A-Z]{2})))"
    sHaystack = "D-22-4BU - London: London is the capital and largest city of England and the United Kingdom." & vbCrLf & _
                "D-48-8AO - Birmingham: Birmingham is a city and metropolitan borough in the West Midlands, England" & vbCrLf & _
                "A-22-9AO - Paris: Paris is the capital and most populous city of France. Also contains A-22-9AP."
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0000()
    ' Pattern property
    On Error GoTo TestFail
    
    Dim sPattern As String
    Dim sHaystack As String
    Dim rx As stdRegex3
    
    SetPatternAndHaystack01 sPattern, sHaystack
    Set rx = stdRegex3.Create(sPattern, "")
    Assert.AreEqual sPattern, rx.pattern
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0010()
    ' Flags property
    On Error GoTo TestFail
    
    Dim sPattern As String
    Dim sHaystack As String
    Dim rx As stdRegex3
    
    SetPatternAndHaystack01 sPattern, sHaystack
    Set rx = stdRegex3.Create(sPattern, "")
    Assert.AreEqual "", rx.flags
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0011()
    ' Flags property
    On Error GoTo TestFail
    
    Assert.AreEqual "", stdRegex3.Create("dummy", "").flags
    Assert.AreEqual "g", stdRegex3.Create("dummy", "G").flags
    Assert.AreEqual "i", stdRegex3.Create("dummy", "iii").flags
    Assert.AreEqual "m", stdRegex3.Create("dummy", "Mmmm").flags
    Assert.AreEqual "gi", stdRegex3.Create("dummy", "Ig").flags
    Assert.AreEqual "gm", stdRegex3.Create("dummy", "gm").flags
    Assert.AreEqual "im", stdRegex3.Create("dummy", "Mi").flags
    Assert.AreEqual "gim", stdRegex3.Create("dummy", "img").flags
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0020()
    ' Case-insensitive matching, matching pattern, method Test
    On Error GoTo TestFail
    
    Dim sPattern As String
    Dim sHaystack As String
    Dim rx As stdRegex3
    
    sPattern = "ABCdef"
    sHaystack = "abCDEF"
    Set rx = stdRegex3.Create(sPattern, "i")
    
    Assert.IsTrue rx.Test(sHaystack)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0021()
    ' Case-insensitive matching, non-matching pattern, method Test
    On Error GoTo TestFail
    
    Dim sPattern As String
    Dim sHaystack As String
    Dim rx As stdRegex3
    
    sPattern = "ABCdef"
    sHaystack = "abCDfe"
    Set rx = stdRegex3.Create(sPattern, "i")
    
    Assert.IsFalse rx.Test(sHaystack)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("stdRegex3")
Private Sub StdRegexThree0022()
    ' Case-insensitive matching, umlauts, matching pattern, method Test
    On Error GoTo TestFail
    
    Dim sPattern As String
    Dim sHaystack As String
    Dim rx As stdRegex3
    
    sPattern = "aÄäÖöÜüß"
    sHaystack = "AäÄöÖüÜß"
    Set rx = stdRegex3.Create(sPattern, "i")
    
    Assert.IsTrue rx.Test(sHaystack)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0023()
    ' Case-insensitive matching, umlauts, non-matching pattern, method Test
    On Error GoTo TestFail
    
    Dim sPattern As String
    Dim sHaystack As String
    Dim rx As stdRegex3
    
    sPattern = "aÄäÓöÜüß"
    sHaystack = "AäÄöÖüÜß"
    Set rx = stdRegex3.Create(sPattern, "i")
    
    Assert.IsFalse rx.Test(sHaystack)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0024()
    ' Case-insensitive matching, ranges, matching pattern, method Test
    On Error GoTo TestFail
    
    Dim sPattern As String
    Dim sHaystack As String
    Dim rx As stdRegex3
    
    sPattern = "^[A-Zäöü\d]+$"
    sHaystack = "ABCÄ239abcöö"
    Set rx = stdRegex3.Create(sPattern, "i")
    
    Assert.IsTrue rx.Test(sHaystack)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0025()
    ' Case-insensitive matching, ranges, non-matching pattern, method Test
    On Error GoTo TestFail
    
    Dim sPattern As String
    Dim sHaystack As String
    Dim rx As stdRegex3
    
    sPattern = "^[A-Zäóü\d]+$"
    sHaystack = "ABCÄ239abcöö"
    Set rx = stdRegex3.Create(sPattern, "i")
    
    Assert.IsFalse rx.Test(sHaystack)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0030()
    ' method Test, multiline matching
    On Error GoTo TestFail
    
    Dim sPattern As String
    Dim sHaystack As String
    
    SetPatternAndHaystack01 sPattern, sHaystack
    
    Assert.IsTrue stdRegex3.Create("^(?<Code>(?<Country>[A-Z])-(?:\d{2}-(\d[A-Z]{2})))", "m").Test(sHaystack)
    Assert.IsFalse stdRegex3.Create("^[A-Z]{3}", "m").Test(sHaystack)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0300()
    ' method Match, example from the doc comment
    
    Dim sPattern As String, sHaystack As String
    Dim rx As stdRegex3
    Dim oMatch As Object

    sPattern = "(?<id>\d{5}-ST[A-Z]\d) - (?<desc>.*)"
    sHaystack = "Some sites were in critical condition" & vbCrLf & _
        "* 12345-STA1 - Large crack through pipe." & vbCrLf & _
        "* 12323-STB9 - Acid leakage polluting watercourse." & vbCrLf & _
        "* and some others were largely ok:" & vbCrLf & _
        "* 23565-STC2" & vbCrLf & _
        "* 62346-STZ9"
        
    Set rx = stdRegex3.Create(sPattern)
    Set oMatch = rx.Match(sHaystack)
    
    Assert.IsTrue oMatch(0) = "12345-STA1 - Large crack through pipe."
    Assert.IsTrue oMatch("id") = "12345-STA1"
    Assert.IsTrue oMatch(1) = "12345-STA1"
    Assert.IsTrue oMatch("desc") = "Large crack through pipe."
    Assert.IsTrue oMatch(2) = "Large crack through pipe."
    Assert.IsTrue oMatch("$COUNT") = 2
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0310()
    ' method Match, multiline matching
    'On Error GoTo TestFail
    
    Dim sPattern As String, sHaystack As String
    Dim rx As stdRegex3
    Dim oMatch As Object
    
    sPattern = "^(?<Code>(?<Country>[A-Z])-(?:\d{2}-(\d[A-Z]{2})))"
    sHaystack = "D-22-4BU - London: London is the capital and largest city of England and the United Kingdom." & vbCrLf & _
                "D-48-8AO - Birmingham: Birmingham is a city and metropolitan borough in the West Midlands, England" & vbCrLf & _
                "A-22-9AO - Paris: Paris is the capital and most populous city of France. Also contains A-22-9AP."
    
    Set rx = stdRegex3.Create(sPattern, "m")
    
    'Match should return a dictionary containing the first match only
    Set oMatch = rx.Match(sHaystack)
    Assert.IsTrue TypeName(oMatch) = "Dictionary", "Match returns Dictionary"
    Assert.IsTrue oMatch(0) = "D-22-4BU", "Match Dictionary contains numbered captures 1"
    Assert.IsTrue oMatch(1) = "D-22-4BU", "Match Dictionary contains numbered captures 2"
    Assert.IsTrue oMatch(2) = "D", "Match Dictionary contains numbered captures 3"
    Assert.IsTrue oMatch(3) = "4BU", "Match Dictionary contains numbered captures 4 & ensure non-capturing group not captured"
    Assert.IsTrue oMatch("$COUNT") = 3, "Match contains count of submatches"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree400()
    ' method MatchAll, example from the doc comment
    
    Dim sPattern As String, sHaystack As String
    Dim rx As stdRegex3
    Dim oMatchAll As Collection

    sPattern = "\d{5}-ST[A-Z]\d"
    sHaystack = "Some sites were in critical condition" & vbCrLf & _
        "* 12345-STA1 - Large crack through pipe." & vbCrLf & _
        "* 12323-STB9 - Acid leakage polluting watercourse." & vbCrLf & _
        "* and some others were largely ok:" & vbCrLf & _
        "* 23565-STC2" & vbCrLf & _
        "* 62346-STZ9"
        
    Set rx = stdRegex3.Create(sPattern)
    Set oMatchAll = rx.MatchAll(sHaystack)
    
    Assert.IsTrue oMatchAll(1)(0) = "12345-STA1"
    Assert.IsTrue oMatchAll(2)(0) = "12323-STB9"
    Assert.IsTrue oMatchAll(3)(0) = "23565-STC2"
    Assert.IsTrue oMatchAll(4)(0) = "62346-STZ9"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0410()
    ' method MatchAll, multiline matching
    On Error GoTo TestFail
    
    Dim sPattern As String, sHaystack As String
    Dim rx As stdRegex3
    Dim oMatches As Object
    
    sPattern = "^(?<Code>(?<Country>[A-Z])-(?:\d{2}-(\d[A-Z]{2})))"
    sHaystack = "D-22-4BU - London: London is the capital and largest city of England and the United Kingdom." & vbCrLf & _
                "D-48-8AO - Birmingham: Birmingham is a city and metropolitan borough in the West Midlands, England" & vbCrLf & _
                "A-22-9AO - Paris: Paris is the capital and most populous city of France. Also contains A-22-9AP."
    
    Set rx = stdRegex3.Create(sPattern, "m")
    
    'Match should return a dictionary containing the first match only
    Set oMatches = rx.MatchAll(sHaystack)
    Assert.IsTrue TypeName(oMatches) = "Collection", "MatchAll returns Collection"
    Assert.IsTrue oMatches.Count = 3, "MatchAll contains all matches"
    Assert.IsTrue TypeName(oMatches(1)) = "Dictionary", "MatchAll contains Dictionaries"
    Assert.IsTrue oMatches(1)(0) = "D-22-4BU", "MatchAll dictionaries are populated 1"
    Assert.IsTrue oMatches(2)(0) = "D-48-8AO", "MatchAll dictionaries are populated 2"
    Assert.IsTrue oMatches(3)(0) = "A-22-9AO", "MatchAll dictionaries are populated 3"
    Assert.IsTrue oMatches(1)("Code") = "D-22-4BU", "MatchAll named capture exists 1"
    Assert.IsTrue oMatches(1)("Country") = "D", "MatchAll named capture exists 1"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0411()
    ' method MatchAll, correctly handles subsequent matches
    On Error GoTo TestFail
    
    Dim sPattern As String, sHaystack As String
    Dim rx As stdRegex3
    Dim oMatches As Object
    
    sPattern = "ABC"
    sHaystack = "abcABCaBC"
    
    Set rx = stdRegex3.Create(sPattern, "i")
    
    'Match should return a dictionary containing the first match only
    Set oMatches = rx.MatchAll(sHaystack)
    Assert.IsTrue oMatches(1)(0) = "abc", "Finds first occurrence"
    Assert.IsTrue oMatches(2)(0) = "ABC", "Finds second occurrence"
    Assert.IsTrue oMatches(3)(0) = "aBC", "Finds third occurrence"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree500()
    ' method Replace, example from the doc comment
    
    Dim sPattern As String, sHaystack As String, sReplacer As String, sResult As String
    Dim rx As stdRegex3

    sPattern = "(?<id>\d{5}-ST[A-Z]\d)\s+(?<count>\d+)\s+(?<date>../../....)"
    sHaystack = "Here is some cool data:" & vbCrLf & _
        "12345-STA1  123    10/02/2019" & vbCrLf & _
        "12323-STB9  2123   01/01/2005" & vbCrLf & _
        "and here is some more:" & vbCrLf & _
        "23565-STC2  23     ??/??/????" & vbCrLf & _
        "62346-STZ9  5      01/05/1932"
    sReplacer = "$<id>,$<date>,$<count>"
        
    Set rx = stdRegex3.Create(sPattern, "g")
    sResult = rx.Replace(sHaystack, sReplacer)
    
    Assert.IsTrue sResult = "Here is some cool data:" & vbCrLf & _
        "12345-STA1,10/02/2019,123" & vbCrLf & _
        "12323-STB9,01/01/2005,2123" & vbCrLf & _
        "and here is some more:" & vbCrLf & _
        "23565-STC2,??/??/????,23" & vbCrLf & _
        "62346-STZ9,01/05/1932,5"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0510()
    ' method Replace, simple exampleH
    'On Error GoTo TestFail
    
    Dim sPattern As String, sHaystack As String
    Dim rx As stdRegex3
    Dim oMatches As Object
    
    sPattern = "abc"
    sHaystack = "abc12abc34abcabc567abc"
    
    Set rx = stdRegex3.Create(sPattern, "g")
    
    'Match should return a dictionary containing the first match only
    Assert.AreEqual "zo12zo34zozo567zo", rx.Replace(sHaystack, "zo")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0511()
    ' method Replace, simple example with $& and $$
    On Error GoTo TestFail
    
    Dim sPattern As String, sHaystack As String
    Dim rx As stdRegex3
    Dim oMatches As Object
    
    sPattern = "a[bde]c"
    sHaystack = "abc12adc34aecabc567abc"
    
    Set rx = stdRegex3.Create(sPattern, "g")
    
    'Match should return a dictionary containing the first match only
    Assert.AreEqual "$abc$12$adc$34$aec$$abc$567$abc$", rx.Replace(sHaystack, "$$$&$$")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("stdRegex3")
Private Sub StdRegexThree0512()
    ' method Replace, simple example with $`
    On Error GoTo TestFail
    
    Dim sPattern As String, sHaystack As String
    Dim rx As stdRegex3
    Dim oMatches As Object
    
    sPattern = "\d"
    sHaystack = "1one2two3three"
    
    Set rx = stdRegex3.Create(sPattern, "g")
    
    'Match should return a dictionary containing the first match only
    Assert.AreEqual "one1onetwo1one2twothree", rx.Replace(sHaystack, "$`")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0513()
    ' method Replace, simple example with $'
    On Error GoTo TestFail
    
    Dim sPattern As String, sHaystack As String
    Dim rx As stdRegex3
    Dim oMatches As Object
    
    sPattern = "\d"
    sHaystack = "1one2two3three"
    
    Set rx = stdRegex3.Create(sPattern, "g")
    
    'Match should return a dictionary containing the first match only
    Assert.AreEqual "one2two3threeonetwo3threetwothreethree", rx.Replace(sHaystack, "$'")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0514()
    ' method Replace, simple example with numbered captures'
    'On Error GoTo TestFail
    
    Dim sPattern As String, sHaystack As String
    Dim rx As stdRegex3
    Dim oMatches As Object
    
    sPattern = "([A-Z]{2})-(\d{2})-(\d{2})"
    sHaystack = "DE-72-11 CH-99-10 US-40-44"
    
    Set rx = stdRegex3.Create(sPattern, "g")
    
    'Match should return a dictionary containing the first match only
    Assert.AreEqual "DE/11/72 CH/10/99 US/44/40", rx.Replace(sHaystack, "$1/$3/$2")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0515()
    ' method Replace, simple example with named captures'
    'On Error GoTo TestFail
    
    Dim sPattern As String, sHaystack As String, bytecode() As Long
    Dim rx As stdRegex3
    Dim oMatches As Object
    
    sPattern = "(?<country>[A-Z]{2})-(?<hcode>\d{2})-(?<lcode>\d{2})"
    sHaystack = "DE-72-11 CH-99-10 US-40-44"
    
    Set rx = stdRegex3.Create(sPattern, "g")
    
    'Match should return a dictionary containing the first match only
    Assert.AreEqual "DE/11/72 CH/10/99 US/44/40", rx.Replace(sHaystack, "$<country>/$<lcode>/$<hcode>")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree600()
    ' method List, example from the doc comment
    
    Dim sPattern As String, sHaystack As String, sFormat As String, sDelimiter As String, sResult As String
    Dim rx As stdRegex3

    sPattern = "(?<id>\d{5}-ST[A-Z]\d)\s+(?<count>\d+)\s+(?<date>../../....)"
    sHaystack = "Here is some cool data:" & vbCrLf & _
        "12345-STA1  123    10/02/2019" & vbCrLf & _
        "12323-STB9  2123   01/01/2005" & vbCrLf & _
        "and here is some more:" & vbCrLf & _
        "23565-STC2  23     ??/??/????" & vbCrLf & _
        "62346-STZ9  5      01/05/1932"
    sFormat = "$<id>,$<date>,$<count>"
    sDelimiter = ";"
    
    Set rx = stdRegex3.Create(sPattern, "g")
    sResult = rx.List(sHaystack, sFormat, sDelimiter)
    
    Assert.IsTrue sResult = "12345-STA1,10/02/2019,123;12323-STB9,01/01/2005,2123;23565-STC2,??/??/????,23;62346-STZ9,01/05/1932,5"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0610()
    ' method List, simple example with named captures'
    'On Error GoTo TestFail
    
    Dim sPattern As String, sHaystack As String, sFormat As String, sExpected As String, bytecode() As Long
    Dim rx As stdRegex3
    
    sPattern = "(?<id>\d{5}-ST[A-Z]\d)\s+(?<count>\d+)\s+(?<date>../../....)"
    sHaystack = "12345-STA1  123    10/02/2019" & vbCrLf & _
        "12323-STB9  2123   01/01/2005" & vbCrLf & _
        "23565-STC2  23     ??/??/????" & vbCrLf & _
        "62346-STZ9  5      01/05/1932"
    sFormat = "$<id>,$<date>,$<count>"
    
    Set rx = stdRegex3.Create(sPattern, "g")
    
    sExpected = "12345-STA1,10/02/2019,123" & vbCrLf & _
        "12323-STB9,01/01/2005,2123" & vbCrLf & _
        "23565-STC2,??/??/????,23" & vbCrLf & _
        "62346-STZ9,01/05/1932,5"
    Assert.IsTrue sExpected = rx.List(sHaystack, sFormat)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Function Make2dStringArray(ByVal nCols As Long, ParamArray strs() As Variant) As String()
    Dim nStrs As Long, nRows As Long, i As Long, j As Long, k As Long
    Dim res() As String
    
    nStrs = UBound(strs) - LBound(strs) + 1
    nRows = (nStrs + nCols - 1) \ nCols
    ReDim res(0 To nRows - 1, 0 To nCols - 1) As String
    
    k = LBound(strs)
    i = 0
    j = 0
    Do While k <= UBound(strs)
        res(i, j) = strs(k)
        j = j + 1: k = k + 1
        If j = nCols Then j = 0: i = i + 1
    Loop
    
    Make2dStringArray = res
End Function

Private Sub AssertEqual2dArrays(ByRef ary1() As String, ByRef ary2() As String)
    Dim nRows1 As Long, nCols1 As Long
    Dim nRows2 As Long, nCols2 As Long
    Dim i As Long, j As Long
    
    nRows1 = UBound(ary1, 1) - LBound(ary1, 1) + 1
    nCols1 = UBound(ary1, 2) - LBound(ary1, 2) + 1
    nRows2 = UBound(ary2, 1) - LBound(ary2, 1) + 1
    nCols2 = UBound(ary2, 2) - LBound(ary2, 2) + 1
    
    Assert.IsTrue nRows1 = nRows2, "First dimension not equal"
    Assert.IsTrue nCols1 = nCols2, "Second dimension not equal"
    For i = 0 To nRows1 - 1
        For j = 0 To nCols1 - 1
            Assert.IsTrue _
                ary1(LBound(ary1, 1) + i, LBound(ary1, 2) + j) = ary2(LBound(ary2, 1) + i, LBound(ary2, 2) + j), _
                "Mismatch for entry (" & i & "," & j & ")"
        Next
    Next
End Sub


'@TestMethod("stdRegex3")
Private Sub StdRegexThreez00()
    ' method ListArray, example from the doc comment
    
    Dim sPattern As String, sHaystack As String, formats() As String, result() As String
    Dim rx As stdRegex3

    sPattern = "(?<id>\d{5}-ST[A-Z]\d)\s+(?<count>\d+)\s+(?<date>../../....)"
    sHaystack = "Here is some cool data:" & vbCrLf & _
        "12345-STA1  123    10/02/2019" & vbCrLf & _
        "12323-STB9  2123   01/01/2005" & vbCrLf & _
        "and here is some more:" & vbCrLf & _
        "23565-STC2  23     ??/??/????" & vbCrLf & _
        "62346-STZ9  5      01/05/1932"
    MakeStringArray formats, "$<date>", "$<id>", "$<count> incidents"
    
    Set rx = stdRegex3.Create(sPattern, "g")
    rx.ListArray result, sHaystack, formats
    
    AssertEqual2dArrays result, Make2dStringArray(3, _
        "10/02/2019", "12345-STA1", "123 incidents", _
        "01/01/2005", "12323-STB9", "2123 incidents", _
        "??/??/????", "23565-STC2", "23 incidents", _
        "01/05/1932", "62346-STZ9", "5 incidents" _
    )
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0710()
    ' method ListArr, simple example with named captures'
    'On Error GoTo TestFail
    
    Dim sPattern As String, sHaystack As String, formats() As String, results() As String
    Dim rx As stdRegex3
    
    sPattern = "(?<id>\d{5}-ST[A-Z]\d)\s+(?<count>\d+)\s+(?<date>../../....)"
    sHaystack = "12345-STA1  123    10/02/2019" & vbCrLf & _
        "12323-STB9  2123   01/01/2005" & vbCrLf & _
        "23565-STC2  23     ??/??/????" & vbCrLf & _
        "62346-STZ9  5      01/05/1932"
    MakeStringArray formats, "$<id>,$<date>", "$<count>", "$&"
    
    Set rx = stdRegex3.Create(sPattern, "g")
    rx.ListArray results, sHaystack, formats
    
    Assert.AreEqual "12345-STA1,10/02/2019", results(0, 0)
    Assert.AreEqual "123", results(0, 1)
    Assert.AreEqual "12345-STA1  123    10/02/2019", results(0, 2)
    
    Assert.AreEqual "12323-STB9,01/01/2005", results(1, 0)
    Assert.AreEqual "2123", results(1, 1)
    Assert.AreEqual "12323-STB9  2123   01/01/2005", results(1, 2)

    Assert.AreEqual "23565-STC2,??/??/????", results(2, 0)
    Assert.AreEqual "23", results(2, 1)
    Assert.AreEqual "23565-STC2  23     ??/??/????", results(2, 2)

    Assert.AreEqual "62346-STZ9,01/05/1932", results(3, 0)
    Assert.AreEqual "5", results(3, 1)
    Assert.AreEqual "62346-STZ9  5      01/05/1932", results(3, 2)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdRegex3")
Private Sub StdRegexThree0711()
    ' method ListArr, zero matches case'
    'On Error GoTo TestFail
    
    Dim sPattern As String, sHaystack As String, formats() As String, results() As String
    Dim rx As stdRegex3
    
    sPattern = "(a)(b)cd"
    sHaystack = "xyz"
    MakeStringArray formats, "$1", "$2"
    
    Set rx = stdRegex3.Create(sPattern)
    rx.ListArray results, sHaystack, formats
    
    Assert.IsTrue UBound(results, 1) - LBound(results, 1) = -1
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


