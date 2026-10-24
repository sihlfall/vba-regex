Attribute VB_Name = "TestRxFunctions"
Option Explicit

Private checks As Long
Private currentCase As String
Private failures As String

Public Function RunRxTests() As Variant
    On Error GoTo Failed
    checks = 0: failures = vbNullString
    TestScalars
    TestNumericResults
    TestSpills
    TestSplitBroadcast
    TestArguments
    TestBroadcast
    TestArrayShapes
    TestWorksheetFormulas
    TestFormatAPI
    If Len(failures) > 0 Then RunRxTests = failures Else RunRxTests = checks
    Exit Function
Failed:
    RunRxTests = currentCase & ": " & Err.Description
End Function

Private Sub TestScalars()
    Check "test hit", RxTest("ab12cd34", "\d+"), True
    Check "test miss", RxTest("abc", "\d+"), False
    Check "match first", RxMatch("ab12cd34", "\d+"), "12"
    Check "match nth", RxMatch("ab12cd34", "\d+", , 2), "34"
    Check "match absent", RxMatch("abc", "\d+"), CVErr(xlErrNA)
    Check "nth absent", RxMatch("a1", "\d", , 2), CVErr(xlErrNA)
    Check "format captures", RxMatch("AB-12", "([A-Z]+)-(\d+)", "$2/$1"), "12/AB"
    Check "named capture", RxMatch("AB-12", "(?<id>[A-Z]+)-\d+", "$<id>"), "AB"
    Check "format empty", RxMatch("ab", "a", ""), ""
    Check "replace all", RxReplace("a12 b34", "\d+", "X"), "aX bX"
    Check "replace first", RxReplace("a12 b34", "\d+", "X", 1), "aX b34"
    Check "replace nth", RxReplace("a12 b34", "\d+", "X", 2), "a12 bX"
    Check "replace absent nth", RxReplace("a12", "\d+", "X", 2), "a12"
    Check "replace miss", RxReplace("abc", "\d+", "X"), "abc"
    Check "replace captures", RxReplace("a1 b2", "([a-z])(\d)", "$2$1"), "1a 2b"
    Check "case sensitive default", RxTest("ABC", "abc"), False
    Check "case flag", RxTest("ABC", "abc", "i"), True
    Check "match case flag", RxMatch("ABC", "abc", , , "i"), "ABC"
    Check "replace case flag", RxReplace("Aa", "a", "x", , "i"), "xx"
    Check "multiline off", RxTest("x" & vbLf & "a", "^a"), False
    Check "multiline on", RxTest("x" & vbLf & "a", "^a", "m"), True
    Check "match multiline", RxMatch("x" & vbLf & "a", "^a", , , "m"), "a"
    Check "replace multiline", RxReplace("a" & vbLf & "a", "^a", "x", , "m"), "x" & vbLf & "x"
    Check "dotAll off", RxTest("a" & vbLf & "b", "a.b"), False
    Check "dotAll on", RxTest("a" & vbLf & "b", "a.b", "s"), True
    Check "match dotAll", RxMatch("a" & vbLf & "b", "a.b", , , "s"), "a" & vbLf & "b"
    Check "replace dotAll", RxReplace("a" & vbLf & "b", "a.b", "x", , "s"), "x"
    Check "combined flags", RxTest("x" & vbLf & "A" & vbLf & "B", "^a.b$", "ims"), True
    Check "reordered repeated flags", RxTest("A" & vbLf & "B", "a.b", "sii"), True
    Check "test startAt", RxTest("a1b2", "\d", , 3), True
    Check "match startAt", RxMatch("a1b2", "\d", , , , 3), "2"
    Check "occurrence after startAt", RxMatch("1a2b3", "\d", , 2, , 3), "3"
    Check "anchor retains original text", RxTest("abc", "^b", , 2), False
    Check "word boundary retains prefix", RxTest("ab", "\bb", , 2), False
    Check "prefix format", RxMatch("ab12", "\d+", "$" & Chr$(96), , , 3), "ab"
    Check "suffix format", RxMatch("ab12cd", "\d+", "$'"), "cd"
    Check "replace preserves prefix", RxReplace("1a2b3", "\d", "X", , , 3), "1aXbX"
    Check "replace counts after startAt", RxReplace("1a2b3", "\d", "X", 2, , 3), "1a2bX"
    Check "empty match at end", RxMatch("ab", "$", , , , 3), ""
    Check "test terminal nonempty", RxTest("ab", ".", , 3), False
    Check "replace terminal match", RxReplace("ab", "$", "X", , , 3), "abX"
    Check "empty input", RxMatch("", "^$"), ""
    Check "replace zero length", RxReplace("ab", "", "-"), "-a-b-"
    Check "replace zero length nth", RxReplace("ab", "", "-", 2), "a-b"
    Check "literal dollar", RxReplace("a", "a", "$$"), "$"
    Check "max result length", RxReplace("a", "a", String$(32767, "x")), String$(32767, "x")
    Check "result length exceeded", RxReplace("aa", "a", String$(16384, "x")), CVErr(xlErrNum)
    Check "test execution limit", RxTest(String$(18, "a"), "(a+)+b"), CVErr(xlErrNum)
    Check "match execution limit", RxMatch(String$(18, "a"), "(a+)+b"), CVErr(xlErrNum)
    Check "replace execution limit", RxReplace(String$(18, "a"), "(a+)+b", "x"), CVErr(xlErrNum)
End Sub

Private Sub TestNumericResults()
    Dim invalid As Variant, emptyValue As Variant, surrogate As String, ws As Worksheet
    Dim inputs(1 To 2, 1 To 1) As Variant, patterns(1 To 1, 1 To 2) As Variant
    Dim starts(1 To 2, 1 To 1) As Variant, occurrences(1 To 2, 1 To 1) As Variant
    Dim flags(1 To 2, 1 To 1) As Variant
    Set ws = ThisWorkbook.Worksheets(1)

    Check "find first", RxFind("ab12cd345", "\d+"), 3&
    Check "find nth", RxFind("ab12cd345", "\d+", 2), 7&
    Check "find missing", RxFind("abc", "\d+"), CVErr(xlErrNA)
    Check "find nth missing", RxFind("a1", "\d", 2), CVErr(xlErrNA)
    Check "length first", RxLength("ab12cd345", "\d+"), 2&
    Check "length nth", RxLength("ab12cd345", "\d+", 2), 3&
    Check "length missing", RxLength("abc", "\d+"), CVErr(xlErrNA)
    Check "length nth missing", RxLength("a1", "\d", 2), CVErr(xlErrNA)
    Check "length entire match", RxLength("AB-123", "([A-Z]+)-(\d+)"), 6&
    Check "find absolute position after start", RxFind("ab12cd345", "\d+", , , 5), 7&
    Check "length after start", RxLength("ab12cd345", "\d+", , , 5), 3&
    Check "find counts after start", RxFind("1a22b333", "\d+", 2, , 3), 6&
    Check "length counts after start", RxLength("1a22b333", "\d+", 2, , 3), 3&
    Check "find anchors retain input", RxFind("ab", "^b", , , 2), CVErr(xlErrNA)
    Check "length boundary retains input", RxLength("ab", "\bb", , , 2), CVErr(xlErrNA)
    Check "find case flags", RxFind("xABC", "abc", , "i"), 2&
    Check "length case flags", RxLength("xABC", "abc", , "i"), 3&
    Check "find multiline", RxFind("x" & vbLf & "a", "^a", , "m"), 3&
    Check "length dotAll", RxLength("a" & vbLf & "b", "a.b", , "s"), 3&
    Check "find empty input", RxFind("", "^$"), 1&
    Check "length empty input", RxLength("", "^$"), 0&
    Check "find terminal empty", RxFind("ab", "$", , , 3), 3&
    Check "length terminal empty", RxLength("ab", "$", , , 3), 0&
    Check "find nth zero length", RxFind("ab", "", 2), 2&
    Check "length nth zero length", RxLength("ab", "", 2), 0&
    surrogate = ChrW$(&HD83D) & ChrW$(&HDE00)
    Check "find UTF16 position", RxFind(surrogate & "a", "a"), 3&
    Check "length UTF16 units", RxLength(surrogate & "a", surrogate), 2&

    Check "count matches", RxCount("ab12cd345", "\d+"), 2&
    Check "count missing", RxCount("abc", "\d"), 0&
    Check "count nonoverlapping", RxCount("banana", "ana"), 1&
    Check "count empty input", RxCount("", "."), 0&
    Check "count empty match", RxCount("", "^$"), 1&
    Check "count zero length", RxCount("ab", ""), 3&
    Check "count terminal empty", RxCount("ab", "$", , 3), 1&
    Check "count terminal nonempty", RxCount("ab", ".", , 3), 0&
    Check "count after start", RxCount("1a2b3", "\d", , 3), 2&
    Check "count from inside match", RxCount("aaa", "aa", , 2), 1&
    Check "count anchors retain input", RxCount("ab", "^b", , 2), 0&
    Check "count default case", RxCount("Aa", "a"), 1&
    Check "count case flag", RxCount("Aa", "a", "i"), 2&
    Check "count multiline", RxCount("a" & vbLf & "a", "^a", "m"), 2&
    Check "count dotAll", RxCount("a" & vbLf & "b", "a.b", "s"), 1&
    Check "count many without spilling", RxCount(String$(5000, "a"), "a"), 5000&
    Check "find empty default occurrence", RxFind("a1", "\d", emptyValue), 2&
    Check "length empty default start", RxLength("a1", "\d", , , emptyValue), 1&
    Check "count blank input", RxCount(emptyValue, "^$"), 1&

    For Each invalid In Array(0, -1, 1.5, "1", True, 4)
        Check "find invalid start " & CStr(invalid), RxFind("ab", ".", , , invalid), CVErr(xlErrValue)
        Check "length invalid start " & CStr(invalid), RxLength("ab", ".", , , invalid), CVErr(xlErrValue)
        Check "count invalid start " & CStr(invalid), RxCount("ab", ".", , invalid), CVErr(xlErrValue)
    Next
    For Each invalid In Array(0, -1, 1.5, "1", True)
        Check "find invalid occurrence " & CStr(invalid), RxFind("ab", ".", invalid), CVErr(xlErrValue)
        Check "length invalid occurrence " & CStr(invalid), RxLength("ab", ".", invalid), CVErr(xlErrValue)
    Next
    Check "find bad pattern", RxFind("a", "["), CVErr(xlErrValue)
    Check "length bad pattern", RxLength("a", "["), CVErr(xlErrValue)
    Check "count bad pattern", RxCount("a", "["), CVErr(xlErrValue)
    Check "find bad flags", RxFind("a", ".", , "x"), CVErr(xlErrValue)
    Check "length bad flags", RxLength("a", ".", , "x"), CVErr(xlErrValue)
    Check "count bad flags", RxCount("a", ".", "x"), CVErr(xlErrValue)
    Check "find nontext input", RxFind(123, "."), CVErr(xlErrValue)
    Check "length nontext input", RxLength(123, "."), CVErr(xlErrValue)
    Check "count nontext input", RxCount(123, "."), CVErr(xlErrValue)
    Check "find input error", RxFind(CVErr(xlErrRef), "."), CVErr(xlErrRef)
    Check "length pattern error", RxLength("a", CVErr(xlErrNA)), CVErr(xlErrNA)
    Check "count input error", RxCount(CVErr(xlErrDiv0), "."), CVErr(xlErrDiv0)
    Check "find option error", RxFind("a", ".", CVErr(xlErrRef)), CVErr(xlErrRef)
    Check "length option error", RxLength("a", ".", , CVErr(xlErrNA)), CVErr(xlErrNA)
    Check "count option error", RxCount("a", ".", , CVErr(xlErrDiv0)), CVErr(xlErrDiv0)
    Check "find execution limit", RxFind(String$(18, "a"), "(a+)+b"), CVErr(xlErrNum)
    Check "length execution limit", RxLength(String$(18, "a"), "(a+)+b"), CVErr(xlErrNum)
    Check "count execution limit", RxCount(String$(18, "a"), "(a+)+b"), CVErr(xlErrNum)
    Check "count fails instead of partial count", RxCount("b" & String$(18, "a"), "b|(a+)+z"), CVErr(xlErrNum)
    Check "find stops after selected match", RxFind("b" & String$(18, "a"), "b|(a+)+z"), 1&
    Check "length stops after selected match", RxLength("b" & String$(18, "a"), "b|(a+)+z"), 1&

    inputs(1, 1) = "ab1cd22": inputs(2, 1) = "z333"
    patterns(1, 1) = "[a-z]+": patterns(1, 2) = "\d+"
    Grid "find cross broadcast", RxFind(inputs, patterns), 2, 2, Array(1&, 3&, 1&, 2&)
    Grid "length cross broadcast", RxLength(inputs, patterns), 2, 2, Array(2&, 1&, 1&, 3&)
    Grid "count cross broadcast", RxCount(inputs, patterns), 2, 2, Array(2&, 2&, 1&, 1&)
    ws.Range("A60:A61").Value2 = inputs
    ws.Range("C60:D60").Value2 = patterns
    Grid "find ranges", RxFind(ws.Range("A60:A61"), ws.Range("C60:D60")), 2, 2, Array(1&, 3&, 1&, 2&)
    Grid "length ranges", RxLength(ws.Range("A60:A61"), ws.Range("C60:D60")), 2, 2, Array(2&, 1&, 1&, 3&)
    Grid "count ranges", RxCount(ws.Range("A60:A61"), ws.Range("C60:D60")), 2, 2, Array(2&, 2&, 1&, 1&)
    occurrences(1, 1) = 2: occurrences(2, 1) = 1
    starts(1, 1) = 1: starts(2, 1) = 3
    Grid "find broadcast options", RxFind("1a22b333", "\d+", occurrences, , starts), 2, 1, Array(3&, 3&)
    Grid "length broadcast options", RxLength("1a22b333", "\d+", occurrences, , starts), 2, 1, Array(2&, 2&)
    Grid "count broadcast start", RxCount("1a22b333", "\d+", , starts), 2, 1, Array(3&, 2&)
    flags(1, 1) = "": flags(2, 1) = "i"
    Grid "find broadcast flags", RxFind("A", "a", , flags), 2, 1, Array(CVErr(xlErrNA), 1&)
    Grid "length broadcast flags", RxLength("A", "a", , flags), 2, 1, Array(CVErr(xlErrNA), 1&)
    Grid "count broadcast flags", RxCount("A", "a", flags), 2, 1, Array(0&, 1&)
    inputs(1, 1) = CVErr(xlErrRef)
    Grid "find local error", RxFind(inputs, "\d+"), 2, 1, Array(CVErr(xlErrRef), 2&)
    Grid "length local error", RxLength(inputs, "\d+"), 2, 1, Array(CVErr(xlErrRef), 3&)
    Grid "count local error", RxCount(inputs, "\d+"), 2, 1, Array(CVErr(xlErrRef), 1&)
    Grid "find literal row", RxFind(Array("ab1", "c22"), "\d+"), 1, 2, Array(3&, 2&)
    Grid "length literal row", RxLength(Array("ab1", "c22"), "\d+"), 1, 2, Array(1&, 2&)
    Grid "count literal row", RxCount(Array("a1b2", "c33"), "\d+"), 1, 2, Array(2&, 1&)
End Sub

Private Sub TestSpills()
    Dim formats(1 To 1, 1 To 2) As Variant, column(1 To 2, 1 To 1) As Variant, ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    Grid "all matches", RxMatches("a12 b34", "\d+"), 2, 1, Array("12", "34")
    Grid "all formatted", RxMatches("a12 b34", "\d+", "[$&]"), 2, 1, Array("[12]", "[34]")
    Grid "all startAt", RxMatches("1a2b3", "\d", , , 3), 2, 1, Array("2", "3")
    Grid "all flags", RxMatches("A" & vbLf & "B", "^.$", , "im"), 2, 1, Array("A", "B")
    Grid "all dotAll", RxMatches("A" & vbLf & "B", "a.b", , "is"), 1, 1, Array("A" & vbLf & "B")
    Check "all missing", RxMatches("abc", "\d"), CVErr(xlErrNA)
    Grid "all empty at end", RxMatches("ab", "$", , , 3), 1, 1, Array("")
    Grid "all zero length", RxMatches("ab", ""), 3, 1, Array("", "", "")
    Grid "all empty input", RxMatches("", ""), 1, 1, Array("")
    formats(1, 1) = "$1": formats(1, 2) = "$2"
    Grid "capture table", RxMatches("AB-12 CD-34", "([A-Z]+)-(\d+)", formats), 2, 2, Array("AB", "12", "CD", "34")
    ws.Range("A10:B10").NumberFormat = "@"
    ws.Range("A10:B10").Value2 = formats
    Grid "capture format range", RxMatches("AB-12 CD-34", "([A-Z]+)-(\d+)", ws.Range("A10:B10")), 2, 2, Array("AB", "12", "CD", "34")
    Grid "split fields", RxSplit("red, green;blue", "[,;]\s*"), 1, 3, Array("red", "green", "blue")
    Grid "split preserves empties", RxSplit(",a,,b,", ","), 1, 5, Array("", "a", "", "b", "")
    Grid "split ignores empties", RxSplit(",a,,b,", ",", True), 1, 2, Array("a", "b")
    Grid "split excludes captures", RxSplit("a,b;c", "([,;])"), 1, 3, Array("a", "b", "c")
    Grid "split no separator", RxSplit("abc", ","), 1, 1, Array("abc")
    Grid "split empty input", RxSplit("", ","), 1, 1, Array("")
    Check "split all ignored", RxSplit(",,", ",", True), CVErr(xlErrNA)
    Check "split empty ignored", RxSplit("", ",", True), CVErr(xlErrNA)
    Grid "split zero length", RxSplit("ab", ""), 1, 4, Array("", "a", "b", "")
    Grid "split zero length ignored", RxSplit("ab", "", True), 1, 2, Array("a", "b")
    Grid "split flags", RxSplit("aXbxc", "x", , "i"), 1, 3, Array("a", "b", "c")
    Grid "split multiline", RxSplit("a" & vbLf & "b", "^", True, "m"), 1, 2, Array("a" & vbLf, "b")
    Grid "split dotAll", RxSplit("a" & vbLf & "b", "a.b", , "s"), 1, 2, Array("", "")
    Check "matches execution limit", RxMatches(String$(18, "a"), "(a+)+b"), CVErr(xlErrNum)
    Check "split execution limit", RxSplit(String$(18, "a"), "(a+)+b"), CVErr(xlErrNum)
    Check "split column limit", RxSplit(String$(16384, ","), ","), CVErr(xlErrNum)
    column(1, 1) = "a": column(2, 1) = "b"
    Check "matches rejects texts", RxMatches(column, "."), CVErr(xlErrValue)
    Check "matches rejects patterns", RxMatches("ab", column), CVErr(xlErrValue)
    Check "matches rejects format column", RxMatches("ab", ".", column), CVErr(xlErrValue)
    Grid "split text column", RxSplit(column, ","), 2, 1, Array("a", "b")
    formats(1, 1) = CVErr(xlErrRef)
    Check "format table error", RxMatches("a", ".", formats), CVErr(xlErrRef)
End Sub

Private Sub TestSplitBroadcast()
    Dim inputs(1 To 3, 1 To 1) As Variant, patterns(1 To 3, 1 To 1) As Variant
    Dim ignore(1 To 3, 1 To 1) As Variant, flags(1 To 3, 1 To 1) As Variant
    Dim short(1 To 2, 1 To 1) As Variant, zero(0 To 1, 0 To 0) As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    inputs(1, 1) = "a,b,c": inputs(2, 1) = "d": inputs(3, 1) = "e,f"
    Grid "split column pads shorter rows", RxSplit(inputs, ","), 3, 3, _
        Array("a", "b", "c", "d", CVErr(xlErrNA), CVErr(xlErrNA), "e", "f", CVErr(xlErrNA))
    ws.Range("A70:A72").Value2 = inputs
    Grid "split range column pads shorter rows", RxSplit(ws.Range("A70:A72"), ","), 3, 3, _
        Array("a", "b", "c", "d", CVErr(xlErrNA), CVErr(xlErrNA), "e", "f", CVErr(xlErrNA))
    patterns(1, 1) = ",": patterns(2, 1) = ";": patterns(3, 1) = ":"
    inputs(1, 1) = "a,b": inputs(2, 1) = "c;d;e": inputs(3, 1) = "f"
    Grid "split per row patterns", RxSplit(inputs, patterns), 3, 3, _
        Array("a", "b", CVErr(xlErrNA), "c", "d", "e", "f", CVErr(xlErrNA), CVErr(xlErrNA))
    ws.Range("C70:C72").Value2 = patterns
    Grid "split scalar text over pattern column", RxSplit("a,b;c:d", ws.Range("C70:C72")), 3, 2, _
        Array("a", "b;c:d", "a,b", "c:d", "a,b;c", "d")
    ignore(1, 1) = False: ignore(2, 1) = True: ignore(3, 1) = False
    Grid "split per row ignoreEmpty", RxSplit(",a,,", ",", ignore), 3, 4, _
        Array("", "a", "", "", "a", CVErr(xlErrNA), CVErr(xlErrNA), CVErr(xlErrNA), "", "a", "", "")
    flags(1, 1) = "": flags(2, 1) = "i": flags(3, 1) = ""
    Grid "split per row flags", RxSplit("aXbxc", "x", , flags), 3, 3, _
        Array("aXb", "c", CVErr(xlErrNA), "a", "b", "c", "aXb", "c", CVErr(xlErrNA))
    flags(1, 1) = "s": flags(2, 1) = "": flags(3, 1) = "s"
    Grid "split per row dotAll", RxSplit("a" & vbLf & "b", "a.b", , flags), 3, 2, _
        Array("", "", "a" & vbLf & "b", CVErr(xlErrNA), "", "")
    flags(1, 1) = "m": flags(2, 1) = "": flags(3, 1) = "m"
    Grid "split per row multiline", RxSplit("a" & vbLf & "b", "^", True, flags), 3, 2, _
        Array("a" & vbLf, "b", "a" & vbLf & "b", CVErr(xlErrNA), "a" & vbLf, "b")
    inputs(1, 1) = "a,,": inputs(2, 1) = "": inputs(3, 1) = "b"
    Grid "split keeps empty fields distinct from padding", RxSplit(inputs, ","), 3, 3, _
        Array("a", "", "", "", CVErr(xlErrNA), CVErr(xlErrNA), "b", CVErr(xlErrNA), CVErr(xlErrNA))
    Grid "split removes empty fields per row", RxSplit(inputs, ",", True), 3, 1, Array("a", CVErr(xlErrNA), "b")
    inputs(1, 1) = ",,": inputs(2, 1) = ",a,b": inputs(3, 1) = ""
    Grid "split empty rows preserve source alignment", RxSplit(inputs, ",", True), 3, 2, _
        Array(CVErr(xlErrNA), CVErr(xlErrNA), "a", "b", CVErr(xlErrNA), CVErr(xlErrNA))
    inputs(1, 1) = "ab": inputs(2, 1) = "c": inputs(3, 1) = ""
    Grid "split zero length column", RxSplit(inputs, ""), 3, 4, _
        Array("", "a", "b", "", "", "c", "", CVErr(xlErrNA), "", "", CVErr(xlErrNA), CVErr(xlErrNA))
    inputs(1, 1) = CVErr(xlErrRef): inputs(2, 1) = "a,b": inputs(3, 1) = "c"
    Grid "split row input error", RxSplit(inputs, ","), 3, 2, _
        Array(CVErr(xlErrRef), CVErr(xlErrRef), "a", "b", "c", CVErr(xlErrNA))
    patterns(1, 1) = "[": patterns(2, 1) = ",": patterns(3, 1) = ","
    Grid "split row pattern error", RxSplit("a,b", patterns), 3, 2, _
        Array(CVErr(xlErrValue), CVErr(xlErrValue), "a", "b", "a", "b")
    patterns(1, 1) = CVErr(xlErrDiv0)
    Grid "split row pattern cell error", RxSplit("a,b", patterns), 3, 2, _
        Array(CVErr(xlErrDiv0), CVErr(xlErrDiv0), "a", "b", "a", "b")
    flags(1, 1) = "x": flags(2, 1) = "": flags(3, 1) = ""
    Grid "split row option error", RxSplit("a,b", ",", , flags), 3, 2, _
        Array(CVErr(xlErrValue), CVErr(xlErrValue), "a", "b", "a", "b")
    ignore(1, 1) = "true"
    Grid "split row boolean error", RxSplit("a,b", ",", ignore), 3, 2, _
        Array(CVErr(xlErrValue), CVErr(xlErrValue), "a", "b", "a", "b")
    inputs(1, 1) = String$(18, "a"): inputs(2, 1) = "b": inputs(3, 1) = "c"
    Grid "split row execution limit", RxSplit(inputs, "(a+)+b"), 3, 1, Array(CVErr(xlErrNum), "b", "c")
    inputs(1, 1) = String$(16384, ","): inputs(2, 1) = "a,b": inputs(3, 1) = "c"
    Grid "split row column limit", RxSplit(inputs, ","), 3, 2, _
        Array(CVErr(xlErrNum), CVErr(xlErrNum), "a", "b", "c", CVErr(xlErrNA))
    inputs(1, 1) = CVErr(xlErrRef): inputs(2, 1) = 1: inputs(3, 1) = CVErr(xlErrDiv0)
    Grid "split all error rows", RxSplit(inputs, ","), 3, 1, Array(CVErr(xlErrRef), CVErr(xlErrValue), CVErr(xlErrDiv0))
    short(1, 1) = ",": short(2, 1) = ","
    Check "split incompatible rows", RxSplit(inputs, short), CVErr(xlErrValue)
    Check "split incompatible rows reversed", RxSplit(short, inputs), CVErr(xlErrValue)
    Check "split rejects text row", RxSplit(Array("a,b", "c,d"), ","), CVErr(xlErrValue)
    Check "split rejects pattern row", RxSplit("a,b", Array(",", ";")), CVErr(xlErrValue)
    Check "split rejects flags row", RxSplit("a,b", ",", , Array("", "i")), CVErr(xlErrValue)
    Check "split rejects multiple areas", RxSplit(Union(ws.Range("A70"), ws.Range("A72")), ","), CVErr(xlErrValue)
    zero(0, 0) = "a,b": zero(1, 0) = "c"
    Grid "split zero based column", RxSplit(zero, ","), 2, 2, Array("a", "b", "c", CVErr(xlErrNA))
End Sub

Private Sub TestArguments()
    Dim invalid As Variant, emptyValue As Variant, ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    For Each invalid In Array(0, -1, 1.5, "1", True, 4, 1E+20)
        Check "invalid start " & CStr(invalid), RxTest("ab", ".", , invalid), CVErr(xlErrValue)
    Next
    For Each invalid In Array(0, -1, 1.5, "1", True)
        Check "invalid occurrence " & CStr(invalid), RxMatch("ab", ".", , invalid), CVErr(xlErrValue)
    Next
    Check "invalid replace occurrence", RxReplace("ab", ".", "x", -1), CVErr(xlErrValue)
    Check "invalid flags", RxTest("a", ".", "x"), CVErr(xlErrValue)
    Check "uppercase flag", RxTest("a", ".", "I"), CVErr(xlErrValue)
    Check "nontext input", RxTest(123, "."), CVErr(xlErrValue)
    Check "nontext pattern", RxMatch("a", 1), CVErr(xlErrValue)
    Check "nontext replacement", RxReplace("a", ".", 1), CVErr(xlErrValue)
    Check "nontext format", RxMatches("a", ".", 1), CVErr(xlErrValue)
    Check "invalid split boolean", RxSplit("a", ".", "true"), CVErr(xlErrValue)
    Check "test bad pattern", RxTest("a", "["), CVErr(xlErrValue)
    Check "match bad pattern", RxMatch("a", "["), CVErr(xlErrValue)
    Check "replace bad pattern", RxReplace("a", "[", "x"), CVErr(xlErrValue)
    Check "matches bad pattern", RxMatches("a", "["), CVErr(xlErrValue)
    Check "split bad pattern", RxSplit("a", "["), CVErr(xlErrValue)
    Check "bad format before search", RxMatch("a", "z", "$"), CVErr(xlErrValue)
    Check "bad replacement before search", RxReplace("a", "z", "$"), CVErr(xlErrValue)
    Check "bad spill format before search", RxMatches("a", "z", "$"), CVErr(xlErrValue)
    Check "text error", RxTest(CVErr(xlErrNA), "."), CVErr(xlErrNA)
    Check "pattern error", RxMatch("a", CVErr(xlErrRef)), CVErr(xlErrRef)
    Check "format error", RxMatch("a", ".", CVErr(xlErrDiv0)), CVErr(xlErrDiv0)
    Check "occurrence error", RxReplace("a", ".", "x", CVErr(xlErrRef)), CVErr(xlErrRef)
    Check "flags error", RxTest("a", ".", CVErr(xlErrDiv0)), CVErr(xlErrDiv0)
    Check "start error", RxTest("a", ".", , CVErr(xlErrNA)), CVErr(xlErrNA)
    Check "matches error", RxMatches(CVErr(xlErrRef), "."), CVErr(xlErrRef)
    Check "split error", RxSplit("a", ".", CVErr(xlErrNA)), CVErr(xlErrNA)
    Check "empty variant text", RxTest(emptyValue, "^$"), True
    Check "empty position defaults", RxTest("a", ".", , emptyValue), True
    ws.Range("A20").ClearContents
    Check "blank range", RxMatch(ws.Range("A20"), "^$"), ""
    ws.Range("A20").Value2 = "ab"
    Grid "matches scalar range", RxMatches(ws.Range("A20"), "."), 2, 1, Array("a", "b")
    Grid "split scalar range", RxSplit(ws.Range("A20"), "b"), 1, 2, Array("a", "")
    Check "multi area range", RxTest(Union(ws.Range("A20"), ws.Range("C20")), "."), CVErr(xlErrValue)
End Sub

Private Sub TestBroadcast()
    Dim a(1 To 2, 1 To 1) As Variant, p(1 To 1, 1 To 2) As Variant, opts(1 To 2, 1 To 1) As Variant
    Dim start(1 To 2, 1 To 1) As Variant, nth(1 To 2, 1 To 1) As Variant
    Dim fmts(1 To 1, 1 To 2) As Variant, ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    a(1, 1) = "a1": a(2, 1) = "b2": p(1, 1) = "[a-z]": p(1, 2) = "\d"
    Grid "broadcast test", RxTest(a, p), 2, 2, Array(True, True, True, True)
    Grid "broadcast match", RxMatch(a, p), 2, 2, Array("a", "1", "b", "2")
    Grid "broadcast replacement", RxReplace(a, p, "X"), 2, 2, Array("X1", "aX", "X2", "bX")
    ws.Range("A30:A31").Value2 = a: ws.Range("C30:D30").Value2 = p
    Grid "range match", RxMatch(ws.Range("A30:A31"), ws.Range("C30:D30")), 2, 2, Array("a", "1", "b", "2")
    Grid "mixed replace", RxReplace(ws.Range("A30:A31"), p, "X"), 2, 2, Array("X1", "aX", "X2", "bX")
    opts(1, 1) = "": opts(2, 1) = "i"
    Grid "flags compilation", RxTest("A", "a", opts), 2, 1, Array(False, True)
    Grid "flags match", RxMatch("A", "a", , , opts), 2, 1, Array(CVErr(xlErrNA), "A")
    Grid "flags replacement", RxReplace("A", "a", "x", , opts), 2, 1, Array("A", "x")
    start(1, 1) = 1: start(2, 1) = 3: nth(1, 1) = 2: nth(2, 1) = 1
    Grid "start positions", RxMatch("1a2b3", "\d", , , , start), 2, 1, Array("1", "2")
    Grid "occurrence positions", RxMatch("1a2b3", "\d", , nth, , start), 2, 1, Array("2", "2")
    Grid "replacement positions", RxReplace("1a2b3", "\d", "X", nth, , start), 2, 1, Array("1aXb3", "1aXb3")
    fmts(1, 1) = "[$&]": fmts(1, 2) = "<$&>"
    Grid "format cross broadcast", RxMatch(a, "\d", fmts), 2, 2, Array("[1]", "<1>", "[2]", "<2>")
    Grid "replacement cross broadcast", RxReplace(a, "\d", fmts), 2, 2, Array("a[1]", "a<1>", "b[2]", "b<2>")
    a(1, 1) = CVErr(xlErrRef)
    Grid "local error", RxMatch(a, "\d"), 2, 1, Array(CVErr(xlErrRef), "2")
    p(1, 1) = "[": p(1, 2) = "\d"
    Grid "compile failure local", RxMatch("a1", p), 1, 2, Array(CVErr(xlErrValue), "1")
    opts(1, 1) = "x": opts(2, 1) = ""
    Grid "option failure local", RxTest("a", ".", opts), 2, 1, Array(CVErr(xlErrValue), True)

End Sub

Private Sub TestArrayShapes()
    Dim zero(0 To 1, 0 To 0) As Variant, cube(1 To 1, 1 To 1, 1 To 1) As Variant
    Dim uninitialized() As Variant, flags(1 To 2, 1 To 1) As Variant
    Dim short(1 To 2, 1 To 1) As Variant, tall(1 To 3, 1 To 1) As Variant
    Dim wide(1 To 1, 1 To 3) As Variant, patternRow(1 To 1, 1 To 2) As Variant
    Dim surrogate As String
    zero(0, 0) = "a1": zero(1, 0) = "b2"
    Grid "zero based 2D", RxMatch(zero, "\d"), 2, 1, Array("1", "2")
    Grid "one dimensional row", RxMatch(Array("a1", "b2"), "\d"), 1, 2, Array("1", "2")
    Grid "one dimensional formats", RxMatches("AB-12", "([A-Z]+)-(\d+)", Array("$1", "$2")), 1, 2, Array("AB", "12")
    Check "reject 3D array", RxTest(cube, "."), CVErr(xlErrValue)
    Check "reject uninitialized array", RxTest(uninitialized, "."), CVErr(xlErrValue)
    short(1, 1) = "a": short(2, 1) = "a"
    tall(1, 1) = "a": tall(2, 1) = "a": tall(3, 1) = "a"
    Check "incompatible rows descending", RxTest(tall, short), CVErr(xlErrValue)
    Check "incompatible rows ascending", RxTest(short, tall), CVErr(xlErrValue)
    flags(1, 1) = "": flags(2, 1) = "i"
    Check "incompatible option rows", RxReplace(tall, "a", "x", , flags), CVErr(xlErrValue)
    wide(1, 1) = "a": wide(1, 2) = "a": wide(1, 3) = "a"
    patternRow(1, 1) = "a": patternRow(1, 2) = "a"
    Check "incompatible columns", RxMatch(wide, patternRow), CVErr(xlErrValue)
    surrogate = ChrW$(&HD83D) & ChrW$(&HDE00) & "a"
    Check "UTF16 startAt", RxMatch(surrogate, "a", , , , 3), "a"
End Sub

Private Sub TestWorksheetFormulas()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    ws.Range("J1:Z20").ClearContents
    ws.Range("J1").Formula2 = "=RxTest(""ABC"",""abc"",""i"")"
    ws.Range("J2").Formula2 = "=RxMatch(""a1b2"",""\d"",,2)"
    ws.Range("J3").Formula2 = "=RxReplace(""a1b2"",""\d"",""X"",1)"
    ws.Range("L1").Formula2 = "=RxMatches(""AB-12 CD-34"",""([A-Z]+)-(\d+)"",{""$1"",""$2""})"
    ws.Range("P1").Formula2 = "=RxSplit(""red, green;blue"",""[,;]\s*"")"
    ws.Range("J5").Formula2 = "=IFNA(RxMatch(""a"",""\d""),""missing"")"
    ws.Range("J6").Formula2 = "=RxMatch(""1a2b3"",""\d"",,2,,3)"
    ws.Range("L5").Formula2 = "=RxMatch({""a1"";""b2""},""\d"")"
    ws.Range("P5").Formula2 = "=RxMatch({""a1"",""b2""},""\d"")"
    ws.Range("L9").Formula2 = "=RxTest(""A"",""a"",{"""",""i""})"
    ws.Range("P9").Formula2 = "=RxMatch(""1a2b3"",""\d"",,{1,2})"
    ws.Range("L13").Formula2 = "=RxMatch(""1a2b3"",""\d"",,,,3)"
    ws.Range("AA1:AJ10").ClearContents
    ws.Range("AA1").Formula2 = "=RxSplit(A70:A72,"","")"
    ws.Range("AE1").Formula2 = "=RxSplit({""a,b,c"";""d"";""e,f""},"","")"
    ws.Range("AI1").Formula2 = "=RxSplit(""aXbxc"",""x"",,{"""";""i""})"
    ws.Range("J15").Formula2 = "=RxFind(""ab12cd345"",""\d+"",2)"
    ws.Range("K15").Formula2 = "=RxLength(""ab12cd345"",""\d+"",2)"
    ws.Range("L15").Formula2 = "=RxCount(""ab12cd345"",""\d+"")"
    ws.Range("N15").Formula2 = "=MID(""ab12cd345"",RxFind(""ab12cd345"",""\d+"",2),RxLength(""ab12cd345"",""\d+"",2))"
    ws.Range("P15").Formula2 = "=RxCount({""a1b2"",""c33""},""\d+"")"
    ws.Range("J17").Formula2 = "=RxFind(""1a22b333"",""\d+"",2,,3)"
    ws.Range("K17").Formula2 = "=RxLength(""1a22b333"",""\d+"",2,,3)"
    ws.Range("L17").Formula2 = "=RxCount(""1a22b333"",""\d+"",,3)"
    ws.Range("N17").Formula2 = "=RxCount(""abc"",""\d+"")"
    ws.Range("P17").Formula2 = "=IFNA(RxFind(""abc"",""\d+""),""missing"")"
    ws.Range("R17").Formula2 = "=RxLength(""ab"",""$"")"
    ws.Calculate
    currentCase = "literal column": AssertValue ws.Range("L5").Value2, "1": AssertValue ws.Range("L6").Value2, "2"
    currentCase = "literal row": AssertValue ws.Range("P5").Value2, "1": AssertValue ws.Range("Q5").Value2, "2"
    currentCase = "literal flags": AssertValue ws.Range("L9").Value2, False: AssertValue ws.Range("M9").Value2, True
    currentCase = "literal occurrences": AssertValue ws.Range("P9").Value2, "1": AssertValue ws.Range("Q9").Value2, "2"
    currentCase = "startAt omitted options": AssertValue ws.Range("L13").Value2, "2"
    currentCase = "worksheet numeric functions"
    AssertValue ws.Range("J15").Value2, 7#: AssertValue ws.Range("K15").Value2, 3#: AssertValue ws.Range("L15").Value2, 2#
    currentCase = "worksheet MID composition": AssertValue ws.Range("N15").Value2, "345"
    currentCase = "worksheet count broadcast": AssertValue ws.Range("P15").Value2, 2#: AssertValue ws.Range("Q15").Value2, 1#
    currentCase = "worksheet numeric startAt"
    AssertValue ws.Range("J17").Value2, 6#: AssertValue ws.Range("K17").Value2, 3#: AssertValue ws.Range("L17").Value2, 2#
    currentCase = "worksheet count absent": AssertValue ws.Range("N17").Value2, 0#
    currentCase = "worksheet find absent": AssertValue ws.Range("P17").Value2, "missing"
    currentCase = "worksheet length empty": AssertValue ws.Range("R17").Value2, 0#
    currentCase = "worksheet split column range"
    AssertValue ws.Range("AA1").Value2, "a": AssertValue ws.Range("AC1").Value2, "c"
    AssertValue ws.Range("AA2").Value2, "d": AssertValue ws.Range("AB2").Value2, CVErr(xlErrNA)
    AssertValue ws.Range("AA3").Value2, "e": AssertValue ws.Range("AB3").Value2, "f"
    currentCase = "worksheet split column literal"
    AssertValue ws.Range("AE1").Value2, "a": AssertValue ws.Range("AG1").Value2, "c"
    AssertValue ws.Range("AE2").Value2, "d": AssertValue ws.Range("AF2").Value2, CVErr(xlErrNA)
    AssertValue ws.Range("AE3").Value2, "e": AssertValue ws.Range("AF3").Value2, "f"
    currentCase = "worksheet split column flags"
    AssertValue ws.Range("AI1").Value2, "aXb": AssertValue ws.Range("AK1").Value2, CVErr(xlErrNA)
    AssertValue ws.Range("AI2").Value2, "a": AssertValue ws.Range("AK2").Value2, "c"
    currentCase = "worksheet test": AssertValue ws.Range("J1").Value2, True
    currentCase = "worksheet nth": AssertValue ws.Range("J2").Value2, "2"
    currentCase = "worksheet replace": AssertValue ws.Range("J3").Value2, "aXb2"
    currentCase = "worksheet capture spill"
    AssertValue ws.Range("L1").Value2, "AB": AssertValue ws.Range("M1").Value2, "12"
    AssertValue ws.Range("L2").Value2, "CD": AssertValue ws.Range("M2").Value2, "34"
    currentCase = "worksheet split spill"
    AssertValue ws.Range("P1").Value2, "red": AssertValue ws.Range("Q1").Value2, "green": AssertValue ws.Range("R1").Value2, "blue"
    currentCase = "worksheet missing": AssertValue ws.Range("J5").Value2, "missing"
    currentCase = "worksheet omitted flags": AssertValue ws.Range("J6").Value2, "3"
    ws.Range("U2").Value2 = "block"
    ws.Range("U1").Formula2 = "=RxMatches(""ab"",""."")"
    ws.Calculate
    currentCase = "blocked spill": AssertValue ws.Evaluate("ERROR.TYPE(U1)"), 9#
End Sub

Private Sub TestFormatAPI()
    Dim regex As StaticRegexSingle.RegexTy, state As StaticRegexSingle.MatcherStateTy
    Dim format As StaticRegexSingle.FormatTy, text As String, caught As Long
    currentCase = "format API"
    text = "ab12cd34"
    StaticRegexSingle.InitializeRegex regex, "(\d+)"
    StaticRegexSingle.InitializeFormat format, regex, "[$1]"
    AssertValue StaticRegexSingle.MatchNext(state, regex, text), True
    AssertValue StaticRegexSingle.FormatMatch(state, text, format), "[12]"
    AssertValue StaticRegexSingle.MatchNext(state, regex, text), True
    AssertValue StaticRegexSingle.FormatMatch(state, text, format), "[34]"
    text = "a"
    StaticRegexSingle.InitializeRegex regex, "a"
    regex.stepsLimit = 1
    StaticRegexSingle.InitializeMatcherState state
    AssertValue StaticRegexSingle.MatchNext(state, regex, text), False
    On Error Resume Next
    StaticRegexSingle.InitializeMatcherState state
    StaticRegexSingle.MatchNext state, regex, text, raiseOnLimit:=True
    caught = Err.Number
    On Error GoTo 0
    AssertValue caught, vbObjectError + 3051
End Sub

Private Sub Check(ByVal label As String, ByVal actual As Variant, ByVal expected As Variant)
    currentCase = label
    If IsArray(actual) Then
        If UBound(actual, 1) <> 0 Or UBound(actual, 2) <> 0 Then Err.Raise 513, , "Expected one cell"
        AssertValue actual(0, 0), expected
    Else
        If Not IsError(actual) Then Err.Raise 513, , "Expected array or error"
        AssertValue actual, expected
    End If
End Sub

Private Sub Grid(ByVal label As String, ByVal actual As Variant, ByVal nr As Long, ByVal nc As Long, ByVal expected As Variant)
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