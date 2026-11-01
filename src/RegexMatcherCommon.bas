Attribute VB_Name = "RegexMatcherCommon"
Option Explicit

Public Enum MatcherCommonConstant
    MATCHER_ENDOFINPUT = -1
End Enum

Public Type StartLengthPair
    start As Long
    Length As Long
End Type

Public Type CapturesTy
    nNumberedCaptures As Long
    nNamedCaptures As Long
    entireMatch As StartLengthPair
    numberedCaptures() As StartLengthPair
    namedCaptures() As Long
End Type

Public Type LiteralMatcherTy
    isInitialized As Boolean
    isLiteral As Boolean
    text As String
    startAnchorOp As Long
    endAnchorOp As Long
End Type

Public Type InputBufferTy
    bytes() As Byte
    Length As Long ' UTF-16 code units, not bytes
    windowStart As Long ' inclusive, zero-based UTF-16 position
    windowEnd As Long ' exclusive
    Capacity As Long ' target window size in bytes
End Type

Public Type LineStartSearchTy
    initialized As Boolean
    nextLf As Long
    nextCr As Long
    nextLs As Long
    nextPs As Long
End Type

Public Const INPUT_BUFFER_INITIAL_BYTES As Long = 1024
Public Const INPUT_BUFFER_MAX_BYTES As Long = 65536

Public Sub InitializeInputBuffer(ByRef inputBuffer As InputBufferTy, ByRef inputStr As String)
    With inputBuffer
        .Length = Len(inputStr)
        .windowStart = 0: .windowEnd = 0
        .Capacity = INPUT_BUFFER_INITIAL_BYTES
        Erase .bytes
    End With
End Sub

' The source stays with the caller and must remain unchanged during a match.
' Copy chunks as raw UTF-16 bytes; StrConv would change the encoding. A backward
' miss grows by at least 1.4 and covers the distance back from the old window's
' end, capped independently for each reader. Using that end conservatively
' includes the last read position without recording every individual read.
Public Sub RefillInputBuffer(ByRef inputBuffer As InputBufferTy, ByRef inputStr As String, ByVal sp As Long, ByVal spDelta As Long)
    Dim n As Long, required As Long, backwardMiss As Boolean
    With inputBuffer
        If sp < 0 Or sp >= .Length Then Exit Sub
        backwardMiss = sp < .windowStart
        If backwardMiss Then
            .Capacity = ((.Capacity * 7 + 4) \ 5 + 1) And Not 1&
            ' Include the requested position and one preceding assertion char.
            If .windowEnd - sp >= INPUT_BUFFER_MAX_BYTES \ 2 Then
                required = INPUT_BUFFER_MAX_BYTES
            Else
                required = 2 * (.windowEnd - sp + 1)
            End If
            If .Capacity < required Then .Capacity = required
            If .Capacity > INPUT_BUFFER_MAX_BYTES Then .Capacity = INPUT_BUFFER_MAX_BYTES
        End If
        n = .Capacity \ 2
        If backwardMiss And .windowEnd - sp < n Then
            ' Retain the previous read region as well as the backward target.
            .windowStart = .windowEnd - n
            If .windowStart < 0 Then .windowStart = 0
        ElseIf backwardMiss Or spDelta < 0 Then
            ' Leave one following character available for reverse assertions.
            .windowEnd = sp + 1
            If .windowEnd < .Length Then .windowEnd = .windowEnd + 1
            .windowStart = .windowEnd - n
            If .windowStart < 0 Then .windowStart = 0
        Else
            ' Keep the preceding character, including across chunk boundaries.
            .windowStart = sp - 1
            If .windowStart > .Length - n Then .windowStart = .Length - n
            If .windowStart < 0 Then .windowStart = 0
            .windowEnd = .windowStart + n
            If .windowEnd > .Length Then .windowEnd = .Length
        End If
        .bytes = Mid$(inputStr, .windowStart + 1, .windowEnd - .windowStart)
    End With
End Sub

Public Function GetBufferCharCode(ByRef inputBuffer As InputBufferTy, ByRef inputStr As String, ByRef sp As Long, ByVal spDelta As Long) As Long
    Dim byteOffset As Long
    If sp < inputBuffer.windowStart Or sp >= inputBuffer.windowEnd Then
        If sp < 0 Or sp >= inputBuffer.Length Then
            GetBufferCharCode = MATCHER_ENDOFINPUT
            Exit Function
        End If
        RefillInputBuffer inputBuffer, inputStr, sp, spDelta
    End If
    byteOffset = (sp - inputBuffer.windowStart) * 2
    GetBufferCharCode = CLng(inputBuffer.bytes(byteOffset)) Or (CLng(inputBuffer.bytes(byteOffset + 1)) * 256&)
    sp = sp + spDelta
End Function

Public Function PeekBufferCharCode(ByRef inputBuffer As InputBufferTy, ByRef inputStr As String, ByVal sp As Long, Optional ByVal spDelta As Long = 1) As Long
    Dim byteOffset As Long
    If sp < inputBuffer.windowStart Or sp >= inputBuffer.windowEnd Then
        If sp < 0 Or sp >= inputBuffer.Length Then
            PeekBufferCharCode = MATCHER_ENDOFINPUT
            Exit Function
        End If
        RefillInputBuffer inputBuffer, inputStr, sp, spDelta
    End If
    byteOffset = (sp - inputBuffer.windowStart) * 2
    PeekBufferCharCode = CLng(inputBuffer.bytes(byteOffset)) Or (CLng(inputBuffer.bytes(byteOffset + 1)) * 256&)
End Function

' Return a zero-based candidate strictly after sp, including EOF after a final
' terminator. Cached searches avoid quadratic rescans on inputs with many lines.
Public Function NextLineStart(ByRef search As LineStartSearchTy, ByRef inputStr As String, ByVal sp As Long) As Long
    Dim from As Long, best As Long
    from = sp + 1 ' InStr is one-based; a terminator here precedes the next start.
    With search
        If Not .initialized Then
            .nextLf = -1: .nextCr = -1: .nextLs = -1: .nextPs = -1
            .initialized = True
        End If
        If .nextLf = -1 Or (.nextLf > 0 And .nextLf < from) Then .nextLf = InStr(from, inputStr, vbLf, vbBinaryCompare)
        If .nextCr = -1 Or (.nextCr > 0 And .nextCr < from) Then .nextCr = InStr(from, inputStr, vbCr, vbBinaryCompare)
        If .nextLs = -1 Or (.nextLs > 0 And .nextLs < from) Then .nextLs = InStr(from, inputStr, ChrW(&H2028), vbBinaryCompare)
        If .nextPs = -1 Or (.nextPs > 0 And .nextPs < from) Then .nextPs = InStr(from, inputStr, ChrW(&H2029), vbBinaryCompare)
        best = .nextLf
        If .nextCr > 0 Then
            If best = 0 Or .nextCr < best Then best = .nextCr
        End If
        If .nextLs > 0 Then
            If best = 0 Or .nextLs < best Then best = .nextLs
        End If
        If .nextPs > 0 Then
            If best = 0 Or .nextPs < best Then best = .nextPs
        End If
    End With
    If best = 0 Then NextLineStart = -1 Else NextLineStart = best
End Function

' Conservatively recognize an unavoidable leading ^, not a ^ in one alternative.
Public Function LeadingStartAnchor(ByRef bytecode() As Long) As Long
    Dim p As Long
    If Not HasSearchPrefix(bytecode) Then Exit Function
    p = bytecode(RegexBytecode.BYTECODE_IDX_START_ACTUAL)
    Do While bytecode(p) = RegexBytecode.REOP_SAVE
        p = p + 2
    Loop
    If (bytecode(p) And RegexBytecode.REOP_OPCODE_MASK) = RegexBytecode.REOP_ASSERT_START Then LeadingStartAnchor = bytecode(p)
End Function

' Recognize the compiler's unanchored-search prefix. Handwritten bytecode with a
' different entry sequence must retain its original VM behavior.
Public Function HasSearchPrefix(ByRef bytecode() As Long) As Boolean
    Dim p As Long
    p = bytecode(RegexBytecode.BYTECODE_IDX_START_WITH_PREFIX)
    If bytecode(RegexBytecode.BYTECODE_IDX_START_ACTUAL) <> p + 5 Then Exit Function
    If bytecode(p) <> RegexBytecode.REOP_SPLIT2 Then Exit Function
    If bytecode(p + 1) <> 3 Then Exit Function
    If bytecode(p + 2) <> (RegexBytecode.REOP_DOT Or RegexBytecode.MODIFIER_S_MASK) Then Exit Function
    If bytecode(p + 3) <> RegexBytecode.REOP_JUMP Then Exit Function
    If bytecode(p + 4) <> -5 Then Exit Function
    HasSearchPrefix = True
End Function

Public Sub InitializeLiteralMatcher(ByRef literal As LiteralMatcherTy, ByRef bytecode() As Long)
    Dim literalPc As Long, n As Long, i As Long
    With literal
        .isInitialized = True: .isLiteral = False
        .text = vbNullString: .startAnchorOp = 0: .endAnchorOp = 0
    End With
    If Not GetLiteralLayout(bytecode, literalPc, n, literal.startAnchorOp, literal.endAnchorOp) Then Exit Sub
    literal.text = String$(n, vbNullChar)
    For i = 1 To n
        Mid$(literal.text, i, 1) = ChrW(bytecode(literalPc + i - 1))
    Next
    literal.isLiteral = True
End Sub

' The compiler uses the same layout check without allocating a literal string.
Public Function IsLiteralBytecode(ByRef bytecode() As Long) As Boolean
    Dim literalPc As Long, n As Long, startAnchorOp As Long, endAnchorOp As Long
    IsLiteralBytecode = GetLiteralLayout(bytecode, literalPc, n, startAnchorOp, endAnchorOp)
End Function

Private Function GetLiteralLayout( _
    ByRef bytecode() As Long, ByRef literalPc As Long, ByRef n As Long, _
    ByRef startAnchorOp As Long, ByRef endAnchorOp As Long _
) As Boolean
    Dim p As Long, op As Long
    ' Captures and general assertions remain with the regular matchers. Binary
    ' string operations cannot implement the regex's Unicode case folding.
    If bytecode(RegexBytecode.BYTECODE_IDX_MAX_PROPER_CAPTURE_SLOT) <> 1 Then Exit Function
    If bytecode(RegexBytecode.BYTECODE_IDX_N_IDENTIFIERS) <> 0 Then Exit Function
    If Not HasSearchPrefix(bytecode) Then Exit Function
    p = bytecode(RegexBytecode.BYTECODE_IDX_START_ACTUAL)
    If bytecode(p) <> RegexBytecode.REOP_SAVE Then Exit Function
    If bytecode(p + 1) <> 0 Then Exit Function
    p = p + 2
    If (bytecode(p) And RegexBytecode.REOP_OPCODE_MASK) = RegexBytecode.REOP_ASSERT_START Then
        startAnchorOp = bytecode(p): p = p + 1
    End If

    op = bytecode(p)
    If op And RegexBytecode.MODIFIER_I_ACTIVE Then Exit Function
    Select Case op And RegexBytecode.REOP_OPCODE_MASK
    Case RegexBytecode.REOP_CHAR
        literalPc = p + 1: n = 1: p = p + 2
    Case RegexBytecode.REOP_STRING
        n = bytecode(p + 1): literalPc = p + 2: p = p + 2 + n
    Case RegexBytecode.REOP_SAVE, RegexBytecode.REOP_ASSERT_END
        n = 0 ' the empty literal, possibly anchored
    Case Else
        Exit Function
    End Select
    If (bytecode(p) And RegexBytecode.REOP_OPCODE_MASK) = RegexBytecode.REOP_ASSERT_END Then
        endAnchorOp = bytecode(p): p = p + 1
    End If
    If bytecode(p) <> RegexBytecode.REOP_SAVE Then Exit Function
    If bytecode(p + 1) <> 1 Then Exit Function
    If bytecode(p + 2) <> RegexBytecode.REOP_MATCH Then Exit Function

    GetLiteralLayout = True
End Function

Public Function LiteralFindMatch( _
    ByRef outMatch As StartLengthPair, ByRef literal As LiteralMatcherTy, _
    ByRef inputStr As String, ByVal matchFrom As Long, ByVal multiline As Boolean _
) As Boolean
    Dim p As Long, n As Long, lastStart As Long, inputLength As Long
    Dim startMultiline As Boolean, endMultiline As Boolean, fixedStart As Boolean
    outMatch.start = 0: outMatch.Length = 0
    inputLength = Len(inputStr): n = Len(literal.text)
    lastStart = inputLength - n + 1
    If matchFrom < 1 Or matchFrom > lastStart Then Exit Function
    startMultiline = (literal.startAnchorOp - (multiline And RegexBytecode.MODIFIER_M_WRITE) And RegexBytecode.MODIFIER_M_ACTIVE) <> 0
    endMultiline = (literal.endAnchorOp - (multiline And RegexBytecode.MODIFIER_M_WRITE) And RegexBytecode.MODIFIER_M_ACTIVE) <> 0

    If literal.endAnchorOp <> 0 And Not endMultiline Then
        p = lastStart: fixedStart = True
    ElseIf literal.startAnchorOp <> 0 And Not startMultiline Then
        If matchFrom <> 1 Then Exit Function
        p = 1: fixedStart = True
    Else
        p = matchFrom
    End If
    If fixedStart Then
        If StrComp(Mid$(inputStr, p, n), literal.text, vbBinaryCompare) <> 0 Then Exit Function
    End If
    Do
        If Not fixedStart And n > 0 Then
            p = InStr(p, inputStr, literal.text, vbBinaryCompare)
            If p = 0 Then Exit Function
        End If
        If literal.startAnchorOp <> 0 And p > 1 Then
            If Not startMultiline Then GoTo NextCandidate
            If Not RegexUnicodeSupport.UnicodeIsLineTerminator(PeekInputCharCode(inputStr, p - 2)) Then GoTo NextCandidate
        End If
        If literal.endAnchorOp <> 0 And p + n - 1 < inputLength Then
            If Not endMultiline Then GoTo NextCandidate
            If Not RegexUnicodeSupport.UnicodeIsLineTerminator(PeekInputCharCode(inputStr, p + n - 1)) Then GoTo NextCandidate
        End If
        outMatch.start = p: outMatch.Length = n
        LiteralFindMatch = True
        Exit Function
NextCandidate:
        If fixedStart Then Exit Function
        p = p + 1
        If p > lastStart Then Exit Function
    Loop
End Function

Public Function GetInputCharCode(ByRef inputStr As String, ByRef sp As Long, ByVal spDelta As Long) As Long
    If sp >= Len(inputStr) Then
        GetInputCharCode = MATCHER_ENDOFINPUT
    ElseIf sp < 0 Then
        GetInputCharCode = MATCHER_ENDOFINPUT
    Else
        GetInputCharCode = AscW(Mid$(inputStr, sp + 1, 1)) And &HFFFF& ' sp is 0-based and Mid$ is 1-based
        sp = sp + spDelta
    End If
End Function

Public Function PeekInputCharCode(ByRef inputStr As String, ByRef sp As Long) As Long
    If sp >= Len(inputStr) Then
        PeekInputCharCode = MATCHER_ENDOFINPUT
    ElseIf sp < 0 Then
        PeekInputCharCode = MATCHER_ENDOFINPUT
    Else
        PeekInputCharCode = AscW(Mid$(inputStr, sp + 1, 1)) And &HFFFF& ' sp is 0-based and Mid$ is 1-based
    End If
End Function

Public Function UnicodeReIsWordchar(c As Long) As Boolean
    If c <= 90 Then ' Z
        If c >= 65 Then UnicodeReIsWordchar = True: Exit Function ' A
        If c > 57 Then UnicodeReIsWordchar = False: Exit Function ' 9
        If c < 48 Then UnicodeReIsWordchar = False: Exit Function ' 0
        UnicodeReIsWordchar = True: Exit Function
    Else
        If c > 122 Then UnicodeReIsWordchar = False: Exit Function ' z
        If c < 97 Then UnicodeReIsWordchar = c = 95: Exit Function ' a, underscore
        UnicodeReIsWordchar = True: Exit Function
    End If
End Function

Public Sub CopyCaptures( _
    ByRef buf() As Long, _
    ByVal start As Long, _
    ByRef captures As RegexMatcherCommon.CapturesTy _
)
    Dim i As Long, pt1 As Long, pt2 As Long
    
    pt1 = buf(start)
    pt2 = buf(start + 1)
    If pt1 = -1 Then
        With captures.entireMatch: .start = 0: .Length = 0: End With
    ElseIf pt2 < pt1 Then
        With captures.entireMatch: .start = 0: .Length = 0: End With
    Else
        With captures.entireMatch: .start = pt1 + 1: .Length = pt2 - pt1: End With
    End If
        
    For i = 1 To captures.nNumberedCaptures
        pt1 = buf(start + 2 * i)
        pt2 = buf(start + 2 * i + 1)
        If pt1 = -1 Then
            With captures.numberedCaptures(i - 1): .start = 0: .Length = 0: End With
        ElseIf pt2 < pt1 Then
            With captures.numberedCaptures(i - 1): .start = 0: .Length = 0: End With
        Else
            With captures.numberedCaptures(i - 1): .start = pt1 + 1: .Length = pt2 - pt1: End With
        End If
    Next
    
    start = start + 2 * captures.nNumberedCaptures + 2
    For i = 0 To captures.nNamedCaptures - 1
        captures.namedCaptures(i) = buf(start + i)
    Next
End Sub

