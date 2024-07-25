Attribute VB_Name = "RegexReplace"
Option Explicit
Option Private Module

Public Const REPL_END As Long = 0
Public Const REPL_DOLLAR As Long = 1
Public Const REPL_SUBSTR As Long = 2
Public Const REPL_PREFIX As Long = 3
Public Const REPL_SUFFIX As Long = 4
Public Const REPL_ACTUAL As Long = 5
Public Const REPL_NUMBERED As Long = 6
Public Const REPL_NAMED As Long = 7

Public Sub ParseFormatString(ByRef parsedFormat As ArrayBuffer.Ty, ByRef formatString As String, ByRef bytecode() As Long, ByRef pattern As String)
    Dim curPos As Long, lastPos As Long, c As Long, formatStringLen As Long, num As Long, substrLen As Long, identifierId As Long
    
    Const UNICODE_DOLLAR As Long = 36
    Const UNICODE_AMP As Long = 38
    Const UNICODE_SQUOTE As Long = 39
    Const UNICODE_DIGIT_0 As Long = 48
    Const UNICODE_DIGIT_9 As Long = 57
    Const UNICODE_LT As Long = 60
    Const UNICODE_BACKTICK As Long = 96
    Const UNICODE_TILDE As Long = 126
    Const MAX_LONG As Long = &H7FFFFFFF
    Const MAX_LONG_DIV_10 As Long = &H7FFFFFFF \ 10
    
    
    formatStringLen = Len(formatString)
    curPos = 1
    lastPos = 1
    Do
        curPos = InStr(curPos, formatString, "$", vbBinaryCompare)
        If curPos = 0 Then Exit Do
        If curPos = formatStringLen Then GoTo InvalidReplacementString
        curPos = curPos + 1
        c = AscW(Mid$(formatString, curPos, 1))
        If c = UNICODE_DOLLAR Then
            If curPos - lastPos = 1 Then
                ArrayBuffer.AppendLong parsedFormat, REPL_DOLLAR
            Else
                ArrayBuffer.AppendThree parsedFormat, REPL_SUBSTR, lastPos, curPos - lastPos
            End If
            curPos = curPos + 1
        Else
            substrLen = curPos - lastPos - 1
            If substrLen > 0 Then ArrayBuffer.AppendThree parsedFormat, REPL_SUBSTR, lastPos, substrLen
            Select Case c
            Case UNICODE_AMP
                ArrayBuffer.AppendLong parsedFormat, REPL_ACTUAL
                curPos = curPos + 1
            Case UNICODE_SQUOTE
                ArrayBuffer.AppendLong parsedFormat, REPL_SUFFIX
                curPos = curPos + 1
            Case UNICODE_LT
                If curPos = formatStringLen Then GoTo InvalidReplacementString
                lastPos = curPos + 1
                curPos = InStr(lastPos, formatString, ">", vbBinaryCompare)
                If curPos = lastPos Then GoTo InvalidReplacementString ' empty identifier
                
                identifierId = RegexBytecode.GetIdentifierId(bytecode, pattern, Mid$(formatString, lastPos, curPos - lastPos))
                If identifierId >= 0 Then ArrayBuffer.AppendTwo parsedFormat, REPL_NAMED, identifierId

                curPos = curPos + 1
            Case UNICODE_BACKTICK
                ArrayBuffer.AppendLong parsedFormat, REPL_PREFIX
                curPos = curPos + 1
            Case UNICODE_TILDE
                ' ignore
                curPos = curPos + 1
            Case Else
                ' Todo: Check whether we can merge this with parsing a number within a regex
                If c < UNICODE_DIGIT_0 Then GoTo InvalidReplacementString
                If c > UNICODE_DIGIT_9 Then GoTo InvalidReplacementString
                num = 0
                Do
                    If num > MAX_LONG_DIV_10 Then GoTo InvalidReplacementString
                    
                    num = 10 * num
                    c = c - UNICODE_DIGIT_0
                    
                    If num > MAX_LONG - c Then GoTo InvalidReplacementString
                    
                    num = num + c
                    
                    curPos = curPos + 1
                    If curPos > formatStringLen Then Exit Do
                    c = AscW(Mid$(formatString, curPos, 1))
                    If c < UNICODE_DIGIT_0 Then Exit Do
                    If c > UNICODE_DIGIT_9 Then Exit Do
                Loop
                ArrayBuffer.AppendTwo parsedFormat, REPL_NUMBERED, num
            End Select
        End If
        lastPos = curPos
    Loop
    
    substrLen = formatStringLen + 1 - lastPos
    If substrLen > 0 Then ArrayBuffer.AppendThree parsedFormat, REPL_SUBSTR, lastPos, substrLen
    ArrayBuffer.AppendLong parsedFormat, REPL_END
    
    Exit Sub
InvalidReplacementString:
    Err.Raise RegexErrors.REGEX_ERR_INVALID_REPLACEMENT_STRING
End Sub


Public Sub AppendFormatted( _
    ByRef sb As StaticStringBuilder.Ty, _
    ByRef sHaystack As String, _
    ByRef captures As RegexDfsMatcher.CapturesTy, _
    ByRef formatString As String, _
    ByRef parsed() As Long, _
    Optional ByVal parsedStartPos As Long = 0 _
)
    Dim j As Long, num As Long

    j = parsedStartPos
    Do
        Select Case parsed(j)
        Case REPL_END
            Exit Do
        Case REPL_DOLLAR
            StaticStringBuilder.AppendStr sb, "$"
            j = j + 1
        Case REPL_SUBSTR
            StaticStringBuilder.AppendStr sb, Mid$(formatString, parsed(j + 1), parsed(j + 2))
            j = j + 3
        Case REPL_PREFIX
            StaticStringBuilder.AppendStr sb, Left$(sHaystack, captures.entireMatch.start - 1)
            j = j + 1
        Case REPL_SUFFIX
            StaticStringBuilder.AppendStr sb, Mid$(sHaystack, captures.entireMatch.start + captures.entireMatch.Length)
            j = j + 1
        Case REPL_ACTUAL
            StaticStringBuilder.AppendStr sb, Mid$(sHaystack, captures.entireMatch.start, captures.entireMatch.Length)
            j = j + 1
        Case REPL_NUMBERED
            num = parsed(j + 1)
            If num <= captures.nNumberedCaptures Then
                With captures.numberedCaptures(num - 1)
                    If .Length > 0 Then StaticStringBuilder.AppendStr sb, Mid$(sHaystack, .start, .Length)
                End With
            End If
            j = j + 2
        Case REPL_NAMED
            num = captures.namedCaptures(parsed(j + 1))
            If num >= 0 Then
                If num <= captures.nNumberedCaptures Then
                    With captures.numberedCaptures(num - 1)
                        If .Length > 0 Then StaticStringBuilder.AppendStr sb, Mid$(sHaystack, .start, .Length)
                    End With
                End If
            End If
            j = j + 2
        Case Else
            Err.Raise RegexErrors.REGEX_ERR_INTERNAL_LOGIC_ERR
        End Select
    Loop


End Sub

