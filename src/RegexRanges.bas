Attribute VB_Name = "RegexRanges"
Option Explicit

Public Sub EmitPredefinedRange( _
    ByRef outBuffer As ArrayBuffer.Ty, ByRef source() As Long, ByVal sourceStart As Long, ByVal sourceLength As Long _
)
    Dim i As Long, j As Long
    
    With outBuffer
        i = .length
        ArrayBuffer.AppendUnspecified outBuffer, sourceLength
        j = sourceStart + sourceLength - 2
        Do While j >= sourceStart
            .Buffer(i) = source(j): i = i + 1
            .Buffer(i) = source(j + 1): i = i + 1
            j = j - 2
        Loop
    End With
End Sub

Public Function UnicodeIsIdentifierPart(x As Long) As Boolean
    UnicodeIsIdentifierPart = False
End Function

Public Sub RegexpGenerateRanges(ByRef outBuffer As ArrayBuffer.Ty, _
    ByVal caseInsensitive As Boolean, ByVal r1 As Long, ByVal r2 As Long _
)
    Dim a As Long, b As Long, m As Long, d As Long, ub As Long, lastDelta As Long
    Dim r As Long, rc As Long

    If Not caseInsensitive Then
        ArrayBuffer.AppendLong outBuffer, r1
        ArrayBuffer.AppendLong outBuffer, r2
        Exit Sub
    End If
        
    rc = RegexUnicodeSupport.ReCanonicalizeChar(r1)
    lastDelta = rc - r1
    ArrayBuffer.AppendLong outBuffer, rc

    a = RegexUnicodeSupport.UNICODE_CANON_RUNS_TABLE_START - 1
    ub = a + RegexUnicodeSupport.UNICODE_CANON_RUNS_TABLE_LENGTH
    
    If RegexUnicodeSupport.StaticData(ub) > r1 Then
        ' Find the index of the first element larger than r1.
        ' The index is guaranteed to be in the interval (a;b].
        b = ub
        Do
            d = b - a
            If d = 1 Then Exit Do
            m = a + d \ 2
            If RegexUnicodeSupport.StaticData(m) > r1 Then b = m Else a = m
        Loop
        
        ' Now b is the index of the first element larger than r1.
        Do
            r = RegexUnicodeSupport.StaticData(b)
            If r > r2 Then Exit Do
            ArrayBuffer.AppendLong outBuffer, r - 1 + lastDelta
            
            rc = RegexUnicodeSupport.ReCanonicalizeChar(r)
            ArrayBuffer.AppendLong outBuffer, rc
            lastDelta = rc - r

            If b = ub Then Exit Do
            b = b + 1
        Loop
    End If
    
    ArrayBuffer.AppendLong outBuffer, r2 + lastDelta
End Sub
