Attribute VB_Name = "TestWorksheetDimensions"
Option Explicit

' Exercise every pair of argument positions, both axes, and both size orders.
' Size 1 must broadcast; sizes 2 and 3 must agree.
Public Function RunDimensionTests() As Variant
    On Error GoTo Failed
    Dim fn As Long, count As Long, first As Long, second As Long
    Dim axis As Long, sizeA As Long, sizeB As Long, ranges As Long
    Dim args(0 To 4) As Variant, result As Variant, expected As Variant
    Dim i As Long, r As Long, c As Long, nr As Long, nc As Long
    Dim ws As Worksheet, checks As Long
    Set ws = ThisWorkbook.Worksheets(1)
    For ranges = 0 To 1
        For fn = 0 To 2
            If fn = 0 Then count = 4 Else count = 5
            For first = 0 To count - 2
                For second = first + 1 To count - 1
                    For axis = 1 To 2
                        For sizeA = 1 To 3
                            For sizeB = 1 To 3
                                For i = 0 To count - 1
                                    If IsObject(args(i)) Then Set args(i) = Nothing
                                    If i < 2 Then
                                        args(i) = "a"
                                    ElseIf count = 5 And i = 2 Then
                                        args(i) = "x"
                                    Else
                                        args(i) = False
                                    End If
                                Next
                                If ranges Then
                                    Set args(first) = SizedArgument(args(first), sizeA, axis, True, ws, 1)
                                    Set args(second) = SizedArgument(args(second), sizeB, axis, True, ws, 5)
                                Else
                                    args(first) = SizedArgument(args(first), sizeA, axis, False, ws, 1)
                                    args(second) = SizedArgument(args(second), sizeB, axis, False, ws, 5)
                                End If
                                Select Case fn
                                Case 0: result = VbaRegexWorksheetFunctions.RegexTest(args(0), args(1), args(2), args(3))
                                Case 1: result = VbaRegexWorksheetFunctions.RegexMatch(args(0), args(1), args(2), args(3), args(4))
                                Case 2: result = VbaRegexWorksheetFunctions.RegexReplace(args(0), args(1), args(2), args(3), args(4))
                                End Select
                                If sizeA > 1 And sizeB > 1 And sizeA <> sizeB Then
                                    If Not IsError(result) Then Err.Raise 513, , "Incompatible dimensions accepted"
                                    If CStr(result) <> CStr(CVErr(xlErrValue)) Then Err.Raise 513, , "Expected #VALUE!"
                                Else
                                    nr = 1: nc = 1
                                    If sizeA > sizeB Then i = sizeA Else i = sizeB
                                    If axis = 1 Then nr = i Else nc = i
                                    If Not IsArray(result) Then Err.Raise 513, , "Expected array: " & fn & "," & first & "," & second & "," & axis & "," & sizeA & "," & sizeB & "," & ranges & " result=" & CStr(result)
                                    If UBound(result, 1) <> nr - 1 Or UBound(result, 2) <> nc - 1 Then Err.Raise 513, , "Wrong dimensions"
                                    If fn = 0 Then expected = True Else expected = "x"
                                    For r = 0 To nr - 1
                                        For c = 0 To nc - 1
                                            If IsError(result(r, c)) Then Err.Raise 513, , "Unexpected cell error"
                                            If result(r, c) <> expected Then Err.Raise 513, , "Wrong broadcast value"
                                        Next
                                    Next
                                End If
                                checks = checks + 1
                            Next
                        Next
                    Next
                Next
            Next
        Next
    Next
    RunDimensionTests = checks
    Exit Function
Failed:
    RunDimensionTests = Err.Description
End Function

Private Function SizedArgument(ByVal value As Variant, ByVal size As Long, ByVal axis As Long, _
    ByVal useRange As Boolean, ByVal ws As Worksheet, ByVal offset As Long) As Variant
    Dim a() As Variant, nr As Long, nc As Long, r As Long, c As Long
    nr = 1: nc = 1
    If axis = 1 Then nr = size Else nc = size
    ReDim a(1 To nr, 1 To nc)
    For r = 1 To nr
        For c = 1 To nc
            a(r, c) = value
        Next
    Next
    If useRange Then
        With ws.Cells(offset, 1).Resize(nr, nc)
            .Value2 = a
            Set SizedArgument = .Cells
        End With
    Else
        SizedArgument = a
    End If
End Function