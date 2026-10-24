Attribute VB_Name = "TestRxDimensions"
Option Explicit

Public Function RunRxDimensionTests() As Variant
    On Error GoTo Failed
    Dim args(0 To 5) As Variant, n As Long, fn As Long, first As Long, second As Long
    Dim axis As Long, sizeA As Long, sizeB As Long, useRange As Long, k As Long, result As Variant
    Dim ws As Worksheet, expected As Variant, nr As Long, nc As Long, size As Long, checks As Long
    Dim context As String
    Set ws = ThisWorkbook.Worksheets(1)
    For useRange = 0 To 1
        For fn = 0 To 5
            Select Case fn
            Case 0, 5: n = 4
            Case 3, 4: n = 5
            Case Else: n = 6
            End Select
            For first = 0 To n - 2
                For second = first + 1 To n - 1
                    For axis = 1 To 2
                        For sizeA = 1 To 3
                            For sizeB = 1 To 3
                                For k = 0 To n - 1
                                    If IsObject(args(k)) Then Set args(k) = Nothing
                                    If k < 2 Then
                                        args(k) = "a"
                                    ElseIf fn = 0 Or fn = 5 Then
                                        If k = 2 Then args(k) = "" Else args(k) = 1
                                    ElseIf fn = 3 Or fn = 4 Then
                                        If k = 3 Then args(k) = "" Else args(k) = 1
                                    Else
                                        Select Case k
                                        Case 2: args(k) = "x"
                                        Case 3, 5: args(k) = 1
                                        Case 4: args(k) = ""
                                        End Select
                                    End If
                                Next
                                SizeArgument args(first), sizeA, axis, useRange, ws, 40
                                SizeArgument args(second), sizeB, axis, useRange, ws, 45
                                context = fn & "," & first & "," & second & "," & axis & "," & sizeA & "," & sizeB & "," & useRange
                                Select Case fn
                                Case 0: result = RxTest(args(0), args(1), args(2), args(3))
                                Case 1: result = RxMatch(args(0), args(1), args(2), args(3), args(4), args(5))
                                Case 2: result = RxReplace(args(0), args(1), args(2), args(3), args(4), args(5))
                                Case 3: result = RxFind(args(0), args(1), args(2), args(3), args(4))
                                Case 4: result = RxLength(args(0), args(1), args(2), args(3), args(4))
                                Case 5: result = RxCount(args(0), args(1), args(2), args(3))
                                End Select
                                If sizeA > 1 And sizeB > 1 And sizeA <> sizeB Then
                                    If Not IsError(result) Then Err.Raise 513, , "Incompatible dimensions accepted"
                                    If CStr(result) <> CStr(CVErr(xlErrValue)) Then Err.Raise 513, , "Expected #VALUE!"
                                Else
                                    nr = 1: nc = 1
                                    If sizeA > sizeB Then size = sizeA Else size = sizeB
                                    If axis = 1 Then nr = size Else nc = size
                                    Select Case fn
                                    Case 0: expected = True
                                    Case 1, 2: expected = "x"
                                    Case Else: expected = 1&
                                    End Select
                                    If Not IsArray(result) Then Err.Raise 513, , "Expected broadcast array"
                                    If UBound(result, 1) <> nr - 1 Or UBound(result, 2) <> nc - 1 Then Err.Raise 513, , "Wrong broadcast shape"
                                    For k = 0 To size - 1
                                        If axis = 1 Then
                                            If IsError(result(k, 0)) Then Err.Raise 513, , "Unexpected cell error"
                                            If result(k, 0) <> expected Then Err.Raise 513, , "Wrong value"
                                        Else
                                            If IsError(result(0, k)) Then Err.Raise 513, , "Unexpected cell error"
                                            If result(0, k) <> expected Then Err.Raise 513, , "Wrong value"
                                        End If
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
    RunRxDimensionTests = checks
    Exit Function
Failed:
    RunRxDimensionTests = context & ": " & Err.Description
End Function

Private Sub SizeArgument(ByRef arg As Variant, ByVal size As Long, ByVal axis As Long, ByVal useRange As Boolean, ByVal ws As Worksheet, ByVal offset As Long)
    Dim a() As Variant, nr As Long, nc As Long, r As Long, c As Long
    nr = 1: nc = 1
    If axis = 1 Then nr = size Else nc = size
    ReDim a(1 To nr, 1 To nc)
    For r = 1 To nr
        For c = 1 To nc
            a(r, c) = arg
        Next
    Next
    If useRange Then
        ws.Cells(offset, 1).Resize(nr, nc).Value2 = a
        Set arg = ws.Cells(offset, 1).Resize(nr, nc)
    Else
        arg = a
    End If
End Sub
