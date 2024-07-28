Attribute VB_Name = "ArrayBuffer"
Public Type Ty
    Buffer() As Long
    Capacity As Long
    length As Long
End Type

Private Const MINIMUM_CAPACITY As Long = 16

Public Sub AppendLong(ByRef lab As Ty, ByVal v As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .length + 1
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        .Buffer(.length) = v
        .length = requiredCapacity
    End With
End Sub

Public Sub AppendTwo(ByRef lab As Ty, ByVal v1 As Long, ByVal v2 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .length + 2
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        .Buffer(.length) = v1
        .Buffer(.length + 1) = v2
        .length = requiredCapacity
    End With
End Sub

Public Sub AppendThree(ByRef lab As Ty, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .length + 3
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        .Buffer(.length) = v1
        .Buffer(.length + 1) = v2
        .Buffer(.length + 2) = v3
        .length = requiredCapacity
    End With
End Sub

Public Sub AppendFour(ByRef lab As Ty, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long, ByVal v4 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .length + 4
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        .Buffer(.length) = v1
        .Buffer(.length + 1) = v2
        .Buffer(.length + 2) = v3
        .Buffer(.length + 3) = v4
        .length = requiredCapacity
    End With
End Sub

Public Sub AppendFive(ByRef lab As Ty, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long, ByVal v4 As Long, v5 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .length + 5
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        .Buffer(.length) = v1
        .Buffer(.length + 1) = v2
        .Buffer(.length + 2) = v3
        .Buffer(.length + 3) = v4
        .Buffer(.length + 4) = v5
        .length = requiredCapacity
    End With
End Sub

Public Sub AppendEight(ByRef lab As Ty, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long, ByVal v4 As Long, v5 As Long, v6 As Long, v7 As Long, v8 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .length + 8
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        .Buffer(.length) = v1
        .Buffer(.length + 1) = v2
        .Buffer(.length + 2) = v3
        .Buffer(.length + 3) = v4
        .Buffer(.length + 4) = v5
        .Buffer(.length + 5) = v6
        .Buffer(.length + 6) = v7
        .Buffer(.length + 7) = v8
        .length = requiredCapacity
    End With
End Sub

Public Sub AppendFill(ByRef lab As Ty, ByVal cnt As Long, ByVal v As Long)
    Dim requiredCapacity As Long, i As Long
    With lab
        requiredCapacity = .length + cnt
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        i = .length
        Do While i < requiredCapacity: .Buffer(i) = v: i = i + 1: Loop
        .length = requiredCapacity
    End With
End Sub

Public Sub AppendSlice(ByRef lab As Ty, ByVal offset As Long, ByVal length As Long)
    Dim requiredCapacity As Long, i As Long, j As Long, upper As Long
    With lab
        requiredCapacity = .length + length
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        upper = offset + length: i = offset: j = .length
        Do While i < upper
            .Buffer(j) = .Buffer(i)
            i = i + 1: j = j + 1
        Loop
        .length = requiredCapacity
    End With
End Sub

Public Sub AppendUnspecified(ByRef lab As Ty, ByVal n As Long)
    Dim requiredCapacity As Long
    With lab
        .length = .length + n
        requiredCapacity = .length
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
    End With
End Sub

Public Sub AppendPrefixedPairsArray(ByRef lab As Ty, ByVal prefix As Long, ByRef ary() As Long, ByVal aryStart As Long, ByVal aryLength As Long)
    ' prefix, number of pairs, pairs
    Dim requiredCapacity As Long, ub As Long, i As Long, j As Long
    With lab
        i = .length
        .length = .length + aryLength + 2
        requiredCapacity = .length
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        
        .Buffer(i) = prefix: i = i + 1
        .Buffer(i) = aryLength \ 2: i = i + 1
        ub = aryStart + aryLength - 1
        For j = aryStart To ub
            .Buffer(i) = ary(j)
            i = i + 1
        Next
    End With
End Sub

Private Sub IncreaseCapacity(ByRef lab As Ty, requiredCapacity As Long)
    Dim cap As Long
    With lab
        cap = .Capacity
        If cap <= MINIMUM_CAPACITY Then cap = MINIMUM_CAPACITY
        Do Until cap >= requiredCapacity
            cap = cap + cap \ 2
        Loop
        ReDim Preserve .Buffer(0 To cap - 1) As Long
        .Capacity = cap
    End With
End Sub
