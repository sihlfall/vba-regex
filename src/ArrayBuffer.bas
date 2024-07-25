Attribute VB_Name = "ArrayBuffer"
Public Type Ty
    Buffer() As Long
    Capacity As Long
    Length As Long
End Type

Private Const MINIMUM_CAPACITY As Long = 16

Public Sub AppendLong(ByRef lab As Ty, ByVal v As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .Length + 1
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        .Buffer(.Length) = v
        .Length = requiredCapacity
    End With
End Sub

Public Sub AppendTwo(ByRef lab As Ty, ByVal v1 As Long, ByVal v2 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .Length + 2
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        .Buffer(.Length) = v1
        .Buffer(.Length + 1) = v2
        .Length = requiredCapacity
    End With
End Sub

Public Sub AppendThree(ByRef lab As Ty, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .Length + 3
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        .Buffer(.Length) = v1
        .Buffer(.Length + 1) = v2
        .Buffer(.Length + 2) = v3
        .Length = requiredCapacity
    End With
End Sub

Public Sub AppendFour(ByRef lab As Ty, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long, ByVal v4 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .Length + 4
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        .Buffer(.Length) = v1
        .Buffer(.Length + 1) = v2
        .Buffer(.Length + 2) = v3
        .Buffer(.Length + 3) = v4
        .Length = requiredCapacity
    End With
End Sub

Public Sub AppendFive(ByRef lab As Ty, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long, ByVal v4 As Long, v5 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .Length + 5
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        .Buffer(.Length) = v1
        .Buffer(.Length + 1) = v2
        .Buffer(.Length + 2) = v3
        .Buffer(.Length + 3) = v4
        .Buffer(.Length + 4) = v5
        .Length = requiredCapacity
    End With
End Sub

Public Sub AppendEight(ByRef lab As Ty, ByVal v1 As Long, ByVal v2 As Long, ByVal v3 As Long, ByVal v4 As Long, v5 As Long, v6 As Long, v7 As Long, v8 As Long)
    Dim requiredCapacity As Long
    With lab
        requiredCapacity = .Length + 8
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        .Buffer(.Length) = v1
        .Buffer(.Length + 1) = v2
        .Buffer(.Length + 2) = v3
        .Buffer(.Length + 3) = v4
        .Buffer(.Length + 4) = v5
        .Buffer(.Length + 5) = v6
        .Buffer(.Length + 6) = v7
        .Buffer(.Length + 7) = v8
        .Length = requiredCapacity
    End With
End Sub

Public Sub AppendFill(ByRef lab As Ty, ByVal cnt As Long, ByVal v As Long)
    Dim requiredCapacity As Long, i As Long
    With lab
        requiredCapacity = .Length + cnt
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        i = .Length
        Do While i < requiredCapacity: .Buffer(i) = v: i = i + 1: Loop
        .Length = requiredCapacity
    End With
End Sub

Public Sub AppendSlice(ByRef lab As Ty, ByVal offset As Long, ByVal Length As Long)
    Dim requiredCapacity As Long, i As Long, j As Long, upper As Long
    With lab
        requiredCapacity = .Length + Length
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        upper = offset + Length: i = offset: j = .Length
        Do While i < upper
            .Buffer(j) = .Buffer(i)
            i = i + 1: j = j + 1
        Loop
        .Length = requiredCapacity
    End With
End Sub

Public Sub AppendUnspecified(ByRef lab As Ty, ByVal n As Long)
    Dim requiredCapacity As Long
    With lab
        .Length = .Length + n
        requiredCapacity = .Length
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
    End With
End Sub

Public Sub AppendPrefixedPairsArray(ByRef lab As Ty, ByVal prefix As Long, ByRef ary() As Long)
    ' prefix, number of pairs, pairs
    Dim requiredCapacity As Long, lb As Long, ub As Long, i As Long, j As Long
    With lab
        lb = LBound(ary)
        ub = UBound(ary)
        i = .Length
        .Length = .Length + ub - lb + 3
        requiredCapacity = .Length
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        
        .Buffer(i) = prefix: i = i + 1
        .Buffer(i) = (ub - lb + 1) \ 2: i = i + 1
        For j = lb To ub
            .Buffer(i) = ary(j)
            i = i + 1
        Next
    End With
End Sub

Public Sub AppendPrefixedPairsArrayBackwards(ByRef lab As Ty, ByVal prefix As Long, ByRef ary() As Long)
    ' prefix, number of pairs, pairs
    Dim requiredCapacity As Long, lb As Long, ub As Long, i As Long, j As Long
    With lab
        lb = LBound(ary)
        ub = UBound(ary)
        i = .Length
        .Length = .Length + ub - lb + 3
        requiredCapacity = .Length
        If requiredCapacity > .Capacity Then IncreaseCapacity lab, requiredCapacity
        
        .Buffer(i) = prefix: i = i + 1
        .Buffer(i) = (ub - lb + 1) \ 2: i = i + 1
        j = ub - 1
        Do
            .Buffer(i) = ary(j): i = i + 1
            .Buffer(i) = ary(j + 1): i = i + 1
            j = j - 2
        Loop Until j < lb
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
