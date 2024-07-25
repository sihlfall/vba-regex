Attribute VB_Name = "StaticStringBuilder"
'
' StaticStringBuilder
' v0.1.1 by Sihlfall
' MIT license
'
' A (relatively) performant, portable StringBuilder user-defined type for VBA / VB6.
'
' Usage:
' Copy this module into your project.
'
' Dim sb As StaticStringBuilder.Ty
' StaticStringBuilder.AppendStr sb, "First"
' StaticStringBuilder.AppendStr sb, "Second"
' StaticStringBuilder.AppendStr sb, "Third"
'
' Dim s As String
' s = StaticStringBuilder.GetStr(sb)
' ' Now s equals "FirstSecondThird"
'
Option Explicit

' Must be at least 2.
Public Const DEFAULT_MINIMUM_CAPACITY As Long = 16

' The StaticStringBuilder type.
' Best do not access its fields directly, but rather use the below subroutines.
Public Type Ty
    Active As Integer            ' index of the currently active buffer (0 or 1)
    Buffer(0 To 1) As String     ' .buffer(.active) is the currently active buffer
    Capacity As Long             ' current allocated capacity in characters
    Length As Long               ' current length of the string, in characters
    MinimumCapacity As Long      ' minimum capacity set (0 or >= 2)
End Type


''' subroutines and functions

' Append s to the string being built.
' s is taken by reference for performance reasons only. s will remain unchanged.
Public Sub AppendStr(ByRef sb As Ty, ByRef s As String)
    Dim Length As Long, nRequired As Long
    With sb
        Length = Len(s)
        If Length = 0 Then Exit Sub
        nRequired = .Length + Length
        If nRequired > .Capacity Then SwitchToLargerBuffer sb, nRequired
        Mid$(.Buffer(.Active), .Length + 1, Length) = s
        .Length = nRequired
    End With
End Sub

Public Sub Clear(ByRef sb As Ty)
    With sb
        .Active = 0
        .Buffer(0) = vbNullString
        .Buffer(1) = vbNullString
        .Capacity = 0
        .Length = 0
    End With
End Sub

Public Function GetLength(ByRef sb As Ty) As Long
    GetLength = sb.Length
End Function

Public Function GetStr(ByRef sb As Ty) As String
    With sb
        GetStr = Left$(.Buffer(.Active), .Length)
    End With
End Function

Public Function GetSubstr(ByRef sb As Ty, ByVal start As Long, ByVal Length As Long) As String
    Dim n As Long
    With sb
        n = .Length - start + 1
        If n <= 0 Then
            GetSubstr = vbNullString
            Exit Function
        End If
        If Length <= n Then n = Length
        GetSubstr = Mid$(.Buffer(.Active), start, n)
    End With
End Function

Public Sub SetMinimumCapacity(ByRef sb As Ty, ByVal MinimumCapacity As Long)
    If MinimumCapacity >= 2 Then sb.MinimumCapacity = MinimumCapacity Else sb.MinimumCapacity = 2
End Sub

''' Private subroutines and functions

Private Sub SwitchToLargerBuffer(ByRef sb As Ty, ByVal nRequired As Long)
    ' Allocate buffer that is able to hold nRequired characters.
    ' The new buffer size is calculated by repeatedly growing the current size by 50%.
    ' Copy string over to the new buffer.
    ' Deallocate the old buffer.
    With sb
        If .MinimumCapacity <= 1 Then .MinimumCapacity = DEFAULT_MINIMUM_CAPACITY
        If .Capacity < .MinimumCapacity Then .Capacity = .MinimumCapacity
        Do
            If .Capacity >= nRequired Then Exit Do
            .Capacity = .Capacity + .Capacity \ 2
        Loop
        .Buffer(1 - .Active) = String(.Capacity, 0)
        Mid$(.Buffer(1 - .Active), 1, .Length) = .Buffer(.Active)
        .Buffer(.Active) = vbNullString
        .Active = 1 - .Active
    End With
End Sub

