VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PortableAssert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public checks As Long
Public failures As String

Public Sub Fail(Optional ByVal message As String = "Assertion failed")
  failures = failures & message & vbLf
End Sub
Public Sub AreEqual(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  checks = checks + 1
  If expected <> actual Then Fail "Expected " & CStr(expected) & ", actual " & CStr(actual) & ". " & message
End Sub
Public Sub IsTrue(ByVal actual As Boolean, Optional ByVal message As String = "")
  checks = checks + 1
  If Not actual Then Fail "Expected True. " & message
End Sub
Public Sub IsFalse(ByVal actual As Boolean, Optional ByVal message As String = "")
  checks = checks + 1
  If actual Then Fail "Expected False. " & message
End Sub
Private Function Dimensions(ByRef value As Variant) As Long
  On Error GoTo Done
  Dim n As Long, bound As Long
  For n = 1 To 10
    bound = UBound(value, n)
    Dimensions = n
  Next
Done:
End Function
Public Sub SequenceEquals(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
  checks = checks + 1
  Dim dims As Long, i As Long, j As Long
  dims = Dimensions(expected)
  If dims <> Dimensions(actual) Then Fail "Array dimensions differ. " & message: Exit Sub
  For i = 1 To dims
    If LBound(expected, i) <> LBound(actual, i) Or UBound(expected, i) <> UBound(actual, i) Then
      Fail "Array bounds differ. " & message: Exit Sub
    End If
  Next
  If dims = 1 Then
    For i = LBound(expected) To UBound(expected)
      If expected(i) <> actual(i) Then Fail "Array element " & i & " differs. " & message: Exit Sub
    Next
  ElseIf dims = 2 Then
    For i = LBound(expected, 1) To UBound(expected, 1)
      For j = LBound(expected, 2) To UBound(expected, 2)
        If expected(i, j) <> actual(i, j) Then Fail "Array element differs. " & message: Exit Sub
      Next
    Next
  ElseIf dims <> 0 Then
    Fail "Unsupported array dimensions in assertion adapter."
  End If
End Sub
