Attribute VB_Name = "CommandLineParser"
Option Explicit

Private Const CLP_EOF As Long = 0
Private Const CLP_ARG As Long = 1
Private Const CLP_UNTERMINATED As Long = -1

Private Type SimpleStringBuilder
    length As Long
    buf As String
End Type

Private Sub ParseArgument(ByRef outToken As Long, ByRef outSb As SimpleStringBuilder, ByRef cmdLine As String, ByRef i As Long)
    Dim cml As Long, ch As String
    
    outSb.length = 0
    cml = Len(cmdLine)
    ' skip leading whitespace
    Do
        If i > cml Then
            outToken = CLP_EOF
            Exit Sub
        End If
        ch = Mid$(cmdLine, i, 1)
        If ch <> " " Then Exit Do
        i = i + 1
    Loop
    GoTo I0
    
T0:
    i = i + 1
    If i > cml Then
        outToken = CLP_ARG
        Exit Sub
    End If
    ch = Mid$(cmdLine, i, 1)
S0:
    If ch = " " Then
        outToken = CLP_ARG
        Exit Sub
    End If
I0:
    If ch = ChrW$(34) Then GoTo T1
    With outSb: Mid$(.buf, .length + 1, 1) = ch: .length = .length + 1: End With
    GoTo T0
    
T1:
    i = i + 1
    If i > cml Then
        outToken = CLP_UNTERMINATED
        Exit Sub
    End If
    ch = Mid$(cmdLine, i, 1)
S1:
    If ch = ChrW$(34) Then GoTo T2
    With outSb: Mid$(.buf, .length + 1, 1) = ch: .length = .length + 1: End With
    GoTo T1
    
T2:
    i = i + 1
    If i > cml Then
        outToken = CLP_ARG
        Exit Sub
    End If
    ch = Mid$(cmdLine, i, 1)
S2:
    If ch = ChrW$(34) Then
        With outSb: Mid$(.buf, .length + 1, 1) = ch: .length = .length + 1: End With
        GoTo T1
    End If
    GoTo S0
End Sub

Function ParseCommandLine(ByRef cmdLine As String) As Collection
    Dim col As Collection, sl As Long, i As Long, token As Long, sb As SimpleStringBuilder
    
    Set col = New Collection
    sl = Len(cmdLine)
    If sl = 0 Then
        Set ParseCommandLine = col
        Exit Function
    End If
    
    With sb: .buf = String(sl, 0): .length = 0: End With
    i = 1
    Do
        ParseArgument token, sb, cmdLine, i
        If token = CLP_ARG Then
            With sb: col.Add Mid$(.buf, 1, .length): End With
        ElseIf token = CLP_UNTERMINATED Then
            Set ParseCommandLine = Nothing
            Exit Function
        Else ' token = CLP_EOF
            Set ParseCommandLine = col
            Exit Function
        End If
    Loop
End Function
