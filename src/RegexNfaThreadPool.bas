Attribute VB_Name = "RegexNfaThreadPool"
Option Explicit

Public Type ThreadSetNode
    conditions As Long ' DFA: next-character categories already explored here
    pc As Long
    strOffset As Long
    
    rbParent As Long
    rbChild(0 To 1) As Long ' 0 = left, 1 = right
    rbIsBlack As Boolean
End Type

Public Type ThreadSetBuffer
    Capacity As Long
    nUsed As Long
    buf() As ThreadSetNode
End Type

' returns
' * index of node if node was found; in that case out parameters will remain unchanged
' * -1 otherwise, then the out parameters will indicate where the new node would have to be inserted
Private Function ThreadSetTreeFindPosition( _
    ByRef outParent As Long, ByRef outAsRightHandChild As Boolean, _
    ByRef buf() As ThreadSetNode, ByVal root As Long, ByVal pc As Long, _
    strOffset As Long _
) As Long
    Dim cmp As Long, cur As Long, p As Long, rhc As Boolean
    
    cur = root: p = -1: rhc = False
    Do Until cur = -1
        ' Comparator
        cmp = pc - buf(cur).pc
        If cmp < 0 Then
            p = cur: rhc = False
            cur = buf(cur).rbChild(0)
        ElseIf cmp = 0 Then
            cmp = strOffset - buf(cur).strOffset
            If cmp < 0 Then
                p = cur: rhc = False
                cur = buf(cur).rbChild(0)
            ElseIf cmp = 0 Then
                ThreadSetTreeFindPosition = cur
                Exit Function
            Else
                p = cur: rhc = True
                cur = buf(cur).rbChild(1)
            End If
        Else
            p = cur: rhc = True
            cur = buf(cur).rbChild(1)
        End If
    Loop
    outParent = p: outAsRightHandChild = rhc: ThreadSetTreeFindPosition = -1
End Function

' Algorithm adapted from https://en.wikipedia.org/w/index.php?title=Red%E2%80%93black_tree&oldid=1150140777
Private Sub ThreadSetTreeInsert( _
    ByRef buf() As ThreadSetNode, ByRef outRoot As Long, _
    ByVal newNode As Long, ByVal parent As Long, ByVal asRightHandChild As Boolean _
)
    Dim g As Long, u As Long, p As Long, n As Long, pIsRhc As Boolean
    Dim gg As Long, b As Long, c As Long, x As Long, y As Long, z As Long, nIsRhc As Boolean
    
    With buf(newNode)
        .rbIsBlack = False
        .rbChild(0) = -1
        .rbChild(1) = -1
        .rbParent = parent
    End With
    If parent = -1 Then
        outRoot = newNode
        Exit Sub
    End If
    
    buf(parent).rbChild(-asRightHandChild) = newNode
    
    n = newNode: p = parent
    Do
        If buf(p).rbIsBlack Then Exit Sub
        ' p red
        g = buf(p).rbParent
        If g = -1 Then ' p red and root
            buf(p).rbIsBlack = True
            Exit Sub
        End If
        ' p red and not root (g exists)
        ' u is supposed to refer to the brother of p
        pIsRhc = buf(g).rbChild(1) = p
        u = buf(g).rbChild(1 + pIsRhc)
        If u = -1 Then GoTo ExitWithRotation
        If buf(u).rbIsBlack Then GoTo ExitWithRotation

        ' p and u red, g exists
        buf(p).rbIsBlack = True
        buf(u).rbIsBlack = True
        buf(g).rbIsBlack = False
        n = g
        p = buf(n).rbParent
    Loop Until p = -1
    Exit Sub

ExitWithRotation: ' p red and u black (or does not exist), g exists
    ' For an explanation of the following, see
    '   https://en.wikibooks.org/w/index.php?title=F_Sharp_Programming/Advanced_Data_Structures&oldid=4052491 ,
    '   Section 3.1 ("Red Black Trees"), second diagram (following the sentence "The center tree is the balanced version.").
    nIsRhc = buf(p).rbChild(1) = n
    If pIsRhc = nIsRhc Then ' outer child
        y = p
        If pIsRhc Then
            b = buf(p).rbChild(0): c = buf(n).rbChild(0): x = g: z = n
        Else
            b = buf(n).rbChild(1): c = buf(p).rbChild(1): x = n: z = g
        End If
    Else ' inner child
        y = n: With buf(n): b = .rbChild(0): c = .rbChild(1): End With
        If pIsRhc Then
            x = g: z = p
        Else
            x = p: z = g
        End If
    End If
    
    gg = buf(g).rbParent
    
    With buf(x): .rbIsBlack = False: .rbParent = y: .rbChild(1) = b: End With
    With buf(y): .rbIsBlack = True: .rbParent = gg: .rbChild(0) = x: .rbChild(1) = z: End With
    With buf(z): .rbIsBlack = False: .rbParent = y: .rbChild(0) = c: End With
    
    If b <> -1 Then buf(b).rbParent = x
    If c <> -1 Then buf(c).rbParent = z
    
    If gg = -1 Then
        outRoot = y
    Else
        With buf(gg): .rbChild(-(.rbChild(1) = g)) = y: End With
    End If
End Sub

' returns True if element was actually inserted, False if element existed already
Public Function ThreadSetInsert( _
    ByRef Buffer As ThreadSetBuffer, ByRef root As Long, _
    ByVal pc As Long, strOffset As Long, Optional ByRef outNode As Long _
) As Boolean
    Dim pos As Long, parent As Long, asRhc As Boolean
    With Buffer
        pos = ThreadSetTreeFindPosition( _
            parent, asRhc, Buffer.buf, root, pc, strOffset _
        )
        If pos <> -1 Then outNode = pos: ThreadSetInsert = False: Exit Function
        outNode = .nUsed
    
        If .Capacity = .nUsed Then
            .Capacity = .Capacity + .Capacity \ 2
            ReDim Preserve .buf(0 To .Capacity - 1) As ThreadSetNode
        End If
        With .buf(.nUsed)
            .conditions = 0
            .pc = pc
            .strOffset = strOffset
        End With
        ThreadSetTreeInsert .buf, root, .nUsed, parent, asRhc
        .nUsed = .nUsed + 1
        ThreadSetInsert = True
    End With
End Function


' Explore a state at most once per possible next-character category. Earlier
' paths dominate only the categories they cover; the caller keeps the remainder.
Public Function ThreadSetVisitConditions( _
    ByRef Buffer As ThreadSetBuffer, ByRef root As Long, _
    ByVal pc As Long, ByVal strOffset As Long, ByRef conditions As Long _
) As Boolean
    Dim node As Long
    ThreadSetInsert Buffer, root, pc, strOffset, node
    With Buffer.buf(node)
        conditions = conditions And Not .conditions
        .conditions = .conditions Or conditions
    End With
    ThreadSetVisitConditions = conditions <> 0
End Function
