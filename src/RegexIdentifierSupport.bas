Attribute VB_Name = "RegexIdentifierSupport"
Option Explicit

Public Type StartLengthPair
    start As Long
    Length As Long
End Type

Public Type IdentifierTreeNode
    rbParent As Long
    rbChild(0 To 1) As Long ' 0 = left, 1 = right
    rbIsBlack As Boolean
    reference As StartLengthPair
    identifierId As Long
End Type

Public Type IdentifierTreeTy
    nEntries As Long
    root As Long
    bufferCapacity As Long
    Buffer() As IdentifierTreeNode
End Type

Public Sub RedBlackDumpTree(ByRef target() As Long, ByVal targetStartIdx As Long, ByRef tree As IdentifierTreeTy)
    Dim targetIdx As Long, currentNode As Long, nextNode As Long
    
    If tree.nEntries = 0 Then Exit Sub
    
    targetIdx = targetStartIdx
    
    currentNode = tree.root
    Do
        nextNode = tree.Buffer(currentNode).rbChild(0)
        If nextNode = -1 Then Exit Do
        currentNode = nextNode
    Loop
    
    Do
        ' handle current node
        target(targetIdx) = tree.Buffer(currentNode).reference.start: targetIdx = targetIdx + 1
        target(targetIdx) = tree.Buffer(currentNode).reference.Length: targetIdx = targetIdx + 1
        target(targetIdx) = tree.Buffer(currentNode).identifierId: targetIdx = targetIdx + 1
        
        nextNode = tree.Buffer(currentNode).rbChild(1)
        If nextNode <> -1 Then
            ' right-hand child exists
            ' -> go to leftmost descendant of right-hand child
            currentNode = nextNode
            Do
                nextNode = tree.Buffer(currentNode).rbChild(0)
                If nextNode = -1 Then Exit Do
                currentNode = nextNode
            Loop
        Else
            ' right-hand child does not exist
            ' -> go to first ancestor for which our subtree is the left-hand subtree
            Do
                nextNode = tree.Buffer(currentNode).rbParent
                If nextNode = -1 Then Exit Sub
                If currentNode = tree.Buffer(nextNode).rbChild(0) Then Exit Do
                currentNode = nextNode
            Loop
            currentNode = nextNode
        End If
    Loop
End Sub

Public Function RedBlackFindOrInsert(ByRef lake As String, ByRef tree As IdentifierTreeTy, ByRef vReference As StartLengthPair) As Long
    Dim parent As Long, found As Long
    Dim asRightHandChild As Boolean
    
    found = RedBlackFindPosition(parent, asRightHandChild, lake, tree, vReference)
    If found = -1 Then
        If tree.nEntries = tree.bufferCapacity Then
            tree.bufferCapacity = tree.bufferCapacity + (tree.bufferCapacity + 16) \ 2
            ReDim Preserve tree.Buffer(0 To tree.bufferCapacity - 1) As IdentifierTreeNode
        End If
        With tree.Buffer(tree.nEntries)
            .reference = vReference
            .identifierId = tree.nEntries
        End With
        RedBlackInsert tree, tree.nEntries, parent, asRightHandChild
        RedBlackFindOrInsert = tree.nEntries
        tree.nEntries = tree.nEntries + 1
    Else
        RedBlackFindOrInsert = tree.Buffer(found).identifierId
    End If
End Function


Public Function RedBlackComparator(ByRef lake As String, ByRef v1 As StartLengthPair, ByRef v2 As StartLengthPair) As Long
    If v1.Length < v2.Length Then RedBlackComparator = -1: Exit Function
    If v1.Length > v2.Length Then RedBlackComparator = 1: Exit Function
    RedBlackComparator = StrComp( _
        Mid$(lake, v1.start, v1.Length), _
        Mid$(lake, v2.start, v2.Length), _
        vbBinaryCompare _
    )
End Function

' returns
' * index of node if node was found; in that case out parameters will remain unchanged
' * -1 otherwise, then the out parameters will indicate where the new node would have to be inserted
Public Function RedBlackFindPosition( _
    ByRef outParent As Long, ByRef outAsRightHandChild As Boolean, _
    ByRef lake As String, ByRef tree As IdentifierTreeTy, _
    ByRef vReference As StartLengthPair _
) As Long
    Dim cmp As Long, cur As Long, p As Long, rhc As Boolean
    
    cur = tree.root: p = -1: rhc = False
    Do Until cur = -1
        cmp = RedBlackComparator(lake, vReference, tree.Buffer(cur).reference)
        If cmp < 0 Then
            p = cur: rhc = False
            cur = tree.Buffer(cur).rbChild(0)
        ElseIf cmp = 0 Then
            RedBlackFindPosition = cur
            Exit Function
        Else
            p = cur: rhc = True
            cur = tree.Buffer(cur).rbChild(1)
        End If
    Loop
    outParent = p: outAsRightHandChild = rhc: RedBlackFindPosition = -1
End Function

' Algorithm adapted from https://en.wikipedia.org/w/index.php?title=Red%E2%80%93black_tree&oldid=1150140777
Public Sub RedBlackInsert( _
    ByRef tree As IdentifierTreeTy, _
    ByVal newNode As Long, ByVal parent As Long, ByVal asRightHandChild As Boolean _
)
    Dim g As Long, u As Long, p As Long, n As Long, pIsRhc As Boolean
    Dim gg As Long, b As Long, c As Long, x As Long, y As Long, z As Long, nIsRhc As Boolean
    
    With tree.Buffer(newNode)
        .rbIsBlack = False
        .rbChild(0) = -1
        .rbChild(1) = -1
        .rbParent = parent
    End With
    If parent = -1 Then
        tree.root = newNode
        Exit Sub
    End If
    
    tree.Buffer(parent).rbChild(-asRightHandChild) = newNode
    
    n = newNode: p = parent
    Do
        If tree.Buffer(p).rbIsBlack Then Exit Sub
        ' p red
        g = tree.Buffer(p).rbParent
        If g = -1 Then ' p red and root
            tree.Buffer(p).rbIsBlack = True
            Exit Sub
        End If
        ' p red and not root (g exists)
        ' u is supposed to refer to the brother of p
        pIsRhc = tree.Buffer(g).rbChild(1) = p
        u = tree.Buffer(g).rbChild(1 + pIsRhc)
        If u = -1 Then GoTo ExitWithRotation
        If tree.Buffer(u).rbIsBlack Then GoTo ExitWithRotation

        ' p and u red, g exists
        tree.Buffer(p).rbIsBlack = True
        tree.Buffer(u).rbIsBlack = True
        tree.Buffer(g).rbIsBlack = False
        n = g
        p = tree.Buffer(n).rbParent
    Loop Until p = -1
    Exit Sub

ExitWithRotation: ' p red and u black (or does not exist), g exists
    ' For an explanation of the following, see
    '   https://en.wikibooks.org/w/index.php?title=F_Sharp_Programming/Advanced_Data_Structures&oldid=4052491 ,
    '   Section 3.1 ("Red Black Trees"), second diagram (following the sentence "The center tree is the balanced version.").
    nIsRhc = tree.Buffer(p).rbChild(1) = n
    If pIsRhc = nIsRhc Then ' outer child
        y = p
        If pIsRhc Then
            b = tree.Buffer(p).rbChild(0): c = tree.Buffer(n).rbChild(0): x = g: z = n
        Else
            b = tree.Buffer(n).rbChild(1): c = tree.Buffer(p).rbChild(1): x = n: z = g
        End If
    Else ' inner child
        y = n: With tree.Buffer(n): b = .rbChild(0): c = .rbChild(1): End With
        If pIsRhc Then
            x = g: z = p
        Else
            x = p: z = g
        End If
    End If
    
    gg = tree.Buffer(g).rbParent
    
    With tree.Buffer(x): .rbIsBlack = False: .rbParent = y: .rbChild(1) = b: End With
    With tree.Buffer(y): .rbIsBlack = True: .rbParent = gg: .rbChild(0) = x: .rbChild(1) = z: End With
    With tree.Buffer(z): .rbIsBlack = False: .rbParent = y: .rbChild(0) = c: End With
    
    If b <> -1 Then tree.Buffer(b).rbParent = x
    If c <> -1 Then tree.Buffer(c).rbParent = z
    
    If gg = -1 Then
        tree.root = y
    Else
        With tree.Buffer(gg): .rbChild(-(.rbChild(1) = g)) = y: End With
    End If
End Sub

