Attribute VB_Name = "RegexNfaThreadPool"
Option Explicit

Public Const NFA_TS_FIRST_ACTUAL As Long = 6

Private Const DEFAULT_MINIMUM_THREADSTACK_CAPACITY As Long = 16

Private Const NFA_TS_UNUSED_SENTINEL_B As Long = 0
Private Const NFA_TS_UNUSED_SENTINEL_E As Long = 1

Public Type NfaThread
    pc As Long
    parent As Long
    qstack As ArrayBuffer.Ty
    saved() As Long
    prev As Long
    nxt As Long
    
    rbParent As Long
    rbChild(0 To 1) As Long ' 0 = left, 1 = right
    rbIsBlack As Boolean
    
    processed As Boolean
End Type

Public Type ThreadPool
    tsCapacity As Long
    tsStack() As NfaThread
    tsActive As Long ' 0 or 1
    rbTreeRoot(0 To 1) As Long
End Type


Public Sub Initialize(ByRef pool As ThreadPool)
    Dim i As Long
    With pool
        .tsCapacity = DEFAULT_MINIMUM_THREADSTACK_CAPACITY
        ReDim .tsStack(0 To .tsCapacity - 1) As NfaThread
        
        For i = NFA_TS_FIRST_ACTUAL To .tsCapacity - 1
            With .tsStack(i)
                .prev = i - 1: .nxt = i + 1
            End With
        Next
        
        .tsStack(NFA_TS_UNUSED_SENTINEL_B).nxt = NFA_TS_FIRST_ACTUAL
        .tsStack(NFA_TS_UNUSED_SENTINEL_E).prev = .tsCapacity - 1
        .tsStack(NFA_TS_FIRST_ACTUAL).prev = NFA_TS_UNUSED_SENTINEL_B
        .tsStack(.tsCapacity - 1).nxt = NFA_TS_UNUSED_SENTINEL_E
        
        .tsStack(2).nxt = 3
        .tsStack(3).prev = 2
        .tsStack(4).nxt = 5
        .tsStack(5).prev = 4
        
        .rbTreeRoot(0) = -1: .rbTreeRoot(1) = -1
    End With
End Sub

Private Sub IncreaseCapacity(ByRef pool As ThreadPool)
    Dim oldCapacity As Long, newCapacity As Long, i As Long
    With pool
        oldCapacity = .tsCapacity
        newCapacity = oldCapacity + oldCapacity \ 2
        ReDim Preserve .tsStack(0 To newCapacity - 1) As NfaThread
        With .tsStack(oldCapacity): .prev = pool.tsStack(NFA_TS_UNUSED_SENTINEL_E).prev: pool.tsStack(.prev).nxt = oldCapacity: .nxt = oldCapacity + 1: End With
        i = oldCapacity + 1
        Do
            .tsStack(i).prev = i - 1
            If i = newCapacity - 1 Then Exit Do
            i = i + 1
            .tsStack(i).nxt = i
        Loop
        .tsStack(i).nxt = NFA_TS_UNUSED_SENTINEL_E
        .tsStack(NFA_TS_UNUSED_SENTINEL_E).prev = i
    End With
End Sub

Public Sub AddFirst(ByRef pool As ThreadPool, ByVal pc As Long, ByVal maxSave As Long)
    ' assert: There is at least one actual unused block.
    
    Dim i As Long, u As Long, idx As Long, nxt As Long, prev As Long
    Const FIRST_THREAD As Long = 0
    
    With pool
        .tsActive = 0
        idx = .tsStack(NFA_TS_UNUSED_SENTINEL_B).nxt
        With .tsStack(idx): prev = .prev: nxt = .nxt: End With
        .tsStack(prev).nxt = nxt
        .tsStack(nxt).prev = prev
        
        .tsStack(2).nxt = idx
        .tsStack(3).prev = idx
        With .tsStack(idx): .prev = 2: .nxt = 3: End With
        
        With .tsStack(idx)
            .pc = pc: .parent = -1
            ReDim .saved(0 To maxSave) As Long
            For i = 0 To maxSave: .saved(i) = -1: Next
        End With
        
        RedBlackInsert .tsStack, .rbTreeRoot(.tsActive), idx, .rbTreeRoot(.tsActive), False
    End With
End Sub

Public Sub InsertAfterCurrent(ByRef pool As ThreadPool, ByVal current As Long, _
    ByVal pc As Long, ByRef qstack As ArrayBuffer.Ty, ByRef saved() As Long _
)
    Dim existingThread As Long, outParent As Long, outAsRightHandChild As Boolean
    Dim newIdx As Long, nxt As Long, prev As Long
    
    With pool
        ' TODO: Consider qstack as well
        existingThread = RedBlackFindPosition(outParent, outAsRightHandChild, .tsStack, .rbTreeRoot(.tsActive), pc)
        If existingThread <> -1 Then
            If .tsStack(existingThread).processed Then Exit Sub ' Higher-priority thread already exists
            
            ' Lower-priority thread already exists -- remove from priority list
            newIdx = existingThread
            
            With .tsStack(newIdx): prev = .prev: nxt = .nxt: End With
            .tsStack(prev).nxt = nxt: .tsStack(nxt).prev = prev
        Else
            If .tsStack(NFA_TS_UNUSED_SENTINEL_B).nxt = NFA_TS_UNUSED_SENTINEL_E Then IncreaseCapacity pool
        
            ' Remove first unused block from the "unused" list
            newIdx = .tsStack(NFA_TS_UNUSED_SENTINEL_B).nxt
            With .tsStack(newIdx): prev = .prev: nxt = .nxt: End With
            .tsStack(prev).nxt = nxt: .tsStack(nxt).prev = prev
        
            ' Insert new block into tree
            RedBlackInsert .tsStack, .rbTreeRoot(.tsActive), newIdx, outParent, outAsRightHandChild
        End If
        
        ' Now the insert the new block after the current one in the priority list
        nxt = .tsStack(current).nxt
        With .tsStack(newIdx): .prev = current: .nxt = nxt: End With
        .tsStack(current).nxt = newIdx: .tsStack(nxt).prev = newIdx
    
        ' Set contents
        With .tsStack(newIdx)
            .pc = pc
            .processed = False
            .saved = saved
            .qstack = qstack
        End With
    End With
End Sub

Public Sub AppendToInactive(ByRef pool As ThreadPool, _
    ByVal pc As Long, ByRef qstack As ArrayBuffer.Ty, ByRef saved() As Long _
)
    Dim existingThread As Long, outParent As Long, outAsRightHandChild As Boolean
    Dim newIdx As Long, nxt As Long, prev As Long
    
    With pool
        ' TODO: Consider qstack as well
        existingThread = RedBlackFindPosition(outParent, outAsRightHandChild, .tsStack, .rbTreeRoot(1 - .tsActive), pc)
        If existingThread <> -1 Then Exit Sub ' Thread already exists
            
        If .tsStack(NFA_TS_UNUSED_SENTINEL_B).nxt = NFA_TS_UNUSED_SENTINEL_E Then IncreaseCapacity pool
        
        ' Remove first unused block from the "unused" list
        newIdx = .tsStack(NFA_TS_UNUSED_SENTINEL_B).nxt
        With .tsStack(newIdx): prev = .prev: nxt = .nxt: End With
        .tsStack(prev).nxt = nxt: .tsStack(nxt).prev = prev
            
        ' Insert new block into tree
        RedBlackInsert .tsStack, .rbTreeRoot(1 - .tsActive), newIdx, outParent, outAsRightHandChild
            
        ' Now insert at the end of the non-active list.
        nxt = 2 + 2 * (1 - .tsActive) + 1 ' end sentinel
        prev = .tsStack(nxt).prev
        With .tsStack(newIdx): .prev = prev: .nxt = nxt: End With
        .tsStack(prev).nxt = newIdx: .tsStack(nxt).prev = newIdx
        
        ' Set contents
        With .tsStack(newIdx)
            .pc = pc
            .processed = False
            .saved = saved
            .qstack = qstack
        End With
    End With
End Sub

Public Sub ClearNonActive(ByRef pool As ThreadPool)
    ' rbTreeRoot of the non-active list has to be cleared
    '
    ' Moreover:
    ' We have two doubly linked lists:
    '    u: NFA_TS_UNUSED_SENTINEL_B <-> uFirst ( <-> [...u...] )
    '    v: vSentinelB <-> vFirst <-> [...v...] <-> vLast <-> vSentinelE
    ' uFirst = NFA_TS_UNUSED_SENTINEL_E is possible.
    ' vFirst = vLast is possible.
    ' v might look like vSentinelB <-> vSentinelE, which is a special case.
    '
    ' We want to end up with:
    '    NFA_TS_UNUSED_SENTINEL_B <-> vFirst <-> [...v...] <-> vLast <-> uFirst ( <-> [...u...])
    '    vSentinelB <-> vSentinelE
    
    Dim uFirst As Long, vSentinelB As Long, vSentinelE As Long, vFirst As Long, vLast As Long
    
    With pool
        .rbTreeRoot(1 - .tsActive) = -1
            
        vSentinelB = 2 + 2 * (1 - .tsActive)
        vSentinelE = vSentinelB + 1
            
        vFirst = .tsStack(vSentinelB).nxt
        If vFirst = vSentinelE Then Exit Sub
        vLast = .tsStack(vSentinelE).prev
        
        uFirst = .tsStack(NFA_TS_UNUSED_SENTINEL_B).nxt
        
        .tsStack(NFA_TS_UNUSED_SENTINEL_B).nxt = vFirst
        .tsStack(vFirst).prev = NFA_TS_UNUSED_SENTINEL_B
        .tsStack(vLast).nxt = uFirst
        .tsStack(uFirst).prev = vLast
        
        .tsStack(vSentinelB).nxt = vSentinelE
        .tsStack(vSentinelE).prev = vSentinelB
    End With
End Sub

' returns
' * index of node if node was found; in that case out parameters will remain unchanged
' * -1 otherwise, then the out parameters will indicate where the new node would have to be inserted
Private Function RedBlackFindPosition(ByRef outParent As Long, ByRef outAsRightHandChild As Boolean, ByRef buf() As NfaThread, ByVal root As Long, ByVal pc As Long) As Long
    Dim cmp As Long, cur As Long, p As Long, rhc As Boolean
    
    cur = root: p = -1: rhc = False
    Do Until cur = -1
        cmp = pc - buf(cur).pc
        If cmp < 0 Then
            p = cur: rhc = False
            cur = buf(cur).rbChild(0)
        ElseIf cmp = 0 Then
            RedBlackFindPosition = cur
            Exit Function
        Else
            p = cur: rhc = True
            cur = buf(cur).rbChild(1)
        End If
    Loop
    outParent = p: outAsRightHandChild = rhc: RedBlackFindPosition = -1
End Function

' Algorithm adapted from https://en.wikipedia.org/w/index.php?title=Red%E2%80%93black_tree&oldid=1150140777
Private Sub RedBlackInsert(ByRef buf() As NfaThread, ByRef outRoot As Long, ByVal newNode As Long, ByVal parent As Long, ByVal asRightHandChild As Boolean)
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
    
    With buf(x): .rbIsBlack = True: .rbParent = y: .rbChild(1) = b: End With
    With buf(y): .rbIsBlack = False: .rbParent = gg: .rbChild(0) = x: .rbChild(1) = z: End With
    With buf(z): .rbIsBlack = True: .rbParent = y: .rbChild(0) = c: End With
    
    If gg = -1 Then
        outRoot = y
    Else
        With buf(gg): .rbChild(-(.rbChild(1) = g)) = y: End With
    End If
End Sub


