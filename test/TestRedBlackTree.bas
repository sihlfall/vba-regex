Attribute VB_Name = "TestRedBlackTree"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub



Private Function AppendNode(ByRef buf As LongArrayBuffer.Ty, ByRef s As String) As Long
    Dim i As Long, pos As Long
    
    pos = buf.length
    LongArrayBuffer.AppendFive buf, -42, -42, -42, -42, -42
    LongArrayBuffer.AppendLong buf, Len(s)
    For i = 1 To Len(s)
        LongArrayBuffer.AppendLong buf, AscW(Mid$(s, i, 1))
    Next
    AppendNode = pos
End Function

'Private Const OFFSET_PARENT As Long = 0
'Private Const OFFSET_CHILD As Long = 1
'Private Const OFFSET_IS_BLACK As Long = 3
'Private Const OFFSET_STR_START As Long = 4
'Private Const OFFSET_STR_LEN As Long = 5
'Private Const OFFSET_STR_CHARACTERS As Long = 6

'@TestMethod("RedBlackTree")
Private Sub RedBlackTree0001()
    Dim pos As Long, parent As Long, root As Long, firstNode As Long, secondNode As Long, thirdNode As Long
    Dim asRightHandChild As Boolean
    Dim buf As LongArrayBuffer.Ty
    
    On Error GoTo TestFail
    
    root = -1
    firstNode = AppendNode(buf, "aa")
        
    If RegexRedBlackTree.RedBlackFindPosition(parent, asRightHandChild, buf.buffer, root, firstNode) = -1 Then
        RegexRedBlackTree.RedBlackInsert buf.buffer, root, firstNode, parent, asRightHandChild
    Else
        buf.length = firstNode
    End If
    Assert.AreEqual 0&, root
    Assert.AreEqual -1&, buf.buffer(firstNode + 0)
    Assert.AreEqual -1&, buf.buffer(firstNode + 1)
    Assert.AreEqual -1&, buf.buffer(firstNode + 2)
    
    secondNode = AppendNode(buf, "a")
        
    If RegexRedBlackTree.RedBlackFindPosition(parent, asRightHandChild, buf.buffer, root, secondNode) = -1 Then
        RegexRedBlackTree.RedBlackInsert buf.buffer, root, secondNode, parent, asRightHandChild
    Else
        buf.length = secondNode
    End If
    Assert.AreEqual 0&, root
    Assert.AreEqual -1&, buf.buffer(firstNode + 0)
    Assert.AreEqual secondNode, buf.buffer(firstNode + 1)
    Assert.AreEqual -1&, buf.buffer(firstNode + 2)
    Assert.AreEqual firstNode, buf.buffer(secondNode + 0)
    Assert.AreEqual -1&, buf.buffer(secondNode + 1)
    Assert.AreEqual -1&, buf.buffer(secondNode + 2)
    
    thirdNode = AppendNode(buf, "aa")
        
    Assert.IsTrue RegexRedBlackTree.RedBlackFindPosition(parent, asRightHandChild, buf.buffer, root, thirdNode) <> -1
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
