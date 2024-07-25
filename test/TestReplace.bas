Attribute VB_Name = "TestReplace"
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


    
Private Sub MakeArray(ByRef outAry() As Long, ParamArray p() As Variant)
    ReDim outAry(0 To UBound(p)) As Long
    Dim i As Long
    For i = 0 To UBound(p)
        outAry(i) = p(i)
    Next
End Sub

'@TestMethod("RegexReplace")
Private Sub RegexReplace0001()
    ' nothing
    Dim replacer As String, expected() As Long, parsed As ArrayBuffer.Ty, dummyBytecode() As Long
    
    replacer = ""
    RegexReplace.ParseFormatString parsed, replacer, dummyBytecode, ""
    ReDim Preserve parsed.Buffer(0 To parsed.Length - 1)
    
    MakeArray expected, RegexReplace.REPL_END
    
    Assert.SequenceEquals expected, parsed.Buffer
End Sub

'@TestMethod("RegexReplace")
Private Sub RegexReplace0002()
    ' single dollar
    Dim replacer As String, expected() As Long, parsed As ArrayBuffer.Ty, dummyBytecode() As Long
    
    replacer = "$$"
    RegexReplace.ParseFormatString parsed, replacer, dummyBytecode, ""
    ReDim Preserve parsed.Buffer(0 To parsed.Length - 1)
    
    MakeArray expected, RegexReplace.REPL_DOLLAR, RegexReplace.REPL_END
    
    Assert.SequenceEquals expected, parsed.Buffer
End Sub

'@TestMethod("RegexReplace")
Private Sub RegexReplace0003()
    ' double dollar
    Dim replacer As String, expected() As Long, parsed As ArrayBuffer.Ty, dummyBytecode() As Long
    
    replacer = "$$$$"
    RegexReplace.ParseFormatString parsed, replacer, dummyBytecode, ""
    ReDim Preserve parsed.Buffer(0 To parsed.Length - 1)
    
    MakeArray expected, RegexReplace.REPL_DOLLAR, RegexReplace.REPL_DOLLAR, RegexReplace.REPL_END
    
    Assert.SequenceEquals expected, parsed.Buffer
End Sub

'@TestMethod("RegexReplace")
Private Sub RegexReplace0004()
    ' double dollar with text
    Dim replacer As String, expected() As Long, parsed As ArrayBuffer.Ty, dummyBytecode() As Long
    
    replacer = "abc$$$$xyz"
    RegexReplace.ParseFormatString parsed, replacer, dummyBytecode, ""
    ReDim Preserve parsed.Buffer(0 To parsed.Length - 1)
    
    MakeArray expected, _
        RegexReplace.REPL_SUBSTR, 1, 4, _
        RegexReplace.REPL_DOLLAR, _
        RegexReplace.REPL_SUBSTR, 8, 3, _
        RegexReplace.REPL_END
        
    Assert.SequenceEquals expected, parsed.Buffer
End Sub

'@TestMethod("RegexReplace")
Private Sub RegexReplace0005()
    ' $`
    Dim replacer As String, expected() As Long, parsed As ArrayBuffer.Ty, dummyBytecode() As Long
    
    replacer = "abc$`xyz"
    RegexReplace.ParseFormatString parsed, replacer, dummyBytecode, ""
    ReDim Preserve parsed.Buffer(0 To parsed.Length - 1)
    
    MakeArray expected, _
        RegexReplace.REPL_SUBSTR, 1, 3, _
        RegexReplace.REPL_PREFIX, _
        RegexReplace.REPL_SUBSTR, 6, 3, _
        RegexReplace.REPL_END
        
    Assert.SequenceEquals expected, parsed.Buffer
End Sub

'@TestMethod("RegexReplace")
Private Sub RegexReplace0006()
    ' $&
    Dim replacer As String, expected() As Long, parsed As ArrayBuffer.Ty, dummyBytecode() As Long
    
    replacer = "$&abcxyz"
    RegexReplace.ParseFormatString parsed, replacer, dummyBytecode, ""
    ReDim Preserve parsed.Buffer(0 To parsed.Length - 1)
    
    MakeArray expected, _
        RegexReplace.REPL_ACTUAL, _
        RegexReplace.REPL_SUBSTR, 3, 6, _
        RegexReplace.REPL_END
        
    Assert.SequenceEquals expected, parsed.Buffer
End Sub


'@TestMethod("RegexReplace")
Private Sub RegexReplace0007()
    ' $'
    Dim replacer As String, expected() As Long, parsed As ArrayBuffer.Ty, dummyBytecode() As Long
    
    replacer = "abcxyz$'"
    RegexReplace.ParseFormatString parsed, replacer, dummyBytecode, ""
    ReDim Preserve parsed.Buffer(0 To parsed.Length - 1)
    
    MakeArray expected, _
        RegexReplace.REPL_SUBSTR, 1, 6, _
        RegexReplace.REPL_SUFFIX, _
        RegexReplace.REPL_END
        
    Assert.SequenceEquals expected, parsed.Buffer
End Sub

'@TestMethod("RegexReplace")
Private Sub RegexReplace0008()
    ' $~ and nothing else
    Dim replacer As String, expected() As Long, parsed As ArrayBuffer.Ty, dummyBytecode() As Long
    
    replacer = "$~"
    RegexReplace.ParseFormatString parsed, replacer, dummyBytecode, ""
    ReDim Preserve parsed.Buffer(0 To parsed.Length - 1)
    
    MakeArray expected, RegexReplace.REPL_END
    
    Assert.SequenceEquals expected, parsed.Buffer
End Sub

'@TestMethod("RegexReplace")
Private Sub RegexReplace0009()
    ' $~
    Dim replacer As String, expected() As Long, parsed As ArrayBuffer.Ty, dummyBytecode() As Long
    
    replacer = "$~abc$~xyz$~uvw$~"
    RegexReplace.ParseFormatString parsed, replacer, dummyBytecode, ""
    ReDim Preserve parsed.Buffer(0 To parsed.Length - 1)
    
    MakeArray expected, _
        RegexReplace.REPL_SUBSTR, 3, 3, _
        RegexReplace.REPL_SUBSTR, 8, 3, _
        RegexReplace.REPL_SUBSTR, 13, 3, _
        RegexReplace.REPL_END
        
    Assert.SequenceEquals expected, parsed.Buffer
End Sub

'@TestMethod("RegexReplace")
Private Sub RegexReplace0010()
    ' $n
    Dim replacer As String, expected() As Long, parsed As ArrayBuffer.Ty, dummyBytecode() As Long
    
    replacer = "abc$12xyz"
    RegexReplace.ParseFormatString parsed, replacer, dummyBytecode, ""
    ReDim Preserve parsed.Buffer(0 To parsed.Length - 1)
    
    MakeArray expected, _
        RegexReplace.REPL_SUBSTR, 1, 3, _
        RegexReplace.REPL_NUMBERED, 12, _
        RegexReplace.REPL_SUBSTR, 7, 3, _
        RegexReplace.REPL_END
        
    Assert.SequenceEquals expected, parsed.Buffer
End Sub

'@TestMethod("RegexReplace")
Private Sub RegexReplace0020()
    ' $<identifier>
    Dim replacer As String, pattern As String, expected() As Long, parsed As ArrayBuffer.Ty, dummyBytecode() As Long
    
    replacer = "abc$<z>$<y>$<x>$<w>abc"
    pattern = "xzyw"
    MakeArray dummyBytecode, _
        0, 4, 0, _
        4, 1, 10, _
        1, 1, 11, _
        3, 1, 12, _
        2, 1, 13
    RegexReplace.ParseFormatString parsed, replacer, dummyBytecode, pattern
    ReDim Preserve parsed.Buffer(0 To parsed.Length - 1)
    
    MakeArray expected, _
        RegexReplace.REPL_SUBSTR, 1, 3, _
        RegexReplace.REPL_NAMED, 13, _
        RegexReplace.REPL_NAMED, 12, _
        RegexReplace.REPL_NAMED, 11, _
        RegexReplace.REPL_NAMED, 10, _
        RegexReplace.REPL_SUBSTR, 20, 3, _
        RegexReplace.REPL_END
        
    Assert.SequenceEquals expected, parsed.Buffer
End Sub

