Attribute VB_Name = "MainModule"
Option Explicit

Private Declare Sub ExitProcess Lib "kernel32" (ByVal exitCode As Long)

Sub Main()
    Dim inFilePath As String, outFilePath As String, engineName As String
    Dim testName As String, testStrs() As String, testRegexs() As String
    Dim resultSb As StaticStringBuilder.Ty
    Dim errorMessage As String
    On Error GoTo Failed
    
    ParseCommandLineArguments inFilePath, outFilePath, engineName
    
    ReadTestDataFromFile testName, testStrs, testRegexs, inFilePath
    
    RunTests resultSb, testStrs, testRegexs, engineName
    
    WriteResultsToFile outFilePath, StaticStringBuilder.GetStr(resultSb)
    Exit Sub
Failed:
    errorMessage = CStr(Err.Number) & ": " & Err.Description
    On Error Resume Next
    If Len(outFilePath) > 0 Then WriteTextFile outFilePath & ".error.txt", errorMessage
    ExitProcess 1
End Sub

Sub ParseCommandLineArguments(ByRef inFileName As String, ByRef outFileName As String, ByRef engineName As String)
    Dim col As Collection
    Set col = CommandLineParser.ParseCommandLine(Command$)
    If col Is Nothing Then Err.Raise 5
    If col.Count < 2 Or col.Count > 3 Then Err.Raise 5
    inFileName = col(1)
    outFileName = col(2)
    engineName = "dfs"
    If col.Count = 3 Then engineName = LCase$(col(3))
End Sub

Function ReadTextFile(ByRef filePath As String)
    Dim fso As Object, objText As Object, Text As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objText = fso.OpenTextFile(filePath)
    Text = objText.ReadAll()
    objText.Close
    
    Set objText = Nothing
    Set fso = Nothing
    
    ReadTextFile = Text
End Function

Sub WriteTextFile(ByRef filePath As String, Text As String)
    Dim fso As Object, objText As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objText = fso.OpenTextFile(filePath, 2, True) ' 2 = for writing, True = permit creation
    objText.Write Text
    objText.Close
    
    Set objText = Nothing
    Set fso = Nothing
End Sub

Sub ReadTestDataFromFile( _
    ByRef outName As String, ByRef outStrs() As String, ByRef outRegexs() As String, _
    ByRef filePath As String _
)
    Dim Text As String
    Dim jsonData As Object
    Dim cnt As Long, i As Long, s As Variant
    
    Text = ReadTextFile(filePath)
    Set jsonData = JSON.Parse(Text)
    
    outName = jsonData("name")
    
    cnt = jsonData("strs").Count
    ReDim outStrs(0 To cnt - 1) As String
    i = 0
    For Each s In jsonData("strs")
        outStrs(i) = s
        i = i + 1
    Next
    
    cnt = jsonData("regexs").Count
    ReDim outRegexs(0 To cnt - 1) As String
    i = 0
    For Each s In jsonData("regexs")
        outRegexs(i) = s
        i = i + 1
    Next
End Sub

Sub WriteResultsToFile(ByRef outFilePath As String, ByRef resultString As String)
    WriteTextFile outFilePath, resultString
End Sub

Private Sub RunTests( _
    ByRef resultSb As StaticStringBuilder.Ty, ByRef testStrs() As String, _
    ByRef testRegexs() As String, ByRef engineName As String _
)
    Dim lRegexs As Long, uRegexs As Long, i As Long
    Dim res() As Long
    
    Dim re As IRegexEngine
    Dim matcher As MatcherRegexEngine
    
    lRegexs = LBound(testRegexs): uRegexs = UBound(testRegexs)
    
    StaticStringBuilder.AppendStr resultSb, "["
    
    i = lRegexs
    Do
        ' Keep each pattern's matcher state independent; reuse only across its strings.
        Set matcher = New MatcherRegexEngine
        matcher.EngineName = engineName
        Set re = matcher
        If re.Compile(testRegexs(i)) Then
            re.Match res, testStrs
            MatrixToJson resultSb, res
        Else
            StaticStringBuilder.AppendStr resultSb, "null" ' Unsupported by this engine.
        End If
                
        If i = uRegexs Then Exit Do
        
        StaticStringBuilder.AppendStr resultSb, ","
        
        i = i + 1
    Loop
    
    StaticStringBuilder.AppendStr resultSb, vbCrLf
    StaticStringBuilder.AppendStr resultSb, "]"
    StaticStringBuilder.AppendStr resultSb, vbCrLf
End Sub

Sub MatrixToJson(ByRef resultSb As StaticStringBuilder.Ty, ByRef m() As Long)
    Dim low1 As Long, up1 As Long, low2 As Long, up2 As Long
    Dim i1 As Long, i2 As Long
    
    low1 = LBound(m, 1): up1 = UBound(m, 1)
    low2 = LBound(m, 2): up2 = UBound(m, 2)
    
    StaticStringBuilder.AppendStr resultSb, vbCrLf
    StaticStringBuilder.AppendStr resultSb, "["
    StaticStringBuilder.AppendStr resultSb, vbCrLf
    
    i1 = low1
    Do
        StaticStringBuilder.AppendStr resultSb, "["
    
        i2 = low2
        Do
            StaticStringBuilder.AppendStr resultSb, CStr(m(i1, i2))
            If i2 = up2 Then Exit Do
            StaticStringBuilder.AppendStr resultSb, ", "
            i2 = i2 + 1
        Loop
    
        StaticStringBuilder.AppendStr resultSb, "]"
    
        If i1 = up1 Then Exit Do
        
        StaticStringBuilder.AppendStr resultSb, ","
        StaticStringBuilder.AppendStr resultSb, vbCrLf
        
        i1 = i1 + 1
    Loop
    
    StaticStringBuilder.AppendStr resultSb, vbCrLf
    StaticStringBuilder.AppendStr resultSb, "]"
End Sub

