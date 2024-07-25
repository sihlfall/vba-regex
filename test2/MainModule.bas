Attribute VB_Name = "MainModule"
Option Explicit

' see https://classicvb.net/samples/Console/


Sub Main()
    Dim sName As String
    Dim fColor As Long, bColor As Long
    Dim sCaption As String
    Dim inFilePath As String, outFilePath As String
    Dim testName As String, testStrs() As String, testRegexs() As String
    Dim resultSb As StaticStringBuilder.Ty
    
    ParseCommandLineArguments inFilePath, outFilePath
    'inFilePath = "C:\Users\Johannes\Documents\git\regex-test-cases\test-cases\EgrepLiterals_FoldCase.json"
    'outFilePath = "C:\Users\Johannes\Documents\git\vba-regex-dev\test2\testout.json"
    
    ReadTestDataFromFile testName, testStrs, testRegexs, inFilePath
    
    RunTests resultSb, testStrs, testRegexs
    
    WriteResultsToFile outFilePath, StaticStringBuilder.GetStr(resultSb)
End Sub

Sub ParseCommandLineArguments(ByRef inFileName As String, ByRef outFileName As String)
    Dim col As Collection
    Set col = CommandLineParser.ParseCommandLine(Command$)
    If col.Count <> 2 Then Err.Raise 5000
    inFileName = col(1)
    outFileName = col(2)
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

Private Sub RunTests(ByRef resultSb As StaticStringBuilder.Ty, ByRef testStrs() As String, ByRef testRegexs() As String)
    Dim lStrs As Long, uStrs As Long, lRegexs As Long, uRegexs As Long
    Dim i As Long, j As Long
    Dim res() As Long
    
    Dim re As IRegexEngine
    Set re = New DfsRegexEngine
    
    lStrs = LBound(testStrs): uStrs = UBound(testStrs)
    lRegexs = LBound(testRegexs): uRegexs = UBound(testRegexs)
    
    StaticStringBuilder.AppendStr resultSb, "["
    
    i = lRegexs
    Do
        re.Compile (testRegexs(i))
        re.Match res, testStrs
        MatrixToJson resultSb, res
                
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

