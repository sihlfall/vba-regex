VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DfsRegexEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IRegexEngine

Private bytecode() As Long

Public Sub IRegexEngine_Compile(ByRef pattern As String)
    
    RegexCompiler.Compile bytecode, pattern
    
    'mRmatch.recursionLimit = 100
End Sub

Public Sub IRegexEngine_Match(ByRef outResult() As Long, ByRef strs() As String)
    Dim matcherResult As Long, lStrs As Long, uStrs As Long, nStrs As Long, nSaved As Long, i As Long, j As Long, captures As RegexDfsMatcher.CapturesTy
    
    lStrs = LBound(strs)
    uStrs = UBound(strs)
    nStrs = uStrs - lStrs + 1
    
     ' Todo: Hack
    nSaved = bytecode(0) + 1
    
    ReDim outResult(0 To nStrs - 1, 0 To nSaved - 1) As Long
    
    For i = lStrs To uStrs
        matcherResult = -2
        
        On Error Resume Next
        matcherResult = RegexDfsMatcher.DfsMatch(captures, bytecode, strs(i))
        On Error GoTo 0
        
        If matcherResult = -2 Then
            For j = 0 To nSaved - 1: outResult(i - lStrs, j) = -2: Next
        ElseIf matcherResult = -1 Then
            For j = 0 To nSaved - 1: outResult(i - lStrs, j) = -1: Next
        Else
            outResult(i - lStrs, 0) = captures.entireMatch.start - 1
            outResult(i - lStrs, 1) = captures.entireMatch.start + captures.entireMatch.Length - 1
            For j = 2 To nSaved - 1 Step 2
                outResult(i - lStrs, j) = captures.numberedCaptures(j \ 2 - 1).start - 1
                outResult(i - lStrs, j + 1) = captures.numberedCaptures(j \ 2 - 1).start + captures.numberedCaptures(j \ 2 - 1).Length - 1
            Next
        End If
    Next
End Sub

Private Sub CopyBytecode(ByRef dst() As Long, ByRef src As ArrayBuffer.Ty)
    Dim i As Long, u As Long
    u = src.Length - 1
    ReDim dst(0 To u) As Long
    
    For i = 0 To u
        dst(i) = src.Buffer(i)
    Next
End Sub

