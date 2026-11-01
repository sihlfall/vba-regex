VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MatcherRegexEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IRegexEngine

Private bytecode() As Long
Private dfaState As RegexDfaMatcher.StateStructure
Public EngineName As String

Public Function IRegexEngine_Compile(ByRef pattern As String) As Boolean
    Select Case EngineName
    Case "dfs", "nfa", "dfa"
    Case Else
        Err.Raise 5, "MatcherRegexEngine", "Unknown matcher: " & EngineName
    End Select
    RegexCompiler.Compile bytecode, pattern
    IRegexEngine_Compile = EngineName = "dfs" Or _
        (bytecode(RegexBytecode.BYTECODE_IDX_PROPERTIES) And RegexBytecode.BYTECODE_PROPERTY_REQUIRES_DFS_MATCHING) = 0
    If EngineName = "dfa" Then RegexDfaMatcher.DfaStateStructureInitialize dfaState
End Function

Public Sub IRegexEngine_Match(ByRef outResult() As Long, ByRef strs() As String)
    Dim matcherResult As Long, lStrs As Long, uStrs As Long, nStrs As Long, nSaved As Long, i As Long, j As Long, captures As RegexMatcherCommon.CapturesTy
    Dim context As RegexDfsMatcher.DfsMatcherContext
    
    lStrs = LBound(strs)
    uStrs = UBound(strs)
    nStrs = uStrs - lStrs + 1
    
    If EngineName = "dfa" Then
        nSaved = 2 ' DFA exposes the complete match span, without capture groups.
    Else
        nSaved = bytecode(RegexBytecode.BYTECODE_IDX_MAX_PROPER_CAPTURE_SLOT) + 1
    End If
    
    ReDim outResult(0 To nStrs - 1, 0 To nSaved - 1) As Long
    
    For i = lStrs To uStrs
        matcherResult = -2
        
        On Error Resume Next
        matcherResult = MatchOne(context, captures, strs(i))
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

Private Function MatchOne( _
    ByRef context As RegexDfsMatcher.DfsMatcherContext, _
    ByRef captures As RegexMatcherCommon.CapturesTy, ByRef text As String _
) As Long
    Dim span As RegexMatcherCommon.StartLengthPair
    Select Case EngineName
    Case "dfs"
        MatchOne = RegexDfsMatcher.DfsMatchFrom(context, captures, bytecode, text, 0, raiseOnLimit:=True)
    Case "nfa"
        MatchOne = RegexNfaMatcher.NfaMatch(captures, bytecode, text)
    Case "dfa"
        MatchOne = -1
        If RegexDfaMatcher.DfaFindMatch(span, dfaState, bytecode, text, 1, False) Then
            captures.entireMatch = span
            MatchOne = span.start + span.Length - 1
        End If
    End Select
End Function

