Attribute VB_Name = "ExamplesForReadme"
Option Explicit

' Examples for the README file
' Strange formatting intentional to facilitate copy/paste

Public Sub RunAllExamples()
    Dim regex As StaticRegex.RegexTy
    Dim exampleString As String
    
    Debug.Print "-------"

    ExamplesInitialize regex, exampleString
    Debug.Print
    Example01 regex, exampleString
    Debug.Print
    Example02 regex, exampleString
    Debug.Print
    Example03 regex, exampleString
    Debug.Print
    Example04 regex, exampleString
    Debug.Print
    Example05 regex, exampleString
    Debug.Print
    Example06 regex, exampleString
End Sub


' ***************************************************************************************************************

Private Sub ExamplesInitialize(ByRef pRegex As StaticRegex.RegexTy, ByRef pExampleString As String)
Debug.Print "Examples Initialize"
' ---
Dim regex As StaticRegex.RegexTy

StaticRegex.InitializeRegex regex, _
   "(?<month>\w{3})-(?<day>\d{1,2})-(?<year>\d{4})"
   
Dim exampleString As String
exampleString = "On Jul-4-1776, independence was declared. " & _
   "On Apr-30-1789, George Washington became the first president."
' ---
pRegex = regex
pExampleString = exampleString
End Sub

' ***************************************************************************************************************

Private Sub Example01(ByRef regex As StaticRegex.RegexTy, ByRef exampleString As String)
Debug.Print "Example 1"
' ---
Dim wasFound As Boolean

wasFound = StaticRegex.Test(regex, exampleString)

Debug.Print wasFound   ' prints: True
' ---
End Sub

' ***************************************************************************************************************

Private Sub Example02(ByRef regex As StaticRegex.RegexTy, ByRef exampleString As String)
Debug.Print "Example 2"
' ---
Dim wasFound As Boolean, matcherState As StaticRegex.MatcherStateTy

wasFound = StaticRegex.Match(matcherState, regex, exampleString)

Debug.Print wasFound ' prints: True
Debug.Print StaticRegex.GetCapture(matcherState, exampleString)
   ' prints: 'Jul-4-1776' (entire match)
Debug.Print StaticRegex.GetCapture(matcherState, exampleString, 2)
   ' prints: '4' (second parenthesis)
Debug.Print StaticRegex.GetCaptureByName(matcherState, regex, exampleString, "month")
   ' prints: 'Jul' (capture named "month")
' ---
End Sub

' ***************************************************************************************************************

Private Sub Example03(ByRef regex As StaticRegex.RegexTy, ByRef exampleString As String)
Debug.Print "Example 3"
' ---
Dim matcherState As MatcherStateTy

Do While StaticRegex.MatchNext(matcherState, regex, exampleString)
   Debug.Print StaticRegex.GetCapture(matcherState, exampleString)
   Debug.Print StaticRegex.GetCaptureByName(matcherState, regex, exampleString, "year")
Loop
' ---
End Sub

' ***************************************************************************************************************

Private Sub Example04(ByRef regex As StaticRegex.RegexTy, ByRef exampleString As String)
Debug.Print "Example 4"
' ---
Debug.Print StaticRegex.MatchThenJoin(regex, exampleString, delimiter:=", ")
' ---
End Sub

' ***************************************************************************************************************

Private Sub Example05(ByRef regex As StaticRegex.RegexTy, ByRef exampleString As String)
Debug.Print "Example 5"
' ---
Debug.Print StaticRegex.MatchThenJoin( _
   regex, exampleString, delimiter:=", ", format:="$<day> $<month> $<year>" _
)
' ---
End Sub

' ***************************************************************************************************************

' Source: http://dailydoseofexcel.com/archives/2015/01/28/joining-two-dimensional-arrays/
Private Function Join2D(ByVal vArray As Variant, Optional ByVal sWordDelim As String = " ", Optional ByVal sLineDelim As String = vbNewLine) As String
    
    Dim i As Long, j As Long
    Dim aReturn() As String
    Dim aLine() As String
    
    ReDim aReturn(LBound(vArray, 1) To UBound(vArray, 1))
    ReDim aLine(LBound(vArray, 2) To UBound(vArray, 2))
    
    For i = LBound(vArray, 1) To UBound(vArray, 1)
        For j = LBound(vArray, 2) To UBound(vArray, 2)
            'Put the current line into a 1d array
            aLine(j) = vArray(i, j)
        Next j
        'Join the current line into a 1d array
        aReturn(i) = Join(aLine, sWordDelim)
    Next i
    
    Join2D = Join(aReturn, sLineDelim)
    
End Function

Private Function MakeStringArray(ParamArray strings() As Variant) As String()
   Dim ary() As String, i As Long
   ReDim ary(0 To UBound(strings) - LBound(strings) + 1) As String
   For i = LBound(strings) To UBound(strings)
      ary(i - LBound(strings)) = strings(i)
   Next
   MakeStringArray = ary
End Function

Private Sub Example06(ByRef regex As StaticRegex.RegexTy, ByRef exampleString As String)
Debug.Print "Example 5"
' ---
Dim results() As String

StaticRegex.MatchThenList results, _
   regex, exampleString, _
   MakeStringArray("$&", "$<day>", "$<month>", "$<year>")
' ---
Debug.Print Join2D(results, ",")
End Sub

