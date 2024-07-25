Attribute VB_Name = "RegexRangeConstants"
Option Explicit

Public RangeTablesInitialized As Boolean ' auto-initialized to False
Public RangeTableDigit(0 To 1) As Long
Public RangeTableWhite(0 To 21) As Long
Public RangeTableWordchar(0 To 7) As Long
Public RangeTableNotDigit(0 To 3) As Long
Public RangeTableNotWhite(0 To 23) As Long
Public RangeTableNotWordChar(0 To 9) As Long

Private Const MIN_LONG = &H80000000
Private Const MAX_LONG = &H7FFFFFFF

Public Sub InitializeRangeTables()
    RangeTableDigit(0) = &H30&
    RangeTableDigit(1) = &H39&
    
    '---------------------------------
    
    RangeTableWhite(0) = &H9&
    RangeTableWhite(1) = &HD&
    RangeTableWhite(2) = &H20&
    RangeTableWhite(3) = &H20&
    RangeTableWhite(4) = &HA0&
    RangeTableWhite(5) = &HA0&
    RangeTableWhite(6) = &H1680&
    RangeTableWhite(7) = &H1680&
    RangeTableWhite(8) = &H180E&
    RangeTableWhite(9) = &H180E&
    RangeTableWhite(10) = &H2000&
    RangeTableWhite(11) = &H200A
    RangeTableWhite(12) = &H2028&
    RangeTableWhite(13) = &H2029&
    RangeTableWhite(14) = &H202F
    RangeTableWhite(15) = &H202F
    RangeTableWhite(16) = &H205F&
    RangeTableWhite(17) = &H205F&
    RangeTableWhite(18) = &H3000&
    RangeTableWhite(19) = &H3000&
    RangeTableWhite(20) = &HFEFF&
    RangeTableWhite(21) = &HFEFF&
    
    '---------------------------------
    
    RangeTableWordchar(0) = &H30&
    RangeTableWordchar(1) = &H39&
    RangeTableWordchar(2) = &H41&
    RangeTableWordchar(3) = &H5A&
    RangeTableWordchar(4) = &H5F&
    RangeTableWordchar(5) = &H5F&
    RangeTableWordchar(6) = &H61&
    RangeTableWordchar(7) = &H7A&
    
    '---------------------------------
    
    RangeTableNotDigit(0) = MIN_LONG
    RangeTableNotDigit(1) = &H2F&
    RangeTableNotDigit(2) = &H3A&
    RangeTableNotDigit(3) = MAX_LONG
    
    '---------------------------------
    
    RangeTableNotWhite(0) = MIN_LONG
    RangeTableNotWhite(1) = &H8&
    RangeTableNotWhite(2) = &HE&
    RangeTableNotWhite(3) = &H1F&
    RangeTableNotWhite(4) = &H21&
    RangeTableNotWhite(5) = &H9F&
    RangeTableNotWhite(6) = &HA1&
    RangeTableNotWhite(7) = &H167F&
    RangeTableNotWhite(8) = &H1681&
    RangeTableNotWhite(9) = &H180D&
    RangeTableNotWhite(10) = &H180F&
    RangeTableNotWhite(11) = &H1FFF&
    RangeTableNotWhite(12) = &H200B&
    RangeTableNotWhite(13) = &H2027&
    RangeTableNotWhite(14) = &H202A&
    RangeTableNotWhite(15) = &H202E&
    RangeTableNotWhite(16) = &H2030&
    RangeTableNotWhite(17) = &H205E&
    RangeTableNotWhite(18) = &H2060&
    RangeTableNotWhite(19) = &H2FFF&
    RangeTableNotWhite(20) = &H3001&
    RangeTableNotWhite(21) = &HFEFE&
    RangeTableNotWhite(22) = &HFF00&
    RangeTableNotWhite(23) = MAX_LONG
    
    '---------------------------------
    
    RangeTableNotWordChar(0) = MIN_LONG
    RangeTableNotWordChar(1) = &H2F&
    RangeTableNotWordChar(2) = &H3A&
    RangeTableNotWordChar(3) = &H40&
    RangeTableNotWordChar(4) = &H5B&
    RangeTableNotWordChar(5) = &H5E&
    RangeTableNotWordChar(6) = &H60&
    RangeTableNotWordChar(7) = &H60&
    RangeTableNotWordChar(8) = &H7B&
    RangeTableNotWordChar(9) = MAX_LONG
    
    '---------------------------------
    
    RangeTablesInitialized = True
End Sub

