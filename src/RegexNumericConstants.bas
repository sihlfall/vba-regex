Attribute VB_Name = "RegexNumericConstants"
Option Explicit

Public Enum NumericConstantLong
    LONG_FIRST_BIT = &H80000000
    LONG_ALL_BUT_FIRST_BIT = Not LONG_FIRST_BIT
    LONG_MIN = &H80000000
    LONG_MAX = &H7FFFFFFF
    LONG_MAX_DIV_10 = LONG_MAX \ 10
End Enum




