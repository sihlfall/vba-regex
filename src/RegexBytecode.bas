Attribute VB_Name = "RegexBytecode"
Option Explicit

Public Enum BytecodeDescriptionConstant
    BYTECODE_IDX_MAX_PROPER_CAPTURE_SLOT = 0
    BYTECODE_IDX_N_IDENTIFIERS = 1
    BYTECODE_IDX_CASE_INSENSITIVE_INDICATOR = 2
    BYTECODE_IDENTIFIER_MAP_BEGIN = 3
    BYTECODE_IDENTIFIER_MAP_ENTRY_SIZE = 3
    BYTECODE_IDENTIFIER_MAP_ENTRY_START_IN_PATTERN = 0
    BYTECODE_IDENTIFIER_MAP_ENTRY_LENGTH_IN_PATTERN = 1
    BYTECODE_IDENTIFIER_MAP_ENTRY_ID = 2
    
    ' Todo: Introduce special value or restrict max. explicit quantifier value to RegexNumericConstants.LONG_MAX - 1
    RE_QUANTIFIER_INFINITE = &H7FFFFFFF
End Enum

' regexp opcodes
Public Enum ReOpType
    REOP_MATCH = 1
    REOP_CHAR = 2
    REOP_PERIOD = 3
    REOP_RANGES = 4 ' nranges [must be >= 1], chfrom, chto, chfrom, chto, ...
    REOP_INVRANGES = 5
    REOP_JUMP = 6
    REOP_SPLIT1 = 7 ' prefer direct
    REOP_SPLIT2 = 8 ' prefer jump
    REOP_SAVE = 11
    REOP_SET_NAMED = 12 ' id, capture num
    REOP_LOOKPOS = 13
    REOP_LOOKNEG = 14
    REOP_BACKREFERENCE = 15
    REOP_ASSERT_START = 16
    REOP_ASSERT_END = 17
    REOP_ASSERT_WORD_BOUNDARY = 18
    REOP_ASSERT_NOT_WORD_BOUNDARY = 19
    REOP_REPEAT_EXACTLY_INIT = 20 ' <none>
    REOP_REPEAT_EXACTLY_START = 21 ' quantity [must be >= 1], offset
    REOP_REPEAT_EXACTLY_END = 22 ' quantity [must be >= 1], offset
    REOP_REPEAT_MAX_HUMBLE_INIT = 23 ' <none>
    REOP_REPEAT_MAX_HUMBLE_START = 24 ' quantity, offset
    REOP_REPEAT_MAX_HUMBLE_END = 25 ' quantitiy, offset
    REOP_REPEAT_GREEDY_MAX_INIT = 26 ' <none>
    REOP_REPEAT_GREEDY_MAX_START = 27 ' quantity, offset
    REOP_REPEAT_GREEDY_MAX_END = 28 ' quantitiy, offset
    REOP_CHECK_LOOKAHEAD = 29 ' <none>
    REOP_CHECK_LOOKBEHIND = 30 ' <none>
    REOP_END_LOOKPOS = 31 ' <none>
    REOP_END_LOOKNEG = 32 ' <none>
    REOP_FAIL = 33
End Enum

Public Function isCaseInsensitive(ByRef bytecode() As Long) As Boolean
    isCaseInsensitive = bytecode(BYTECODE_IDX_CASE_INSENSITIVE_INDICATOR) <> 0
End Function

Public Function GetIdentifierId( _
    ByRef bytecode() As Long, _
    ByRef lake As String, _
    ByRef identifier As String _
) As Long
    Dim aa As Long, bb As Long, mm As Long, compare As Long, identifierLength As Long
    
    identifierLength = Len(identifier)
    
    aa = RegexBytecode.BYTECODE_IDENTIFIER_MAP_BEGIN
    bb = RegexBytecode.BYTECODE_IDENTIFIER_MAP_BEGIN + RegexBytecode.BYTECODE_IDENTIFIER_MAP_ENTRY_SIZE * bytecode(RegexBytecode.BYTECODE_IDX_N_IDENTIFIERS)
    
    ' Find numeric id for identifier name.
    ' We are doing a binary search here.
    ' Invariant: Value we are looking for, if it exists, is always contained in the interval [aa;bb).
    Do
        If aa >= bb Then GetIdentifierId = -1: Exit Function ' identifier not found
        
        mm = aa + 3 * ((bb - aa) \ 6)
        If identifierLength < bytecode(mm + 1) Then
            bb = mm
        ElseIf identifierLength > bytecode(mm + 1) Then
            aa = mm + 3
        Else
            compare = StrComp( _
                identifier, _
                Mid$(lake, bytecode(mm), bytecode(mm + 1)), _
                vbBinaryCompare _
            )
            If compare < 0 Then
                bb = mm
            ElseIf compare > 0 Then
                aa = mm + 3
            Else
                ' found
                GetIdentifierId = bytecode(mm + 2)
                Exit Function
            End If
        End If
    Loop

End Function
