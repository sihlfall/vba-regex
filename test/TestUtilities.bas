Attribute VB_Name = "TestUtilities"
Option Explicit

Public Type BytecodeParameters
    maxProperCaptureSlot As Long
    nIdentifiers As Long
    properties As Long
    startDfaBwRunOffset As Long
End Type

Public Function MakeBytecodeParameters( _
    Optional ByVal maxProperCaptureSlot As Long = 1, _
    Optional ByVal nIdentifiers As Long = 0, _
    Optional ByVal properties As Long = 0, _
    Optional ByVal startDfaBwRunOffset As Long = -1 _
) As BytecodeParameters
    With MakeBytecodeParameters
        .maxProperCaptureSlot = maxProperCaptureSlot
        .nIdentifiers = nIdentifiers
        .properties = properties
        .startDfaBwRunOffset = startDfaBwRunOffset
    End With
End Function

' Supply identifier-map triples (if any), then the instruction words. The fixed
' prologue and implicit search prefix always use the current compiler layout.
' The reverse-program offset is relative to the first instruction after that prefix.
Public Sub MakeBytecode( _
    ByRef outAry() As Long, _
    ByRef bytecodeParams As BytecodeParameters, _
    ParamArray bytecode() As Variant _
)
    Dim i As Long, j As Long, mapLength As Long, startWithPrefix As Long, startActual As Long
    Dim startAfterCaptures As Long
    mapLength = BYTECODE_IDENTIFIER_MAP_ENTRY_SIZE * bytecodeParams.nIdentifiers
    startWithPrefix = BYTECODE_IDENTIFIER_MAP_BEGIN + mapLength
    startActual = startWithPrefix + 5
    ReDim outAry(0 To startActual + UBound(bytecode) - mapLength) As Long
    
    outAry(BYTECODE_IDX_MAX_PROPER_CAPTURE_SLOT) = bytecodeParams.maxProperCaptureSlot
    outAry(BYTECODE_IDX_N_IDENTIFIERS) = bytecodeParams.nIdentifiers
    outAry(BYTECODE_IDX_PROPERTIES) = bytecodeParams.properties
    outAry(BYTECODE_IDX_START_WITH_PREFIX) = startWithPrefix
    outAry(BYTECODE_IDX_START_ACTUAL) = startActual
    If bytecodeParams.startDfaBwRunOffset > 0 Then
        outAry(BYTECODE_IDX_START_DFA_BW_RUN) = startActual + _
            bytecodeParams.startDfaBwRunOffset
    Else
        outAry(BYTECODE_IDX_START_DFA_BW_RUN) = -1
    End If
    outAry(BYTECODE_IDX_REOP_PERFORM_ASSERT) = REOP_PERFORM_ASSERT
    outAry(BYTECODE_IDX_REOP_MATCH_PREV) = REOP_MATCH_PREV
    outAry(BYTECODE_IDX_REOP_MATCH_OLD) = REOP_MATCH_OLD
    
    For i = 0 To mapLength - 1
        outAry(BYTECODE_IDENTIFIER_MAP_BEGIN + i) = bytecode(i)
    Next

    j = startWithPrefix
    outAry(j) = REOP_SPLIT2: j = j + 1
    outAry(j) = 3: j = j + 1
    outAry(j) = REOP_DOT Or MODIFIER_S_MASK: j = j + 1
    outAry(j) = REOP_JUMP: j = j + 1
    outAry(j) = -5: j = j + 1
    
    startAfterCaptures = startActual
    For i = mapLength To UBound(bytecode)
        If j = startAfterCaptures Then
            Select Case CLng(bytecode(i)) And REOP_OPCODE_MASK
            Case REOP_SAVE
                startAfterCaptures = j + 2
            Case REOP_SET_NAMED
                startAfterCaptures = j + 3
            End Select
        End If
        outAry(j) = bytecode(i): j = j + 1
    Next
    outAry(BYTECODE_IDX_START_AFTER_CAPTURES) = startAfterCaptures
End Sub



