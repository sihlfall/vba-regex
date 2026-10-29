Attribute VB_Name = "TestUtilities"
Option Explicit

Public Type BytecodeParameters
    maxProperCaptureSlot As Long
    nIdentifiers As Long
    properties As Long
    startActual As Long
    startDfaBwRunOffset As Long
End Type

Public Function MakeBytecodeParameters( _
    Optional ByVal maxProperCaptureSlot As Long = 1, _
    Optional ByVal nIdentifiers As Long = 0, _
    Optional ByVal properties As Long = 0, _
    Optional ByVal startActual As Long = BYTECODE_FIXED_PROLOGUE_LAST_IDX + 6, _
    Optional ByVal startDfaBwRunOffset As Long = -1 _
) As BytecodeParameters
    With MakeBytecodeParameters
        .maxProperCaptureSlot = maxProperCaptureSlot
        .nIdentifiers = nIdentifiers
        .properties = properties
        .startActual = startActual
        .startDfaBwRunOffset = startDfaBwRunOffset
    End With
End Function

Public Sub MakeBytecode( _
    ByRef outAry() As Long, _
    ByRef bytecodeParams As BytecodeParameters, _
    ParamArray bytecode() As Variant _
)
    ReDim outAry(0 To BYTECODE_FIXED_PROLOGUE_LAST_IDX + UBound(bytecode) + 6) As Long
    Dim i As Long, j As Long
    
    outAry(BYTECODE_IDX_MAX_PROPER_CAPTURE_SLOT) = bytecodeParams.maxProperCaptureSlot
    outAry(BYTECODE_IDX_N_IDENTIFIERS) = bytecodeParams.nIdentifiers
    outAry(BYTECODE_IDX_PROPERTIES) = bytecodeParams.properties
    outAry(BYTECODE_IDX_START_WITH_PREFIX) = BYTECODE_FIXED_PROLOGUE_LAST_IDX + 1
    outAry(BYTECODE_IDX_START_ACTUAL) = bytecodeParams.startActual
    If bytecodeParams.startDfaBwRunOffset > 0 Then
        outAry(BYTECODE_IDX_START_DFA_BW_RUN) = bytecodeParams.startActual + _
            bytecodeParams.startDfaBwRunOffset
    Else
        outAry(BYTECODE_IDX_START_DFA_BW_RUN) = -1
    End If
    outAry(BYTECODE_IDX_REOP_PERFORM_ASSERT) = REOP_PERFORM_ASSERT
    outAry(BYTECODE_IDX_REOP_MATCH_PREV) = REOP_MATCH_PREV
    outAry(BYTECODE_IDX_REOP_MATCH_OLD) = REOP_MATCH_OLD
    
    j = BYTECODE_FIXED_PROLOGUE_LAST_IDX + 1
    outAry(j) = REOP_SPLIT2: j = j + 1
    outAry(j) = 3: j = j + 1
    outAry(j) = REOP_DOT Or MODIFIER_S_MASK: j = j + 1
    outAry(j) = REOP_JUMP: j = j + 1
    outAry(j) = -5: j = j + 1
    
    For i = 0 To UBound(bytecode)
        outAry(j) = bytecode(i): j = j + 1
    Next
End Sub



