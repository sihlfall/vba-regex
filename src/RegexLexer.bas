Attribute VB_Name = "RegexLexer"
Option Explicit

' TODO: Make sure unexpected end of input is treated as an error everywhere.

' inputStr: the pattern being compiled.
' iEnd points to the last character, i.e. iEnd = Len(inputStr) - 1.
' iCurrent points to the current character.
' currentCharacter contains the value of the current character (16 bit between 0 and 2^16 - 1 = &HFFFF&, or LEXER_ENDOFINPUT).
' After the end of the input has been reached, currentCharacter = LEXER_ENDOFINPUT and iCurrent = iEnd.
' All fields are intended to be private to this module.
Public Type Ty
    iCurrent As Long
    iEnd As Long
    inputStr As String
    currentCharacter As Long
    identifierTree As RegexIdentifierSupport.IdentifierTreeTy
End Type

Public Type ReToken
    t As Long ' token type
    quantifierType As Long
    num As Long ' numeric value (character, count, id for named capture group, -1 for non-named capture group, modifiers for non-capture group)
    qmin As Long
    qmax As Long
End Type

Public Enum TokenQuantifierType
    QUANTIFIER_GREEDY = 1
    QUANTIFIER_HUMBLE = 2
    QUANTIFIER_POSSESSIVE = 3
End Enum

Public Enum TokenTypeIdType
    RETOK_EOF = 0
    RETOK_DISJUNCTION = 1
    RETOK_QUANTIFIER = 2
    RETOK_ASSERT_START = 3
    RETOK_ASSERT_END = 4
    RETOK_ASSERT_WORD_BOUNDARY = 5
    RETOK_ASSERT_NOT_WORD_BOUNDARY = 6
    RETOK_ASSERT_START_POS_LOOKAHEAD = 7
    RETOK_ASSERT_START_NEG_LOOKAHEAD = 8
    RETOK_ATOM_PERIOD = 9
    RETOK_ATOM_CHAR = 10
    RETOK_ATOM_DIGIT = 11                   ' assumptions in regexp compiler
    RETOK_ATOM_NOT_DIGIT = 12               ' -""-
    RETOK_ATOM_WHITE = 13                   ' -""-
    RETOK_ATOM_NOT_WHITE = 14               ' -""-
    RETOK_ATOM_WORD_CHAR = 15               ' -""-
    RETOK_ATOM_NOT_WORD_CHAR = 16           ' -""-
    RETOK_ATOM_BACKREFERENCE = 17
    RETOK_ATOM_START_CAPTURE_GROUP = 18
    RETOK_ATOM_START_NONCAPTURE_GROUP = 19
    RETOK_ATOM_START_CHARCLASS = 20
    RETOK_ATOM_START_CHARCLASS_INVERTED = 21
    RETOK_ASSERT_START_POS_LOOKBEHIND = 22
    RETOK_ASSERT_START_NEG_LOOKBEHIND = 23
    RETOK_ATOM_START_ATOMIC_GROUP = 24
    RETOK_ATOM_END = 25 ' closing parenthesis
        ' (ends (POS|NEG)_LOOK(AHEAD|BEHIND), CAPTURE_GROUP, NONCAPTURE_GROUP, ATOMIC_GROUP)
    RETOK_UNBOUNDED_MODIFIER = 26
End Enum

' Returned by input reading function after end of input has been reached
' Since our characters are 16 bit, converted to a positive Long value, and Long is 32 bit, -1 is free to use for us.
Private Const LEXER_ENDOFINPUT As Long = -1

Private Enum LexerUnicodeCodepointConstant
    UNICODE_EXCLAMATION = 33  ' !
    UNICODE_DOLLAR = 36  ' $
    UNICODE_LPAREN = 40  ' (
    UNICODE_RPAREN = 41  ' )
    UNICODE_STAR = 42  ' *
    UNICODE_PLUS = 43  ' +
    UNICODE_COMMA = 44  ' ,
    UNICODE_MINUS = 45  ' -
    UNICODE_PERIOD = 46  ' .
    UNICODE_0 = 48  ' 0
    UNICODE_1 = 49  ' 1
    UNICODE_7 = 55  ' 7
    UNICODE_9 = 57  ' 9
    UNICODE_COLON = 58  ' :
    UNICODE_LT = 60  ' <
    UNICODE_EQUALS = 61  ' =
    UNICODE_GT = 62  ' >
    UNICODE_QUESTION = 63  ' ?
    UNICODE_UC_A = 65  ' A
    UNICODE_UC_B = 66  ' B
    UNICODE_UC_D = 68  ' D
    UNICODE_UC_F = 70  ' F
    UNICODE_UC_S = 83  ' S
    UNICODE_UC_W = 87  ' W
    UNICODE_UC_Z = 90  ' Z
    UNICODE_LBRACKET = 91  ' [
    UNICODE_BACKSLASH = 92  ' \
    UNICODE_RBRACKET = 93  ' ]
    UNICODE_CARET = 94  ' ^
    UNICODE_LC_A = 97  ' a
    UNICODE_LC_B = 98  ' b
    UNICODE_LC_C = 99  ' c
    UNICODE_LC_D = 100  ' d
    UNICODE_LC_F = 102  ' f
    UNICODE_LC_I = 105  ' i
    UNICODE_LC_M = 109  ' m
    UNICODE_LC_N = 110  ' n
    UNICODE_LC_R = 114  ' r
    UNICODE_LC_S = 115  ' s
    UNICODE_LC_T = 116  ' t
    UNICODE_LC_U = 117  ' u
    UNICODE_LC_V = 118  ' v
    UNICODE_LC_W = 119  ' w
    UNICODE_LC_X = 120  ' x
    UNICODE_LC_Z = 122  ' z
    UNICODE_LCURLY = 123  ' {
    UNICODE_PIPE = 124  ' |
    UNICODE_RCURLY = 125  ' }
    UNICODE_CP_ZWNJ = &H200C& ' zero-width non-joiner
    UNICODE_CP_ZWJ = &H200D&  ' zero-width joiner
End Enum

Public Sub Initialize(ByRef lexCtx As Ty, ByRef inputStr As String)
    With lexCtx
        .inputStr = inputStr
        .iEnd = Len(.inputStr)
        .iCurrent = 0
        .currentCharacter = Not LEXER_ENDOFINPUT ' value does not matter, as long as it does not equal LEXER_ENDOFINPUT
        
        .identifierTree.nEntries = 0
        .identifierTree.root = -1
    End With
    Advance lexCtx
End Sub

'
'  Parse a RegExp token.  The grammar is described in E5 Section 15.10.
'  Terminal constructions (such as quantifiers) are parsed directly here.
'
'  0xffffffffU is used as a marker for "infinity" in quantifiers.  Further,
'  _MAX_RE_QUANT_DIGITS limits the maximum number of digits that
'  will be accepted for a quantifier.
'
Public Sub ParseReToken(ByRef lexCtx As Ty, ByRef outToken As ReToken)
    Dim x As Long
    
    ' used only locally
    Dim i As Long, val1 As Long, val2 As Long, digits As Long, tmp As Long
    
    ' Todo: remove
    Dim slp As RegexIdentifierSupport.StartLengthPair

    Dim emptyReToken As ReToken ' effectively a constant -- zeroed out by default
    outToken = emptyReToken

    x = Advance(lexCtx)
    
    If x <= 32 Then GoTo L_any_or_EOF
    If x >= 126 Then GoTo L_any

    On x - 32 GoTo _
        L_any, L_any, L_any, L_036, L_any, L_any, L_any, _
        L_040, L_041, L_042, L_043, L_any, L_any, L_046, L_any, _
        L_any, L_any, L_any, L_any, L_any, L_any, L_any, L_any, _
        L_any, L_any, L_any, L_any, L_any, L_any, L_any, L_063, _
        L_any, L_any, L_any, L_any, L_any, L_any, L_any, L_any, _
        L_any, L_any, L_any, L_any, L_any, L_any, L_any, L_any, _
        L_any, L_any, L_any, L_any, L_any, L_any, L_any, L_any, _
        L_any, L_any, L_any, L_091, L_092, L_093, L_094, L_any, _
        L_any, L_any, L_any, L_any, L_any, L_any, L_any, L_any, _
        L_any, L_any, L_any, L_any, L_any, L_any, L_any, L_any, _
        L_any, L_any, L_any, L_any, L_any, L_any, L_any, L_any, _
        L_any, L_any, L_any, L_123, L_124, L_125
    
    ' unreachable
    
L_any_or_EOF:
    If x = LEXER_ENDOFINPUT Then
        ' EOF
        outToken.t = RETOK_EOF
        Exit Sub
    End If
    ' Fallthrough for other characters

L_any:
    ' Pattern character
    outToken.t = RETOK_ATOM_CHAR
    outToken.num = x
    Exit Sub
    
L_036: ' $
    outToken.t = RETOK_ASSERT_END
    Exit Sub
        
L_040: ' (
    If lexCtx.currentCharacter <> UNICODE_QUESTION Then ' (
        outToken.t = RETOK_ATOM_START_CAPTURE_GROUP
        outToken.num = -1
        Exit Sub
    End If
    
    Advance lexCtx
    x = Advance(lexCtx)
    
    If x <= 32 Then GoTo G_any
    If x >= 63 Then GoTo G_any
    
    On x - 32 GoTo _
        G_033, G_any, G_any, G_any, G_any, G_any, G_any, _
        G_any, G_any, G_any, G_any, G_any, G_any, G_any, G_any, _
        G_any, G_any, G_any, G_any, G_any, G_any, G_any, G_any, _
        G_any, G_any, G_058, G_any, G_060, G_061, G_062
    
    ' unreachable
    Err.Raise RegexErrors.REGEX_ERR_INTERNAL_LOGIC_ERR
    
G_033: ' (?!
    outToken.t = RETOK_ASSERT_START_NEG_LOOKAHEAD
    Exit Sub
    
G_058: ' (?:
    outToken.t = RETOK_ATOM_START_NONCAPTURE_GROUP
    outToken.num = 0 ' no modifiers
    Exit Sub
    
G_060: ' (?<
    x = Advance(lexCtx)
    If x = UNICODE_EQUALS Then
        outToken.t = RETOK_ASSERT_START_POS_LOOKBEHIND
    ElseIf x = UNICODE_EXCLAMATION Then
        outToken.t = RETOK_ASSERT_START_NEG_LOOKBEHIND
    ElseIf IsIdentifierChar(x) Then
        With lexCtx
            val1 = .identifierTree.nEntries
            val2 = .iCurrent - 1
            Do
                x = Advance(lexCtx)
                If x = UNICODE_GT Then Exit Do
                ' Todo: Allow unicode escape sequences
                If Not IsIdentifierChar(x) Then Err.Raise REGEX_ERR_INVALID_IDENTIFIER
            Loop
            outToken.t = RETOK_ATOM_START_CAPTURE_GROUP
            With slp
                .start = val2: .Length = lexCtx.iCurrent - 1 - val2
            End With
            outToken.num = RegexIdentifierSupport.RedBlackFindOrInsert( _
                lexCtx.inputStr, _
                lexCtx.identifierTree, _
                slp _
            )
        End With
    Else
        Err.Raise REGEX_ERR_INVALID_REGEXP_GROUP
    End If
    Exit Sub
    
G_061: ' (?=
    outToken.t = RETOK_ASSERT_START_POS_LOOKAHEAD
    Exit Sub
    
G_062: ' (?>
    outToken.t = RETOK_ATOM_START_ATOMIC_GROUP
    outToken.num = 0 ' no modifiers
    Exit Sub
    
G_any: ' (?modifier
    tmp = 0
    Do
        If x = UNICODE_LC_I Then
            If tmp And RegexBytecode.MODIFIER_I_WRITE Then Err.Raise RegexErrors.REGEX_ERR_INVALID_MODIFIER
            tmp = tmp Or (RegexBytecode.MODIFIER_I_WRITE Or RegexBytecode.MODIFIER_I_ACTIVE)
        ElseIf x = UNICODE_LC_M Then
            If tmp And RegexBytecode.MODIFIER_M_WRITE Then Err.Raise RegexErrors.REGEX_ERR_INVALID_MODIFIER
            tmp = tmp Or (RegexBytecode.MODIFIER_M_WRITE Or RegexBytecode.MODIFIER_M_ACTIVE)
        ElseIf x = UNICODE_LC_S Then
            If tmp And RegexBytecode.MODIFIER_S_WRITE Then Err.Raise RegexErrors.REGEX_ERR_INVALID_MODIFIER
            tmp = tmp Or (RegexBytecode.MODIFIER_S_WRITE Or RegexBytecode.MODIFIER_S_ACTIVE)
        ElseIf x = UNICODE_MINUS Then
            Exit Do
        ElseIf x = UNICODE_COLON Then
            GoTo EndOfBoundedModifier
        ElseIf x = UNICODE_RPAREN Then
            GoTo EndOfUnboundedModifier
        Else
            Err.Raise RegexErrors.REGEX_ERR_INVALID_MODIFIER
        End If
        
        x = Advance(lexCtx)
    Loop
    
    x = Advance(lexCtx) ' skip dash
    
    Do
        If x = UNICODE_LC_I Then
            If tmp And RegexBytecode.MODIFIER_I_WRITE Then Err.Raise RegexErrors.REGEX_ERR_INVALID_MODIFIER
            tmp = tmp Or RegexBytecode.MODIFIER_I_WRITE
        ElseIf x = UNICODE_LC_M Then
            If tmp And RegexBytecode.MODIFIER_M_WRITE Then Err.Raise RegexErrors.REGEX_ERR_INVALID_MODIFIER
            tmp = tmp Or RegexBytecode.MODIFIER_M_WRITE
        ElseIf x = UNICODE_LC_S Then
            If tmp And RegexBytecode.MODIFIER_S_WRITE Then Err.Raise RegexErrors.REGEX_ERR_INVALID_MODIFIER
            tmp = tmp Or RegexBytecode.MODIFIER_S_WRITE
        ElseIf x = UNICODE_COLON Then
            GoTo EndOfBoundedModifier
        ElseIf x = UNICODE_RPAREN Then
            GoTo EndOfUnboundedModifier
        Else
            Err.Raise RegexErrors.REGEX_ERR_INVALID_MODIFIER
        End If
        
        x = Advance(lexCtx)
    Loop
    
    ' unreachable
    Err.Raise RegexErrors.REGEX_ERR_INTERNAL_LOGIC_ERR
    
EndOfBoundedModifier:
    outToken.t = RETOK_ATOM_START_NONCAPTURE_GROUP
    outToken.num = tmp ' no modifiers
    Exit Sub

EndOfUnboundedModifier:
    outToken.t = RETOK_UNBOUNDED_MODIFIER
    outToken.num = tmp
    Exit Sub

L_041: ' )
    outToken.t = RETOK_ATOM_END
    Exit Sub

L_042: ' *
    With outToken
        .qmin = 0
        .qmax = RE_QUANTIFIER_INFINITE
        If lexCtx.currentCharacter = UNICODE_QUESTION Then
            Advance lexCtx
            .t = RETOK_QUANTIFIER
            .quantifierType = QUANTIFIER_HUMBLE
        ElseIf lexCtx.currentCharacter = UNICODE_PLUS Then
            Advance lexCtx
            .t = RETOK_QUANTIFIER
            .quantifierType = QUANTIFIER_POSSESSIVE
        Else
            .t = RETOK_QUANTIFIER
            .quantifierType = QUANTIFIER_GREEDY
        End If
    End With
    Exit Sub

L_043: ' +
    With outToken
        .qmin = 1
        .qmax = RE_QUANTIFIER_INFINITE
        If lexCtx.currentCharacter = UNICODE_QUESTION Then
            Advance lexCtx
            .t = RETOK_QUANTIFIER
            .quantifierType = QUANTIFIER_HUMBLE
        ElseIf lexCtx.currentCharacter = UNICODE_PLUS Then
            Advance lexCtx
            .t = RETOK_QUANTIFIER
            .quantifierType = QUANTIFIER_POSSESSIVE
        Else
            .t = RETOK_QUANTIFIER
            .quantifierType = QUANTIFIER_GREEDY
        End If
    End With
    Exit Sub

L_046: ' .
    outToken.t = RETOK_ATOM_PERIOD
    Exit Sub
    
L_063: ' ?
    With outToken
        .qmin = 0
        .qmax = 1
        If lexCtx.currentCharacter = UNICODE_QUESTION Then
            Advance lexCtx
            .t = RETOK_QUANTIFIER
            .quantifierType = QUANTIFIER_HUMBLE
        ElseIf lexCtx.currentCharacter = UNICODE_PLUS Then
            Advance lexCtx
            .t = RETOK_QUANTIFIER
            .quantifierType = QUANTIFIER_POSSESSIVE
        Else
            .t = RETOK_QUANTIFIER
            .quantifierType = QUANTIFIER_GREEDY
        End If
    End With
    Exit Sub
    
L_091: ' [
    ' To avoid creating a heavy intermediate value for the list of ranges,
    ' only the start token ('[' or '[^') is parsed here.  The regexp
    ' compiler parses the ranges itself.
    If lexCtx.currentCharacter = UNICODE_CARET Then
        Advance lexCtx
        outToken.t = RETOK_ATOM_START_CHARCLASS_INVERTED
    Else
        outToken.t = RETOK_ATOM_START_CHARCLASS
    End If
    Exit Sub
    
L_092: ' \
    ' The E5.1 specification does not seem to allow IdentifierPart characters
    ' to be used as identity escapes.  Unfortunately this includes '$', which
    ' cannot be escaped as '\$'; it needs to be escaped e.g. as '\u0024'.
    ' Many other implementations (including V8 and Rhino, for instance) do
    ' accept '\$' as a valid identity escape, which is quite pragmatic, and
    ' ES2015 Annex B relaxes the rules to allow these (and other) real world forms.
    '
    ' Here, we use a Perl-like convention:
    '   \<character> is an identity escape (= represents the character itself)
    '   unless the character is within [0-9A-Za-z];
    '   we hope that this will not require any breaking changes in the future.
    x = Advance(lexCtx)
    
    If x <= 47 Then GoTo IdentityEscape_or_EOF
    If x >= 123 Then GoTo IdentityEscape
    
    On x - 47 GoTo _
        E_Digit0, E_Digit, E_Digit, E_Digit, E_Digit, E_Digit, E_Digit, E_Digit, _
        E_Digit, E_Digit, IdentityEscape, IdentityEscape, IdentityEscape, IdentityEscape, IdentityEscape, IdentityEscape, _
        IdentityEscape, E_Invalid, E_066, E_Invalid, E_068, E_Invalid, E_Invalid, E_Invalid, _
        E_Invalid, E_Invalid, E_Invalid, E_Invalid, E_Invalid, E_Invalid, E_Invalid, E_Invalid, _
        E_Invalid, E_Invalid, E_Invalid, E_083, E_Invalid, E_Invalid, E_Invalid, E_087, _
        E_Invalid, E_Invalid, E_Invalid, IdentityEscape, IdentityEscape, IdentityEscape, IdentityEscape, IdentityEscape, _
        IdentityEscape, E_Invalid, E_098, E_099, E_100, E_Invalid, E_102, E_Invalid, _
        E_Invalid, E_Invalid, E_Invalid, E_Invalid, E_Invalid, E_Invalid, E_110, E_Invalid, _
        E_Invalid, E_Invalid, E_114, E_115, E_116, E_117, E_118, E_119, _
        E_120, E_Invalid, E_Invalid
        
    ' unreachable
    Err.Raise RegexErrors.REGEX_ERR_INTERNAL_LOGIC_ERR
            
IdentityEscape_or_EOF:
    If x < 0 Then Err.Raise RegexErrors.REGEX_ERR_INVALID_REGEXP_ESCAPE
    ' else fallthrough
    
IdentityEscape:
    outToken.num = x
    outToken.t = RETOK_ATOM_CHAR
    Exit Sub
            
E_Digit0: ' \0
    x = lexCtx.currentCharacter
    
    ' Reject a following digit without consuming the next character.
    If x >= UNICODE_0 And x <= UNICODE_9 Then Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
    outToken.num = 0
    outToken.t = RETOK_ATOM_CHAR
    Exit Sub
    
E_Digit: ' \1... to \9...
    val1 = 0
    i = 0
    Do
        ' We have to be careful here to make sure there will be no overflow.
        ' 2^31 - 1 backreferences is a bit ridiculous, though.
        If val1 > RegexNumericConstants.LONG_MAX_DIV_10 Then Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
        val1 = val1 * 10
        tmp = x - UNICODE_0
        If RegexNumericConstants.LONG_MAX - val1 < tmp Then Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
        val1 = val1 + tmp
        x = lexCtx.currentCharacter
        If x < UNICODE_0 Or x > UNICODE_9 Then Exit Do
        Advance lexCtx
        i = i + 1
    Loop
    outToken.t = RETOK_ATOM_BACKREFERENCE
    outToken.num = val1
    Exit Sub
    
E_066: ' \B
    outToken.t = RETOK_ASSERT_NOT_WORD_BOUNDARY
    Exit Sub
    
E_068: ' \D
    outToken.t = RETOK_ATOM_NOT_DIGIT
    Exit Sub
    
E_083: ' \S
    outToken.t = RETOK_ATOM_NOT_WHITE
    Exit Sub
    
E_087: ' \W
    outToken.t = RETOK_ATOM_NOT_WORD_CHAR
    Exit Sub

E_098: ' \b
    outToken.t = RETOK_ASSERT_WORD_BOUNDARY
    Exit Sub
    
E_099: ' \c
    x = Advance(lexCtx)
    If x < UNICODE_UC_A Then GoTo E_Invalid
    If x > UNICODE_LC_Z Then GoTo E_Invalid
    If x > UNICODE_UC_Z And x < UNICODE_LC_A Then GoTo E_Invalid
    ' subtract 64, for lowercase deliver corresponding uppercase value
    outToken.num = (x - 64) And &H1F&
    outToken.t = RETOK_ATOM_CHAR
    Exit Sub

E_100: ' \d
    outToken.t = RETOK_ATOM_DIGIT
    Exit Sub
    
E_102: ' \f
    outToken.num = &HC&
    outToken.t = RETOK_ATOM_CHAR
    Exit Sub
    
E_110: ' \n
    outToken.num = &HA&
    outToken.t = RETOK_ATOM_CHAR
    Exit Sub
    
E_114: ' \r
    outToken.num = &HD&
    outToken.t = RETOK_ATOM_CHAR
    Exit Sub
    
E_115: ' \s
    outToken.t = RETOK_ATOM_WHITE
    Exit Sub

E_116: ' \t
    outToken.num = &H9&
    outToken.t = RETOK_ATOM_CHAR
    Exit Sub
    
E_117: ' \u
    ' Todo: What does the following mean?
    ' The token value is the Unicode codepoint without
    ' it being decode into surrogate pair characters
    ' here.  The \u{H+} is only allowed in Unicode mode
    ' which we don't support yet.
    outToken.num = LexerParseEscapeU(lexCtx)
    outToken.t = RETOK_ATOM_CHAR
    Exit Sub
    
E_118: ' \v
    outToken.num = &HB&
    outToken.t = RETOK_ATOM_CHAR
    Exit Sub
    
E_119: ' \w
    outToken.t = RETOK_ATOM_WORD_CHAR
    Exit Sub

E_120: ' \x
    outToken.num = LexerParseEscapeX(lexCtx)
    outToken.t = RETOK_ATOM_CHAR
    Exit Sub
    
E_Invalid:
    Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
    
L_093: ' ]
L_125: ' }
    GoTo L_any

L_094: ' ^
    outToken.t = RETOK_ASSERT_START
    Exit Sub
    
L_123: ' {
    ' Production allows 'DecimalDigits', including leading zeroes
    val1 = 0
    val2 = RE_QUANTIFIER_INFINITE
    
    digits = 0

    Do
        x = Advance(lexCtx)
        If (x >= UNICODE_0) And (x <= UNICODE_9) Then
            digits = digits + 1
            ' Be careful to prevent overflow
            If val1 > RegexNumericConstants.LONG_MAX_DIV_10 Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER
            val1 = val1 * 10
            tmp = x - UNICODE_0
            If RegexNumericConstants.LONG_MAX - val1 < tmp Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER
            val1 = val1 + tmp
        ElseIf x = UNICODE_COMMA Then
            If val2 <> RE_QUANTIFIER_INFINITE Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER
            If lexCtx.currentCharacter = UNICODE_RCURLY Then
                ' form: { DecimalDigits , }, val1 = min count
                If digits = 0 Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER
                outToken.qmin = val1
                outToken.qmax = RE_QUANTIFIER_INFINITE
                Advance lexCtx
                Exit Do
            End If
            val2 = val1
            val1 = 0
            digits = 0 ' not strictly necessary because of lookahead '}' above
        ElseIf x = UNICODE_RCURLY Then
            If digits = 0 Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER
            If val2 <> RE_QUANTIFIER_INFINITE Then
                ' val2 = min count, val1 = max count
                outToken.qmin = val2
                outToken.qmax = val1
            Else
                ' val1 = count
                outToken.qmin = val1
                outToken.qmax = val1
            End If
            Exit Do
        Else
            Err.Raise REGEX_ERR_INVALID_QUANTIFIER
        End If
    Loop
    If lexCtx.currentCharacter = UNICODE_QUESTION Then
        outToken.quantifierType = QUANTIFIER_HUMBLE
        Advance lexCtx
    ElseIf lexCtx.currentCharacter = UNICODE_PLUS Then
        outToken.quantifierType = QUANTIFIER_POSSESSIVE
        Advance lexCtx
    Else
        outToken.quantifierType = QUANTIFIER_GREEDY
    End If
    outToken.t = RETOK_QUANTIFIER
    Exit Sub

L_124: ' |
    outToken.t = RETOK_DISJUNCTION
    Exit Sub
End Sub

Public Sub ParseReRanges(lexCtx As Ty, ByRef outBuffer As ArrayBuffer.Ty, ByRef nranges As Long, ByVal ignoreCase As Boolean)
    Dim start As Long, ch As Long, x As Long, dash As Boolean, y As Long, bufferStart As Long
    
    bufferStart = outBuffer.Length
    
    ' start is -2 at the very beginning of the range expression,
    '   -1 when we have not seen a possible "start" character,
    '   and it equals the possible start character if we have seen one
    start = -2
    dash = False
    
    ' begin loop
ContinueLoop:
        x = Advance(lexCtx)

        If x < 0 Then Err.Raise REGEX_ERR_UNTERMINATED_CHARCLASS
        
        Select Case x
        Case UNICODE_RBRACKET: GoTo RngRBracket
        Case UNICODE_MINUS: GoTo RngMinus
        Case UNICODE_BACKSLASH: GoTo RngBackslash
        Case Else: ch = x: GoTo RngFinalizeLiteral ' character represents itself
        End Select
        
        ' unreachable
        Err.Raise RegexErrors.REGEX_ERR_INTERNAL_LOGIC_ERR
        
RngRBracket:
        If start >= 0 Then
            RegexpGenerateRanges outBuffer, ignoreCase, start, start
            GoTo ExitLoop
        ElseIf start = -1 Then
            GoTo ExitLoop
        Else ' start = -2
            ' ] at the very beginning of a range expression is interpreted literally,
            '   since empty ranges are not permitted.
            '   This corresponds to what RE2 does.
            ch = x
        End If
        GoTo RngFinalizeLiteral
        
RngMinus:
        If start >= 0 Then
            If Not dash Then
                If lexCtx.currentCharacter <> UNICODE_RBRACKET Then
                    ' '-' as a range indicator
                    dash = True
                    GoTo ContinueLoop
                End If
            End If
        End If
        ' '-' verbatim
        ch = x
        GoTo RngFinalizeLiteral
        
RngBackslash:
        '
        '  The escapes are same as outside a character class, except that \b has a
        '  different meaning, and \B and backreferences are prohibited (see E5
        '  Section 15.10.2.19).  However, it's difficult to share code because we
        '  handle e.g. "\n" very differently: here we generate a single character
        '  range for it.
        '

        ' XXX: ES2015 surrogate pair handling.

        x = Advance(lexCtx)

        If x <= 47 Then GoTo RngIdentityEscape_or_EOF
        If x >= 123 Then GoTo RngIdentityEscape
        
        On x - 47 GoTo _
            RngE_OctDigit, RngE_OctDigit, RngE_OctDigit, RngE_OctDigit, RngE_OctDigit, RngE_OctDigit, RngE_OctDigit, RngE_OctDigit, _
            RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, _
            RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngE_068, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, _
            RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, _
            RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngE_083, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngE_087, _
            RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, _
            RngIdentityEscape, RngIdentityEscape, RngE_098, RngE_099, RngE_100, RngIdentityEscape, RngE_102, RngIdentityEscape, _
            RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngIdentityEscape, RngE_110, RngIdentityEscape, _
            RngIdentityEscape, RngIdentityEscape, RngE_114, RngE_115, RngE_116, RngE_117, RngE_118, RngE_119, _
            RngE_120, RngIdentityEscape, RngIdentityEscape
            
        ' unreachable
        Err.Raise RegexErrors.REGEX_ERR_INTERNAL_LOGIC_ERR
        
RngIdentityEscape_or_EOF:
        If x < 0 Then Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
        ' else fallthrough
RngIdentityEscape:
        ' IdentityEscape: ES2015 Annex B allows almost all
        ' source characters here.
        ch = x
        GoTo RngFinalizeLiteral
RngE_OctDigit: ' \0... to \7... -> \0 or octal escape from \0 up to \377
        ch = LexerParseLegacyOctal(lexCtx, x)
        GoTo RngFinalizeLiteral
RngE_068: ' \D
        RegexRanges.EmitPredefinedRange outBuffer, RegexUnicodeSupport.StaticData, RegexUnicodeSupport.RANGE_TABLE_NOTDIGIT_START, RegexUnicodeSupport.RANGE_TABLE_NOTDIGIT_LENGTH
        GoTo RngFinalizePredefined
RngE_083: ' \S
        RegexRanges.EmitPredefinedRange outBuffer, RegexUnicodeSupport.StaticData, RegexUnicodeSupport.RANGE_TABLE_NOTWHITE_START, RegexUnicodeSupport.RANGE_TABLE_NOTWHITE_LENGTH
        GoTo RngFinalizePredefined
RngE_087: ' \W
        RegexRanges.EmitPredefinedRange outBuffer, RegexUnicodeSupport.StaticData, RegexUnicodeSupport.RANGE_TABLE_NOTWORDCHAR_START, RegexUnicodeSupport.RANGE_TABLE_NOTWORDCHAR_LENGTH
        GoTo RngFinalizePredefined
RngE_098: ' \b
        ' Note: '\b' in char class is different than outside (assertion),
        ' '\B' is not allowed and is caught by the duk_unicode_is_identifier_part()
        ' check below.
        ch = &H8&
        GoTo RngFinalizeLiteral
RngE_099: ' \c
        x = Advance(lexCtx)
        If x < UNICODE_UC_A Then GoTo RngE_099_FailEscape
        If x > UNICODE_LC_Z Then GoTo RngE_099_FailEscape
        If x > UNICODE_UC_Z And x < UNICODE_LC_A Then GoTo RngE_099_FailEscape
        ' subtract 64, for lowercase deliver corresponding uppercase value
        ch = (x - 64) And &H1F&
        GoTo RngFinalizeLiteral
RngE_099_FailEscape: ' \cX with X not a letter
        Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
RngE_100: ' \d
        RegexRanges.EmitPredefinedRange outBuffer, RegexUnicodeSupport.StaticData, RegexUnicodeSupport.RANGE_TABLE_DIGIT_START, RegexUnicodeSupport.RANGE_TABLE_DIGIT_LENGTH
        GoTo RngFinalizePredefined
RngE_102: ' \f
        ch = &HC&
        GoTo RngFinalizeLiteral
RngE_110: ' \n
        ch = &HA&
        GoTo RngFinalizeLiteral
RngE_114: ' \r
        ch = &HD&
        GoTo RngFinalizeLiteral
RngE_115: ' \s
        RegexRanges.EmitPredefinedRange outBuffer, RegexUnicodeSupport.StaticData, RegexUnicodeSupport.RANGE_TABLE_WHITE_START, RegexUnicodeSupport.RANGE_TABLE_WHITE_LENGTH
        GoTo RngFinalizePredefined
RngE_116: ' \t
        ch = &H9&
        GoTo RngFinalizeLiteral
RngE_117: ' \u
        ch = LexerParseEscapeU(lexCtx)
        GoTo RngFinalizeLiteral
RngE_118: ' \v
        ch = &HB&
        GoTo RngFinalizeLiteral
RngE_119: ' \w
        RegexRanges.EmitPredefinedRange outBuffer, RegexUnicodeSupport.StaticData, RegexUnicodeSupport.RANGE_TABLE_WORDCHAR_START, RegexUnicodeSupport.RANGE_TABLE_WORDCHAR_LENGTH
        GoTo RngFinalizePredefined
RngE_120: ' \x
        ch = LexerParseEscapeX(lexCtx)
        GoTo RngFinalizeLiteral

RngFinalizeLiteral:
        ' final part of loop body if parsed entity was a literal character
        If start >= 0 Then
            If dash Then
                If start > ch Then Err.Raise REGEX_ERR_INVALID_RANGE
                RegexpGenerateRanges outBuffer, ignoreCase, start, ch
                start = -1
                dash = 0
            Else
                RegexpGenerateRanges outBuffer, ignoreCase, start, start
                start = ch
                ' dash is already 0
            End If
        Else
            start = ch
        End If
        GoTo ContinueLoop
        
RngFinalizePredefined:
        ' final part of loop body if parsed entity was a predefined range such as \s
            
        ' multi-character sets not allowed as part of ranges, see
        ' E5 Section 15.10.2.15, abstract operation CharacterRange.
        '
        If start >= 0 Then
            If dash Then
                Err.Raise REGEX_ERR_INVALID_RANGE
            Else
                RegexpGenerateRanges outBuffer, ignoreCase, start, start
            End If
        End If
        start = -1
        ' dash is already 0
        GoTo ContinueLoop

    ' end loop
    
ExitLoop:
    If outBuffer.Length - 2 > bufferStart Then
        ' We have at least 2 intervals.
        HeapsortPairs outBuffer.Buffer, bufferStart, outBuffer.Length - 2
        outBuffer.Length = 2 + Unionize(outBuffer.Buffer, bufferStart, outBuffer.Length - 2)
    End If
    
    nranges = (outBuffer.Length - bufferStart) \ 2
End Sub

' input: array of pairs [x0, y0, x1, y1, ..., x(n-1), y(n-1)]
'   sorts pairs by first entry, second entry irrelevant for order
Private Sub HeapsortPairs(ByRef ary() As Long, ByVal b As Long, ByVal t As Long)
    Dim bb As Long
    Dim parent As Long, child As Long
    Dim smallestValueX As Long, smallestValueY As Long, tmpX As Long, tmpY As Long
    
    ' build heap
    ' bb marks the next element to be added to the heap
    bb = t - 2
    Do Until bb < b
        child = bb
        Do Until child = t
            parent = child + 2 + 2 * ((t - child) \ 4)
            If ary(parent) <= ary(child) Then Exit Do
            tmpX = ary(parent): tmpY = ary(parent + 1)
            ary(parent) = ary(child): ary(parent + 1) = ary(child + 1)
            ary(child) = tmpX: ary(child + 1) = tmpY
            child = parent
        Loop
        bb = bb - 2
    Loop

    ' demount heap
    ' bb marks the lower end of the remaining heap
    bb = b
    Do While bb < t
        smallestValueX = ary(t): smallestValueY = ary(t + 1)
        
        parent = t
        Do
            child = parent - t + parent - 2
            
            ' if there are no children, we are finished
            If child < bb Then Exit Do
            
            ' if there are two children, prefer the one with the smaller value
            If child > bb Then child = child + 2 * (ary(child - 2) <= ary(child))
            
            ary(parent) = ary(child): ary(parent + 1) = ary(child + 1)
            parent = child
        Loop
        
        ' now position parent is free
        
        ' if parent <> bb, free bb rather than parent
        ' by swapping the values in parent and bb and repairing the heap bottom-up
        If parent > bb Then
            ary(parent) = ary(bb): ary(parent + 1) = ary(bb + 1)
            child = parent
            Do Until child = t
                parent = child + 2 + 2 * ((t - child) \ 4)
                If ary(parent) <= ary(child) Then Exit Do
                tmpX = ary(parent): tmpY = ary(parent + 1)
                ary(parent) = ary(child): ary(parent + 1) = ary(child + 1)
                ary(child) = tmpX: ary(child + 1) = tmpY
                child = parent
            Loop
        End If
        
        ' now position bb is free
        
        ary(bb) = smallestValueX: ary(bb + 1) = smallestValueY
        bb = bb + 2
    Loop
End Sub

' assert: t > b
' return: index of first element of last pair
Private Function Unionize(ByRef ary() As Long, ByVal b As Long, ByVal t As Long) As Long
    Dim i As Long, j As Long, lower As Long, upper As Long, nextLower As Long, nextUpper As Long
    
    lower = ary(b): upper = ary(b + 1)
    j = b
    For i = b + 2 To t Step 2
        nextLower = ary(i): nextUpper = ary(i + 1)
        If nextLower <= upper + 1 Then
            If nextUpper > upper Then upper = nextUpper
        Else
            ary(j) = lower: j = j + 1: ary(j) = upper: j = j + 1
            lower = nextLower: upper = nextUpper
        End If
    Next
    ary(j) = lower: ary(j + 1) = upper
    Unionize = j
End Function

' Parse a Unicode escape of the form \xHH.
Private Function LexerParseEscapeX(ByRef lexCtx As Ty) As Long
    Dim dig As Long, escval As Long, x As Long
    
    x = Advance(lexCtx)
    dig = HexvalValidate(x)
    If dig < 0 Then GoTo FailEscape
    escval = dig
    
    x = Advance(lexCtx)
    dig = HexvalValidate(x)
    If dig < 0 Then GoTo FailEscape
    escval = escval * 16 + dig
    
    LexerParseEscapeX = escval
    Exit Function
    
FailEscape:
    Err.Raise REGEX_ERR_INVALID_ESCAPE
End Function

' Parse a Unicode escape of the form \uHHHH, or \u{H+}.
Private Function LexerParseEscapeU(ByRef lexCtx As Ty) As Long
    Dim dig As Long, escval As Long, x As Long
    
    If lexCtx.currentCharacter = UNICODE_LCURLY Then
        Advance lexCtx
        
        escval = 0
        x = Advance(lexCtx)
        If x = UNICODE_RCURLY Then GoTo FailEscape ' Empty escape \u{}
        Do
            dig = HexvalValidate(x)
            If dig < 0 Then GoTo FailEscape
            If escval > &H10FFF Then GoTo FailEscape
            escval = escval * 16 + dig
            
            x = Advance(lexCtx)
        Loop Until x = UNICODE_RCURLY
        LexerParseEscapeU = escval
    Else
        dig = HexvalValidate(Advance(lexCtx))
        If dig < 0 Then GoTo FailEscape
        escval = dig
        
        dig = HexvalValidate(Advance(lexCtx))
        If dig < 0 Then GoTo FailEscape
        escval = escval * 16 + dig
        
        dig = HexvalValidate(Advance(lexCtx))
        If dig < 0 Then GoTo FailEscape
        escval = escval * 16 + dig
        
        dig = HexvalValidate(Advance(lexCtx))
        If dig < 0 Then GoTo FailEscape
        escval = escval * 16 + dig
        
        LexerParseEscapeU = escval
    End If
    Exit Function
    
FailEscape:
    Err.Raise REGEX_ERR_INVALID_ESCAPE
End Function

' If ch is a hex digit, return its value.
' If ch is not a hex digit, return -1.
Private Function HexvalValidate(ByVal ch As Long) As Long
    Const HEX_DELTA_L As Long = UNICODE_LC_A - 10
    Const HEX_DELTA_U As Long = UNICODE_UC_A - 10

    HexvalValidate = -1
    If ch <= UNICODE_UC_F Then
        If ch <= UNICODE_9 Then
            If ch >= UNICODE_0 Then HexvalValidate = ch - UNICODE_0
        Else
            If ch >= UNICODE_UC_A Then HexvalValidate = ch - HEX_DELTA_U
        End If
    Else
        If ch <= UNICODE_LC_F Then
            If ch >= UNICODE_LC_A Then HexvalValidate = ch - HEX_DELTA_L
        End If
    End If
End Function

' Parse legacy octal escape of the form \N{1,3}, e.g. \0, \5, \0377.  Maximum
' allowed value is \0377 (U+00FF), longest match is used.  Used for both string
' RegExp octal escape parsing.
' x is the first digit, which must have already been validated to be in [0-7] by the caller.
'
Private Function LexerParseLegacyOctal(ByRef lexCtx As Ty, ByVal x As Long) As Long
    Dim cp As Long, tmp As Long, i As Long

    cp = x - UNICODE_0

    tmp = lexCtx.currentCharacter
    If tmp < UNICODE_0 Then GoTo ExitFunction
    If tmp > UNICODE_7 Then GoTo ExitFunction

    cp = cp * 8 + (tmp - UNICODE_0)
    Advance lexCtx

    If cp > 31 Then GoTo ExitFunction
    
    tmp = lexCtx.currentCharacter
    If tmp < UNICODE_0 Then GoTo ExitFunction
    If tmp > UNICODE_7 Then GoTo ExitFunction

    cp = cp * 8 + (tmp - UNICODE_0)
    Advance lexCtx

ExitFunction:
    LexerParseLegacyOctal = cp
End Function

Private Function IsIdentifierChar(ByVal c As Long) As Boolean
    ' Todo: Temporary Hack.
    IsIdentifierChar = ((c >= AscW("A")) And (c <= AscW("Z"))) Or ((c >= AscW("a")) And (c <= AscW("z")))
End Function

Private Function Advance(ByRef lexCtx As Ty) As Long
    Dim lower As Long, upper As Long

    With lexCtx
        Advance = .currentCharacter
        If .currentCharacter = LEXER_ENDOFINPUT Then Exit Function
        If .iCurrent = .iEnd Then
            .currentCharacter = LEXER_ENDOFINPUT
        Else
            .iCurrent = .iCurrent + 1
            .currentCharacter = AscW(Mid$(.inputStr, .iCurrent, 1)) And &HFFFF&
        End If
    End With
End Function

