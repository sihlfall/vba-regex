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
    greedy As Boolean
    num As Long ' numeric value (character, count, id for named capture group, -1 for non-named capture group)
    qmin As Long
    qmax As Long
End Type

Public Const RETOK_EOF As Long = 0
Public Const RETOK_DISJUNCTION As Long = 1
Public Const RETOK_QUANTIFIER As Long = 2
Public Const RETOK_ASSERT_START As Long = 3
Public Const RETOK_ASSERT_END As Long = 4
Public Const RETOK_ASSERT_WORD_BOUNDARY As Long = 5
Public Const RETOK_ASSERT_NOT_WORD_BOUNDARY As Long = 6
Public Const RETOK_ASSERT_START_POS_LOOKAHEAD As Long = 7
Public Const RETOK_ASSERT_START_NEG_LOOKAHEAD As Long = 8
Public Const RETOK_ATOM_PERIOD As Long = 9
Public Const RETOK_ATOM_CHAR As Long = 10
Public Const RETOK_ATOM_DIGIT As Long = 11                   ' assumptions in regexp compiler
Public Const RETOK_ATOM_NOT_DIGIT As Long = 12               ' -""-
Public Const RETOK_ATOM_WHITE As Long = 13                   ' -""-
Public Const RETOK_ATOM_NOT_WHITE As Long = 14               ' -""-
Public Const RETOK_ATOM_WORD_CHAR As Long = 15               ' -""-
Public Const RETOK_ATOM_NOT_WORD_CHAR As Long = 16           ' -""-
Public Const RETOK_ATOM_BACKREFERENCE As Long = 17
Public Const RETOK_ATOM_START_CAPTURE_GROUP As Long = 18
Public Const RETOK_ATOM_START_NONCAPTURE_GROUP As Long = 19
Public Const RETOK_ATOM_START_CHARCLASS As Long = 20
Public Const RETOK_ATOM_START_CHARCLASS_INVERTED As Long = 21
Public Const RETOK_ASSERT_START_POS_LOOKBEHIND As Long = 22
Public Const RETOK_ASSERT_START_NEG_LOOKBEHIND As Long = 23
Public Const RETOK_ATOM_END As Long = 24 ' closing parenthesis (ends (POS|NEG)_LOOK(AHEAD|BEHIND), CAPTURE_GROUP, NONCAPTURE_GROUP)

' Returned by input reading function after end of input has been reached
' Since our characters are 16 bit, converted to a positive Long value, and Long is 32 bit, -1 is free to use for us.
Private Const LEXER_ENDOFINPUT As Long = -1

Private Const MAX_LONG As Long = &H7FFFFFFF
Private Const MAX_LONG_DIV_10 As Long = MAX_LONG \ 10

Private Const UNICODE_EXCLAMATION As Long = 33  ' !
Private Const UNICODE_DOLLAR As Long = 36  ' $
Private Const UNICODE_LPAREN As Long = 40  ' (
Private Const UNICODE_RPAREN As Long = 41  ' )
Private Const UNICODE_STAR As Long = 42  ' *
Private Const UNICODE_PLUS As Long = 43  ' +
Private Const UNICODE_COMMA As Long = 44  ' ,
Private Const UNICODE_MINUS As Long = 45  ' -
Private Const UNICODE_PERIOD As Long = 46  ' .
Private Const UNICODE_0 As Long = 48  ' 0
Private Const UNICODE_1 As Long = 49  ' 1
Private Const UNICODE_7 As Long = 55  ' 7
Private Const UNICODE_9 As Long = 57  ' 9
Private Const UNICODE_COLON As Long = 58  ' :
Private Const UNICODE_LT As Long = 60  ' <
Private Const UNICODE_EQUALS As Long = 61  ' =
Private Const UNICODE_GT As Long = 62  ' >
Private Const UNICODE_QUESTION As Long = 63  ' ?
Private Const UNICODE_UC_A As Long = 65  ' A
Private Const UNICODE_UC_B As Long = 66  ' B
Private Const UNICODE_UC_D As Long = 68  ' D
Private Const UNICODE_UC_F As Long = 70  ' F
Private Const UNICODE_UC_S As Long = 83  ' S
Private Const UNICODE_UC_W As Long = 87  ' W
Private Const UNICODE_UC_Z As Long = 90  ' Z
Private Const UNICODE_LBRACKET As Long = 91  ' [
Private Const UNICODE_BACKSLASH As Long = 92  ' \
Private Const UNICODE_RBRACKET As Long = 93  ' ]
Private Const UNICODE_CARET As Long = 94  ' ^
Private Const UNICODE_LC_A As Long = 97  ' a
Private Const UNICODE_LC_B As Long = 98  ' b
Private Const UNICODE_LC_C As Long = 99  ' c
Private Const UNICODE_LC_D As Long = 100  ' d
Private Const UNICODE_LC_F As Long = 102  ' f
Private Const UNICODE_LC_N As Long = 110  ' n
Private Const UNICODE_LC_R As Long = 114  ' r
Private Const UNICODE_LC_S As Long = 115  ' s
Private Const UNICODE_LC_T As Long = 116  ' t
Private Const UNICODE_LC_U As Long = 117  ' u
Private Const UNICODE_LC_V As Long = 118  ' v
Private Const UNICODE_LC_W As Long = 119  ' w
Private Const UNICODE_LC_X As Long = 120  ' x
Private Const UNICODE_LC_Z As Long = 122  ' z
Private Const UNICODE_LCURLY As Long = 123  ' {
Private Const UNICODE_PIPE As Long = 124  ' |
Private Const UNICODE_RCURLY As Long = 125  ' }
Private Const UNICODE_CP_ZWNJ As Long = &H200C& ' zero-width non-joiner
Private Const UNICODE_CP_ZWJ As Long = &H200D&  ' zero-width joiner

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
    Select Case x
    Case UNICODE_PIPE
        outToken.t = RETOK_DISJUNCTION
    Case UNICODE_CARET
        outToken.t = RETOK_ASSERT_START
    Case UNICODE_DOLLAR
        outToken.t = RETOK_ASSERT_END
    Case UNICODE_QUESTION
        With outToken
            .qmin = 0
            .qmax = 1
            If lexCtx.currentCharacter = UNICODE_QUESTION Then
                Advance lexCtx
                .t = RETOK_QUANTIFIER
                .greedy = False
            Else
                .t = RETOK_QUANTIFIER
                .greedy = True
            End If
        End With
    Case UNICODE_STAR
        With outToken
            .qmin = 0
            .qmax = RE_QUANTIFIER_INFINITE
            If lexCtx.currentCharacter = UNICODE_QUESTION Then
                Advance lexCtx
                .t = RETOK_QUANTIFIER
                .greedy = False
            Else
                .t = RETOK_QUANTIFIER
                .greedy = True
            End If
        End With
    Case UNICODE_PLUS
        With outToken
            .qmin = 1
            .qmax = RE_QUANTIFIER_INFINITE
            If lexCtx.currentCharacter = UNICODE_QUESTION Then
                Advance lexCtx
                .t = RETOK_QUANTIFIER
                .greedy = False
            Else
                .t = RETOK_QUANTIFIER
                .greedy = True
            End If
        End With
    Case UNICODE_LCURLY
        ' Production allows 'DecimalDigits', including leading zeroes
        val1 = 0
        val2 = RE_QUANTIFIER_INFINITE
        
        digits = 0

        Do
            x = Advance(lexCtx)
            If (x >= UNICODE_0) And (x <= UNICODE_9) Then
                digits = digits + 1
                ' Be careful to prevent overflow
                If val1 > MAX_LONG_DIV_10 Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER
                val1 = val1 * 10
                tmp = x - UNICODE_0
                If MAX_LONG - val1 < tmp Then Err.Raise REGEX_ERR_INVALID_QUANTIFIER
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
            outToken.greedy = False
            Advance lexCtx
        Else
            outToken.greedy = True
        End If
        outToken.t = RETOK_QUANTIFIER
    Case UNICODE_PERIOD
        outToken.t = RETOK_ATOM_PERIOD
    Case UNICODE_BACKSLASH
        ' The E5.1 specification does not seem to allow IdentifierPart characters
        ' to be used as identity escapes.  Unfortunately this includes '$', which
        ' cannot be escaped as '\$'; it needs to be escaped e.g. as '\u0024'.
        ' Many other implementations (including V8 and Rhino, for instance) do
        ' accept '\$' as a valid identity escape, which is quite pragmatic, and
        ' ES2015 Annex B relaxes the rules to allow these (and other) real world forms.
        x = Advance(lexCtx)
        Select Case x
        Case UNICODE_LC_B
            outToken.t = RETOK_ASSERT_WORD_BOUNDARY
        Case UNICODE_UC_B
            outToken.t = RETOK_ASSERT_NOT_WORD_BOUNDARY
        Case UNICODE_LC_F
            outToken.num = &HC&
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_N
            outToken.num = &HA&
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_T
            outToken.num = &H9&
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_R
            outToken.num = &HD&
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_V
            outToken.num = &HB&
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_C
            x = Advance(lexCtx)
            If (x >= UNICODE_LC_A And x <= UNICODE_LC_Z) Or (x >= UNICODE_UC_A And x <= UNICODE_UC_Z) Then
                outToken.num = x \ 32
                outToken.t = RETOK_ATOM_CHAR
            Else
                Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
            End If
        Case UNICODE_LC_X
            outToken.num = LexerParseEscapeX(lexCtx)
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_U
            ' Todo: What does the following mean?
            ' The token value is the Unicode codepoint without
            ' it being decode into surrogate pair characters
            ' here.  The \u{H+} is only allowed in Unicode mode
            ' which we don't support yet.
            outToken.num = LexerParseEscapeU(lexCtx)
            outToken.t = RETOK_ATOM_CHAR
        Case UNICODE_LC_D
            outToken.t = RETOK_ATOM_DIGIT
        Case UNICODE_UC_D
            outToken.t = RETOK_ATOM_NOT_DIGIT
        Case UNICODE_LC_S
            outToken.t = RETOK_ATOM_WHITE
        Case UNICODE_UC_S
            outToken.t = RETOK_ATOM_NOT_WHITE
        Case UNICODE_LC_W
            outToken.t = RETOK_ATOM_WORD_CHAR
        Case UNICODE_UC_W
            outToken.t = RETOK_ATOM_NOT_WORD_CHAR
        Case UNICODE_0
            x = Advance(lexCtx)
            
            ' E5 Section 15.10.2.11
            If x >= UNICODE_0 And x <= UNICODE_9 Then Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
            outToken.num = 0
            outToken.t = RETOK_ATOM_CHAR
        Case Else
            If x >= UNICODE_1 And x <= UNICODE_9 Then
                val1 = 0
                i = 0
                Do
                    ' We have to be careful here to make sure there will be no overflow.
                    ' 2^31 - 1 backreferences is a bit ridiculous, though.
                    If val1 > MAX_LONG_DIV_10 Then Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
                    val1 = val1 * 10
                    tmp = x - UNICODE_0
                    If MAX_LONG - val1 < tmp Then Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
                    val1 = val1 + tmp
                    x = lexCtx.currentCharacter
                    If x < UNICODE_0 Or x > UNICODE_9 Then Exit Do
                    Advance lexCtx
                    i = i + 1
                Loop
                outToken.t = RETOK_ATOM_BACKREFERENCE
                outToken.num = val1
            ElseIf (x >= 0 And Not UnicodeIsIdentifierPart(0)) Or x = UNICODE_CP_ZWNJ Or x = UNICODE_CP_ZWJ Then
                ' For ES5.1 identity escapes are not allowed for identifier
                ' parts.  This conflicts with a lot of real world code as this
                ' doesn't e.g. allow escaping a dollar sign as /\$/.
                outToken.num = x
                outToken.t = RETOK_ATOM_CHAR
            Else
                Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE
            End If
        End Select
    Case UNICODE_LPAREN
        If lexCtx.currentCharacter = UNICODE_QUESTION Then
            Advance lexCtx
            x = Advance(lexCtx)
            Select Case x
            Case UNICODE_EQUALS
                ' (?=
                outToken.t = RETOK_ASSERT_START_POS_LOOKAHEAD
            Case UNICODE_EXCLAMATION
                ' (?!
                outToken.t = RETOK_ASSERT_START_NEG_LOOKAHEAD
            Case UNICODE_COLON
                ' (?:
                outToken.t = RETOK_ATOM_START_NONCAPTURE_GROUP
            Case UNICODE_LT
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
            Case Else
                Err.Raise REGEX_ERR_INVALID_REGEXP_GROUP
            End Select
        Else
            ' (
            outToken.t = RETOK_ATOM_START_CAPTURE_GROUP
            outToken.num = -1
        End If
    Case UNICODE_RPAREN
        outToken.t = RETOK_ATOM_END
    Case UNICODE_LBRACKET
        ' To avoid creating a heavy intermediate value for the list of ranges,
        ' only the start token ('[' or '[^') is parsed here.  The regexp
        ' compiler parses the ranges itself.
        If lexCtx.currentCharacter = UNICODE_CARET Then
            Advance lexCtx
            outToken.t = RETOK_ATOM_START_CHARCLASS_INVERTED
        Else
            outToken.t = RETOK_ATOM_START_CHARCLASS
        End If
    Case UNICODE_RCURLY, UNICODE_RBRACKET
        ' Although these could be parsed as PatternCharacters unambiguously (here),
        ' * E5 Section 15.10.1 grammar explicitly forbids these as PatternCharacters.
        Err.Raise REGEX_ERR_INVALID_REGEXP_CHARACTER
    Case LEXER_ENDOFINPUT
        ' EOF
        outToken.t = RETOK_EOF
    Case Else
        ' PatternCharacter, all excluded characters are matched by cases above
        outToken.t = RETOK_ATOM_CHAR
        outToken.num = x
    End Select
End Sub

Public Sub ParseReRanges(lexCtx As Ty, ByRef outBuffer As ArrayBuffer.Ty, ByRef nranges As Long, ByVal ignoreCase As Boolean)
    Dim start As Long, ch As Long, x As Long, dash As Boolean, y As Long, bufferStart As Long
    
    bufferStart = outBuffer.Length
    
    ' start is -2 at the very beginning of the range expression,
    '   -1 when we have not seen a possible "start" character,
    '   and it equals the possible start character if we have seen one
    start = -2
    dash = False
    
    Do
ContinueLoop:
        x = Advance(lexCtx)

        If x < 0 Then GoTo FailUntermCharclass
        
        Select Case x
        Case UNICODE_RBRACKET
            If start >= 0 Then
                RegexpGenerateRanges outBuffer, ignoreCase, start, start
                Exit Do
            ElseIf start = -1 Then
                Exit Do
            Else ' start = -2
                ' ] at the very beginning of a range expression is interpreted literally,
                '   since empty ranges are not permitted.
                '   This corresponds to what RE2 does.
                ch = x
            End If
        Case UNICODE_MINUS
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
        Case UNICODE_BACKSLASH
            '
            '  The escapes are same as outside a character class, except that \b has a
            '  different meaning, and \B and backreferences are prohibited (see E5
            '  Section 15.10.2.19).  However, it's difficult to share code because we
            '  handle e.g. "\n" very differently: here we generate a single character
            '  range for it.
            '

            ' XXX: ES2015 surrogate pair handling.

            x = Advance(lexCtx)

            Select Case x
            Case UNICODE_LC_B
                ' Note: '\b' in char class is different than outside (assertion),
                ' '\B' is not allowed and is caught by the duk_unicode_is_identifier_part()
                ' check below.
                '
                ch = &H8&
            Case x = UNICODE_LC_F
                ch = &HC&
            Case UNICODE_LC_N
                ch = &HA&
            Case UNICODE_LC_T
                ch = &H9&
            Case UNICODE_LC_R
                ch = &HD&
            Case UNICODE_LC_V
                ch = &HB&
            Case UNICODE_LC_C
                x = Advance(lexCtx)
                If ((x >= UNICODE_LC_A And x <= UNICODE_LC_Z) Or (x >= UNICODE_UC_A And x <= UNICODE_UC_Z)) Then
                    ch = x Mod 32
                Else
                    GoTo FailEscape
                End If
            Case UNICODE_LC_X
                ch = LexerParseEscapeX(lexCtx)
            Case UNICODE_LC_U
                ch = LexerParseEscapeU(lexCtx)
            Case UNICODE_LC_D
                RegexRanges.EmitPredefinedRange outBuffer, RegexUnicodeSupport.StaticData, RegexUnicodeSupport.RANGE_TABLE_DIGIT_START, RegexUnicodeSupport.RANGE_TABLE_DIGIT_LENGTH
                ch = -1
            Case UNICODE_UC_D
                RegexRanges.EmitPredefinedRange outBuffer, RegexUnicodeSupport.StaticData, RegexUnicodeSupport.RANGE_TABLE_NOTDIGIT_START, RegexUnicodeSupport.RANGE_TABLE_NOTDIGIT_LENGTH
                ch = -1
            Case UNICODE_LC_S
                RegexRanges.EmitPredefinedRange outBuffer, RegexUnicodeSupport.StaticData, RegexUnicodeSupport.RANGE_TABLE_WHITE_START, RegexUnicodeSupport.RANGE_TABLE_WHITE_LENGTH
                ch = -1
            Case UNICODE_UC_S
                RegexRanges.EmitPredefinedRange outBuffer, RegexUnicodeSupport.StaticData, RegexUnicodeSupport.RANGE_TABLE_NOTWHITE_START, RegexUnicodeSupport.RANGE_TABLE_NOTWHITE_LENGTH
                ch = -1
            Case UNICODE_LC_W
                RegexRanges.EmitPredefinedRange outBuffer, RegexUnicodeSupport.StaticData, RegexUnicodeSupport.RANGE_TABLE_WORDCHAR_START, RegexUnicodeSupport.RANGE_TABLE_WORDCHAR_LENGTH
                ch = -1
            Case UNICODE_UC_W
                RegexRanges.EmitPredefinedRange outBuffer, RegexUnicodeSupport.StaticData, RegexUnicodeSupport.RANGE_TABLE_NOTWORDCHAR_START, RegexUnicodeSupport.RANGE_TABLE_NOTWORDCHAR_LENGTH
                ch = -1
            Case Else
                If x < 0 Then GoTo FailEscape
                If x <= UNICODE_7 Then
                    If x >= UNICODE_0 Then
                        ' \0 or octal escape from \0 up to \377
                        ch = LexerParseLegacyOctal(lexCtx, x)
                    Else
                        ' IdentityEscape: ES2015 Annex B allows almost all
                        ' source characters here.  Match anything except
                        ' EOF here.
                        ch = x
                    End If
                Else
                    ' IdentityEscape: ES2015 Annex B allows almost all
                    ' source characters here.  Match anything except
                    ' EOF here.
                    ch = x
                End If
            End Select
        Case Else
            ' character represents itself
            ch = x
        End Select

        ' ch is a literal character here or -1 if parsed entity was
        ' an escape such as "\s".
        '

        If ch < 0 Then
            ' multi-character sets not allowed as part of ranges, see
            ' E5 Section 15.10.2.15, abstract operation CharacterRange.
            '
            If start >= 0 Then
                If dash Then
                    GoTo FailRange
                Else
                    RegexpGenerateRanges outBuffer, ignoreCase, start, start
                End If
            End If
            start = -1
            ' dash is already 0
        Else
            If start >= 0 Then
                If dash Then
                    If start > ch Then GoTo FailRange
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
        End If
    Loop

    If outBuffer.Length - 2 > bufferStart Then
        ' We have at least 2 intervals.
        HeapsortPairs outBuffer.Buffer, bufferStart, outBuffer.Length - 2
        outBuffer.Length = 2 + Unionize(outBuffer.Buffer, bufferStart, outBuffer.Length - 2)
    End If
    
    nranges = (outBuffer.Length - bufferStart) \ 2
    
    Exit Sub

FailEscape:
    Err.Raise REGEX_ERR_INVALID_REGEXP_ESCAPE

FailRange:
    Err.Raise REGEX_ERR_INVALID_RANGE

FailUntermCharclass:
    Err.Raise REGEX_ERR_UNTERMINATED_CHARCLASS
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
Private Function Unionize(ByRef ary() As Long, ByVal b As Long, ByVal t As Long)
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
Private Function LexerParseLegacyOctal(ByRef lexCtx As Ty, ByVal x As Long)
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

