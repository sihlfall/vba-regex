Attribute VB_Name = "RegexNfaMatcher"
Option Explicit

Private Const RE_FLAG_MULTILINE As Long = 1
Type NfaMatcher
    reFlags As Long
    input As String
    bytecode() As Long
    recursionDepth As Long
    recursionLimit As Long
    stepsCount As Long
    stepsLimit As Long
    saved() As Long
End Type

Private Function GetBc(ByRef bytecode() As Long, ByRef pc As Long) As Long
    If pc > UBound(bytecode) Then
        GetBc = 0 ' Todo: ??????
        Exit Function
    End If
    GetBc = bytecode(pc)
    pc = pc + 1
End Function

Private Function PeekInputCharCode(ByRef inputStr As String, ByRef sp As Long) As Long
    If sp >= Len(inputStr) Then
        PeekInputCharCode = -1
        Exit Function
    End If
    PeekInputCharCode = AscW(Mid$(inputStr, sp + 1, 1))
End Function


Private Function UnicodeIsLineTerminator(c As Long)
    UnicodeIsLineTerminator = False
End Function

Private Function UnicodeReIsWordchar(c As Long)
    'TODO: Temporary hack
    UnicodeReIsWordchar = ((c >= AscW("A")) And (c <= AscW("Z"))) Or ((c >= AscW("a") And (c <= AscW("z"))))
End Function


Function NfaMatch(ByRef reCtx As NfaMatcher) As Long
    Dim maxSave As Long
    maxSave = reCtx.bytecode(0)
    ReDim reCtx.saved(0 To maxSave) As Long
    NfaMatch = NfaDoMatch(reCtx, reCtx.saved, 1, 0, maxSave)
End Function

Private Function NfaDoMatch(ByRef reCtx As NfaMatcher, ByRef outCaptures() As Long, ByVal pc As Long, ByVal sp As Long, ByVal maxSave As Long) As Long
    Dim pool As RegexNfaThreadPool.ThreadPool
    RegexNfaThreadPool.Initialize pool
    RegexNfaThreadPool.AddFirst pool, pc, maxSave
    NfaDoMatch = NfaRunThreads(outCaptures, pool, reCtx.bytecode, reCtx.input, sp, maxSave, reCtx.reFlags)
End Function

Private Function NfaRunThreads( _
    ByRef outCaptures() As Long, _
    ByRef pool As RegexNfaThreadPool.ThreadPool, _
    ByRef bytecode() As Long, ByRef inputStr As String, ByVal sp As Long, ByVal maxSave As Long, _
    ByVal reFlags As Long _
) As Long
    Dim pc As Long, currentChar As Long
    
    Dim op As Long
    Dim c1 As Long
    Dim n As Long
    Dim successfulMatch As Boolean
    Dim r1 As Long, r2 As Long
    Dim b1 As Boolean, b2 As Boolean
    Dim idx As Long
    Dim currentThread As Long
        
    currentThread = 2 + 2 * pool.tsActive
    currentThread = pool.tsStack(currentThread).nxt

    currentChar = PeekInputCharCode(inputStr, sp)
    sp = sp + 1

    GoTo LoopStart
    
    ' BEGIN LOOP
ContinueLoop:
        pool.tsStack(currentThread).processed = True
        currentThread = pool.tsStack(currentThread).nxt
        
LoopStart: ' If we can be sure that we have at least one thread at start, we can move this to below
        If currentThread < NFA_TS_FIRST_ACTUAL Then ' reached a sentinel
        
ContinueWithInactive:
            pool.tsActive = 1 - pool.tsActive
            currentThread = pool.tsStack(2 + 2 * pool.tsActive).nxt
            If currentThread < NFA_TS_FIRST_ACTUAL Then ' List empty
                NfaRunThreads = -1
                Exit Function
            End If
            
            RegexNfaThreadPool.ClearNonActive pool
            currentChar = PeekInputCharCode(inputStr, sp)
            sp = sp + 1
        End If
                    
        With pool.tsStack(currentThread): pc = .pc: End With

        op = GetBc(bytecode, pc)
    
        Select Case op
        Case REOP_MATCH
            If pool.tsStack(currentThread).prev < NFA_TS_FIRST_ACTUAL Then GoTo Match
            RegexNfaThreadPool.AppendToInactive pool, pc - 1, pool.tsStack(currentThread).qstack, pool.tsStack(currentThread).saved
            GoTo ContinueWithInactive ' Lookahead not yet supported ' We continue only with higher priority threads
        Case REOP_CHAR
            '
            '  Byte-based matching would be possible for case-sensitive
            '  matching but not for case-insensitive matching.  So, we
            '  match by decoding the input and bytecode character normally.
            '
            '  Bytecode characters are assumed to be already canonicalized.
            '  Input characters are canonicalized automatically by
            '  duk__inp_get_cp() if necessary.
            '
            '  There is no opcode for matching multiple characters.  The
            '  regexp compiler has trouble joining strings efficiently
            '  during compilation.  See doc/regexp.rst for more discussion.

            c1 = GetBc(bytecode, pc)
            If c1 <> currentChar Then GoTo ContinueLoop
            RegexNfaThreadPool.AppendToInactive pool, pc, pool.tsStack(currentThread).qstack, pool.tsStack(currentThread).saved
            GoTo ContinueLoop
        Case REOP_PERIOD
            If currentChar < 0 Then GoTo ContinueLoop
            If UnicodeIsLineTerminator(c1) Then GoTo ContinueLoop
            RegexNfaThreadPool.AppendToInactive pool, pc, pool.tsStack(currentThread).qstack, pool.tsStack(currentThread).saved
            GoTo ContinueLoop
        Case REOP_RANGES, REOP_INVRANGES
            n = GetBc(bytecode, pc)
            If currentChar < 0 Then GoTo ContinueLoop

            successfulMatch = False
            Do While n > 0
                r1 = GetBc(bytecode, pc)
                r2 = GetBc(bytecode, pc)
                successfulMatch = successfulMatch Or (currentChar >= r1 And currentChar <= r2)
                ' Note: don't bail out early, we must read all the ranges from
                ' bytecode.  Another option is to skip them efficiently after
                ' breaking out of here.  Prefer smallest code.
                ' TODO: bail out and replace by skip
                n = n - 1
            Loop

            If (op = REOP_RANGES) <> successfulMatch Then GoTo ContinueLoop

            RegexNfaThreadPool.AppendToInactive pool, pc, pool.tsStack(currentThread).qstack, pool.tsStack(currentThread).saved
            GoTo ContinueLoop
        Case REOP_ASSERT_START
            If sp > 1 Then
                If Not (reFlags And RE_FLAG_MULTILINE) Then GoTo ContinueLoop
                ' E5 Sections 15.10.2.8, 7.3
                If Not UnicodeIsLineTerminator(currentChar) Then GoTo ContinueLoop
            End If
            RegexNfaThreadPool.InsertAfterCurrent pool, currentThread, pc, pool.tsStack(currentThread).qstack, pool.tsStack(currentThread).saved
            GoTo ContinueLoop
        Case REOP_ASSERT_END
            If currentChar < 0 Then
                RegexNfaThreadPool.InsertAfterCurrent pool, currentThread, pc, pool.tsStack(currentThread).qstack, pool.tsStack(currentThread).saved
                GoTo ContinueLoop
            End If
            If Not (reFlags And RE_FLAG_MULTILINE) Then GoTo ContinueLoop
            If Not UnicodeIsLineTerminator(c1) Then GoTo ContinueLoop
            RegexNfaThreadPool.InsertAfterCurrent pool, currentThread, pc, pool.tsStack(currentThread).qstack, pool.tsStack(currentThread).saved
            GoTo ContinueLoop
        Case REOP_ASSERT_WORD_BOUNDARY, REOP_ASSERT_NOT_WORD_BOUNDARY
            '
            '  E5 Section 15.10.2.6.  The previous and current character
            '  should -not- be canonicalized as they are now.  However,
            '  canonicalization does not affect the result of IsWordChar()
            '  (which depends on Unicode characters never canonicalizing
            '  into ASCII characters) so this does not matter.
            If sp = 1 Then
                b1 = False  ' not a wordchar
            Else
                c1 = PeekInputCharCode(inputStr, sp - 2)
                b1 = UnicodeReIsWordchar(c1)
            End If
            If sp > Len(inputStr) Then
                b2 = False ' not a wordchar
            Else
                b2 = UnicodeReIsWordchar(currentChar)
            End If

            If (op = REOP_ASSERT_WORD_BOUNDARY) = (b1 = b2) Then GoTo ContinueLoop

            RegexNfaThreadPool.InsertAfterCurrent pool, currentThread, pc, pool.tsStack(currentThread).qstack, pool.tsStack(currentThread).saved
            GoTo ContinueLoop
        Case REOP_JUMP
            n = GetBc(bytecode, pc)
            RegexNfaThreadPool.InsertAfterCurrent pool, currentThread, pc + n, pool.tsStack(currentThread).qstack, pool.tsStack(currentThread).saved
            GoTo ContinueLoop
        Case REOP_SPLIT1
            ' split1: prefer direct execution (no jump)
            n = GetBc(bytecode, pc)
            RegexNfaThreadPool.InsertAfterCurrent pool, currentThread, pc, pool.tsStack(currentThread).qstack, pool.tsStack(currentThread).saved
            RegexNfaThreadPool.InsertAfterCurrent pool, currentThread, pc + n, pool.tsStack(currentThread).qstack, pool.tsStack(currentThread).saved
            GoTo ContinueLoop
        Case REOP_SPLIT2
            ' split2: prefer jump execution (not direct)
            n = GetBc(bytecode, pc)
            RegexNfaThreadPool.InsertAfterCurrent pool, currentThread, pc + n, pool.tsStack(currentThread).qstack, pool.tsStack(currentThread).saved
            RegexNfaThreadPool.InsertAfterCurrent pool, currentThread, pc, pool.tsStack(currentThread).qstack, pool.tsStack(currentThread).saved
            GoTo ContinueLoop
        Case REOP_REPEAT_EXACTLY_INIT, REOP_REPEAT_MAX_HUMBLE_INIT, REOP_REPEAT_GREEDY_MAX_INIT
            Err.Raise 5000
        Case REOP_REPEAT_EXACTLY_START
            Err.Raise 5000
        Case REOP_REPEAT_EXACTLY_END
            Err.Raise 5000
        Case REOP_REPEAT_MAX_HUMBLE_START
            Err.Raise 5000
        Case REOP_REPEAT_GREEDY_MAX_START
            Err.Raise 5000
        Case REOP_REPEAT_MAX_HUMBLE_END, REOP_REPEAT_GREEDY_MAX_END
            Err.Raise 5000
        Case REOP_SAVE
            idx = GetBc(bytecode, pc)
            If idx > maxSave Then GoTo InternalError
            RegexNfaThreadPool.InsertAfterCurrent pool, currentThread, pc, pool.tsStack(currentThread).qstack, pool.tsStack(currentThread).saved
            pool.tsStack(pool.tsStack(currentThread).nxt).saved(idx) = sp - 1
            GoTo ContinueLoop
        Case REOP_CHECK_LOOKAHEAD
            Err.Raise 3000
        Case REOP_LOOKPOS
            Err.Raise 3000
        Case REOP_LOOKNEG
            Err.Raise 3000
        Case REOP_BACKREFERENCE
            Err.Raise 3000
        Case Else
            'DUK_D(DUK_DPRINT("internal error, regexp opcode error: %ld", (long) op));
            GoTo InternalError
        End Select
    ' END LOOP

Match:
    outCaptures = pool.tsStack(currentThread).saved
    NfaRunThreads = pool.tsStack(currentThread).saved(1)
    Exit Function
    
InternalError:
    ' TODO: Raise correct exception
    Err.Raise 3000
    ' DUK_ERROR_INTERNAL(re_ctx->thr);
    ' DUK_WO_NORETURN(return -1;);
End Function

